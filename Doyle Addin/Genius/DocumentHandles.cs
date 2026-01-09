#nullable enable

#region

using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using MessageBox = System.Windows.MessageBox;

#endregion

namespace Doyle_Addin.Genius;

public static class DocumentHandles
{
	private const string ErrorMessageBoxTitle = "Error";

	public static async Task<(PartInfo? calculatedPartInfo, bool hasCalculatedValues)> HandlePurchasedPart(
		Application mInventorApp, DataGrid geniusPropertiesGrid, IDictionary<string, string> databaseProperties)
	{
		// Initialize return values
		PartInfo? calculatedPartInfo  = null;
		var       hasCalculatedValues = false;

		try
		{
			var doc = mInventorApp.ActiveDocument;
			if (doc == null) return (null, false);

			// Calculate mass for purchased part
			var partInfo = new PartInfo();

			// Use appropriate mass calculation based on document type
			var success = doc.DocumentType switch
			{
				kPartDocumentObject     => partInfo.CalculateMassOnly(doc),
				kAssemblyDocumentObject => partInfo.CalculateMassOnly(doc),
				_                       => false
			};

			if (success)
			{
				// Store calculated values in memory (don't update document yet)
				calculatedPartInfo  = partInfo;
				hasCalculatedValues = true;

				// Update Family property from database if available - store in memory only
				if (databaseProperties.TryGetValue("Family", out var familyValue) && !string.IsNullOrEmpty(familyValue))
				{
					// Store Family value in memory without modifying the document
					partInfo.CostCenter = familyValue;
					Debug.WriteLine($"[FAMILY_UPDATE] Stored Family in memory: '{familyValue}'");
				}

				// Refresh the data grids to show calculated values (but preserve database properties)
				await PartInfo.RefreshDataGridsWithCalculatedValues(mInventorApp, geniusPropertiesGrid,
					hasCalculatedValues,
					calculatedPartInfo, null, databaseProperties, doc);
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error handling purchased part: {ex.Message}");
			MessageBox.Show($"Error loading properties for purchased part: {ex.Message}",
				ErrorMessageBoxTitle, MessageBoxButton.OK, MessageBoxImage.Error);
		}

		return (calculatedPartInfo, hasCalculatedValues);
	}

	public static async Task<(PartInfo? calculatedPartInfo, bool hasCalculatedValues, RawStockInfo? calculatedRawStock)>
		HandleSheetMetalPart(PartDocument partDoc, Application inventorApp, DataGrid geniusPropertiesGrid,
			IDictionary<string, string> databaseProperties)
	{
		// Create PartInfo instance and calculate dimensions
		var partInfo = new PartInfo();

		// Only calculate for sheet metal parts
		var success = partInfo.CalculateSheetMetalDimensions(partDoc);

		if (!success)
		{
			MessageBox.Show("This function only works on sheet metal parts with a valid flat pattern.",
				ErrorMessageBoxTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
			return (null, false, null);
		}

		// Get calculated thickness
		var thickness = PartInfo.GetSheetMetalThickness(partDoc);
		Debug.WriteLine($"[RM_DEBUG] Calculated thickness: {thickness:F4} in");

		// Try to find raw stock based on thickness and material
		RawStockInfo? rawStock = null;
		try
		{
			// Get material from part - try to get from material property
			var material = PartInfo.GetNormalizedMaterial(partDoc);

			Debug.WriteLine($"[RM_DEBUG] Final material: '{material}', Thickness: {thickness:F4}");

			if (!string.IsNullOrEmpty(material) && thickness > 0)
			{
				Debug.WriteLine(
					$"[RM_DEBUG] Searching database for raw stock with material='{material}' and thickness={thickness:F4}");
				var service = new DatabaseService();
				rawStock = await service.GetRawStockByThicknessAndMaterialAsync(thickness,
					material, partInfo.CalculatedLength, partInfo.CalculatedWidth);

				Debug.WriteLine(rawStock != null
					? $"[RM_DEBUG] Raw stock found: {rawStock.Stock}, RMUNIT: '{rawStock.RMUNIT}'"
					: "[RM_DEBUG] No raw stock found in database");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[RM_DEBUG] Error searching for raw stock: {ex.Message}");
		}

		// Store calculated values in memory (don't update document yet)
		var calculatedPartInfo = partInfo;
		calculatedPartInfo.Thickness = thickness.ToString("F3") + " in";
		const bool hasCalculatedValues = true;

		// Refresh the data grids to show calculated values (but preserve database properties)
		await PartInfo.RefreshDataGridsWithCalculatedValues(inventorApp, geniusPropertiesGrid, hasCalculatedValues,
			calculatedPartInfo, rawStock, databaseProperties, partDoc as Document);


		return (calculatedPartInfo, hasCalculatedValues, rawStock);
	}

	public static async Task<(PartInfo? calculatedPartInfo, bool hasCalculatedValues, List<PartInfo> componentParts)>
		HandleAssembly(
			AssemblyDocument assyDoc, Application inventorApp, DataGrid geniusPropertiesGrid,
			IDictionary<string, string> databaseProperties)
	{
		var partInfo       = new PartInfo();
		var componentParts = new List<PartInfo>();

		// Calculate assembly mass first
		var success = partInfo.CalculateAssemblyMass(assyDoc);

		if (!success)
		{
			MessageBox.Show("Could not calculate mass for this assembly.",
				ErrorMessageBoxTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
			return (null, false, componentParts);
		}

		// Process all assembly occurrences to get correct BOM structure context
		try
		{
			var occurrences = assyDoc.ComponentDefinition.Occurrences;
			Debug.WriteLine($"[ASSEMBLY] Found {occurrences.Count} assembly occurrences");

			await ProcessAssemblyOccurrences(occurrences, componentParts);
			Debug.WriteLine($"[ASSEMBLY] Successfully processed {componentParts.Count} components");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[ASSEMBLY] Error processing assembly occurrences: {ex.Message}");
			MessageBox.Show($"Error processing assembly components: {ex.Message}",
				ErrorMessageBoxTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
		}

		// Store calculated values in memory (don't update document yet)
		var        calculatedPartInfo  = partInfo;
		const bool hasCalculatedValues = true;

		// Refresh the data grids to show calculated values (but preserve database properties)
		await PartInfo.RefreshDataGridsWithCalculatedValues(inventorApp, geniusPropertiesGrid, hasCalculatedValues,
			calculatedPartInfo, null, databaseProperties, assyDoc as Document);

		return (calculatedPartInfo, hasCalculatedValues, componentParts);
	}

	private static async Task ProcessAssemblyOccurrences(ComponentOccurrences occurrences,
		List<PartInfo> componentParts)
	{
		foreach (ComponentOccurrence occurrence in occurrences)
		{
			Debug.WriteLine(
				$"[ASSEMBLY] Processing occurrence: {occurrence.Name} (Type: {occurrence.DefinitionDocumentType})");

			// Check if this is a Reference BOM component and ignore it using assembly context
			if (ShouldIgnoreReferenceComponentByOccurrence(occurrence))
			{
				Debug.WriteLine($"[ASSEMBLY] Ignoring Reference component: {occurrence.Name}");
				continue;
			}

			// Get the referenced document from the occurrence
			if (occurrence.ReferencedDocumentDescriptor?.ReferencedDocument is not Document refDoc) continue;
			var componentInfo = await ProcessReferencedDocument(refDoc);
			if (componentInfo == null) continue;
			componentParts.Add(componentInfo);

			// If this is a sub-assembly, recursively add its components to the main list
			if (componentInfo.SubAssemblyComponents is not { Count: > 0 }) continue;
			Debug.WriteLine(
				$"[ASSEMBLY] Adding {componentInfo.SubAssemblyComponents.Count} components from sub-assembly {componentInfo.DocumentName}");
			FlattenSubAssemblyComponents(componentInfo.SubAssemblyComponents, componentParts, 1);
		}
	}

	private static void FlattenSubAssemblyComponents(List<PartInfo> subAssemblyComponents,
		List<PartInfo> componentParts, int depth)
	{
		const int maxDepth = 10; // Prevent infinite recursion

		if (depth > maxDepth)
		{
			Debug.WriteLine(
				$"[ASSEMBLY] Maximum recursion depth ({maxDepth}) reached, stopping to prevent infinite loop");
			return;
		}

		foreach (var component in subAssemblyComponents)
		{
			// Add the component to the main list
			componentParts.Add(component);
			Debug.WriteLine(
				$"[ASSEMBLY] Added component from depth {depth}: {component.DocumentName} ({component.DocumentType})");

			// If this component is also a sub-assembly, recursively flatten its components
			if (component.SubAssemblyComponents is not { Count: > 0 }) continue;
			Debug.WriteLine(
				$"[ASSEMBLY] Recursively processing {component.SubAssemblyComponents.Count} components from sub-assembly {component.DocumentName}");
			FlattenSubAssemblyComponents(component.SubAssemblyComponents, componentParts, depth + 1);
		}
	}

	private static async Task<PartInfo?> ProcessReferencedDocument(Document refDoc)
	{
		return refDoc.DocumentType switch
		{
			kPartDocumentObject when refDoc is PartDocument partDoc => await ProcessPartDocument(partDoc),
			kAssemblyDocumentObject when refDoc is AssemblyDocument subAssyDoc => await ProcessAssemblyDocument(
				subAssyDoc),
			_ => null
		};
	}

	private static async Task<PartInfo?> ProcessPartDocument(PartDocument partDoc)
	{
		var partInfo = new PartInfo
		{
			PartNumber = PartInfo.GetDesignTrackingProperty(partDoc as Document, "Part Number")
		};
		try
		{
			// Check if it's a sheet metal part
			if (partDoc.ComponentDefinition is SheetMetalComponentDefinition)
			{
				var success = partInfo.CalculateSheetMetalDimensions(partDoc);
				if (success)
				{
					partInfo.Thickness    = PartInfo.GetSheetMetalThickness(partDoc).ToString("F3") + " in";
					partInfo.DocumentName = partDoc.DisplayName;
					partInfo.DocumentType = "Sheet Metal Part";

					// Try to find raw stock
					await TryFindRawStockForSheetMetalPart(partDoc, partInfo);

					return partInfo;
				}
			}
			else
			{
				// Regular part
				var success = partInfo.CalculatePartDimensions(partDoc);
				if (success)
				{
					partInfo.DocumentName = partDoc.DisplayName;
					partInfo.DocumentType = "Regular Part";
					return partInfo;
				}
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[ASSEMBLY] Error processing part {partDoc.DisplayName}: {ex.Message}");
		}

		return null;
	}

	private static async Task TryFindRawStockForSheetMetalPart(PartDocument partDoc, PartInfo partInfo)
	{
		var thickness = PartInfo.GetSheetMetalThickness(partDoc);
		var material  = PartInfo.GetNormalizedMaterial(partDoc);

		if (!string.IsNullOrEmpty(material) && thickness > 0)
			try
			{
				var service = new DatabaseService();
				var rawStock = await service.GetRawStockByThicknessAndMaterialAsync(thickness,
					material, partInfo.CalculatedLength, partInfo.CalculatedWidth);

				if (rawStock != null)
				{
					partInfo.RawStockInfo = rawStock;
					Debug.WriteLine($"[ASSEMBLY] Raw stock found for {partDoc.DisplayName}: {rawStock.Stock}");
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[ASSEMBLY] Error finding raw stock for {partDoc.DisplayName}: {ex.Message}");
			}
	}

	private static async Task<PartInfo?> ProcessAssemblyDocument(AssemblyDocument assyDoc)
	{
		var partInfo = new PartInfo { PartNumber = PartInfo.GetDesignTrackingProperty(assyDoc as Document, "Part Number") };

		try
		{
			var success = partInfo.CalculateAssemblyMass(assyDoc);
			if (success)
			{
				partInfo.DocumentName = assyDoc.DisplayName;
				partInfo.DocumentType = "Sub-Assembly";

				// Recursively process sub-assembly components
				try
				{
					var subOccurrences = assyDoc.ComponentDefinition.Occurrences;
					Debug.WriteLine(
						$"[ASSEMBLY] Processing {subOccurrences.Count} components in sub-assembly {assyDoc.DisplayName}");

					// Create a temporary list to hold sub-assembly components
					var subAssemblyComponents = new List<PartInfo>();
					await ProcessAssemblyOccurrences(subOccurrences, subAssemblyComponents);

					// Store sub-assembly components in the PartInfo for later use
					partInfo.SubAssemblyComponents = subAssemblyComponents;
					Debug.WriteLine(
						$"[ASSEMBLY] Found {subAssemblyComponents.Count} components in sub-assembly {assyDoc.DisplayName}");
				}
				catch (Exception ex)
				{
					Debug.WriteLine(
						$"[ASSEMBLY] Error processing components in sub-assembly {assyDoc.DisplayName}: {ex.Message}");
				}

				return partInfo;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[ASSEMBLY] Error processing sub-assembly {assyDoc.DisplayName}: {ex.Message}");
		}

		return null;
	}

	private static bool ShouldIgnoreReferenceComponentByOccurrence(ComponentOccurrence occurrence)
	{
		try
		{
			// Check the BOM structure from the component occurrence in the assembly context
			if (occurrence.BOMStructure == BOMStructureEnum.kReferenceBOMStructure) return true;

			switch (occurrence.ReferencedDocumentDescriptor?.ReferencedDocument)
			{
				// Check for Content Center components
				case PartDocument { ComponentDefinition.IsContentMember: true }:
					Debug.WriteLine($"[ASSEMBLY] Ignoring Content Center component: {occurrence.Name}");
					return true;
				// Check for Factory Member components
				case PartDocument { ComponentDefinition.IsiPartMember: true }:
					Debug.WriteLine($"[ASSEMBLY] Ignoring Factory Member component: {occurrence.Name}");
					return true;
				default:
					return false;
			}
		}
		catch
		{
			// If we can't determine properties, assume it's not to be ignored
			return false;
		}
	}

	public static async Task<(PartInfo? calculatedPartInfo, bool hasCalculatedValues)> HandleRegularPart(
		PartDocument partDoc, Application inventorApp, DataGrid geniusPropertiesGrid,
		IDictionary<string, string> databaseProperties)
	{
		// Create PartInfo instance and calculate dimensions
		var partInfo = new PartInfo();

		var success = partInfo.CalculatePartDimensions(partDoc);

		if (!success)
		{
			MessageBox.Show("Could not calculate dimensions for this part.",
				ErrorMessageBoxTitle, MessageBoxButton.OK, MessageBoxImage.Warning);
			return (null, false);
		}

		// Store calculated values in memory (don't update document yet)
		var        calculatedPartInfo  = partInfo;
		const bool hasCalculatedValues = true;

		// Refresh the data grids to show calculated values (but preserve database properties)
		await PartInfo.RefreshDataGridsWithCalculatedValues(inventorApp, geniusPropertiesGrid, hasCalculatedValues,
			calculatedPartInfo, null, databaseProperties, partDoc as Document);


		return (calculatedPartInfo, hasCalculatedValues);
	}
}
