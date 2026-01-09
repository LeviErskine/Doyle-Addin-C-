#nullable enable

#region

using System.ComponentModel;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using MessageBox = System.Windows.MessageBox;

#endregion

namespace Doyle_Addin.Genius;

public partial class PartInfo
{
	private const string THICKNESS_PROPERTY = "Thickness";
	private const string EXTENT_LENGTH_PROPERTY = "Extent_Length";
	private const string EXTENT_WIDTH_PROPERTY = "Extent_Width";
	private const string EXTENT_AREA_PROPERTY = "Extent_Area";
	private const string INVENTOR_USER_DEFINED_PROPERTIES = "Inventor User Defined Properties";
	private const string DESIGN_TRACKING_PROPERTIES = "Design Tracking Properties";
	private const string FAMILY_PROPERTY = "Family";
	private const string PART_NUMBER_PROPERTY = "Part Number";
	private const string GENIUS_MASS_PROPERTY = "GeniusMass";
	private const string RMQTY_PROPERTY = "RMQTY";
	private const string RMUNIT_PROPERTY = "RMUNIT";
	private const string DESCRIPTION_PROPERTY = "Description";

	public string? CostCenter { get; set; }   // Maps to "Cost Center" column
	public string? PartNumber { get; init; }  // Maps to part number column
	public string? Description { get; init; } // Maps to description column
	public string? Stock { get; init; }       // Maps to stock from vgIcoBillOfMaterials

	// ReSharper disable once InconsistentNaming
	public string? Extent_Area { get; init; } // Maps to Diameter from vgMfiItems

	// ReSharper disable once InconsistentNaming
	public string? Extent_Length { get; init; } // Maps to Length from vgMfiItems

	// ReSharper disable once InconsistentNaming
	public string? Extent_Width { get; init; } // Maps to Width from vgMfiItems
	public string? GeniusMass { get; init; }   // Maps to Mass from vgMfiItems

	// ReSharper disable once InconsistentNaming
	public string? RMQTY { get; init; }    // Maps to RMQTY from vgIcoBillOfMaterials
	public string? RMUNIT { get; init; }   // Maps to RMUNIT from vgIcoBillOfMaterials
	public string? Thickness { get; set; } // Maps to Thickness
	public string? Material { get; init; } // Maps to Specification6 (MS, SS, etc.)

	// Properties for calculated dimensions from Inventor part
	public double CalculatedLength { get; set; }
	public double CalculatedWidth { get; set; }
	public double CalculatedArea { get; set; }
	public double CalculatedMass { get; private set; }

	// Properties for assembly component tracking
	public string? DocumentName { get; set; }
	public string? DocumentType { get; set; }
	public RawStockInfo? RawStockInfo { get; set; }
	public List<PartInfo>? SubAssemblyComponents { get; set; }

	/// <summary>
	///     Calculates the length, width, and area of a sheet metal part from its flat pattern
	/// </summary>
	/// <param name="partDocument">The Inventor part document</param>
	/// <returns>True if calculation was successful, false otherwise</returns>
	public bool CalculateSheetMetalDimensions(PartDocument? partDocument)
	{
		var                            wasUnfolded         = false;
		SheetMetalComponentDefinition? sheetMetalComponent = null;
		try
		{
			if (partDocument == null || !partDocument.SubType.Contains("{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}"))
				// Not a sheet metal part
				return false;

			sheetMetalComponent = partDocument.ComponentDefinition as SheetMetalComponentDefinition;
			if (sheetMetalComponent == null)
				return false;

			// Try to get flat pattern
			if (!sheetMetalComponent.HasFlatPattern)
			{
				// Create flat pattern if it doesn't exist
				sheetMetalComponent.Unfold();
				wasUnfolded = true;
			}

			var flatPattern = sheetMetalComponent.FlatPattern;
			if (flatPattern == null)
				return false;

			// Get dimensions directly from flat pattern properties
			// Convert to desired units (assuming inches for sheet metal)
			var unitsOfMeasure = partDocument.UnitsOfMeasure;
			CalculatedLength = unitsOfMeasure.ConvertUnits(flatPattern.Length, UnitsTypeEnum.kCentimeterLengthUnits,
				UnitsTypeEnum.kInchLengthUnits);
			CalculatedWidth = unitsOfMeasure.ConvertUnits(flatPattern.Width, UnitsTypeEnum.kCentimeterLengthUnits,
				UnitsTypeEnum.kInchLengthUnits);
			CalculatedArea = CalculatedLength * CalculatedWidth;

			// Calculate mass
			CalculatedMass = CalculateMass(partDocument);

			return true;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating sheet metal dimensions: {ex.Message}");
			return false;
		}
		finally
		{
			// Only refold if we unfolded and the component is still valid
			if (wasUnfolded && sheetMetalComponent != null)
				try
				{
					sheetMetalComponent.FlatPattern.ExitEdit();
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"Error refolding sheet metal: {ex.Message}");
				}
		}
	}


	/// <summary>
	///     Gets the calculated area in square feet for RMQTY usage
	/// </summary>
	/// <returns>Area in square feet, rounded to 4 decimal places</returns>
	public double GetCalculatedAreaInSquareFeet()
	{
		// Convert from square inches to square feet
		// 1 square foot = 144 square inches
		return Math.Round(CalculatedArea / 144, 8);
	}

	/// <summary>
	///     Gets the sheet metal thickness from the part
	/// </summary>
	/// <param name="partDocument">The Inventor part document</param>
	/// <returns>Thickness in inches, or 0 if not available</returns>
	public static double GetSheetMetalThickness(PartDocument partDocument)
	{
		try
		{
			var sheetMetalComp = (SheetMetalComponentDefinition)partDocument.ComponentDefinition;

			// Get thickness from the Parameter table
			var thicknessParam = sheetMetalComp.Parameters[THICKNESS_PROPERTY];
			if (thicknessParam == null)
				return 0;

			// Enable Export option if not already enabled
			if (!thicknessParam.ExposedAsProperty) thicknessParam.ExposedAsProperty = true;

			return partDocument.UnitsOfMeasure.ConvertUnits(Convert.ToDouble(thicknessParam.Value),
				UnitsTypeEnum.kCentimeterLengthUnits, UnitsTypeEnum.kInchLengthUnits);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error getting sheet metal thickness: {ex.Message}");
			return 0;
		}
	}

	/// <summary>
	///     Gets the mass of a sheet metal part from the Inventor API
	/// </summary>
	/// <param name="partDocument">The Inventor part document</param>
	/// <returns>Mass in pounds, or 0 if retrieval failed</returns>
	private static double CalculateMass(PartDocument partDocument)
	{
		try
		{
			// Get mass directly from the part document's mass properties
			var massProperties = partDocument.ComponentDefinition.MassProperties;
			if (massProperties == null) return 0;
			// Mass is returned in database units (typically kg for imperial units)
			// Convert to pounds using UnitsOfMeasure
			var massKg = massProperties.Mass;
			var massLb = partDocument.UnitsOfMeasure.ConvertUnits(massKg, UnitsTypeEnum.kKilogramMassUnits,
				UnitsTypeEnum.kLbMassMassUnits);
			return Math.Round(massLb, 4);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error getting mass from API: {ex.Message}");
			return 0;
		}
	}

	/// <summary>
	///     Calculates the bounding box dimensions for a regular (non-sheet metal) part
	/// </summary>
	/// <param name="partDocument">The Inventor part document</param>
	/// <returns>True if calculation was successful, false otherwise</returns>
	public bool CalculatePartDimensions(PartDocument partDocument)
	{
		try
		{
			var partDef = partDocument.ComponentDefinition;

			// Get bounding box
			var rangeBox = partDef.RangeBox;
			if (rangeBox == null)
				return false;

			var unitsOfMeasure = partDocument.UnitsOfMeasure;

			// Calculate dimensions from range box
			var minPoint = rangeBox.MinPoint;
			var maxPoint = rangeBox.MaxPoint;

			if (unitsOfMeasure != null)
			{
				var length = unitsOfMeasure.ConvertUnits(
					maxPoint.X - minPoint.X,
					UnitsTypeEnum.kCentimeterLengthUnits,
					UnitsTypeEnum.kInchLengthUnits);
				var width = unitsOfMeasure.ConvertUnits(
					maxPoint.Y - minPoint.Y,
					UnitsTypeEnum.kCentimeterLengthUnits,
					UnitsTypeEnum.kInchLengthUnits);
				var height = unitsOfMeasure.ConvertUnits(
					maxPoint.Z - minPoint.Z,
					UnitsTypeEnum.kCentimeterLengthUnits,
					UnitsTypeEnum.kInchLengthUnits);

				// Sort dimensions to get Length (largest), Width (middle), Height/Thickness (smallest)
				var dims = new[] { length, width, height }.OrderByDescending(d => d).ToArray();

				CalculatedLength = dims[0];
				CalculatedWidth  = dims[1];
			}

			CalculatedArea = CalculatedLength * CalculatedWidth;

			// Calculate mass
			CalculatedMass = CalculateMass(partDocument);

			return true;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating part dimensions: {ex.Message}");
			return false;
		}
	}

	/// <summary>
	///     Calculates mass for an assembly
	/// </summary>
	/// <param name="assemblyDocument">The Inventor assembly document</param>
	/// <returns>True if calculation was successful, false otherwise</returns>
	public bool CalculateAssemblyMass(AssemblyDocument assemblyDocument)
	{
		try
		{
			var massProperties = assemblyDocument.ComponentDefinition.MassProperties;
			if (massProperties == null)
				return false;

			var massKg = massProperties.Mass;
			var massLb = assemblyDocument.UnitsOfMeasure.ConvertUnits(
				massKg,
				UnitsTypeEnum.kKilogramMassUnits,
				UnitsTypeEnum.kLbMassMassUnits);

			CalculatedMass = Math.Round(massLb, 4);

			return true;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating assembly mass: {ex.Message}");
			return false;
		}
	}

	/// <summary>
	///     Calculates only mass for purchased parts (no dimensions)
	/// </summary>
	/// <param name="document">The Inventor document (part or assembly)</param>
	/// <returns>True if calculation was successful, false otherwise</returns>
	public bool CalculateMassOnly(Document document)
	{
		try
		{
			switch (document.DocumentType)
			{
				case kPartDocumentObject:
					if (document is not PartDocument partDoc) return false;
					CalculatedMass = CalculateMass(partDoc);
					return true;
				case kAssemblyDocumentObject:
					if (document is AssemblyDocument assemblyDoc)
						return CalculateAssemblyMass(assemblyDoc);
					return false;
				default:
					return false;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating mass only: {ex.Message}");
			return false;
		}
	}

	// Helper method to get property value, using calculated values if available
	public static string GetPropertyValueWithCalculated(Document doc, string propertyName, bool hasCalculatedValues,
		PartInfo? calculatedPartInfo, RawStockInfo? calculatedRawStock)
	{
		if (!hasCalculatedValues || calculatedPartInfo == null)
			return GetUserDefinedProperty(doc, propertyName);

		return propertyName switch
		{
			FAMILY_PROPERTY        => calculatedPartInfo.CostCenter ?? GetDesignTrackingProperty(doc, "Cost Center"),
			EXTENT_LENGTH_PROPERTY => $"{calculatedPartInfo.CalculatedLength:F4} in",
			EXTENT_WIDTH_PROPERTY  => $"{calculatedPartInfo.CalculatedWidth:F4} in",
			EXTENT_AREA_PROPERTY   => $"{calculatedPartInfo.CalculatedArea:F4} in^2",
			THICKNESS_PROPERTY     => calculatedPartInfo.Thickness ?? string.Empty,
			GENIUS_MASS_PROPERTY   => $"{calculatedPartInfo.CalculatedMass:F4}",
			RMQTY_PROPERTY         => $"{calculatedPartInfo.GetCalculatedAreaInSquareFeet():F8}",
			"RM"                   => calculatedRawStock?.Stock ?? GetUserDefinedProperty(doc, propertyName),
			RMUNIT_PROPERTY        => calculatedRawStock?.RMUNIT ?? GetUserDefinedProperty(doc, propertyName),
			_                      => GetUserDefinedProperty(doc, propertyName)
		};
	}


	public static async Task LoadServerProperties(List<DocumentProperty> geniusProperties, Document doc,
		Dictionary<string, string> propertyMap, bool isPurchasedPart, bool isAssembly, DatabaseService? databaseService,
		CalculatedData calculatedData)
	{
		// Define consistent property order
		var propertyOrder = new[]
		{
			PART_NUMBER_PROPERTY, DESCRIPTION_PROPERTY, FAMILY_PROPERTY, GENIUS_MASS_PROPERTY, THICKNESS_PROPERTY,
			EXTENT_LENGTH_PROPERTY, EXTENT_WIDTH_PROPERTY, EXTENT_AREA_PROPERTY, RMQTY_PROPERTY, "RM", RMUNIT_PROPERTY
		};

		try
		{
			var validationResult = ValidateServerLookup(doc, databaseService);
			if (!validationResult.IsValid)
			{
				geniusProperties.Add(validationResult.ErrorProperty);
				return;
			}

			var partInfo = await databaseService!.GetPartByNumberAsync(validationResult.PartNumber!);
			if (partInfo == null || string.IsNullOrEmpty(partInfo.PartNumber))
			{
				var notFoundProp = new DocumentProperty();
				notFoundProp.SetPropertyAndValue("Not Found", $"Part '{validationResult.PartNumber}' not in Genius");
				geniusProperties.Add(notFoundProp);
				return;
			}

			var context = new ProcessPropertiesContext
			{
				Document        = doc,
				PropertyMap     = propertyMap,
				PropertyOrder   = propertyOrder,
				PartInfo        = partInfo,
				isPurchasedPart = isPurchasedPart,
				IsAssembly      = isAssembly,
				CalculatedData  = calculatedData
			};

			ProcessServerProperties(geniusProperties, context);
		}
		catch (Exception ex)
		{
			var errorProp = new DocumentProperty();
			errorProp.SetPropertyAndValue("Error", $"Error loading server properties: {ex.Message}");
			geniusProperties.Add(errorProp);
		}
	}

	private static ValidationResult ValidateServerLookup(Document doc, DatabaseService? databaseService)
	{
		var partNumber = GetDesignTrackingProperty(doc, PART_NUMBER_PROPERTY);
		if (string.IsNullOrWhiteSpace(partNumber))
		{
			var noPartProp = new DocumentProperty();
			noPartProp.SetPropertyAndValue("Warning", "No part number found for server lookup");
			return new ValidationResult(false, null, noPartProp);
		}

		if (databaseService != null) return new ValidationResult(true, partNumber, null!);
		var noDbProp = new DocumentProperty();
		noDbProp.SetPropertyAndValue("Warning", "Genius not available");
		return new ValidationResult(false, null, noDbProp);
	}

	private static void ProcessServerProperties(List<DocumentProperty> geniusProperties,
		ProcessPropertiesContext context)
	{
		foreach (var propName in context.PropertyOrder)
		{
			// Skip properties that shouldn't be shown for assemblies
			if (ShouldSkipPropertyForAssembly(propName, context.IsAssembly))
				continue;

			if (propName == FAMILY_PROPERTY)
			{
				ProcessFamilyProperty(geniusProperties, context.Document, context.PropertyMap, propName,
					context.PartInfo, context.CalculatedData);
				continue;
			}

			var (dbValue, suffix, shouldSkip) = GetPropertyMapping(propName, context.PartInfo, context.isPurchasedPart);
			if (shouldSkip || string.IsNullOrEmpty(dbValue))
				continue;

			var fullValue = dbValue + suffix;
			var prop      = new DocumentProperty();
			prop.SetPropertyAndValue(propName, fullValue);

			var documentValue = GetDocumentValueForProperty(context.Document, propName,
				context.CalculatedData.HasCalculatedValues,
				context.CalculatedData.CalculatedPartInfo,
				context.CalculatedData.CalculatedRawStock);
			prop.HasDifference = CalculateHasDifference(propName, fullValue, documentValue);

			geniusProperties.Add(prop);
			context.PropertyMap[propName] = fullValue;
		}
	}

	private static void ProcessFamilyProperty(List<DocumentProperty> geniusProperties, Document doc,
		Dictionary<string, string> propertyMap, string propName, PartInfo partInfo, CalculatedData calculatedData)
	{
		var calculatedFamily = calculatedData is { HasCalculatedValues: true, CalculatedPartInfo.CostCenter: not null }
			? calculatedData.CalculatedPartInfo.CostCenter
			: partInfo.CostCenter;

		if (string.IsNullOrEmpty(calculatedFamily))
		{
			Debug.WriteLine("[FAMILY_DEBUG] Family property SKIPPED - no calculated or database value available");
			return;
		}

		var familyProp = new DocumentProperty();
		familyProp.SetPropertyAndValue(propName, calculatedFamily);

		var familyDocumentValue = GetDocumentValueForProperty(doc, propName, calculatedData.HasCalculatedValues,
			calculatedData.CalculatedPartInfo,
			calculatedData.CalculatedRawStock);
		familyProp.HasDifference = CalculateHasDifference(propName, calculatedFamily, familyDocumentValue);

		geniusProperties.Add(familyProp);
		propertyMap[propName] = calculatedFamily;

		Debug.WriteLine(
			$"[FAMILY_DEBUG] Family property ADDED to grid - calculatedValue: '{calculatedFamily}', documentValue: '{familyDocumentValue}', HasDifference: {familyProp.HasDifference}");
	}

	private static (string? dbValue, string suffix, bool shouldSkip) GetPropertyMapping(string propName,
		PartInfo partInfo,
		bool isPurchasedPart)
	{
		return propName switch
		{
			PART_NUMBER_PROPERTY   => (partInfo.PartNumber, "", false),
			DESCRIPTION_PROPERTY   => (partInfo.Description, "", false),
			FAMILY_PROPERTY        => (partInfo.CostCenter, "", false),
			GENIUS_MASS_PROPERTY   => (partInfo.GeniusMass, "", false),
			THICKNESS_PROPERTY     => (partInfo.Thickness, " in", isPurchasedPart),
			EXTENT_LENGTH_PROPERTY => (partInfo.Extent_Length, " in", isPurchasedPart),
			EXTENT_WIDTH_PROPERTY  => (partInfo.Extent_Width, " in", isPurchasedPart),
			EXTENT_AREA_PROPERTY   => (partInfo.Extent_Area, " in^2", isPurchasedPart),
			RMQTY_PROPERTY         => (partInfo.RMQTY, "", isPurchasedPart),
			"RM"                   => (partInfo.Stock, "", isPurchasedPart),
			RMUNIT_PROPERTY        => (partInfo.RMUNIT, "", isPurchasedPart),
			_                      => (null, "", true)
		};
	}

	public static string GetDatabaseValueForProperty(string propertyName, Dictionary<string, string> databaseProperties)
	{
		return databaseProperties.GetValueOrDefault(propertyName, string.Empty);
	}

	internal static string GetDocumentValueForProperty(Document doc, string propertyName, bool hasCalculatedValues,
		PartInfo? calculatedPartInfo, RawStockInfo? calculatedRawStock)
	{
		return propertyName switch
		{
			FAMILY_PROPERTY => hasCalculatedValues && calculatedPartInfo?.CostCenter != null
				? calculatedPartInfo.CostCenter
				: GetDesignTrackingProperty(doc, "Cost Center"),
			PART_NUMBER_PROPERTY or DESCRIPTION_PROPERTY => GetDesignTrackingProperty(doc, propertyName),
			GENIUS_MASS_PROPERTY or THICKNESS_PROPERTY or EXTENT_LENGTH_PROPERTY or EXTENT_WIDTH_PROPERTY
				or EXTENT_AREA_PROPERTY or RMQTY_PROPERTY or "RM"
				or RMUNIT_PROPERTY
				=> GetPropertyValueWithCalculated(doc, propertyName, hasCalculatedValues, calculatedPartInfo,
					calculatedRawStock),
			_ => GetUserDefinedProperty(doc, propertyName)
		};
	}

	[GeneratedRegex(@"([a-zA-Z^]+.*)$")]
	private static partial Regex MyRegex();

	[GeneratedRegex(@"^([+-]?\d*\.?\d+)")]
	private static partial Regex MyRegex1();

	private static string NormalizeNumericValue(string value)
	{
		if (string.IsNullOrWhiteSpace(value))
			return string.Empty;

		// Trim whitespace
		var normalized = value.Trim();

		// Extract numeric part and units separately
		var numericMatch = MyRegex1().Match(normalized);
		var unitsMatch   = MyRegex().Match(normalized);

		if (!numericMatch.Success) return normalized;
		var numericPart = numericMatch.Groups[1].Value;
		var unitsPart   = unitsMatch.Success ? unitsMatch.Groups[1].Value : string.Empty;

		// Remove trailing zeros after decimal point
		if (numericPart.Contains('.')) numericPart = numericPart.TrimEnd('0').TrimEnd('.');

		return numericPart + (string.IsNullOrEmpty(unitsPart) ? "" : " " + unitsPart);
	}

	internal static bool CalculateHasDifference(string propertyName, string databaseValue, string documentValue)
	{
		if (string.IsNullOrWhiteSpace(databaseValue))
			return string.IsNullOrWhiteSpace(documentValue);

		return propertyName is FAMILY_PROPERTY or "RM" or RMUNIT_PROPERTY
			? CompareValuesExact(databaseValue, documentValue)
			: CompareValues(databaseValue, documentValue);
	}

	public static bool CompareValues(string dbValue, string partValue)
	{
		var normalizedDb   = NormalizeNumericValue(dbValue);
		var normalizedPart = NormalizeNumericValue(partValue);

		var hasDiff = !string.Equals(normalizedDb, normalizedPart, StringComparison.OrdinalIgnoreCase);
		Debug.WriteLine(
			$"Comparing - DB: '{dbValue}' -> '{normalizedDb}' | Part: '{partValue}' -> '{normalizedPart}' | Different: {hasDiff}");
		return hasDiff;
	}

	public static void UpdateDocumentPropertiesWithCalculatedDimensions(Document doc, PartInfo partInfo,
		double thickness, RawStockInfo? calculatedRawStock)
	{
		try
		{
			var propertySet = doc.PropertySets[INVENTOR_USER_DEFINED_PROPERTIES];
			Debug.WriteLine(
				$"Starting property creation. Length: {partInfo.CalculatedLength}, Width: {partInfo.CalculatedWidth}, Area: {partInfo.CalculatedArea}, Thickness: {thickness}");

			// Check if this is a purchased part
			var isPurchasedPart = IsPurchasedPart(doc);
			var isAssembly      = doc.DocumentType == kAssemblyDocumentObject;

			if (isPurchasedPart)
				SetPurchasedPartProperties(propertySet, partInfo);
			else if (isAssembly)
				// For assemblies, only set mass - don't set Extents or RMQTY
				SetAssemblyProperties(propertySet, partInfo);
			else
				SetNonPurchasedPartProperties(propertySet, partInfo, thickness, calculatedRawStock);

			Debug.WriteLine("Property creation completed successfully");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error updating document properties: {ex.Message}");
		}
	}

	private static void SetAssemblyProperties(PropertySet propertySet, PartInfo partInfo)
	{
		// For assemblies, only save GeniusMass
		if (partInfo.CalculatedMass > 0)
			SetOrCreateProperty(propertySet, nameof(GeniusMass), $"{partInfo.CalculatedMass:F4}");
	}

	private static bool IsPurchasedPart(Document doc)
	{
		try
		{
			return doc switch
			{
				PartDocument partDoc => partDoc.ComponentDefinition.BOMStructure ==
				                        BOMStructureEnum.kPurchasedBOMStructure,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.BOMStructure ==
				                            BOMStructureEnum.kPurchasedBOMStructure,
				_ => false
			};
		}
		catch
		{
			// If we can't determine BOM structure, assume it's not purchased
			return false;
		}
	}

	private static bool ShouldSkipPropertyForAssembly(string propName, bool isAssembly)
	{
		if (!isAssembly) return false;

		// For assemblies, only show: Part number, Description, Family, and GeniusMass
		var allowedProperties = new[]
			{ PART_NUMBER_PROPERTY, DESCRIPTION_PROPERTY, FAMILY_PROPERTY, GENIUS_MASS_PROPERTY };
		return !allowedProperties.Contains(propName);
	}

	private static void SetPurchasedPartProperties(PropertySet propertySet, PartInfo partInfo)
	{
		// For purchased parts, only save GeniusMass
		if (partInfo.CalculatedMass > 0)
			SetOrCreateProperty(propertySet, nameof(GeniusMass), $"{partInfo.CalculatedMass:F4}");
	}

	private static void SetNonPurchasedPartProperties(PropertySet propertySet, PartInfo partInfo, double thickness,
		RawStockInfo? calculatedRawStock)
	{
		// Update dimensional properties
		SetOrCreateProperty(propertySet, EXTENT_LENGTH_PROPERTY, $"{partInfo.CalculatedLength:F4} in");
		SetOrCreateProperty(propertySet, EXTENT_WIDTH_PROPERTY, $"{partInfo.CalculatedWidth:F4} in");
		SetOrCreateProperty(propertySet, EXTENT_AREA_PROPERTY, $"{partInfo.CalculatedArea:F4} in^2");

		// Update thickness if valid
		if (thickness > 0)
			SetOrCreateProperty(propertySet, nameof(Thickness), $"{thickness:F3} in");

		// Update quantity and mass
		SetOrCreateProperty(propertySet, nameof(RMQTY), partInfo.GetCalculatedAreaInSquareFeet());

		if (partInfo.CalculatedMass > 0)
			SetOrCreateProperty(propertySet, nameof(GeniusMass), $"{partInfo.CalculatedMass:F4}");

		// Update raw stock information if available
		SetRawStockProperties(propertySet, calculatedRawStock);
	}

	private static void SetRawStockProperties(PropertySet propertySet, RawStockInfo? calculatedRawStock)
	{
		if (calculatedRawStock != null)
		{
			Debug.WriteLine(
				$"[RM_DEBUG] Saving raw stock - Stock: '{calculatedRawStock.Stock}', RMUNIT: '{calculatedRawStock.RMUNIT}'");

			if (!string.IsNullOrEmpty(calculatedRawStock.Stock))
				SetOrCreateProperty(propertySet, "RM", calculatedRawStock.Stock);
			else
				Debug.WriteLine("[RM_DEBUG] RM property not saved - Stock is null or empty");

			if (!string.IsNullOrEmpty(calculatedRawStock.RMUNIT))
				SetOrCreateProperty(propertySet, RMUNIT_PROPERTY, calculatedRawStock.RMUNIT);
			else
				Debug.WriteLine("[RM_DEBUG] RMUNIT property not saved - RMUNIT is null or empty");
		}
		else
		{
			Debug.WriteLine("[RM_DEBUG] Raw stock properties not saved - calculatedRawStock is null");
		}
	}

	public static void SaveUserEditedProperties(Application mInventorApp, DataGrid propertiesDataGrid,
		Dictionary<string, string> databaseProperties, bool hasCalculatedValues, PartInfo? calculatedPartInfo,
		RawStockInfo? calculatedRawStock)
	{
		try
		{
			var doc = mInventorApp.ActiveDocument;
			if (doc == null) return;

			var propertySet = doc.PropertySets[INVENTOR_USER_DEFINED_PROPERTIES];

			if (propertiesDataGrid.ItemsSource is not List<DocumentProperty> documentProperties) return;

			Debug.WriteLine($"[SAVE_DEBUG] Checking {documentProperties.Count} properties for changes");

			var changedProperties = documentProperties.Where(prop => prop.HasChanged).ToList();
			Debug.WriteLine($"[SAVE_DEBUG] Found {changedProperties.Count} changed properties");

			foreach (var prop in changedProperties)
			{
				Debug.WriteLine(
					$"Saving user-edited property: {prop.Property} = '{prop.Value}' (original: '{prop.OriginalValue}')");

				// Special handling for Family property - save to Cost Center in Design Tracking Properties
				if (prop.Property == FAMILY_PROPERTY)
					try
					{
						var designTrackingSet = doc.PropertySets[DESIGN_TRACKING_PROPERTIES];
						SetOrCreateProperty(designTrackingSet, "Cost Center", prop.Value);
						Debug.WriteLine(
							"[PROP_DEBUG] Family property saved to Cost Center in Design Tracking Properties");
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"[PROP_DEBUG] Error saving Family to Cost Center: {ex.Message}");
						// Fallback to saving as Family in User Defined Properties
						SetOrCreateProperty(propertySet, prop.Property, prop.Value);
					}
				else
					SetOrCreateProperty(propertySet, prop.Property, prop.Value);

				// Refresh highlighting for this specific property
				RefreshPropertyDifference(prop, doc, hasCalculatedValues, calculatedPartInfo,
					calculatedRawStock, databaseProperties);
			}

			if (changedProperties.Count == 0)
				Debug.WriteLine("[SAVE_DEBUG] No properties were changed - nothing to save");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error saving user-edited properties: {ex.Message}");
			MessageBox.Show($"Error saving properties: {ex.Message}", "Error",
				MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	public static void SetOrCreateProperty(PropertySet propertySet, string propertyName, object value)
	{
		try
		{
			Debug.WriteLine($"[PROP_DEBUG] Attempting to set {propertyName} = '{value}'");

			var prop = propertySet[propertyName];
			Debug.WriteLine(
				$"[PROP_DEBUG] Updated existing {propertyName} property: '{prop.Value ?? "null"}' -> '{value}'");
			prop.Value = value;
		}
		catch (COMException ex)
		{
			Debug.WriteLine(
				$"[PROP_DEBUG] Property '{propertyName}' not found, attempting to create it. COM error: {ex.Message}");
			try
			{
				propertySet.Add(value, propertyName);
				Debug.WriteLine($"[PROP_DEBUG] Created new {propertyName} property with value '{value}'");
				Debug.WriteLine(
					$"[PROP_DEBUG] Verification - {propertyName} now equals '{propertySet[propertyName].Value ?? "null"}'");
			}
			catch (COMException ex2)
			{
				Debug.WriteLine($"[PROP_DEBUG] COM error creating {propertyName} property: {ex2.Message}");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[PROP_DEBUG] Non-COM error with {propertyName} property: {ex.Message}");
		}
	}

	public static bool CompareValuesExact(string dbValue, string partValue)
	{
		var hasDiff = !string.Equals(dbValue.Trim(), partValue.Trim(), StringComparison.OrdinalIgnoreCase);
		Debug.WriteLine(
			$"Comparing Exact - DB: '{dbValue}' | Part: '{partValue}' | Different: {hasDiff}");
		return hasDiff;
	}

	public static void AddDocumentProperty(List<DocumentProperty> properties, string propertyName, string value,
		Dictionary<string, string> databaseProperties, bool partFoundInGenius, bool compareExact = false)
	{
		var prop = new DocumentProperty();
		prop.SetPropertyAndValue(propertyName, value);

		if (string.IsNullOrWhiteSpace(value))
		{
			prop.HasDifference = true; // Red highlight when empty
		}
		else if (partFoundInGenius)
		{
			var dbValue = databaseProperties.GetValueOrDefault(propertyName, "");
			prop.HasDifference = compareExact
				? CompareValuesExact(dbValue, value)
				: CompareValues(dbValue, value);
		}

		properties.Add(prop);
	}

	public static void RefreshPropertyDifference(DocumentProperty prop, Document doc, bool hasCalculatedValues,
		PartInfo? calculatedPartInfo, RawStockInfo? calculatedRawStock, Dictionary<string, string> databaseProperties)
	{
		try
		{
			Debug.WriteLine($"[HIGHLIGHT_DEBUG] Refreshing difference for {prop.Property}");

			var currentValue = GetDocumentValueForProperty(doc, prop.Property, hasCalculatedValues, calculatedPartInfo,
				calculatedRawStock);
			var databaseValue = GetDatabaseValueForProperty(prop.Property, databaseProperties);

			prop.HasDifference = CalculateHasDifference(prop.Property, databaseValue, currentValue);

			Debug.WriteLine(
				$"[HIGHLIGHT_DEBUG] {prop.Property} - Database: '{databaseValue}', Document: '{currentValue}', HasDifference: {prop.HasDifference}");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[HIGHLIGHT_DEBUG] Error refreshing difference for {prop.Property}: {ex.Message}");
		}
	}

	public static string GetNormalizedMaterial(PartDocument partDoc)
	{
		try
		{
			Debug.WriteLine("[RM_DEBUG] Attempting to get material property...");
			var materialProp = partDoc.PropertySets[DESIGN_TRACKING_PROPERTIES]["Material"];
			if (materialProp is not { Value: not null })
			{
				Debug.WriteLine("[RM_DEBUG] Material property is null or has no value");
				return string.Empty;
			}

			var material = materialProp.Value.ToString() ?? "";
			Debug.WriteLine($"[RM_DEBUG] Raw material property value: '{material}'");

			// Extract material type (MS, SS, etc.) from material name
			var normalized = material switch
			{
				_ when material.Contains("Steel, Mild") || material.Contains("MS")     => "MS",
				_ when material.Contains("Stainless Steel") || material.Contains("SS") => "SS",
				_ when material.Contains("Aluminum") || material.Contains("AL")        => "AL",
				_                                                                      => material
			};

			Debug.WriteLine($"[RM_DEBUG] Normalized material code: '{normalized}'");
			return normalized;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[RM_DEBUG] Error getting material property: {ex.Message}");
			return string.Empty;
		}
	}

	public static bool ShouldSetAllocatedProperty(Document doc)
	{
		try
		{
			var propertySet = doc.PropertySets[INVENTOR_USER_DEFINED_PROPERTIES];

			try
			{
				var allocatedProp = propertySet["ALLOCATED"];
				var currentValue  = allocatedProp.Value?.ToString() ?? string.Empty;

				if (string.IsNullOrWhiteSpace(currentValue))
				{
					Debug.WriteLine("ALLOCATED property should be set to 'Yes'");
					return true;
				}

				Debug.WriteLine($"ALLOCATED property already has value: '{currentValue}'");
				return false;
			}
			catch
			{
				Debug.WriteLine("ALLOCATED property does not exist, should be created");
				return true;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error checking ALLOCATED property: {ex.Message}");
			return false;
		}
	}

	public static void SetAllocatedPropertyIfNotSet(Document doc)
	{
		try
		{
			var propertySet = doc.PropertySets[INVENTOR_USER_DEFINED_PROPERTIES];

			try
			{
				var allocatedProp = propertySet["ALLOCATED"];
				var currentValue  = allocatedProp.Value?.ToString() ?? string.Empty;

				if (string.IsNullOrWhiteSpace(currentValue))
				{
					allocatedProp.Value = "Yes";
					Debug.WriteLine("Set ALLOCATED property to 'Yes'");
				}
				else
				{
					Debug.WriteLine($"ALLOCATED property already has value: '{currentValue}'");
				}
			}
			catch
			{
				propertySet.Add("Yes", "ALLOCATED");
				Debug.WriteLine("Created ALLOCATED property and set to 'Yes'");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error setting ALLOCATED property: {ex.Message}");
		}
	}

	public static void SetDesignTrackingProperty(Document doc, string propName, string value)
	{
		try
		{
			var set  = doc.PropertySets[DESIGN_TRACKING_PROPERTIES];
			var prop = set[propName];
			prop.Value = value;
			Debug.WriteLine($"Set {propName} = '{value}' in Design Tracking Properties");
		}
		catch (COMException ex)
		{
			Debug.WriteLine($"COM exception setting {propName}: {ex.Message}");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Exception setting {propName}: {ex.Message}");
		}
	}

	private static bool IsFactoryDocument(Document doc)
	{
		try
		{
			return doc.DocumentType switch
			{
				kPartDocumentObject when doc is PartDocument partDoc => partDoc.ComponentDefinition.IsiPartFactory,
				kAssemblyDocumentObject when doc is AssemblyDocument assyDoc => assyDoc.ComponentDefinition.IsiAssemblyFactory,
				_ => false
			};
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Error checking if document is factory: {ex.Message}");
			return false;
		}
	}

	private static string? GetFactoryMemberFileName(string factoryFileName, string memberName)
	{
		try
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Getting factory member file for factory: {factoryFileName}, member: {memberName}");
			
			// Get the directory and factory name without extension
			var factoryDir = Path.GetDirectoryName(factoryFileName);
			var factoryName = Path.GetFileNameWithoutExtension(factoryFileName);
			
			if (string.IsNullOrEmpty(factoryDir) || string.IsNullOrEmpty(factoryName))
			{
				Debug.WriteLine($"[THUMBNAIL_DEBUG] Invalid factory file path: {factoryFileName}");
				return null;
			}
			
			// Construct the member folder path (same name as factory)
			var memberFolder = Path.Combine(factoryDir, factoryName);
			
			// Construct the member file path (using member name as filename)
			var memberFile = Path.Combine(memberFolder, $"{memberName}.ipt");
			
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Looking for member file: {memberFile}");
			
			// Check if the member file exists
			if (File.Exists(memberFile))
			{
				Debug.WriteLine($"[THUMBNAIL_DEBUG] Found member file: {memberFile}");
				return memberFile;
			}
			
			// Try with .iam extension for assembly members
			var assemblyMemberFile = Path.Combine(memberFolder, $"{memberName}.iam");
			if (File.Exists(assemblyMemberFile))
			{
				Debug.WriteLine($"[THUMBNAIL_DEBUG] Found assembly member file: {assemblyMemberFile}");
				return assemblyMemberFile;
			}
			
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Member file not found for: {memberName}");
			return null;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Error getting factory member file name: {ex.Message}");
			return null;
		}
	}

	private static BitmapImage? ConvertBitmapToBitmapImage(Bitmap? thumbnail)
	{
		if (thumbnail == null) return null;
		
		using var stream = new MemoryStream();
		thumbnail.Save(stream, ImageFormat.Png);
		stream.Position = 0;

		var bitmapImage = new BitmapImage();
		bitmapImage.BeginInit();
		bitmapImage.StreamSource = stream;
		bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
		bitmapImage.EndInit();
		bitmapImage.Freeze();

		Debug.WriteLine($"[THUMBNAIL_DEBUG] Created BitmapImage: {bitmapImage.Width}x{bitmapImage.Height}");
		return bitmapImage;
	}

	private static bool TryLoadFactoryMemberThumbnail(Document doc, string memberName, System.Windows.Controls.Image thumbnailImage)
	{
		if (string.IsNullOrEmpty(memberName) || !IsFactoryDocument(doc))
			return false;

		Debug.WriteLine($"[THUMBNAIL_DEBUG] Factory document detected, loading thumbnail for member: {memberName}");
		var fileName = doc.FullFileName;
		var memberFileName = GetFactoryMemberFileName(fileName, memberName);
		
		if (string.IsNullOrEmpty(memberFileName) || !File.Exists(memberFileName))
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Factory member file not found: {memberFileName}");
			return false;
		}

		var thumbnail = GetThumbnail(memberFileName);
		Debug.WriteLine($"[THUMBNAIL_DEBUG] GetThumbnail returned: {(thumbnail != null ? $"Bitmap {thumbnail.Width}x{thumbnail.Height}" : "null")}");
		
		if (thumbnail == null) return false;

		var bitmapImage = ConvertBitmapToBitmapImage(thumbnail);
		if (bitmapImage == null) return false;

		thumbnailImage.Source = bitmapImage;
		Debug.WriteLine($"[THUMBNAIL_DEBUG] Set thumbnailImage.Source to BitmapImage");
		Debug.WriteLine("[THUMBNAIL_DEBUG] Factory member thumbnail loaded successfully");
		return true;
	}

	private static bool TryLoadRegularThumbnail(Document doc, System.Windows.Controls.Image thumbnailImage)
	{
		var fileName = doc.FullFileName;
		if (string.IsNullOrEmpty(fileName))
			return false;

		var thumbnail = GetThumbnail(fileName);
		Debug.WriteLine($"[THUMBNAIL_DEBUG] GetThumbnail returned: {(thumbnail != null ? $"Bitmap {thumbnail.Width}x{thumbnail.Height}" : "null")}");
		
		if (thumbnail == null) return false;

		var bitmapImage = ConvertBitmapToBitmapImage(thumbnail);
		if (bitmapImage == null) return false;

		thumbnailImage.Source = bitmapImage;
		Debug.WriteLine($"[THUMBNAIL_DEBUG] Set thumbnailImage.Source to BitmapImage");
		Debug.WriteLine("[THUMBNAIL_DEBUG] Thumbnail loaded successfully");
		return true;
	}

	public static void TryLoadThumbnail(Document doc, System.Windows.Controls.Image thumbnailImage, string? memberName)
	{
		try
		{
			var fileName = doc.FullFileName;
			Debug.WriteLine($"[THUMBNAIL_DEBUG] TryLoadThumbnail called for file: {fileName}, memberName: {memberName}");
			
			// Try factory member thumbnail first
			if (!string.IsNullOrEmpty(memberName) && TryLoadFactoryMemberThumbnail(doc, memberName, thumbnailImage))
				return;
			
			// Try regular thumbnail
			if (TryLoadRegularThumbnail(doc, thumbnailImage))
				return;

			// Thumbnail property doesn't exist or failed - use placeholder
			Debug.WriteLine("[THUMBNAIL_DEBUG] Thumbnail not available, using placeholder");
			thumbnailImage.Source = CreatePlaceholderImage();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Error loading thumbnail: {ex.Message}");
			thumbnailImage.Source = CreatePlaceholderImage();
		}
	}

	public static void TryLoadThumbnail(Document doc, System.Windows.Controls.Image thumbnailImage)
	{
		TryLoadThumbnail(doc, thumbnailImage, null);
	}

	public static Bitmap? GetThumbnail(string strFullFileName)
	{
		Bitmap? imgTemp = null;
		
		using var shFile = Microsoft.WindowsAPICodePack.Shell.ShellFile.FromFilePath(strFullFileName);
		try
		{
			imgTemp = shFile.Thumbnail.MediumBitmap;
		}
		catch (Microsoft.WindowsAPICodePack.Shell.ShellException)
		{
			// Handle Shell exception
		}
		catch (Exception)
		{
			// Handle general exception
		}
		finally
		{
			if (imgTemp == null)
			{
				// Create a simple placeholder bitmap
				imgTemp = new Bitmap(128, 128);
				using var g = Graphics.FromImage(imgTemp);
				g.Clear(Color.LightGray);
				g.DrawString("No Image", new Font("Arial", 8), System.Drawing.Brushes.Black, 5, 40);
			}
		}
		
		return imgTemp;
	}

	public static BitmapImage ConvertBitmapToImageSource(DrawingImage image)
	{
		using var stream = new MemoryStream();
		try
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] Converting DrawingImage: {image.Width}x{image.Height}");

			// Convert DrawingImage to Bitmap
			var drawingVisual = new DrawingVisual();
			using (var drawingContext = drawingVisual.RenderOpen())
			{
				// Use integer dimensions to avoid decimal issues
				var width  = (int)Math.Ceiling(image.Width);
				var height = (int)Math.Ceiling(image.Height);
				drawingContext.DrawImage(image, new Rect(0, 0, width, height));
			}

			// Use integer dimensions for RenderTargetBitmap
			var bitmap = new RenderTargetBitmap(
				(int)Math.Ceiling(image.Width),
				(int)Math.Ceiling(image.Height),
				128, 128,
				PixelFormats.Pbgra32);
			bitmap.Render(drawingVisual);

			var encoder = new PngBitmapEncoder();
			encoder.Frames.Add(BitmapFrame.Create(bitmap));
			encoder.Save(stream);

			stream.Position = 0;

			var imageSource = new BitmapImage();
			imageSource.BeginInit();
			imageSource.StreamSource = stream;
			imageSource.CacheOption  = BitmapCacheOption.OnLoad;
			imageSource.EndInit();
			imageSource.Freeze();

			Debug.WriteLine($"[THUMBNAIL_DEBUG] Final BitmapImage: {imageSource.Width}x{imageSource.Height}");
			return imageSource;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[THUMBNAIL_DEBUG] ConvertBitmapToImageSource failed: {ex.Message}");
			return CreatePlaceholderImage();
		}
	}

	public static BitmapImage CreatePlaceholderImage()
	{
		using var bitmap = new Bitmap(128, 128);
		using var g      = Graphics.FromImage(bitmap);
		g.Clear(Color.Aquamarine);

		using var stream = new MemoryStream();
		bitmap.Save(stream, ImageFormat.Png);
		stream.Position = 0;

		var imageSource = new BitmapImage();
		imageSource.BeginInit();
		imageSource.StreamSource = stream;
		imageSource.CacheOption  = BitmapCacheOption.OnLoad;
		imageSource.EndInit();
		imageSource.Freeze();

		return imageSource;
	}

	internal static string GetDesignTrackingProperty(Document doc, string propName, bool forceNativeProperties = false)
	{
		try
		{
			Debug.WriteLine(
				$"[PROP_DEBUG] GetDesignTrackingProperty: {doc.DisplayName}, Type: {doc.DocumentType}, Property: {propName}, ForceNative: {forceNativeProperties}");

			// When forceNativeProperties is true, we want the part's native properties
			// without any assembly-level overrides
			if (forceNativeProperties && doc.DocumentType == kPartDocumentObject)
				// For parts in assembly context, create a new document reference to get native properties
				// This ensures we get the part's own properties, not assembly overrides
				try
				{
					// Get the part document's full path and reopen it to get native properties
					if (doc is PartDocument partDoc && !string.IsNullOrEmpty(partDoc.FullFileName))
					{
						// Access the property set directly from the part document
						// This bypasses any assembly-level property overrides
						var set   = doc.PropertySets[DESIGN_TRACKING_PROPERTIES];
						var prop  = set[propName];
						var value = prop?.Value?.ToString() ?? string.Empty;
						Debug.WriteLine($"[PROP_DEBUG] Native property value: '{value}'");
						return value;
					}
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"[PROP_DEBUG] Error getting native property: {ex.Message}");
					// Fall back to normal property access
				}

			// Normal property access (may include assembly overrides for components)
			var normalSet   = doc.PropertySets[DESIGN_TRACKING_PROPERTIES];
			var normalProp  = normalSet[propName];
			var normalValue = normalProp?.Value?.ToString() ?? string.Empty;
			Debug.WriteLine($"[PROP_DEBUG] Normal property value: '{normalValue}'");
			return normalValue;
		}
		catch (COMException)
		{
			// Silently handle COM exceptions - these are common when properties don't exist
			return string.Empty;
		}
		catch (Exception ex)
		{
			// Log other exceptions but still return empty string
			Debug.WriteLine($"Non-COM exception in GetDesignTrackingProperty for '{propName}': {ex.Message}");
			return string.Empty;
		}
	}

	public static string GetUserDefinedProperty(Document doc, string propName, bool forceNativeProperties = false)
	{
		try
		{
			// When forceNativeProperties is true, we want the part's native properties
			// without any assembly-level overrides
			if (forceNativeProperties && doc.DocumentType == kPartDocumentObject)
			{
				// Access the property set directly from the part document
				// This ensures we get the native properties, not assembly overrides
				var set  = doc.PropertySets[INVENTOR_USER_DEFINED_PROPERTIES];
				var prop = set[propName];
				return prop?.Value?.ToString() ?? string.Empty;
			}

			// Normal property access (may include assembly overrides for components)
			var normalSet  = doc.PropertySets[INVENTOR_USER_DEFINED_PROPERTIES];
			var normalProp = normalSet[propName];
			return normalProp?.Value?.ToString() ?? string.Empty;
		}
		catch (COMException)
		{
			// Silently handle COM exceptions - these are common when properties don't exist
			// or when Inventor is in an inconsistent state
			return string.Empty;
		}
		catch (Exception ex)
		{
			// Log other exceptions but still return empty string
			Debug.WriteLine($"Non-COM exception in GetUserDefinedProperty for '{propName}': {ex.Message}");
			return string.Empty;
		}
	}

	public static Task RefreshDataGridsWithCalculatedValues(Application inventorApp, DataGrid geniusPropertiesGrid,
		bool hasCalculatedValues, PartInfo? calculatedPartInfo, RawStockInfo? calculatedRawStock,
		IDictionary<string, string> databaseProperties, Document? targetDocument = null)
	{
		// Refresh the data grids to show calculated values (but preserve database properties)
		var savedDatabaseProperties = new Dictionary<string, string>(databaseProperties);
		// Clear and restore database properties to force refresh
		databaseProperties.Clear();
		foreach (var kvp in savedDatabaseProperties) databaseProperties[kvp.Key] = kvp.Value;

		// Update Genius DataGrid highlighting to reflect new calculated values
		System.Windows.Application.Current.Dispatcher.Invoke(() =>
		{
			if (geniusPropertiesGrid?.ItemsSource is not List<DocumentProperty> geniusItems) return;

			var documentToUse = targetDocument ?? inventorApp.ActiveDocument;
			foreach (var geniusProp in geniusItems)
			{
				var documentValue = GetDocumentValueForProperty(documentToUse, geniusProp.Property,
					hasCalculatedValues, calculatedPartInfo, calculatedRawStock);
				var databaseValue =
					databaseProperties.TryGetValue(geniusProp.Property, out var dbValue) ? dbValue : string.Empty;

				geniusProp.HasDifference =
					CalculateHasDifference(geniusProp.Property, databaseValue, documentValue);
			}

			// Force complete rebind for UI refresh
			var newGeniusItems = new List<DocumentProperty>(geniusItems);
			geniusPropertiesGrid.ItemsSource = null;
			geniusPropertiesGrid.ItemsSource = newGeniusItems;
		});
		return Task.CompletedTask;
	}

	public class CalculatedData
	{
		public bool HasCalculatedValues { get; init; }
		public PartInfo? CalculatedPartInfo { get; init; }
		public RawStockInfo? CalculatedRawStock { get; init; }
	}

	private sealed class ProcessPropertiesContext
	{
		public required Document Document { get; init; }
		public Dictionary<string, string> PropertyMap { get; init; } = new();
		public string[] PropertyOrder { get; init; } = [];
		public PartInfo PartInfo { get; init; } = null!;
		public bool isPurchasedPart { get; init; }
		public bool IsAssembly { get; init; }
		public CalculatedData CalculatedData { get; init; } = null!;
	}

	private sealed class ValidationResult(bool isValid, string? partNumber, DocumentProperty errorProperty)
	{
		public bool IsValid { get; } = isValid;
		public string? PartNumber { get; } = partNumber;
		public DocumentProperty ErrorProperty { get; } = errorProperty;
	}

	public sealed class DocumentProperty : INotifyPropertyChanged
	{
		private string value = string.Empty;

		public bool HasChanged => value != OriginalValue;

		public string OriginalValue { get; private set; } = string.Empty;

		public bool HasDifference
		{
			get;
			set
			{
				if (field == value) return;
				field = value;
				Debug.WriteLine($"[PROPERTY_CHANGED] HasDifference changed to {value} for property '{Property}'");
				PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasDifference)));
				// Also force a PropertyChanged for a dummy property to ensure UI updates
				PropertyChanged?.Invoke(this, new PropertyChangedEventArgs("ForceUIUpdate"));
			}
		}

		public string Property
		{
			get;
			set
			{
				if (field == value) return;
				field = value;
				PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Property)));
			}
		} = string.Empty;

		public string Value
		{
			get => value;
			set
			{
				if (value == this.value) return;
				// Only update originalValue if this is the first time setting the value
				// Don't update originalValue when the current value equals originalValue (this prevents tracking changes)
				if (string.IsNullOrEmpty(OriginalValue)) OriginalValue = this.value;

				this.value = value;
				PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Value)));
				PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(HasChanged)));

				Debug.WriteLine(
					$"Value changed: {Property} = '{value}' (original: '{OriginalValue}', HasChanged: {HasChanged})");
			}
		}

		public event PropertyChangedEventHandler? PropertyChanged;

		// Dummy property to force UI updates when needed

		private void SetOriginalValue(string propertyValue)
		{
			OriginalValue = propertyValue;
			value         = propertyValue;
		}

		public void SetPropertyAndValue(string propertyName, string propertyValue)
		{
			Property = propertyName;
			SetOriginalValue(propertyValue);
		}
	}
}

public class RawStockInfo
{
	public string? Stock { get; init; }     // Raw stock item number (maps to m.Item)
	public string? Length { get; init; }    // Extent Length (maps to m.Length)
	public string? Width { get; init; }     // Extent Width (maps to m.Width)
	public string? Thickness { get; init; } // Thickness (maps to m.Thickness)
	public string? Material { get; init; }  // Material type (maps to m.Specification6)
	public string? RMUNIT { get; init; }    // Raw material unit (maps to b.ConversionUnit)
}

/// <summary>
///     Helper Class to convert ActiveX images to .NET Images
/// </summary>
internal class AxHostConverter : AxHost
{
	private AxHostConverter() : base("")
	{
	}

	public static System.Drawing.Image? PictureDispToImage(object pictureDisp)
	{
		return GetPictureFromIPicture(pictureDisp);
	}
}