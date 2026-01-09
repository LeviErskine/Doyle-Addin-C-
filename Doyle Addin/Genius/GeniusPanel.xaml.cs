#region

#nullable enable
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using MessageBox = System.Windows.MessageBox;

#endregion

namespace Doyle_Addin.Genius;

public partial class GeniusPanel
{
	private const string PropertyName = "Part Number";
	private const string Description = "Description";
	private const string Family = "Family";
	private const string Geniusmass = "GeniusMass";
	private const string RMUNIT = "RMUNIT";

	// Stores database properties for comparison during editing
	private readonly Dictionary<string, string> databaseProperties = new();
	// Win32 interop for key translation

	// Database service for SQL Server operations
	private readonly DatabaseService? databaseService;

	// Holds a reference to the Inventor Application object
	private readonly Application mInventorApp;

	// Reference to the parent NewGenius window for closing
	private readonly NewGenius? parentWindow;

	// Stores calculated dimension values in memory (without dirtying document)
	private PartInfo? calculatedPartInfo;

	private RawStockInfo? calculatedRawStock;

	// Localized editor handling
	private bool hasCalculatedValues;

	// Tracks whether ALLOCATED property should be set when saving
	private bool shouldSetAllocatedProperty;

	// --- PUBLIC CONSTRUCTOR ---
	public GeniusPanel(Application inventorApp, DatabaseService? databaseService, NewGenius? parentWindow = null)
	{
		InitializeComponent();
		mInventorApp         = inventorApp;
		this.databaseService = databaseService;
		this.parentWindow    = parentWindow;

		// Load properties immediately on window creation
		_ = InitializePanel();
	}


	// --- PUBLIC INITIALIZATION METHOD ---
	private async Task InitializePanel()
	{
		await LoadDocumentProperties();
		await Dispatcher.BeginInvoke(() => PropertiesDataGrid.Focus(), DispatcherPriority.Loaded);
	}

	private async void Window_Activated(object sender, EventArgs e)
	{
		// Avoid reloading while user is editing a cell, which would steal focus
		await InitializePanel();
	}

	private async void SaveButton_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			// Save user-edited properties first
			PartInfo.SaveUserEditedProperties(mInventorApp, PropertiesDataGrid, databaseProperties, hasCalculatedValues,
				calculatedPartInfo,
				calculatedRawStock);

			// Set Family to "D-PTS" for purchased parts
			if (mInventorApp.ActiveDocument != null && await IsPurchasedPart(mInventorApp.ActiveDocument))
				PartInfo.SetDesignTrackingProperty(mInventorApp.ActiveDocument, "Cost Center", "D-PTS");

			// Apply calculated values to document if they exist
			if (hasCalculatedValues && calculatedPartInfo != null && mInventorApp.ActiveDocument != null)
			{
				var doc       = mInventorApp.ActiveDocument;
				var thickness = double.Parse(calculatedPartInfo.Thickness?.Replace(" in", "") ?? "0");
				PartInfo.UpdateDocumentPropertiesWithCalculatedDimensions(doc, calculatedPartInfo, thickness,
					calculatedRawStock);
			}

			// Set ALLOCATED property if needed
			if (shouldSetAllocatedProperty && mInventorApp.ActiveDocument != null)
				PartInfo.SetAllocatedPropertyIfNotSet(mInventorApp.ActiveDocument);

			DialogResult = true;
		}
		catch
		{
			/* ignore if not shown as dialog */
		}

		mInventorApp.ActiveDocument?.Save2();
		Debug.WriteLine($"[SAVE_DEBUG] parentWindow is null: {parentWindow == null}");
		if (parentWindow != null)
		{
			Debug.WriteLine("[SAVE_DEBUG] Calling CloseDockableWindow");
			parentWindow.CloseDockableWindow();
		}
		else
		{
			Debug.WriteLine("[SAVE_DEBUG] Cannot close window - parentWindow is null");
		}
	}

	private void CancelButton_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			DialogResult = false;
		}
		catch
		{
			/* ignore */
		}

		parentWindow?.CloseDockableWindow();
	}


	private async Task<bool> IsPurchasedPart(Document doc)
	{
		try
		{
			Debug.WriteLine($"[PURCHASED_DEBUG] IsPurchasedPart called for: {doc.FullFileName}");

			// First check if BOM structure is already set to purchased
			var isPurchasedByBom = doc switch
			{
				PartDocument partDoc => partDoc.ComponentDefinition.BOMStructure ==
				                        BOMStructureEnum.kPurchasedBOMStructure,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.BOMStructure ==
				                            BOMStructureEnum.kPurchasedBOMStructure,
				_ => false
			};

			Debug.WriteLine($"[PURCHASED_DEBUG] Is purchased by BOM: {isPurchasedByBom}");

			if (isPurchasedByBom)
				return true;

			// Check if file is in purchased directories or has purchased family
			var shouldAskUser = ShouldAskAboutPurchased(doc);
			Debug.WriteLine($"[PURCHASED_DEBUG] Should ask user: {shouldAskUser}");

			if (!shouldAskUser)
				return false;

			// Ask the user if this is a purchased part using a simple message box
			Debug.WriteLine("[PURCHASED_DEBUG] Showing user prompt");
			var result = MessageBox.Show(
				"Is this a Purchased Part?",
				"Purchased Part Confirmation",
				MessageBoxButton.YesNo,
				MessageBoxImage.Question);

			Debug.WriteLine($"[PURCHASED_DEBUG] User response: {result}");

			if (result != MessageBoxResult.Yes) return false;
			await SetBomStructureToPurchased(doc);

			return true;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[PURCHASED_DEBUG] Exception in IsPurchasedPart: {ex.Message}");
			// If we can't determine BOM structure, assume it's not purchased
			return false;
		}
	}

	private async Task SetBomStructureToPurchased(Document doc)
	{
		// Change the BOM structure to Purchased
		try
		{
			switch (doc)
			{
				case PartDocument partDoc:
					partDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure;
					break;
				case AssemblyDocument assyDoc:
					assyDoc.ComponentDefinition.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure;
					break;
			}

			Debug.WriteLine("[PURCHASED_DEBUG] BOM structure set to Purchased");

			// Reload the table to show proper properties
			await PartInfo.RefreshDataGridsWithCalculatedValues(mInventorApp, GeniusProperties, hasCalculatedValues,
				calculatedPartInfo, calculatedRawStock, databaseProperties, doc);
			Debug.WriteLine("[PURCHASED_DEBUG] Table reloaded");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[PURCHASED_DEBUG] Error setting BOM structure: {ex.Message}");
		}
	}

	private static bool ShouldAskAboutPurchased(Document doc)
	{
		try
		{
			var fileName = doc.FullFileName;
			Debug.WriteLine($"[PURCHASED_DEBUG] Checking file: {fileName}");

			// Check if file is in purchased directories
			var isInPurchasedDirectory =
				fileName.Contains(@"\Doyle_Vault\Designs\purchased\", StringComparison.OrdinalIgnoreCase) ||
				fileName.Contains(@"\Doyle_Vault\Designs\riverview\RIVERVIEW PURCHASED PARTS\",
					StringComparison.OrdinalIgnoreCase);

			Debug.WriteLine($"[PURCHASED_DEBUG] In purchased directory: {isInPurchasedDirectory}");

			if (isInPurchasedDirectory)
				return true;

			// Check if part family indicates purchased part
			var purchasedFamilies = new[] { "D-HDWR", "D-PTO", "D-PTS", "R-PTO", "R-PTS" };

			try
			{
				var designProps    = doc.PropertySets["Design Tracking Properties"];
				var costCenterProp = designProps["Cost Center"];
				var familyValue    = costCenterProp?.Value?.ToString();

				Debug.WriteLine($"[PURCHASED_DEBUG] Family (Cost Center) value: {familyValue}");

				if (!string.IsNullOrEmpty(familyValue) && purchasedFamilies.Contains(familyValue))
					return true;
			}
			catch
			{
				// Cost Center property not found or inaccessible
				Debug.WriteLine("[PURCHASED_DEBUG] Cost Center property not found or inaccessible");
			}

			Debug.WriteLine("[PURCHASED_DEBUG] Should not ask about purchased");
			return false;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"[PURCHASED_DEBUG] Exception in ShouldAskAboutPurchased: {ex.Message}");
			return false;
		}
	}


	private async void CalculateProps_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			var doc = mInventorApp.ActiveDocument;
			if (doc == null) return;

			// Check for purchased parts (both parts and assemblies)
			if (await IsPurchasedPart(doc))
			{
				var result =
					await DocumentHandles.HandlePurchasedPart(mInventorApp, GeniusProperties, databaseProperties);
				calculatedPartInfo  = result.calculatedPartInfo;
				hasCalculatedValues = result.hasCalculatedValues;

				// Refresh the UI to show calculated values
				await LoadDocumentProperties();
				return;
			}

			switch (doc.DocumentType)
			{
				case kPartDocumentObject:
					var partDoc = doc as PartDocument;

					switch (partDoc)
					{
						// Regular sheet metal part logic
						case { ComponentDefinition: SheetMetalComponentDefinition }:
							(calculatedPartInfo, hasCalculatedValues, calculatedRawStock) =
								await DocumentHandles.HandleSheetMetalPart(partDoc, mInventorApp, GeniusProperties,
									databaseProperties);

							// Set Cost Center to D-RMTO for sheetmetal parts in memory
							if (calculatedPartInfo != null)
							{
								calculatedPartInfo.CostCenter = "D-RMTO";
								// Update the databaseProperties to reflect the change
								databaseProperties["Family"] = "D-RMTO";
							}

							break;
						// Regular (non-sheet metal) part
						default:
						{
							if (partDoc != null)
								(calculatedPartInfo, hasCalculatedValues) =
									await DocumentHandles.HandleRegularPart(partDoc, mInventorApp, GeniusProperties,
										databaseProperties);
						}
							break;
					}

					break;

				case kAssemblyDocumentObject:
					// Regular assembly - just calculate mass
					if (doc is AssemblyDocument assyDoc)
						(calculatedPartInfo, hasCalculatedValues, _) =
							await DocumentHandles.HandleAssembly(assyDoc, mInventorApp, GeniusProperties,
								databaseProperties);
					break;
			}

			// Refresh the UI to show calculated values
			await LoadDocumentProperties();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating dimensions: {ex.Message}");
		}
	}

	private static bool CheckIfPurchasedPart(Document doc)
	{
		try
		{
			return doc switch
			{
				// Cast to appropriate document type to access BOMStructure
				PartDocument partDoc => partDoc.ComponentDefinition.BOMStructure ==
				                        BOMStructureEnum.kPurchasedBOMStructure,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.BOMStructure ==
				                            BOMStructureEnum.kPurchasedBOMStructure,
				_ => false
			};
		}
		catch
		{
			return false;
		}
	}

	private static void LoadFilteredProperties(List<PartInfo.DocumentProperty> documentProperties,
		List<PartInfo.DocumentProperty> geniusProperties)
	{
		// For purchased parts and assemblies, only show: Part number, Description, Family, and Geniusmass
		var allowedProperties = new[] { PropertyName, Description, Family, Geniusmass };

		// Filter documentProperties to only include allowed properties
		var filteredDocumentProperties = documentProperties
		                                 .Where(prop => allowedProperties.Contains(prop.Property))
		                                 .ToList();

		// Filter geniusProperties to only include allowed properties
		var filteredGeniusProperties = geniusProperties
		                               .Where(prop => allowedProperties.Contains(prop.Property))
		                               .ToList();

		// Update the original lists
		documentProperties.Clear();
		documentProperties.AddRange(filteredDocumentProperties);
		geniusProperties.Clear();
		geniusProperties.AddRange(filteredGeniusProperties);
	}

	private void LoadGeniusFoundProperties(Document doc, List<PartInfo.DocumentProperty> documentProperties)
	{
		// For non-purchased parts found in Genius, show all properties that the server returned
		foreach (var dbProp in databaseProperties)
		{
			// Skip properties already handled above
			if (dbProp.Key is PropertyName or Description or Family or Geniusmass or "Thickness"
			    or "Extent_Length" or "Extent_Width" or "Extent_Area" or "RMQTY" or "RM" or RMUNIT)
				continue;

			// Show all properties that the server returned, even if they don't exist in the document
			var val  = PartInfo.GetUserDefinedProperty(doc, dbProp.Key);
			var prop = new PartInfo.DocumentProperty();
			prop.SetPropertyAndValue(dbProp.Key, val);
			prop.HasDifference = PartInfo.CompareValues(dbProp.Value, val);
			Debug.WriteLine($"Property: {dbProp.Key} | HasDifference: {prop.HasDifference}");
			documentProperties.Add(prop);
		}
	}

	private static void LoadAllDocumentProperties(Document doc, List<PartInfo.DocumentProperty> documentProperties)
	{
		// Part not found in Genius - show all document properties
		Debug.WriteLine("Part not found in Genius - showing all document properties");

		// Get all user-defined properties
		try
		{
			var userDefinedProps = doc.PropertySets["Inventor User Defined Properties"];
			foreach (Property prop in userDefinedProps)
			{
				// Skip properties already handled above
				if (prop.Name is Geniusmass or "ALLOCATED" or "Thickness" or "Extent_Length"
				    or "Extent_Width" or "Extent_Area" or "RMQTY" or "RM" or RMUNIT)
					continue;

				var val = prop.Value?.ToString() ?? "";
				if (string.IsNullOrWhiteSpace(val)) continue;
				var docProp = new PartInfo.DocumentProperty();
				docProp.SetPropertyAndValue(prop.Name, val);
				// No comparison needed since part is not in Genius
				documentProperties.Add(docProp);
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error loading user-defined properties: {ex.Message}");
		}
	}

	private async Task UpdateUIWithProperties(List<PartInfo.DocumentProperty> geniusProperties,
		List<PartInfo.DocumentProperty> documentProperties)
	{
		// Set data sources for both grids on UI thread
		await Dispatcher.InvokeAsync(() =>
		{
			GeniusProperties.ItemsSource   = geniusProperties;
			PropertiesDataGrid.ItemsSource = documentProperties;

			// Debug: Show what's actually being displayed
			Debug.WriteLine($"[UI_DEBUG] Displaying {documentProperties.Count} properties in Document DataGrid:");
			foreach (var prop in documentProperties)
				Debug.WriteLine(
					$"[UI_DEBUG] Document Property: {prop.Property} = '{prop.Value}' (HasDifference: {prop.HasDifference})");

			// Debug: Force UI refresh
			GeniusProperties.Items.Refresh();
			PropertiesDataGrid.Items.Refresh();
		});
	}

	private async Task LoadDocumentProperties()
	{
		try
		{
			var geniusProperties   = new List<PartInfo.DocumentProperty>();
			var documentProperties = new List<PartInfo.DocumentProperty>();

			if (!await InitializePropertyLoading(documentProperties))
				return;

			var doc             = mInventorApp.ActiveDocument;
			var isPurchasedPart = CheckIfPurchasedPart(doc);
			var isAssembly      = doc.DocumentType == kAssemblyDocumentObject;

			var calculatedData = new PartInfo.CalculatedData
			{
				HasCalculatedValues = hasCalculatedValues,
				CalculatedPartInfo  = calculatedPartInfo,
				CalculatedRawStock  = calculatedRawStock
			};
			await PartInfo.LoadServerProperties(geniusProperties, doc, databaseProperties, isPurchasedPart, isAssembly,
				databaseService, calculatedData);
			var partFoundInGenius = databaseProperties.Count > 0;

			LoadBasicDocumentProperties(doc, documentProperties, partFoundInGenius);
			LoadAdditionalDocumentProperties(doc, documentProperties, partFoundInGenius, isPurchasedPart);
			LoadRemainingProperties(doc, documentProperties, geniusProperties, isPurchasedPart, partFoundInGenius);

			shouldSetAllocatedProperty = PartInfo.ShouldSetAllocatedProperty(doc);
			await UpdateUIWithProperties(geniusProperties, documentProperties);
		}
		catch (Exception ex)
		{
			await HandlePropertyLoadingError(ex);
		}
	}

	private async Task<bool> InitializePropertyLoading(List<PartInfo.DocumentProperty> documentProperties)
	{
		databaseProperties.Clear();

		if (ValidateActiveDocument(out documentProperties)) return true;
		await Dispatcher.InvokeAsync(() => { PropertiesDataGrid.ItemsSource = documentProperties; });
		return true;
	}

	private void LoadBasicDocumentProperties(Document doc, List<PartInfo.DocumentProperty> documentProperties,
		bool partFoundInGenius)
	{
		var partNumber = PartInfo.GetDesignTrackingProperty(doc, PropertyName);
		PartInfo.AddDocumentProperty(documentProperties, PropertyName, partNumber,
			databaseProperties, partFoundInGenius, true);

		PartInfo.TryLoadThumbnail(doc, ThumbnailImage);

		var description = PartInfo.GetDesignTrackingProperty(doc, Description);
		PartInfo.AddDocumentProperty(documentProperties, Description, description,
			databaseProperties, partFoundInGenius, true);
	}

	private void LoadAdditionalDocumentProperties(Document doc, List<PartInfo.DocumentProperty> documentProperties,
		bool partFoundInGenius, bool isPurchasedPart)
	{
		var propertyOrder = GetPropertyOrder();
		var isAssembly    = doc.DocumentType == kAssemblyDocumentObject;

		foreach (var propName in propertyOrder.Skip(2)) // Skip first 2 already handled
		{
			if (ShouldSkipPropertyForPurchasedPart(propName, isPurchasedPart))
				continue;

			if (ShouldSkipPropertyForAssembly(propName, isAssembly))
				continue;

			var value = GetPropertyValueWithFamilyOverride(doc, propName);
			PartInfo.AddDocumentProperty(documentProperties, propName, value, databaseProperties, partFoundInGenius,
				propName is "RM" or RMUNIT);
		}
	}

	private static string[] GetPropertyOrder()
	{
		return
		[
			PropertyName, Description, Family, Geniusmass, "Thickness",
			"Extent_Length", "Extent_Width", "Extent_Area", "RMQTY", "RM", RMUNIT
		];
	}

	private static bool ShouldSkipPropertyForPurchasedPart(string propName, bool isPurchasedPart)
	{
		if (!isPurchasedPart) return false;

		// For purchased parts, only show: Part number, Description, Family, and geniusmass
		var allowedProperties = new[] { PropertyName, Description, Family, Geniusmass };
		return !allowedProperties.Contains(propName);
	}

	private static bool ShouldSkipPropertyForAssembly(string propName, bool isAssembly)
	{
		if (!isAssembly) return false;

		// For assemblies, only show: Part number, Description, Family, and GeniusMass
		var allowedProperties = new[] { PropertyName, Description, Family, Geniusmass };
		return !allowedProperties.Contains(propName);
	}

	private string GetPropertyValueWithFamilyOverride(Document doc, string propName)
	{
		var value = PartInfo.GetPropertyValueWithCalculated(doc, propName, hasCalculatedValues,
			calculatedPartInfo, calculatedRawStock);

		switch (propName)
		{
			// If we already have a valid Family value (from PartInfo.cs), preserve it
			case Family when !string.IsNullOrEmpty(value):
				Debug.WriteLine($"[FAMILY_DEBUG] Preserving existing Family value: '{value}'");
				break;
			// Otherwise try to get calculated value if available
			case Family when (hasCalculatedValues && calculatedPartInfo?.CostCenter != null):
				value = calculatedPartInfo.CostCenter;
				Debug.WriteLine(
					$"[FAMILY_DEBUG] Using calculated Family value: '{value}', hasCalculatedValues: {hasCalculatedValues}");
				break;
			case Family:
				// Fallback: try to get Cost Center from Design Tracking Properties
				value = PartInfo.GetDesignTrackingProperty(doc, "Cost Center");
				Debug.WriteLine(
					$"[FAMILY_DEBUG] Family property fallback - hasCalculatedValues: {hasCalculatedValues}, calculatedPartInfo?.CostCenter: '{calculatedPartInfo?.CostCenter}', fallback value: '{value}'");
				break;
			case RMUNIT:
				Debug.WriteLine(
					$"[RMUNIT_DEBUG_LOAD] rmunit value: '{value}', hasCalculatedValues: {hasCalculatedValues}");
				break;
		}

		return value;
	}

	private void LoadRemainingProperties(Document doc, List<PartInfo.DocumentProperty> documentProperties,
		List<PartInfo.DocumentProperty> geniusProperties, bool isPurchasedPart, bool partFoundInGenius)
	{
		if (documentProperties.Count == 0)
		{
			var statusProp = new PartInfo.DocumentProperty();
			statusProp.SetPropertyAndValue("Status", "No tracked properties found.");
			documentProperties.Add(statusProp);
			return;
		}

		var isAssembly = doc.DocumentType == kAssemblyDocumentObject;

		if (isPurchasedPart || isAssembly)
			LoadFilteredProperties(documentProperties, geniusProperties);
		else
			LoadNonPurchasedPartProperties(doc, documentProperties, partFoundInGenius);
	}

	private void LoadNonPurchasedPartProperties(Document doc, List<PartInfo.DocumentProperty> documentProperties,
		bool partFoundInGenius)
	{
		if (partFoundInGenius)
			LoadGeniusFoundProperties(doc, documentProperties);
		else
			LoadAllDocumentProperties(doc, documentProperties);
	}

	private async Task HandlePropertyLoadingError(Exception ex)
	{
		var errorProperties = new List<PartInfo.DocumentProperty>();
		var errorProp       = new PartInfo.DocumentProperty();
		errorProp.SetPropertyAndValue("Error", $"Error loading properties: {ex.Message}");
		errorProperties.Add(errorProp);

		await Dispatcher.InvokeAsync(() =>
		{
			PropertiesDataGrid.ItemsSource = errorProperties;
			GeniusProperties.ItemsSource   = errorProperties;
		});
	}

	private bool ValidateActiveDocument(out List<PartInfo.DocumentProperty> documentProperties)
	{
		documentProperties = [];

		if (mInventorApp.ActiveDocument != null) return true;
		var statusProp = new PartInfo.DocumentProperty();
		statusProp.SetPropertyAndValue("Status", "No active document.");
		documentProperties.Add(statusProp);
		return false;
	}

	// --- DATAGRID DOUBLE-CLICK EDIT HANDLER ---
	private void PropertiesDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		var dependencyObject = (DependencyObject)e.OriginalSource;

		// Find the cell that was double-clicked
		while (dependencyObject != null && dependencyObject is not DataGridCell)
			dependencyObject = VisualTreeHelper.GetParent(dependencyObject);

		if (dependencyObject is not DataGridCell { Column: DataGridTextColumn column } cell) return;
		// Only allow editing the "Value" column
		if (column.Header?.ToString() != "Value" || cell.DataContext is not PartInfo.DocumentProperty property) return;
		var editDialog = new EditValueDialog(property.Property, property.Value);

		if (editDialog.ShowDialog() != true) return;
		property.Value = editDialog.PropertyValue;
		Debug.WriteLine($"[EDIT_DEBUG] Property updated: {property.Property} = '{property.Value}'");

		// Update highlighting by comparing new value directly with database
		var databaseValue = PartInfo.GetDatabaseValueForProperty(property.Property, databaseProperties);
		property.HasDifference = PartInfo.CalculateHasDifference(property.Property, databaseValue, property.Value);

		Debug.WriteLine(
			$"[HIGHLIGHT_DEBUG] {property.Property} - Database: '{databaseValue}', New Value: '{property.Value}', HasDifference: {property.HasDifference}");

		// Also update the corresponding property in the Genius DataGrid if it exists
		Dispatcher.BeginInvoke(() =>
		{
			if (GeniusProperties.ItemsSource is not List<PartInfo.DocumentProperty> geniusItems) return;
			var geniusProperty = geniusItems.FirstOrDefault(gp => gp.Property == property.Property);
			if (geniusProperty == null) return;
			// Update the HasDifference for the Genius property to match
			geniusProperty.HasDifference = property.HasDifference;
			Debug.WriteLine(
				$"[GENIUS_DEBUG] Updated Genius property {property.Property} HasDifference to {property.HasDifference}");
		}, DispatcherPriority.DataBind);

		// Force the DataGrid to completely rebind by recreating the ItemsSource
		Dispatcher.BeginInvoke(() =>
		{
			if (PropertiesDataGrid.ItemsSource is not List<PartInfo.DocumentProperty> currentItems) return;
			// Create a new list to force rebinding
			var newItems = new List<PartInfo.DocumentProperty>(currentItems);
			PropertiesDataGrid.ItemsSource = null;
			PropertiesDataGrid.ItemsSource = newItems;
		}, DispatcherPriority.DataBind);
	}
}