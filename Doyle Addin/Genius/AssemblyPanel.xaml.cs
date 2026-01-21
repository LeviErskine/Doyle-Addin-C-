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

// ReSharper disable once InconsistentNaming
public partial class AssemblyPanel
{
	// Constant for error message titles
	private const string ErrorTitle = Caption;
	private const string Summary = "Summary";
	private const string RMUNIT = "RMUNIT";
	private const string Separator = "Separator";

	// Property name constants
	private const string PartNumberProperty = "Part Number";
	private const string DescriptionProperty = "Description";
	private const string FamilyProperty = "Family";
	private const string GeniusMassProperty = "GeniusMass";
	private const string ExtentLengthProperty = "Extent_Length";
	private const string ExtentWidthProperty = "Extent_Width";
	private const string ExtentAreaProperty = "Extent_Area";
	private const string Caption = "Error";

	// Stores calculated dimension values in memory (without dirtying document)
	private readonly PartInfo? calculatedPartInfo;

	private readonly RawStockInfo? calculatedRawStock;

	// List of component parts from assembly
	private readonly List<PartInfo> componentParts;

	// Stores database properties for comparison during editing
	private readonly Dictionary<string, string> databaseProperties = [];

	// Database service for loading Genius properties
	private readonly DatabaseService databaseService;

	// Localized editor handling
	private readonly bool hasCalculatedValues;

	// Holds a reference to the Inventor Application object
	private readonly Application mInventorApp;

	// Stores the original assembly document this panel was opened on
	private readonly Document? originalAssemblyDocument;
	private readonly NewGenius? parentWindow;

	public AssemblyPanel(Application inventorApp, PartInfo? calculatedPartInfo, RawStockInfo? calculatedRawStock,
		bool hasCalculatedValues, List<PartInfo> componentParts, NewGenius? parentWindow = null)
	{
		mInventorApp             = inventorApp;
		this.calculatedPartInfo  = calculatedPartInfo;
		this.calculatedRawStock  = calculatedRawStock;
		this.hasCalculatedValues = hasCalculatedValues;
		this.componentParts      = componentParts;
		this.parentWindow        = parentWindow;
		originalAssemblyDocument = inventorApp.ActiveDocument;
		databaseService          = new DatabaseService();
		InitializeComponent();
		Loaded += FmiPartAssembly_Load;
	}

	public async Task InitializePanel()
	{
		await LoadDocumentProperties();
		LoadiPartMembers();
		PartInfo.TryLoadThumbnail(mInventorApp.ActiveDocument, ThumbnailImage);
		_ = Dispatcher.BeginInvoke(() => DocProps.Focus(), DispatcherPriority.Loaded);
	}

	private async void Window_Activated(object sender, EventArgs e)
	{
		// Avoid reloading while user is editing a cell, which would steal focus
		await InitializePanel();
	}

	private async void FmiPartAssembly_Load(object sender, RoutedEventArgs e)
	{
		try
		{
			await InitializePanel();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error loading panel: {ex.Message}", ErrorTitle, MessageBoxButton.OK,
				MessageBoxImage.Error);
		}
	}

	private void LoadiPartMembers()
	{
		try
		{
			var doc = originalAssemblyDocument ?? mInventorApp.ActiveDocument;
			if (doc == null) return;

			var members = new List<object>();

			// Add the assembly itself first
			AddAssemblyToMembers(doc, members);

			// Add assembly components if this is an assembly and we have processed components
			if (doc.DocumentType == kAssemblyDocumentObject && componentParts.Count > 0)
				AddComponentsToMembers(members);

			AssemblyItems.ItemsSource = members;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error loading assembly members: {ex.Message}");
		}
	}

	private async void CalculateProps_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			var doc = mInventorApp.ActiveDocument;
			if (doc == null) return;

			// Check if this is a sheetmetal part and set family to D-RMTO
			if (doc.DocumentType == kPartDocumentObject && doc is PartDocument partDoc)
			{
				// Check if it's a sheetmetal part using the SubType GUID
				if (partDoc.SubType.Contains("{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}") ||
				    partDoc.ComponentDefinition is SheetMetalComponentDefinition)
				{
					// Set Cost Center to D-RMTO for sheetmetal parts in memory
					if (calculatedPartInfo != null)
					{
						calculatedPartInfo.CostCenter = "D-RMTO";
						// Update the databaseProperties to reflect the change
						databaseProperties[FamilyProperty] = "D-RMTO";
					}

					MessageBox.Show("Sheetmetal family set to D-RMTO", "Info",
						MessageBoxButton.OK, MessageBoxImage.Information);
				}
				else
				{
					MessageBox.Show("Not a sheetmetal part - family not changed", "Info",
						MessageBoxButton.OK, MessageBoxImage.Information);
				}
			}
			else
			{
				MessageBox.Show("Not a part document - family not changed", "Info",
					MessageBoxButton.OK, MessageBoxImage.Information);
			}

			// Refresh data grids to show any calculated values
			await PartInfo.RefreshDataGridsWithCalculatedValues(mInventorApp, GeniusProps, hasCalculatedValues,
				calculatedPartInfo, calculatedRawStock, databaseProperties, doc);

			// Reload document properties to show the updated family
			await LoadDocumentProperties();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error calculating properties: {ex.Message}", ErrorTitle, MessageBoxButton.OK,
				MessageBoxImage.Error);
		}
	}

	private void SaveButton_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			// Save user-edited properties first
			PartInfo.SaveUserEditedProperties(mInventorApp, DocProps, databaseProperties, hasCalculatedValues,
				calculatedPartInfo,
				calculatedRawStock);
			// Apply calculated values to document if they exist
			if (hasCalculatedValues && calculatedPartInfo != null && mInventorApp.ActiveDocument != null)
			{
				var doc       = mInventorApp.ActiveDocument;
				var thickness = double.Parse(calculatedPartInfo.Thickness?.Replace(" in", "") ?? "0");
				PartInfo.UpdateDocumentPropertiesWithCalculatedDimensions(doc, calculatedPartInfo, thickness,
					calculatedRawStock);
			}

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

	private void OpenComponent_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			var selectedItem = GetSelectedItemFromContext(sender);
			if (selectedItem == null)
			{
				MessageBox.Show("No component selected to open.", "Info", MessageBoxButton.OK,
					MessageBoxImage.Information);
				return;
			}

			if (!ValidateSelectedItem(selectedItem))
				return;

			var componentDoc = GetComponentDocument(selectedItem);
			if (componentDoc == null)
				return;

			OpenComponentDocument(componentDoc);
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error handling open component request: {ex.Message}", Caption,
				MessageBoxButton.OK, MessageBoxImage.Error);
			Debug.WriteLine($"[OPEN_COMPONENT] General error: {ex.Message}");
		}
	}

	private void OpenAll_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			if (AssemblyItems.ItemsSource == null)
			{
				MessageBox.Show("No assembly items available to open.", "Info", MessageBoxButton.OK,
					MessageBoxImage.Information);
				return;
			}

			var items = AssemblyItems.ItemsSource.Cast<object>().ToList();
			if (items.Count == 0)
			{
				MessageBox.Show("No assembly items available to open.", "Info", MessageBoxButton.OK,
					MessageBoxImage.Information);
				return;
			}

			foreach (var item in items)
				try
				{
					if (!ValidateSelectedItem(item, false))
						continue;

					var componentDoc = GetComponentDocument(item);
					if (componentDoc != null) OpenComponentDocument(componentDoc);
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"[OPEN_ALL] Error opening item: {ex.Message}");
				}
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error opening all components: {ex.Message}", Caption,
				MessageBoxButton.OK, MessageBoxImage.Error);
			Debug.WriteLine($"[OPEN_ALL] General error: {ex.Message}");
		}
	}

	private object? GetSelectedItemFromContext(object sender)
	{
		// Try to get the selected item from the AssemblyItems DataGrid first
		var selectedItem = AssemblyItems.SelectedItem;

		// If nothing is selected, try to get the right-clicked item
		if (selectedItem == null)
			try
			{
				// Get the context menu that triggered this event
				if (sender is MenuItem { Parent: ContextMenu { PlacementTarget: DataGrid dataGrid } })
				{
					// Get the current mouse position relative to the DataGrid
					var position = Mouse.GetPosition(dataGrid);

					// Find the element under the mouse
					var inputElement = dataGrid.InputHitTest(position);
					if (inputElement is DependencyObject dependencyObj)
					{
						// Try to get the DataGridRow
						var row = FindVisualParent<DataGridRow>(dependencyObj);
						if (row is { Item: { } clickedItem }) selectedItem = clickedItem;
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"[CONTEXT_MENU] Error getting right-clicked item: {ex.Message}");
			}

		Debug.WriteLine($"[CONTEXT_MENU] Selected item: {selectedItem?.GetType().Name ?? "null"}");
		return selectedItem;
	}

	private static bool ValidateSelectedItem(object selectedItem, bool showMessage = true)
	{
		var itemType = selectedItem.GetType().GetProperty("Type")?.GetValue(selectedItem)?.ToString();

		// Don't try to open separators or summary items
		if (itemType is Separator or Summary)
		{
			if (showMessage)
				MessageBox.Show("Cannot open this item type.", "Info", MessageBoxButton.OK,
					MessageBoxImage.Information);
			return false;
		}

		// Check if it's the assembly itself
		if (selectedItem.GetType().GetProperty("Document")?.GetValue(selectedItem) is Document assemblyDoc)
		{
			// The assembly is already open, just activate it
			assemblyDoc.Activate();
			Debug.WriteLine($"[OPEN_COMPONENT] Activated assembly: {assemblyDoc.DisplayName}");
			return false;
		}

		// Check if it's a component part
		if (selectedItem.GetType().GetProperty("PartInfo")?.GetValue(selectedItem) is PartInfo partInfo &&
		    !string.IsNullOrEmpty(partInfo.DocumentName)) return true;

		if (showMessage)
			MessageBox.Show("Invalid component selection.", Caption, MessageBoxButton.OK, MessageBoxImage.Error);
		return false;
	}

	private Document? GetComponentDocument(object selectedItem)
	{
		if (selectedItem.GetType().GetProperty("PartInfo")?.GetValue(selectedItem) is not PartInfo partInfo ||
		    string.IsNullOrEmpty(partInfo.DocumentName))
			return null;

		// Find and open the component document
		var componentDoc = FindComponentDocumentByName(partInfo.DocumentName);
		if (componentDoc == null)
		{
			MessageBox.Show($"Could not find component document: {partInfo.DocumentName}", Caption,
				MessageBoxButton.OK, MessageBoxImage.Error);
			return null;
		}

		// Check if this is a Reference BOM component
		if (!ShouldIgnoreReferenceComponent(componentDoc)) return componentDoc;
		MessageBox.Show($"Cannot open Reference component: {partInfo.DocumentName}", "Info",
			MessageBoxButton.OK, MessageBoxImage.Information);
		return null;
	}

	private void OpenComponentDocument(Document componentDoc)
	{
		try
		{
			// Open the document using the application's Documents collection
			mInventorApp.Documents.Open(componentDoc.FullFileName);
			Debug.WriteLine($"[OPEN_COMPONENT] Opened document: {componentDoc.DisplayName}");
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error opening component: {ex.Message}", Caption,
				MessageBoxButton.OK, MessageBoxImage.Error);
			Debug.WriteLine($"[OPEN_COMPONENT] Error opening component: {ex.Message}");
		}
	}

	private static T? FindVisualParent<T>(DependencyObject child) where T : DependencyObject
	{
		while (true)
		{
			var parentObject = VisualTreeHelper.GetParent(child);

			switch (parentObject)
			{
				case null:
					return null;
				case T parent:
					return parent;
				default:
					child = parentObject;
					break;
			}
		}
	}

	private void DocProps_MouseDoubleClick(object sender, MouseButtonEventArgs e)
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
			if (GeniusProps.ItemsSource is not List<PartInfo.DocumentProperty> geniusItems) return;
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
			if (DocProps.ItemsSource is not List<PartInfo.DocumentProperty> currentItems) return;
			// Create a new list to force rebinding
			var newItems = new List<PartInfo.DocumentProperty>(currentItems);
			DocProps.ItemsSource = null;
			DocProps.ItemsSource = newItems;
		}, DispatcherPriority.DataBind);
	}

	private async void AssemblyItems_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		try
		{
			if (AssemblyItems.SelectedItem == null) return;

			var selectedItem = AssemblyItems.SelectedItem;
			var itemType     = selectedItem.GetType().GetProperty("Type")?.GetValue(selectedItem)?.ToString();

			if (itemType is Separator or Summary) return;

			// Check if it's the assembly itself
			if (selectedItem.GetType().GetProperty("Document")?.GetValue(selectedItem) is Document document)
			{
				// Load assembly properties
				await LoadDocumentPropertiesForDocument(document, overridePartInfo: calculatedPartInfo);
				PartInfo.TryLoadThumbnail(document, ThumbnailImage);
				return;
			}

			// Check if it's a component part
			if (selectedItem.GetType().GetProperty("PartInfo")?.GetValue(selectedItem) is not PartInfo partInfo)
				return;

			// Find the component document by part number or document name from the original assembly
			Document? componentDoc;
			if (!string.IsNullOrEmpty(partInfo.PartNumber))
				componentDoc = FindComponentDocumentByPartNumber(partInfo.PartNumber);
			else if (!string.IsNullOrEmpty(partInfo.DocumentName))
				componentDoc = FindComponentDocumentByName(partInfo.DocumentName);
			else
				return; // Both part number and document name are null

			if (componentDoc == null) return;

			// Check if this is a Reference BOM component and ignore it
			if (ShouldIgnoreReferenceComponent(componentDoc))
			{
				Debug.WriteLine(
					$"[REFERENCE_DEBUG] Ignoring selection of Reference component: {partInfo.DocumentName}");
				return;
			}

			try
			{
				// Load component properties using the assembly reference (not native properties)
				// This ensures we get the assembly's properties for the component, not the part's native properties
				await LoadDocumentPropertiesForDocument(componentDoc, overridePartInfo: partInfo);
				PartInfo.TryLoadThumbnail(componentDoc, ThumbnailImage);
			}
			finally
			{
				// Close the document if we opened it directly for native properties
				// Note: Inventor Document doesn't have IsDirty property, check if it's not the active document
				if (componentDoc != mInventorApp.ActiveDocument)
					try
					{
						componentDoc.Close();
						Debug.WriteLine($"[COMPONENT_DEBUG] Closed native part document: {componentDoc.DisplayName}");
					}
					catch (Exception ex)
					{
						Debug.WriteLine($"[COMPONENT_DEBUG] Error closing native part document: {ex.Message}");
					}
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error handling assembly item selection: {ex.Message}");
		}
	}

	private async Task LoadDocumentProperties()
	{
		// Load properties for the original assembly document (or active if original is null)
		var targetDoc = originalAssemblyDocument ?? mInventorApp.ActiveDocument;
		if (targetDoc != null)
			await LoadDocumentPropertiesForDocument(targetDoc, overridePartInfo: calculatedPartInfo);
	}

	private async Task LoadDocumentPropertiesForDocument(Document doc, bool forceNativeProperties = false,
		PartInfo? overridePartInfo = null)
	{
		try
		{
			var geniusProperties   = new List<PartInfo.DocumentProperty>();
			var documentProperties = new List<PartInfo.DocumentProperty>();

			// Clear and refresh database properties
			databaseProperties.Clear();

			// Load genius properties from server
			var isPurchasedPart = CheckIfPurchasedPart(doc);
			var isAssembly      = doc.DocumentType == kAssemblyDocumentObject;
			await LoadServerPropertiesForDocument(geniusProperties, doc, isPurchasedPart, isAssembly, overridePartInfo);
			var partFoundInGenius = databaseProperties.Count > 0;

			// Add basic document properties
			var partNumber = PartInfo.GetDesignTrackingProperty(doc, PartNumberProperty, forceNativeProperties);
			PartInfo.AddDocumentProperty(documentProperties, PartNumberProperty, partNumber, databaseProperties,
				partFoundInGenius);
			var description = PartInfo.GetDesignTrackingProperty(doc, DescriptionProperty, forceNativeProperties);
			PartInfo.AddDocumentProperty(documentProperties, DescriptionProperty, description, databaseProperties,
				partFoundInGenius);

			var family = PartInfo.GetDesignTrackingProperty(doc, "Cost Center", forceNativeProperties);
			PartInfo.AddDocumentProperty(documentProperties, FamilyProperty, family, databaseProperties,
				partFoundInGenius);

			// Load additional properties based on document type
			LoadAdditionalPropertiesForDocument(doc, documentProperties, partFoundInGenius, isPurchasedPart,
				forceNativeProperties, overridePartInfo);

			await Dispatcher.InvokeAsync(() =>
			{
				DocProps.ItemsSource    = documentProperties;
				GeniusProps.ItemsSource = geniusProperties;
			});
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error loading document properties: {ex.Message}", ErrorTitle, MessageBoxButton.OK,
				MessageBoxImage.Error);
		}
	}

	private async Task LoadServerPropertiesForDocument(List<PartInfo.DocumentProperty> geniusProperties, Document doc,
		bool isPurchasedPart, bool isAssembly, PartInfo? overridePartInfo = null)
	{
		var calculatedData = new PartInfo.CalculatedData
		{
			HasCalculatedValues =
				overridePartInfo != null || (hasCalculatedValues &&
				                             doc == (originalAssemblyDocument ?? mInventorApp.ActiveDocument)),
			CalculatedPartInfo = overridePartInfo ?? (doc == (originalAssemblyDocument ?? mInventorApp.ActiveDocument)
				? calculatedPartInfo
				: null),
			CalculatedRawStock = overridePartInfo?.RawStockInfo ??
			                     (doc == (originalAssemblyDocument ?? mInventorApp.ActiveDocument)
				                     ? calculatedRawStock
				                     : null)
		};
		await PartInfo.LoadServerProperties(geniusProperties, doc, databaseProperties, isPurchasedPart, isAssembly,
			databaseService, calculatedData);
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
			// If we can't determine BOM structure, assume it's not purchased
			return false;
		}
	}

	private static void AddAssemblyToMembers(Document doc, List<object> members)
	{
		if (doc.DocumentType != kAssemblyDocumentObject) return;
		var partNumber = PartInfo.GetDesignTrackingProperty(doc, PartNumberProperty);
		members.Add(new { Name = partNumber, Document = doc, PartNumber = partNumber });
	}

	private void AddComponentsToMembers(List<object> members)
	{
		// Add a separator
		members.Add(new { Name = "---", Type = Separator, Status = "", PartNumber = "---" });

		// Track unique display names to avoid duplicates
		var uniqueDisplayNames = new HashSet<string>();

		// Add all processed components sorted alphabetically by display name (part number)
		foreach (var component in componentParts.OrderBy(c => c.PartNumber ?? c.DocumentName))
		{
			if (!ShouldProcessComponent(component))
				continue;

			var componentInfo = GetComponentDisplayInfo(component);

			// Skip if we've already added this display name or if display name is null
			if (string.IsNullOrEmpty(componentInfo.Name) || !uniqueDisplayNames.Add(componentInfo.Name))
			{
				Debug.WriteLine($"[DUPLICATE_DEBUG] Skipping duplicate part: {componentInfo.Name}");
				continue;
			}

			members.Add(new
			{
				componentInfo.Name,
				Type = component.DocumentType,
				componentInfo.Status,
				PartInfo   = component,
				PartNumber = componentInfo.Name
			});
		}
	}

	private bool ShouldProcessComponent(PartInfo component)
	{
		// Get the display name (part number) that will be shown in the grid
		var displayName = component.PartNumber ?? component.DocumentName;

		// Skip components with null or empty display names
		if (string.IsNullOrEmpty(displayName))
			return false;

		// Find the component document to check for Reference BOM structure
		// Use part number for finding the document, fall back to document name if part number is null
		Document? componentDoc;
		if (!string.IsNullOrEmpty(component.PartNumber))
			componentDoc = FindComponentDocumentByPartNumber(component.PartNumber);
		else if (!string.IsNullOrEmpty(component.DocumentName))
			componentDoc = FindComponentDocumentByName(component.DocumentName);
		else
			return false; // Both part number and document name are null

		if (componentDoc == null || !ShouldIgnoreReferenceComponent(componentDoc)) return true;
		Debug.WriteLine($"[REFERENCE_DEBUG] Skipping Reference component: {displayName}");
		return false; // Skip Reference BOM components
	}

	private static (string Name, string Status) GetComponentDisplayInfo(PartInfo component)
	{
		var status         = "Processed";
		var additionalInfo = "";

		switch (component.DocumentType)
		{
			case "Sheet Metal Part" when component.RawStockInfo != null:
				additionalInfo = $" (RS: {component.RawStockInfo.Stock})";
				status         = "Raw Stock Found";
				break;
			case "Sheet Metal Part":
				status = "No Raw Stock";
				break;
		}

		return (component.PartNumber ?? component.DocumentName + additionalInfo, status);
	}

	private static bool ShouldIgnoreReferenceComponent(Document doc)
	{
		try
		{
			return doc switch
			{
				// Cast to appropriate document type to access BOMStructure
				PartDocument partDoc => partDoc.ComponentDefinition.BOMStructure ==
				                        BOMStructureEnum.kReferenceBOMStructure,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.BOMStructure ==
				                            BOMStructureEnum.kReferenceBOMStructure,
				_ => false
			};
		}
		catch
		{
			// If we can't determine BOM structure, assume it's not a reference
			return false;
		}
	}

	private static bool ShouldIgnoreReferenceComponentByOccurrence(ComponentOccurrence occurrence)
	{
		try
		{
			// Check the BOM structure from the component occurrence in the assembly context
			if (occurrence.BOMStructure == BOMStructureEnum.kReferenceBOMStructure)
			{
				Debug.WriteLine($"[REFERENCE_DEBUG] Ignoring Reference BOM component: {occurrence.Name}");
				return true;
			}

			switch (occurrence.ReferencedDocumentDescriptor?.ReferencedDocument)
			{
				// Check for Content Center components
				case PartDocument partDoc when
					partDoc.ComponentDefinition.IsContentMember:
					Debug.WriteLine($"[REFERENCE_DEBUG] Ignoring Content Center component: {occurrence.Name}");
					return true;
				// Check for Factory Member components
				case PartDocument factoryDoc when
					factoryDoc.ComponentDefinition.IsiPartFactory:
					Debug.WriteLine($"[REFERENCE_DEBUG] Ignoring Factory Member component: {occurrence.Name}");
					return true;
				default:
					// If we can't determine properties, assume it's not to be ignored
					return false;
			}
		}
		catch
		{
			// If we can't determine properties, assume it's not to be ignored
			return false;
		}
	}

	private void LoadAdditionalPropertiesForDocument(Document doc, List<PartInfo.DocumentProperty> documentProperties,
		bool partFoundInGenius, bool isPurchasedPart, bool forceNativeProperties = false,
		PartInfo? overridePartInfo = null)
	{
		var propertyOrder = GetPropertyOrder();
		var isAssembly    = doc.DocumentType == kAssemblyDocumentObject;

		foreach (var propName in
		         propertyOrder.Skip(3)) // Skip first 3 already handled (Part Number, Description, Family)
		{
			if (ShouldSkipPropertyForPurchasedPart(propName, isPurchasedPart))
				continue;

			if (ShouldSkipPropertyForAssembly(propName, isAssembly))
				continue;

			var value = GetPropertyValueWithFamilyOverride(doc, propName, forceNativeProperties, overridePartInfo);
			PartInfo.AddDocumentProperty(documentProperties, propName, value, databaseProperties, partFoundInGenius,
				propName is "RM" or RMUNIT);
		}
	}

	private static string[] GetPropertyOrder()
	{
		return
		[
			PartNumberProperty, DescriptionProperty, FamilyProperty, GeniusMassProperty, "Thickness",
			ExtentLengthProperty, ExtentWidthProperty, ExtentAreaProperty, "RMQTY", "RM", "RMUNIT"
		];
	}

	private static bool ShouldSkipPropertyForPurchasedPart(string propName, bool isPurchasedPart)
	{
		if (!isPurchasedPart) return false;

		// For purchased parts, only show: Part number, Description, Family, and GeniusMass
		var allowedProperties = new[] { PartNumberProperty, DescriptionProperty, FamilyProperty, GeniusMassProperty };
		return !allowedProperties.Contains(propName);
	}

	private static bool ShouldSkipPropertyForAssembly(string propName, bool isAssembly)
	{
		if (!isAssembly) return false;

		// For assemblies, only show: Part number, Description, Family, and GeniusMass
		var allowedProperties = new[] { PartNumberProperty, DescriptionProperty, FamilyProperty, GeniusMassProperty };
		return !allowedProperties.Contains(propName);
	}

	private string GetPropertyValueWithFamilyOverride(Document doc, string propName, bool forceNativeProperties = false,
		PartInfo? overridePartInfo = null)
	{
		var targetPartInfo = GetTargetPartInfo(doc, overridePartInfo);
		var targetRawStock = GetTargetRawStock(doc, overridePartInfo);
		var hasValues      = HasCalculatedValues(doc, overridePartInfo);

		if (hasValues && targetPartInfo != null)
			return GetCalculatedPropertyValue(doc, propName, forceNativeProperties, targetPartInfo, targetRawStock);

		return ShouldTryGeometryCalculation(doc, propName)
			? GetGeometryCalculatedPropertyValue(doc, propName, forceNativeProperties)
			: PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties);
	}

	private PartInfo? GetTargetPartInfo(Document doc, PartInfo? overridePartInfo)
	{
		var activeOrOriginal = originalAssemblyDocument ?? mInventorApp.ActiveDocument;
		return overridePartInfo ?? (doc == activeOrOriginal ? calculatedPartInfo : null);
	}

	private RawStockInfo? GetTargetRawStock(Document doc, PartInfo? overridePartInfo)
	{
		var activeOrOriginal = originalAssemblyDocument ?? mInventorApp.ActiveDocument;
		return overridePartInfo?.RawStockInfo ?? (doc == activeOrOriginal ? calculatedRawStock : null);
	}

	private bool HasCalculatedValues(Document doc, PartInfo? overridePartInfo)
	{
		var activeOrOriginal = originalAssemblyDocument ?? mInventorApp.ActiveDocument;
		return overridePartInfo != null || (hasCalculatedValues && doc == activeOrOriginal);
	}

	private static string GetCalculatedPropertyValue(Document doc, string propName, bool forceNativeProperties,
		PartInfo targetPartInfo, RawStockInfo? targetRawStock)
	{
		return propName switch
		{
			GeniusMassProperty => $"{targetPartInfo.CalculatedMass:F4}",
			"Thickness" => targetPartInfo.Thickness ??
			               PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties),
			ExtentLengthProperty => $"{targetPartInfo.CalculatedLength:F4} in",
			ExtentWidthProperty  => $"{targetPartInfo.CalculatedWidth:F4} in",
			ExtentAreaProperty   => $"{targetPartInfo.CalculatedArea:F4} in^2",
			"RMQTY"              => $"{targetPartInfo.GetCalculatedAreaInSquareFeet():F8}",
			"RM" => targetRawStock?.Stock ??
			        PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties),
			"RMUNIT" => targetRawStock?.RMUNIT ??
			            PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties),
			_ => PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties)
		};
	}

	private static bool ShouldTryGeometryCalculation(Document doc, string propName)
	{
		return doc.DocumentType == kPartDocumentObject &&
		       doc is PartDocument &&
		       propName is ExtentLengthProperty or ExtentWidthProperty or ExtentAreaProperty or GeniusMassProperty;
	}

	private static string GetGeometryCalculatedPropertyValue(Document doc, string propName, bool forceNativeProperties)
	{
		if (doc is not PartDocument partDoc)
			return PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties);

		var tempPartInfo = new PartInfo();
		var success      = CalculatePartDimensions(partDoc, tempPartInfo);

		if (success)
			return propName switch
			{
				ExtentLengthProperty => $"{tempPartInfo.CalculatedLength:F4} in",
				ExtentWidthProperty  => $"{tempPartInfo.CalculatedWidth:F4} in",
				ExtentAreaProperty   => $"{tempPartInfo.CalculatedArea:F4} in^2",
				GeniusMassProperty   => $"{tempPartInfo.CalculatedMass:F4}",
				_                    => PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties)
			};

		return PartInfo.GetUserDefinedProperty(doc, propName, forceNativeProperties);
	}

	private static bool CalculatePartDimensions(PartDocument partDoc, PartInfo tempPartInfo)
	{
		return partDoc.ComponentDefinition is SheetMetalComponentDefinition
			? tempPartInfo.CalculateSheetMetalDimensions(partDoc)
			: tempPartInfo.CalculatePartDimensions(partDoc);
	}

	private Document? FindComponentDocumentByPartNumber(string partNumber)
	{
		if (string.IsNullOrEmpty(partNumber))
			return null;

		try
		{
			// Use the original assembly document instead of the active document
			var assemblyDoc = originalAssemblyDocument as AssemblyDocument ??
			                  mInventorApp.ActiveDocument as AssemblyDocument;
			if (assemblyDoc == null)
				return null;

			var componentOccurrence = FindComponentOccurrenceByPartNumber(assemblyDoc, partNumber);

			if (componentOccurrence?.ReferencedDocumentDescriptor?.ReferencedDocument is not Document partDoc)
				return null;

			if (!ShouldIgnoreReferenceComponentByOccurrence(componentOccurrence))
				return TryOpenNativePartDocument(partDoc) ?? partDoc;
			Debug.WriteLine(
				$"[REFERENCE_DEBUG] Component with part number {partNumber} is Reference BOM in assembly context");
			return null;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error finding component document by part number: {ex.Message}");
			return null;
		}
	}

	private static ComponentOccurrence? FindComponentOccurrenceByPartNumber(AssemblyDocument assemblyDoc,
		string partNumber)
	{
		// Search top-level occurrences first
		var topLevelOccurrence = assemblyDoc.ComponentDefinition.Occurrences
		                                    .Cast<ComponentOccurrence>()
		                                    .FirstOrDefault(occ =>
		                                    {
			                                    // Get the part number from the referenced document
			                                    if (occ.ReferencedDocumentDescriptor?.ReferencedDocument is not Document
			                                        doc) return false;
			                                    var docPartNumber =
				                                    PartInfo.GetDesignTrackingProperty(doc, PartNumberProperty);
			                                    return docPartNumber == partNumber;
		                                    });

		return topLevelOccurrence ??
		       // If not found at top level, recursively search sub-assemblies
		       FindComponentOccurrenceRecursiveByPartNumber(assemblyDoc.ComponentDefinition.Occurrences, partNumber, 0);
	}

	private static ComponentOccurrence? FindComponentOccurrenceRecursiveByPartNumber(ComponentOccurrences occurrences,
		string partNumber, int depth)
	{
		const int maxDepth = 10; // Prevent infinite recursion

		if (depth > maxDepth)
		{
			Debug.WriteLine(
				$"[FIND_COMPONENT] Maximum recursion depth ({maxDepth}) reached while searching for part number {partNumber}");
			return null;
		}

		foreach (ComponentOccurrence occurrence in occurrences)
		{
			// Check if this occurrence matches the part number
			if (occurrence.ReferencedDocumentDescriptor?.ReferencedDocument is Document doc)
			{
				var docPartNumber = PartInfo.GetDesignTrackingProperty(doc, PartNumberProperty);
				if (docPartNumber == partNumber)
					return occurrence;
			}

			// If this is a sub-assembly, recursively search within it
			if (occurrence.ReferencedDocumentDescriptor?.ReferencedDocument is not AssemblyDocument subAssemblyDoc)
				continue;
			var nestedOccurrence = FindComponentOccurrenceRecursiveByPartNumber(
				subAssemblyDoc.ComponentDefinition.Occurrences,
				partNumber, depth + 1);
			if (nestedOccurrence != null)
				return nestedOccurrence;
		}

		return null;
	}

	private Document? FindComponentDocumentByName(string documentName)
	{
		if (string.IsNullOrEmpty(documentName))
			return null;

		try
		{
			// Use the original assembly document instead of the active document
			// This ensures the panel keeps working with the original assembly even when switching documents
			var assemblyDoc = originalAssemblyDocument as AssemblyDocument ??
			                  mInventorApp.ActiveDocument as AssemblyDocument;
			if (assemblyDoc == null)
				return null;

			var componentOccurrence = FindComponentOccurrence(assemblyDoc, documentName);

			if (componentOccurrence?.ReferencedDocumentDescriptor?.ReferencedDocument is not Document partDoc)
				return null;

			if (!ShouldIgnoreReferenceComponentByOccurrence(componentOccurrence))
				return TryOpenNativePartDocument(partDoc) ?? partDoc;
			Debug.WriteLine($"[REFERENCE_DEBUG] Component {documentName} is Reference BOM in assembly context");
			return null;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error finding component document: {ex.Message}");
			return null;
		}
	}

	private static ComponentOccurrence? FindComponentOccurrence(AssemblyDocument assemblyDoc, string documentName)
	{
		// Search top-level occurrences first
		var topLevelOccurrence = assemblyDoc.ComponentDefinition.Occurrences
		                                    .Cast<ComponentOccurrence>()
		                                    .FirstOrDefault(occ =>
			                                    occ.ReferencedDocumentDescriptor?.DisplayName == documentName);

		return topLevelOccurrence ??
		       // If not found at top level, recursively search sub-assemblies
		       FindComponentOccurrenceRecursive(assemblyDoc.ComponentDefinition.Occurrences, documentName, 0);
	}

	private static ComponentOccurrence? FindComponentOccurrenceRecursive(ComponentOccurrences occurrences,
		string documentName, int depth)
	{
		const int maxDepth = 10; // Prevent infinite recursion

		if (depth > maxDepth)
		{
			Debug.WriteLine(
				$"[FIND_COMPONENT] Maximum recursion depth ({maxDepth}) reached while searching for {documentName}");
			return null;
		}

		foreach (ComponentOccurrence occurrence in occurrences)
		{
			// Check if this occurrence matches the document name
			if (occurrence.ReferencedDocumentDescriptor?.DisplayName == documentName)
				return occurrence;

			// If this is a sub-assembly, recursively search within it
			if (occurrence.ReferencedDocumentDescriptor?.ReferencedDocument is not AssemblyDocument subAssemblyDoc)
				continue;
			var nestedOccurrence = FindComponentOccurrenceRecursive(subAssemblyDoc.ComponentDefinition.Occurrences,
				documentName, depth + 1);
			if (nestedOccurrence != null)
				return nestedOccurrence;
		}

		return null;
	}

	private Document? TryOpenNativePartDocument(Document partDoc)
	{
		if (partDoc.DocumentType != kPartDocumentObject || string.IsNullOrEmpty(partDoc.FullFileName))
			return null;

		try
		{
			var nativePartDoc = mInventorApp.Documents.Open(partDoc.FullFileName, false);
			Debug.WriteLine(
				$"[COMPONENT_DEBUG] Opened native part document: {nativePartDoc.DisplayName}, Type: {nativePartDoc.DocumentType}");
			return nativePartDoc;
		}
		catch (Exception ex)
		{
			Debug.WriteLine(
				$"[COMPONENT_DEBUG] Could not open native part document, using assembly reference: {ex.Message}");
			return null;
		}
	}
}