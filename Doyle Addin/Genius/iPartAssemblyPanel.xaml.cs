#region

#nullable enable
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using MessageBox = System.Windows.MessageBox;

#endregion

namespace Doyle_Addin.Genius;

// ReSharper disable once InconsistentNaming
public partial class iPartAssemblyPanel
{
	// Constant for error message titles
	private const string ErrorTitle = "Error";

	// Stores calculated dimension values in memory (without dirtying document)
	private readonly PartInfo? calculatedPartInfo;

	private readonly RawStockInfo? calculatedRawStock;

	// Stores database properties for comparison during editing
	private readonly Dictionary<string, string> databaseProperties = new();

	// Database service for loading Genius properties
	private readonly DatabaseService databaseService;

	// Localized editor handling
	private readonly bool hasCalculatedValues;

	// Holds a reference to the Inventor Application object
	private readonly Application mInventorApp;

	public iPartAssemblyPanel(Application inventorApp, PartInfo? calculatedPartInfo, RawStockInfo? calculatedRawStock,
		bool hasCalculatedValues)
	{
		mInventorApp             = inventorApp;
		this.calculatedPartInfo  = calculatedPartInfo;
		this.calculatedRawStock  = calculatedRawStock;
		this.hasCalculatedValues = hasCalculatedValues;
		databaseService          = new DatabaseService();
		InitializeComponent();
		Loaded += FmiPartAssembly_Load;
	}

	public async Task InitializePanel()
	{
		await LoadDocumentProperties();
		LoadiPartMembers();
		PartInfo.TryLoadThumbnail(mInventorApp.ActiveDocument, ThumbnailImage);
		await Dispatcher.BeginInvoke(() => DocProps.Focus(), DispatcherPriority.Loaded);
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
			var doc = mInventorApp.ActiveDocument;
			if (doc == null) return;

			var members = new List<object>();

			switch (doc.DocumentType)
			{
				case kPartDocumentObject:
					if (doc is PartDocument partDoc && partDoc.ComponentDefinition.IsiPartFactory)
					{
						var factory = partDoc.ComponentDefinition.iPartFactory;
						for (var i = 1; i <= factory.TableRows.Count; i++)
						{
							var member = factory.TableRows[i];
							members.Add(new { Name = member.MemberName, Type = "iPart", Status = "Active" });
						}
					}

					break;

				case kAssemblyDocumentObject:
					if (doc is AssemblyDocument assyDoc && assyDoc.ComponentDefinition.IsiAssemblyFactory)
					{
						var factory = assyDoc.ComponentDefinition.iAssemblyFactory;
						for (var i = 1; i <= factory.TableRows.Count; i++)
						{
							var member = factory.TableRows[i];
							members.Add(new { Name = member.MemberName, Type = "iAssembly", Status = "Active" });
						}
					}

					break;
			}

			Factorymembers.ItemsSource = members;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error loading iPart members: {ex.Message}");
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
						databaseProperties["Family"] = "D-RMTO";
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
			// Save user-edited properties using the shared method
			PartInfo.SaveUserEditedProperties(mInventorApp, DocProps, databaseProperties, hasCalculatedValues,
				calculatedPartInfo,
				calculatedRawStock);

			DialogResult = true;
			mInventorApp.ActiveDocument?.Save2();
			Close();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error saving: {ex.Message}", ErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	private void CancelButton_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			DialogResult = false;
			Close();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error closing: {ex.Message}", ErrorTitle, MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	private void DocProps_MouseDoubleClick(object sender, MouseButtonEventArgs e)
	{
		try
		{
			if (sender is not DataGrid { SelectedItem: PartInfo.DocumentProperty selectedProperty } dataGrid) return;
			var editDialog = new EditValueDialog(selectedProperty.Property, selectedProperty.Value);
			var ownerWindow = GetWindow(this);
			if (ownerWindow != null) editDialog.Owner = ownerWindow;

			if (editDialog.ShowDialog() != true) return;
			selectedProperty.Value = editDialog.PropertyValue;

			// Update the database properties dictionary
			databaseProperties[selectedProperty.Property] = editDialog.PropertyValue;

			// Refresh the data grid to show the updated value
			dataGrid.Items.Refresh();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error editing property: {ex.Message}", ErrorTitle, MessageBoxButton.OK,
				MessageBoxImage.Error);
		}
	}

	private void GeniusProps_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		try
		{
			if (sender is not DataGrid { SelectedItem: PartInfo.DocumentProperty selectedProperty } dataGrid) return;
			var editDialog = new EditValueDialog(selectedProperty.Property, selectedProperty.Value);
			var ownerWindow = GetWindow(this);
			if (ownerWindow != null) editDialog.Owner = ownerWindow;

			if (editDialog.ShowDialog() != true) return;
			selectedProperty.Value = editDialog.PropertyValue;

			// Update the database properties dictionary
			databaseProperties[selectedProperty.Property] = editDialog.PropertyValue;

			// Refresh the data grid to show the updated value
			dataGrid.Items.Refresh();
		}
		catch (Exception ex)
		{
			MessageBox.Show($"Error editing property: {ex.Message}", ErrorTitle, MessageBoxButton.OK,
				MessageBoxImage.Error);
		}
	}

	private async void Factorymembers_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		try
		{
			if (Factorymembers.SelectedItem == null) return;

			var selectedItem = Factorymembers.SelectedItem;
			var memberName   = selectedItem.GetType().GetProperty("Name")?.GetValue(selectedItem)?.ToString();

			if (string.IsNullOrEmpty(memberName)) return;

			var activeDoc = mInventorApp.ActiveDocument;
			if (activeDoc == null) return;

			await SwitchToFactoryMember(activeDoc, memberName);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error handling factory member selection: {ex.Message}");
		}
	}

	private async Task SwitchToFactoryMember(Document activeDoc, string memberName)
	{
		switch (activeDoc.DocumentType)
		{
			case kPartDocumentObject
				when activeDoc is PartDocument partDoc && partDoc.ComponentDefinition.IsiPartFactory:
				await SwitchiPartMember(partDoc.ComponentDefinition.iPartFactory, memberName);
				break;

			case kAssemblyDocumentObject
				when activeDoc is AssemblyDocument assyDoc && assyDoc.ComponentDefinition.IsiAssemblyFactory:
				await SwitchiAssemblyMember(assyDoc.ComponentDefinition.iAssemblyFactory, memberName);
				break;
		}
	}

	private async Task SwitchiPartMember(iPartFactory factory, string memberName)
	{
		try
		{
			// Find the target row in the table
			iPartTableRow? targetRow = null;
			for (var i = 1; i <= factory.TableRows.Count; i++)
			{
				var row = factory.TableRows[i];
				if (row.MemberName != memberName) continue;
				targetRow = row;
				break;
			}

			if (targetRow != null)
			{
				// Set the target row as the default row - this switches the active member
				factory.DefaultRow = targetRow;

				// Access the Excel worksheet to trigger property refresh
				try
				{
					var workSheet = factory.ExcelWorkSheet;
					// Force a refresh of the Excel worksheet
					var workbook = workSheet.GetType().GetProperty("Parent")?.GetValue(workSheet);
					workbook?.GetType().GetMethod("Save")?.Invoke(workbook, null);
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"Excel worksheet access error: {ex.Message}");
					// Continue even if Excel access fails
				}

				// Reload properties for the switched member
				await LoadDocumentProperties();
				PartInfo.TryLoadThumbnail(mInventorApp.ActiveDocument, ThumbnailImage, memberName);
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error switching iPart member: {ex.Message}");
		}
	}

	private async Task SwitchiAssemblyMember(iAssemblyFactory factory, string memberName)
	{
		try
		{
			// Find the target row in the table
			iAssemblyTableRow? targetRow = null;
			for (var i = 1; i <= factory.TableRows.Count; i++)
			{
				var row = factory.TableRows[i];
				if (row.MemberName != memberName) continue;
				targetRow = row;
				break;
			}

			if (targetRow != null)
			{
				// Set the target row as the default row - this switches the active member
				factory.DefaultRow = targetRow;

				// Access the Excel worksheet to trigger property refresh
				try
				{
					var workSheet = factory.ExcelWorkSheet;
					// Force a refresh of the Excel worksheet
					var workbook = workSheet.GetType().GetProperty("Parent")?.GetValue(workSheet);
					workbook?.GetType().GetMethod("Save")?.Invoke(workbook, null);
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"Excel worksheet access error: {ex.Message}");
					// Continue even if Excel access fails
				}

				// Reload properties for the switched member
				await LoadDocumentProperties();
				PartInfo.TryLoadThumbnail(mInventorApp.ActiveDocument, ThumbnailImage, memberName);
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error switching iAssembly member: {ex.Message}");
		}
	}

	private async Task LoadDocumentProperties()
	{
		// Load properties for the active document
		var targetDoc = mInventorApp.ActiveDocument;
		if (targetDoc != null)
			await LoadDocumentPropertiesForDocument(targetDoc);
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

	private async Task LoadDocumentPropertiesForDocument(Document doc)
	{
		try
		{
			var geniusProperties   = new List<PartInfo.DocumentProperty>();
			var documentProperties = new List<PartInfo.DocumentProperty>();

			// Clear and refresh database properties
			databaseProperties.Clear();

			// Load genius properties from server
			var isPurchasedPart = IsPurchasedPart(doc);
			var isAssembly      = doc.DocumentType == kAssemblyDocumentObject;

			var calculatedData = new PartInfo.CalculatedData
			{
				HasCalculatedValues = hasCalculatedValues,
				CalculatedPartInfo  = calculatedPartInfo,
				CalculatedRawStock  = calculatedRawStock
			};
			await PartInfo.LoadServerProperties(geniusProperties, doc, databaseProperties, isPurchasedPart, isAssembly,
				databaseService, calculatedData);

			// Check if part was found in Genius (databaseProperties will be empty if not found)
			var partFoundInGenius = databaseProperties.Count > 0;

			// Add basic document properties
			var partNumber = PartInfo.GetDesignTrackingProperty(doc, "Part Number");
			PartInfo.AddDocumentProperty(documentProperties, "Part Number", partNumber, databaseProperties,
				partFoundInGenius);
			var description = PartInfo.GetDesignTrackingProperty(doc, "Description");
			PartInfo.AddDocumentProperty(documentProperties, "Description", description, databaseProperties,
				partFoundInGenius);

			var family = PartInfo.GetDesignTrackingProperty(doc, "Cost Center");
			PartInfo.AddDocumentProperty(documentProperties, "Family", family, databaseProperties, partFoundInGenius);

			// Set data sources for both grids on UI thread after all properties are loaded
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
}