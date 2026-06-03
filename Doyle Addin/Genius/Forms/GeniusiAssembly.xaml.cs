namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Inventor;
using ThemeManager = Options.Themes.ThemeManager;

/// <summary>
///     Represents a user interface control designed to facilitate interaction with
///     and management of Genius-related components within the application.
/// </summary>
public partial class GeniusiAssembly
{
	private readonly PropertyComparator _propertyComparator;
	private Document _selectedDocument;

	/// <summary>
	///     Represents a user interface control designed to facilitate interaction with
	///     and management of Genius-related components within the application.
	///     This control is responsible for initializing necessary services, applying themes,
	///     and loading members and assembly children associated with the Genius system.
	/// </summary>
	public GeniusiAssembly()
	{
		try
		{
			var sqlDataManager = new SqlDataManager(Geniusinfo.DefaultConnectionString);
			_propertyComparator = new PropertyComparator(sqlDataManager);

			InitializeComponent();
			ThemeManager.ApplyTheme(this);

			_ = LoadMembers();
			_ = LoadAssemblyChildren();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiAssembly: Exception in constructor: {ex.Message}");
		}
	}

	private async Task LoadMembers()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc) return;

			var compDef = assemblyDoc.ComponentDefinition;
			var members = new List<MemberInfo>();

			if (compDef.iAssemblyFactory != null)
				members.AddRange(compDef.iAssemblyFactory.TableRows.Cast<iAssemblyTableRow>()
				                        .Select(row => new MemberInfo
				                        {
					                        PartNumber  = row[ColumnNames.PartNumber].Value,
					                        Description = row[ColumnNames.Description].Value
				                        }));
			else if (compDef.ModelStates?.Count > 1)
				members.AddRange(compDef.ModelStates.ModelStateTable.TableRows.Cast<ModelStateTableRow>()
				                        .Select(row => new MemberInfo
				                        {
					                        PartNumber  = row[ColumnNames.PartNumber].Value,
					                        Description = row[ColumnNames.Description].Value
				                        }));

			await Dispatcher.InvokeAsync(() => { Members.ItemsSource = members; });
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiAssembly: Error loading members: {ex.Message}");
		}
	}

	private async Task LoadAssemblyChildren()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument) return;

			var childrenTable = await Task.Run(Geniusinfo.GetAllAssemblyChildren);
			var subassemblies = new List<ChildInfo>();
			var parts         = new List<ChildInfo>();
			var purchased     = new List<ChildInfo>();

			foreach (DataRow row in childrenTable.Rows)
			{
				var child = new ChildInfo
				{
					PartNumber   = row["PartNumber"].ToString() ?? "",
					Description  = row["Description"].ToString() ?? "",
					DocumentType = row["DocumentType"].ToString(),
					IsPurchased  = (bool)row["IsPurchased"]
				};

				if (child.IsPurchased)
				{
					purchased.Add(child);
					continue;
				}

				switch (child.DocumentType)
				{
					case "kAssemblyDocumentObject":
						subassemblies.Add(child);
						break;
					case "kPartDocumentObject":
						parts.Add(child);
						break;
				}
			}

			await Dispatcher.InvokeAsync(() =>
			{
				MembersChildren.ItemsSource   = subassemblies.Concat(parts).ToList();
				PurchasedDataGrid.ItemsSource = purchased;

				MembersTab.Visibility = (Members.ItemsSource as IEnumerable<MemberInfo>)?.Any() == true
					? Visibility.Visible
					: Visibility.Collapsed;
				AssemblyChildrenTab.Visibility =
					subassemblies.Count + parts.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
				PurchasedTab.Visibility = purchased.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiAssembly: Error loading children: {ex.Message}");
		}
	}

	private async void ChildrenDataGrid1_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (MembersChildren.SelectedItem is not ChildInfo selectedChild)
		{
			_selectedDocument = null;
			return;
		}

		_selectedDocument = PropertyExtractor.FindDocumentByPartNumber(selectedChild.PartNumber);
		var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(selectedChild.PartNumber,
			_propertyComparator.SqlDataManager, _selectedDocument);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
		if (_selectedDocument != null) await UpdateThumbnail(_selectedDocument);

		// Zoom to selection
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc ||
			    _selectedDocument == null) return;
			var occurrence = ThisApplication.Documents.ItemByName[_selectedDocument.FullFileName];
			if (occurrence == null) return;

			var asmOcc = assemblyDoc.ComponentDefinition.Occurrences.AllReferencedOccurrences[occurrence];
			if (asmOcc.Count <= 0) return;
			assemblyDoc.SelectSet.Clear();
			assemblyDoc.SelectSet.Select(asmOcc[1]);
			ThisApplication.CommandManager.ControlDefinitions["AppZoomSelectCmd"].Execute();
		}
		catch
		{
			/* ignore selection errors */
		}
	}

	private async void PurchasedDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (PurchasedDataGrid.SelectedItem is not ChildInfo selectedChild)
		{
			_selectedDocument = null;
			return;
		}

		_selectedDocument = PropertyExtractor.FindDocumentByPartNumber(selectedChild.PartNumber);
		var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(selectedChild.PartNumber,
			_propertyComparator.SqlDataManager, _selectedDocument);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
		if (_selectedDocument != null) await UpdateThumbnail(_selectedDocument);

		// Zoom to selection
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc ||
			    _selectedDocument == null) return;
			var occurrence = ThisApplication.Documents.ItemByName[_selectedDocument.FullFileName];
			if (occurrence == null) return;

			var asmOcc = assemblyDoc.ComponentDefinition.Occurrences.AllReferencedOccurrences[occurrence];
			if (asmOcc.Count <= 0) return;
			assemblyDoc.SelectSet.Clear();
			assemblyDoc.SelectSet.Select(asmOcc[1]);
			ThisApplication.CommandManager.ControlDefinitions["AppZoomSelectCmd"].Execute();
		}
		catch
		{
			/* ignore selection errors */
		}
	}

	private async void ChildrenDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (Members.SelectedItem is not MemberInfo selectedMember ||
		    ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc) return;

		try
		{
			var      compDef   = assemblyDoc.ComponentDefinition;
			Document memberDoc = null;

			if (compDef.iAssemblyFactory != null)
			{
				var row = compDef.iAssemblyFactory.TableRows.Cast<iAssemblyTableRow>().FirstOrDefault(r =>
					string.Equals(r[ColumnNames.PartNumber].Value, selectedMember.PartNumber,
						StringComparison.OrdinalIgnoreCase));
				if (row != null)
				{
					compDef.iAssemblyFactory.CreateMember(row);
					memberDoc = PropertyExtractor.FindDocumentByPartNumber(selectedMember.PartNumber);
				}
			}
			else if (compDef.ModelStates?.Count > 1)
			{
				var state = compDef.ModelStates.Cast<ModelState>()
				                   .FirstOrDefault(s => s.Name == selectedMember.PartNumber);
				if (state != null)
				{
					state.Activate();
					memberDoc = ThisApplication.ActiveDocument;
				}
			}

			_selectedDocument = memberDoc;
			var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(selectedMember.PartNumber,
				_propertyComparator.SqlDataManager, memberDoc);
			await Dispatcher.InvokeAsync(() =>
			{
				SqlDataGrid.ItemsSource      = geniusRows;
				InventorDataGrid.ItemsSource = invRows;
			});
			if (memberDoc != null) await UpdateThumbnail(memberDoc);
			await LoadAssemblyChildren();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiAssembly: Error in selection changed: {ex.Message}");
		}
	}


	private async void CalculatePropsButton_Click(object sender, RoutedEventArgs e)
	{
		CalculatePropsButton.IsEnabled = false;
		try
		{
			var doc = _selectedDocument ?? ThisApplication?.ActiveDocument;
			if (doc == null) return;

			var calculated = await CalculateProps.CalculateAllPropertiesAsync((_Document)doc);
			if (calculated.Count <= 0) return;
			var current = PropertyExtractor.GetPropertiesFromDocumentStatic(doc);
			var updates = new Dictionary<string, string>();

			foreach (var calc in calculated)
			{
				var key = calc.Key == "Mass" ? "GeniusMass" : calc.Key;
				if (!GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(key, ""), calc.Value))
					updates[key] = calc.Value;
			}

			if (updates.Count <= 0) return;
			GeniusFormsHelper.UpdateInventorProperties((_Document)doc, updates, "GeniusiAssembly");
			var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(
				current.GetValueOrDefault("Part Number", doc.DisplayName),
				_propertyComparator.SqlDataManager, doc);
			await Dispatcher.InvokeAsync(() =>
			{
				SqlDataGrid.ItemsSource      = geniusRows;
				InventorDataGrid.ItemsSource = invRows;
			});
		}
		finally
		{
			CalculatePropsButton.IsEnabled = true;
		}
	}


	private async Task UpdateThumbnail(Document document)
	{
		try
		{
			var bitmapImage = await Dispatcher.InvokeAsync(() =>
			{
				var pictureDisp = ThumbnailHelper.GetThumbnailRaw(document);
				if (pictureDisp == null) return null;
				var drawingImage = ThumbnailHelper.ConvertIPictureToImage(pictureDisp);
				return drawingImage != null ? ThumbnailHelper.ConvertToBitmapImage(drawingImage) : null;
			});
			image.Source = bitmapImage;
		}
		catch
		{
			image.Source = null;
		}
	}

	private static class ColumnNames
	{
		public const string PartNumber = "Part Number [Project]";
		public const string Description = "Description [Project]";
	}
}