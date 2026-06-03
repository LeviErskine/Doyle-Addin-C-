namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Inventor;
using Path = Path;
using ThemeManager = Options.Themes.ThemeManager;

/// <summary>
///     Represents a user interface control designed to facilitate interaction with
///     and management of Genius-related components within the application.
/// </summary>
public partial class GeniusAssembly
{
	private readonly PropertyComparator _propertyComparator;
	private Document _currentTargetDocument;

	/// <summary>
	///     Represents a WPF user interface control designed to interact with
	///     and manage Genius-related components within the application.
	/// </summary>
	public GeniusAssembly()
	{
		try
		{
			var sqlDataManager = new SqlDataManager(Geniusinfo.DefaultConnectionString);
			_propertyComparator = new PropertyComparator(sqlDataManager);

			InitializeComponent();
			ThemeManager.ApplyTheme(this);

			_ = RefreshData();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Exception in constructor: {ex.Message}");
		}
	}

	private async Task RefreshData(Document targetDocument = null, bool forceRefreshChildren = false)
	{
		try
		{
			if (ThisApplication == null) return;

			_currentTargetDocument = targetDocument ?? ThisApplication.ActiveDocument;
			var comparisonTable = await _propertyComparator.ComparePropertiesAsync(_currentTargetDocument);

			await Dispatcher.InvokeAsync(() =>
			{
				PopulateDataGrids(comparisonTable);

				if (targetDocument != null && !forceRefreshChildren)
					return; // Only refresh children list when refreshing active doc or forced
				var childrenTable = Geniusinfo.GetAllAssemblyChildren();

				var subassemblies = new List<ChildInfo>();
				var parts         = new List<ChildInfo>();
				var purchased     = new List<ChildInfo>();

				foreach (DataRow row in childrenTable.Rows)
				{
					var child = new ChildInfo
					{
						Level         = (int)row["Level"],
						PartNumber    = row["PartNumber"].ToString(),
						Description   = row["Description"].ToString(),
						DocumentType  = row["DocumentType"].ToString(),
						HasDifference = (bool)row["HasDifference"],
						FullPath      = row["FullPath"].ToString(),
						IsPurchased   = (bool)row["IsPurchased"]
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

				Assemblies.ItemsSource        = subassemblies;
				PartsDataGrid.ItemsSource     = parts;
				PurchasedDataGrid.ItemsSource = purchased;

				Assemblies.Items.Refresh();
				PartsDataGrid.Items.Refresh();
				PurchasedDataGrid.Items.Refresh();

				// Manage tab visibility
				var hasSubassemblies = subassemblies.Count > 0;
				var hasParts         = parts.Count > 0;
				var hasPurchased     = purchased.Count > 0;

				AssembliesTab.Visibility = hasSubassemblies ? Visibility.Visible : Visibility.Collapsed;
				PartsTab.Visibility      = hasParts ? Visibility.Visible : Visibility.Collapsed;
				PurchasedTab.Visibility  = hasPurchased ? Visibility.Visible : Visibility.Collapsed;

				if (tabControl.SelectedItem is not TabItem { Visibility: Visibility.Visible })
					tabControl.SelectedItem = hasSubassemblies ? AssembliesTab :
						hasParts                               ? PartsTab :
						hasPurchased                           ? PurchasedTab : tabControl.SelectedItem;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Refresh error: {ex.Message}");
		}
	}

	private void PopulateDataGrids(DataTable comparisonTable)
	{
		var sqlRows = comparisonTable.Rows.Cast<DataRow>().Select(row =>
		{
			var invVal   = row["Inventor Value"].ToString();
			var sqlVal   = row["SQL Value"].ToString();
			var areEqual = GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
			return new PropertyRow
			{
				Property           = row["Property"].ToString(),
				["SQL Value"]      = sqlVal,
				["Inventor Value"] = invVal,
				["Status"]         = areEqual ? "Match" : "Mismatch",
				HasDifference      = !areEqual
			};
		}).ToList();

		var invRows = comparisonTable.Rows.Cast<DataRow>().Select(row =>
		{
			var invVal   = row["Inventor Value"].ToString();
			var sqlVal   = row["SQL Value"].ToString();
			var areEqual = GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
			return new PropertyRow
			{
				Property           = row["Property"].ToString(),
				["Inventor Value"] = invVal,
				["SQL Value"]      = sqlVal,
				["Status"]         = areEqual ? "Match" : "Mismatch",
				HasDifference      = !areEqual
			};
		}).ToList();

		SqlDataGrid.ItemsSource      = sqlRows;
		InventorDataGrid.ItemsSource = invRows;
		SqlDataGrid.Items.Refresh();
		InventorDataGrid.Items.Refresh();
	}

	private async void CalculatePropsButton_Click(object _sender, RoutedEventArgs _e)
	{
		CalculatePropsButton.IsEnabled = false;
		try
		{
			var target = _currentTargetDocument ?? ThisApplication.ActiveDocument;
			if (target == null) return;

			var calculated = await CalculateProps.CalculateAllPropertiesAsync(target);
			if (calculated.Count <= 0) return;
			var current = PropertyExtractor.GetPropertiesFromDocumentStatic(target);
			var updates = new Dictionary<string, string>();

			foreach (var calc in calculated.Where(calc =>
				         !GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(calc.Key, ""), calc.Value)))
				updates[calc.Key] = calc.Value;

			if (updates.Count <= 0) return;
			GeniusFormsHelper.UpdateInventorProperties(target, updates, "GeniusAssembly");
			await RefreshData(target, true);
		}
		finally
		{
			CalculatePropsButton.IsEnabled = true;
		}
	}

	private void Open_All_Parts_Click(object _sender, RoutedEventArgs _e)
	{
		try
		{
			if (ThisApplication == null) return;

			var childrenTable = Geniusinfo.GetAllAssemblyChildren();
			var openDocsLookup = ThisApplication.Documents.Cast<Document>()
			                                    .ToDictionary(doc => doc.FullFileName, doc => doc,
				                                    StringComparer.OrdinalIgnoreCase);

			foreach (DataRow row in childrenTable.Rows)
			{
				var path = row["FullPath"].ToString();
				if (string.IsNullOrEmpty(path) || path.StartsWith("virtual:", StringComparison.OrdinalIgnoreCase))
					continue;

				if (openDocsLookup.TryGetValue(path, out var foundDocument))
					foundDocument.Activate();
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: OpenAllParts error: {ex.Message}");
		}
	}

	private async void SubassembliesDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (Assemblies.SelectedItem is not ChildInfo selectedChild) return;
		_ = UpdateSelectedChildData(selectedChild);

		_currentTargetDocument = PropertyExtractor.FindDocumentByPartNumber(selectedChild.PartNumber);
		var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(selectedChild.PartNumber,
			_propertyComparator.SqlDataManager, _currentTargetDocument);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
		if (_currentTargetDocument != null) await UpdateThumbnail(_currentTargetDocument);

		// Zoom to selection
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc ||
			    _currentTargetDocument == null) return;
			var occurrence = ThisApplication.Documents.ItemByName[_currentTargetDocument.FullFileName];
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

	private async void PartsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (PartsDataGrid.SelectedItem is not ChildInfo selectedChild) return;
		_ = UpdateSelectedChildData(selectedChild);

		_currentTargetDocument = PropertyExtractor.FindDocumentByPartNumber(selectedChild.PartNumber);
		var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(selectedChild.PartNumber,
			_propertyComparator.SqlDataManager, _currentTargetDocument);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
		if (_currentTargetDocument != null) await UpdateThumbnail(_currentTargetDocument);

		// Zoom to selection
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc ||
			    _currentTargetDocument == null) return;
			var occurrence = ThisApplication.Documents.ItemByName[_currentTargetDocument.FullFileName];
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
		if (PurchasedDataGrid.SelectedItem is not ChildInfo selectedChild) return;
		_ = UpdateSelectedChildData(selectedChild);

		_currentTargetDocument = PropertyExtractor.FindDocumentByPartNumber(selectedChild.PartNumber);
		var (geniusRows, invRows) = await PropertyExtractor.LoadPropertiesForPart(selectedChild.PartNumber,
			_propertyComparator.SqlDataManager, _currentTargetDocument);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
		if (_currentTargetDocument != null) await UpdateThumbnail(_currentTargetDocument);

		// Zoom to selection
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc ||
			    _currentTargetDocument == null) return;
			var occurrence = ThisApplication.Documents.ItemByName[_currentTargetDocument.FullFileName];
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

	private async Task UpdateSelectedChildData(ChildInfo selectedChild)
	{
		try
		{
			var target = ThisApplication.Documents.Cast<Document>().FirstOrDefault(doc =>
				string.Equals(doc.FullFileName, selectedChild.FullPath, StringComparison.OrdinalIgnoreCase));

			// Fallback if not found by full path (e.g., virtual components)
			target ??= ThisApplication.Documents.Cast<Document>().FirstOrDefault(doc =>
				string.Equals(doc.DisplayName, selectedChild.PartNumber, StringComparison.OrdinalIgnoreCase) ||
				string.Equals(Path.GetFileNameWithoutExtension(doc.FullFileName), selectedChild.PartNumber,
					StringComparison.OrdinalIgnoreCase));

			_currentTargetDocument = target;

			if (target != null)
			{
				await UpdateThumbnail(target);
				await RefreshData(target);
			}
			else
			{
				image.Source                 = null;
				SqlDataGrid.ItemsSource      = null;
				InventorDataGrid.ItemsSource = null;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: UpdateChild error: {ex.Message}");
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

	private void InventorDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
	{
		if (e.EditAction != DataGridEditAction.Commit) return;
		if (e.Row.Item is not PropertyRow invRow) return;

		// Use ContextIdle priority to ensure edit transaction completes first
		Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, () =>
		{
			var invVal  = invRow.TryGetValue("Inventor Value", out var inv) ? inv?.ToString() : "";
			var sqlVal  = invRow.TryGetValue("SQL Value", out var sql) ? sql?.ToString() : "";
			var hasDiff = !GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
			invRow.HasDifference = hasDiff;
			invRow["Status"]     = hasDiff ? "Mismatch" : "Match";

			// Find and update matching row in SqlDataGrid
			var propertyName = invRow.Property;
			if (SqlDataGrid.ItemsSource is IEnumerable<PropertyRow> sqlRows)
				foreach (var sqlRow in sqlRows.Where(r => r.Property == propertyName))
				{
					sqlRow.HasDifference     = hasDiff;
					sqlRow["Status"]         = hasDiff ? "Mismatch" : "Match";
					sqlRow["Inventor Value"] = invVal;
				}

			InventorDataGrid.Items.Refresh();
			SqlDataGrid.Items.Refresh();
		});
	}
}