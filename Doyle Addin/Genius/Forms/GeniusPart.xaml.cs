namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ThemeManager = Options.Themes.ThemeManager;

/// <summary>
///     Represents a user interface control designed to facilitate interaction with
///     and management of Genius-related components within the application.
/// </summary>
public partial class GeniusPart
{
	private readonly Dictionary<string, string> _pendingProperties = [];
	private readonly PropertyComparator _propertyComparator;

	/// <summary>
	///     Represents a user interface control designed to facilitate interaction with
	///     and management of Genius-related components within the application.
	/// </summary>
	public GeniusPart()
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
			Debug.WriteLine($"GeniusPart: Exception in constructor: {ex.Message}");
		}
	}

	private async void CalculatePropsButton_Click(object _sender, RoutedEventArgs _e)
	{
		await CalculatePropertiesPreview();
	}

	private void SaveButton_Click(object _sender, RoutedEventArgs _e)
	{
		SavePendingProperties();
		Window.GetWindow(this)?.Close();
	}

	private void CancelButton_Click(object _sender, RoutedEventArgs _e)
	{
		_pendingProperties.Clear();
		SaveButton.IsEnabled   = false;
		CancelButton.IsEnabled = false;
		_                      = RefreshData();
	}

	private async Task RefreshData()
	{
		try
		{
			if (ThisApplication?.ActiveDocument == null) return;

			var inventorProperties = PropertyExtractor.GetAllProperties();
			var partNumber         = inventorProperties.GetValueOrDefault("Part Number", "");
			var sqlData            = await _propertyComparator.SqlDataManager.GetSqlDataAsync(partNumber);

			// Map SQL to rows
			var geniusProperties = sqlData.Select(kvp =>
			{
				var invName = GeniusFormsHelper.MapSqlColumnToInventorProperty(kvp.Key);
				var displayValue = _pendingProperties.GetValueOrDefault(invName) ??
				                   inventorProperties.GetValueOrDefault(invName, "");
				return new PropertyRow
				{
					Property      = invName,
					["SQL Value"] = kvp.Value,
					HasDifference = !GeniusFormsHelper.ValuesAreEqual(kvp.Value, displayValue)
				};
			}).ToList();

			if (geniusProperties.Count == 0)
				geniusProperties.Add(new PropertyRow { Property = "Info", ["SQL Value"] = "No data found" });

			// Map Inventor to rows
			var inventorPropertyRows = inventorProperties.Select(kvp =>
			{
				var displayValue = _pendingProperties.GetValueOrDefault(kvp.Key) ?? kvp.Value;
				var sqlValue     = sqlData.GetValueOrDefault(Geniusinfo.GetSqlColumnName(kvp.Key), "");
				return new PropertyRow
				{
					Property           = kvp.Key,
					["Inventor Value"] = displayValue,
					HasDifference      = !GeniusFormsHelper.ValuesAreEqual(sqlValue, displayValue)
				};
			}).ToList();
			inventorPropertyRows.AddRange(
				from pending in _pendingProperties.Where(p => !inventorProperties.ContainsKey(p.Key))
				let sqlValue = sqlData.GetValueOrDefault(Geniusinfo.GetSqlColumnName(pending.Key), "")
				select new PropertyRow
				{
					Property      = pending.Key, ["Inventor Value"] = pending.Value,
					HasDifference = !GeniusFormsHelper.ValuesAreEqual(sqlValue, pending.Value)
				});

			// Add pending new properties

			await Dispatcher.InvokeAsync(() =>
			{
				SqlDataGrid.ItemsSource      = geniusProperties;
				InventorDataGrid.ItemsSource = inventorPropertyRows;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusPart: Refresh error: {ex.Message}");
		}
	}

	private async Task CalculatePropertiesPreview()
	{
		CalculatePropsButton.IsEnabled = false;
		try
		{
			var calculated = await CalculateProps.CalculateAllPropertiesAsync();
			if (calculated.Count == 0) return;

			var current = PropertyExtractor.GetAllProperties();
			_pendingProperties.Clear();

			foreach (var calc in calculated.Where(calc =>
				         !GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(calc.Key, ""), calc.Value)))
				_pendingProperties[calc.Key] = calc.Value;

			await RefreshData();
			SaveButton.IsEnabled   = _pendingProperties.Count > 0;
			CancelButton.IsEnabled = _pendingProperties.Count > 0;
		}
		finally
		{
			CalculatePropsButton.IsEnabled = true;
		}
	}

	private void SavePendingProperties()
	{
		if (_pendingProperties.Count == 0) return;
		GeniusFormsHelper.UpdateInventorProperties(ThisApplication.ActiveDocument, _pendingProperties, "GeniusPart");
		_pendingProperties.Clear();
		SaveButton.IsEnabled   = false;
		CancelButton.IsEnabled = false;
	}

	private async void InventorDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
	{
		if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not PropertyRow rowData) return;

		var propertyName = rowData.Property;
		var newValue     = e.EditingElement is TextBox tb ? tb.Text : "";

		_pendingProperties[propertyName] = newValue;
		SaveButton.IsEnabled             = true;
		CancelButton.IsEnabled           = true;

		// Refresh to update highlights
		await RefreshData();
	}
}