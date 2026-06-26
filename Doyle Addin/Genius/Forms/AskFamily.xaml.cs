namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Windows;
using System.Windows.Media.Animation;

public partial class AskFamily
{
	private readonly string _geniusValue;
	private readonly ISqlDataManager _sqlDataManager;
	private bool _isOtherPanelOpen;

	public AskFamily(string inventorValue, string geniusValue, List<string> costCenters,
		ISqlDataManager sqlDataManager, Document document = null)
	{
		InitializeComponent();

		_geniusValue    = geniusValue;
		_sqlDataManager = sqlDataManager;

		InventorValueText.Text         = inventorValue;
		GeniusValueText.Text           = geniusValue;
		CostCenterComboBox.ItemsSource = costCenters;

		if (costCenters.FirstOrDefault(cc =>
			    cc.Equals(geniusValue, StringComparison.OrdinalIgnoreCase)) is { } match)
			CostCenterComboBox.Text = match;

		SearchAllFamiliesCheckBox.Checked   += OnSearchAllFamiliesChanged;
		SearchAllFamiliesCheckBox.Unchecked += OnSearchAllFamiliesChanged;

		SetThumbnail(document);
	}

	/// <summary>
	///     The cost center to apply. Null means no change.
	/// </summary>
	public string SelectedCostCenter { get; private set; }

	private void SetThumbnail(Document document)
	{
		try
		{
			image.Source = document != null
				? ThumbnailHelper.GetThumbnail(document)
				: null;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"AskFamily: Error setting thumbnail: {ex.Message}");
			image.Source = null;
		}
	}

	private void YesButton_Click(object sender, RoutedEventArgs e)
	{
		SelectedCostCenter = _geniusValue;
		DialogResult       = true;
	}

	private void NoButton_Click(object sender, RoutedEventArgs e)
	{
		SelectedCostCenter = null;
		DialogResult       = false;
	}

	private void OtherButton_Click(object sender, RoutedEventArgs e)
	{
		var key = _isOtherPanelOpen ? "CloseTransition" : "OpenTransition";
		BeginStoryboard((Storyboard)FindResource(key));
		_isOtherPanelOpen = !_isOtherPanelOpen;
	}

	private void SelectFamilyButton_Click(object sender, RoutedEventArgs e)
	{
		var selected = CostCenterComboBox.Text?.Trim();
		if (string.IsNullOrWhiteSpace(selected)) return;
		SelectedCostCenter = selected;
		DialogResult       = true;
	}

	private void OnSearchAllFamiliesChanged(object sender, RoutedEventArgs e)
	{
		try
		{
			if (_sqlDataManager == null) return;

			var searchAll = SearchAllFamiliesCheckBox.IsChecked == true;
			var costCenters = searchAll
				? _sqlDataManager.GetCostCentersAsync(false).GetAwaiter().GetResult()
				: _sqlDataManager.GetCostCentersAsync().GetAwaiter().GetResult();

			CostCenterComboBox.ItemsSource = costCenters;

			if (costCenters.FirstOrDefault(cc =>
				    cc.Equals(_geniusValue, StringComparison.OrdinalIgnoreCase)) is { } match)
				CostCenterComboBox.Text = match;

			Debug.WriteLine(
				$"AskFamily: Reloaded {costCenters.Count} cost centers (Search All Families = {searchAll})");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"AskFamily: Error reloading cost centers: {ex.Message}");
		}
	}
}