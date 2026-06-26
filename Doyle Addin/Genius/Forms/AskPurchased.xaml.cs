namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Windows;
using System.Windows.Media.Animation;

public partial class AskPurchased
{
	private readonly ISqlDataManager _sqlDataManager;
	private bool _isPurchasedPanelOpen;

	public AskPurchased(string partNumber, List<string> costCenters,
		ISqlDataManager sqlDataManager, Document document = null, string currentCostCenter = null)
	{
		InitializeComponent();

		_sqlDataManager = sqlDataManager;

		PartNumberVal.Text             = !string.IsNullOrWhiteSpace(partNumber) ? partNumber : "(unknown)";
		CostCenterVal.Text             = !string.IsNullOrWhiteSpace(currentCostCenter) ? currentCostCenter : "(none)";
		CostCenterComboBox.ItemsSource = FilterDefaultCostCenters(costCenters);

		SearchAllFamiliesCheckBox.Checked   += OnSearchAllFamiliesChanged;
		SearchAllFamiliesCheckBox.Unchecked += OnSearchAllFamiliesChanged;

		SetThumbnail(document);
	}

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
			Debug.WriteLine($"AskPurchased: Error setting thumbnail: {ex.Message}");
			image.Source = null;
		}
	}

	private static List<string> FilterDefaultCostCenters(List<string> costCenters)
	{
		return costCenters?
		       .Where(cc => cc.StartsWith("D-P", StringComparison.OrdinalIgnoreCase) ||
		                    cc.StartsWith("R-P", StringComparison.OrdinalIgnoreCase))
		       .ToList() ?? [];
	}

	private void YesButton_Click(object sender, RoutedEventArgs e)
	{
		var key = _isPurchasedPanelOpen ? "CloseTransition" : "OpenTransition";
		BeginStoryboard((Storyboard)FindResource(key));
		_isPurchasedPanelOpen = !_isPurchasedPanelOpen;
	}

	private void NoButton_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
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
				: FilterDefaultCostCenters(
					_sqlDataManager.GetCostCentersAsync().GetAwaiter().GetResult());

			CostCenterComboBox.ItemsSource = costCenters;

			if (costCenters.FirstOrDefault(cc =>
				    cc.Equals(CostCenterComboBox.Text, StringComparison.OrdinalIgnoreCase)) is null)
				CostCenterComboBox.Text = null;

			Debug.WriteLine(
				$"AskPurchased: Reloaded {costCenters.Count} cost centers (Search All Families = {searchAll})");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"AskPurchased: Error reloading cost centers: {ex.Message}");
		}
	}
}