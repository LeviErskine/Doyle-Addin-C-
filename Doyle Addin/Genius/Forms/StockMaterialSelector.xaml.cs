namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using ThemeManager = Options.Themes.ThemeManager;

/// <summary>
///     Interaction logic for StockMaterialSelector.xaml
/// </summary>
public partial class StockMaterialSelector
{
	private static readonly RoundStockTypeOption[] RoundStockTypeOptions =
	[
		new("BarStock", "Bar Stock"),
		new("RoundTube", "Round Tube / Pipe")
	];

	private readonly bool _allowRoundTypeSelection;
	private readonly List<Dictionary<string, string>> _barMatches;

	private readonly Dictionary<string, string> _dimensions;
	private readonly List<Dictionary<string, string>> _initialMatches;
	private readonly ISqlDataManager _sqlDataManager;
	private readonly List<Dictionary<string, string>> _tubeMatches;
	private string _activeStockTypeKey;
	private List<Dictionary<string, string>> _allItems;
	private List<Dictionary<string, string>> _allMaterials;
	private List<string> _costCenters;
	private bool _isLoadingRoundTypeMatches;

	private List<Dictionary<string, string>> _originalMatches;

	// Tracks whether a selection or cancel was already completed so window-close can be ignored
	private bool _selectionFinalized;

	/// <summary>
	///     Initializes a new instance of the StockMaterialSelector control.
	/// </summary>
	/// <param name="stockType">The detected stock material type (e.g., Square Tube, Bar Stock).</param>
	/// <param name="dimensions">They detected dimensions from the part.</param>
	/// <param name="matches">The list of matching SQL records to display.</param>
	/// <param name="costCenters">The list of available cost centers.</param>
	/// <param name="sqlDataManager">The SQL data manager for reloading cost centers.</param>
	/// <param name="partsOnly">If true, only PARTS families are loaded. If false, all families are loaded.</param>
	/// <param name="allowRoundTypeSelection">If true, show a dropdown to choose bar stock vs. round tube/pipe.</param>
	/// <param name="barMatches">Bar stock matches (for client-side filtering when allowRoundTypeSelection is true).</param>
	/// <param name="tubeMatches">Round tube matches (for client-side filtering when allowRoundTypeSelection is true).</param>
	/// <param name="showAllMaterials">If true, automatically check 'Show All Materials' and load all materials.</param>
	/// <param name="partNumber">The part number to display in the selector header.</param>
	public StockMaterialSelector(string stockType, Dictionary<string, string> dimensions,
		List<Dictionary<string, string>> matches, List<string> costCenters, ISqlDataManager sqlDataManager,
		bool partsOnly = true, bool allowRoundTypeSelection = false,
		List<Dictionary<string, string>> barMatches = null, List<Dictionary<string, string>> tubeMatches = null,
		bool showAllMaterials = false, string partNumber = null)
	{
		try
		{
			_costCenters    = costCenters;
			_sqlDataManager = sqlDataManager;
			_dimensions = dimensions != null
				? new Dictionary<string, string>(dimensions, StringComparer.OrdinalIgnoreCase)
				: new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			_allowRoundTypeSelection = allowRoundTypeSelection;
			_activeStockTypeKey      = allowRoundTypeSelection ? "BarStock" : stockType;
			_initialMatches = matches != null
				? [.. matches]
				: [];
			_barMatches = barMatches != null
				? [.. barMatches]
				: [];
			_tubeMatches = tubeMatches != null
				? [.. tubeMatches]
				: [];
			InitializeComponent();
			ThemeManager.ApplyTheme(this);
			Loaded += (_, _) =>
			{
				try
				{
					var parentWindow = Window.GetWindow(this);
					if (parentWindow == null) return;
					ThemeManager.ApplyTheme(parentWindow);
					// If the user closes the containing window (X), treat it as a cancellation unless selection already finalized
					parentWindow.Closing += ParentWindow_Closing;

					// Find the GeniusAssembly form and select the part we're working on
					if (!string.IsNullOrWhiteSpace(partNumber))
						foreach (Window window in Application.Current.Windows)
							if (window.Content is GeniusAssembly ga)
							{
								ga.SelectPartByPartNumber(partNumber);
								break;
							}
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"StockMaterialSelector: Error selecting part in GeniusAssembly: {ex.Message}");
				}
			};

			if (_allowRoundTypeSelection)
			{
				StockTypeText.Text                 = "Round profile (confirm material type)";
				RoundStockTypePanel.Visibility     = Visibility.Visible;
				RoundStockTypeComboBox.ItemsSource = RoundStockTypeOptions;
				RoundStockTypeComboBox.SelectedItem =
					RoundStockTypeOptions.FirstOrDefault(o => o.Key == _activeStockTypeKey) ??
					RoundStockTypeOptions[0];
			}
			else
			{
				StockTypeText.Text = InsertSpacesBeforeCaps(stockType);
			}

			PartNumberText.Text = !string.IsNullOrWhiteSpace(partNumber) ? partNumber : "(unknown)";

			SearchAllFamiliesCheckBox.IsChecked = !partsOnly;

			Debug.WriteLine($"StockMaterialSelector: Constructor called with {_costCenters.Count} cost centers");

			UpdateDimensionsDisplay();
			ApplyRoundStockColumnVisibility(_activeStockTypeKey);

			// Handle DataGrid selection changed to autopopulate cost center
			MatchesDataGrid.SelectionChanged += OnMatchSelectionChanged;

			// Load matches into the DataGrid - filter to active round type if applicable
			MatchesDataGrid.ItemsSource = _allowRoundTypeSelection ? [.. _barMatches] : _initialMatches;

			// Bind cost centers to ComboBox
			CostCenterComboBox.ItemsSource = _costCenters;
			Debug.WriteLine($"StockMaterialSelector: ComboBox ItemsSource set with {_costCenters.Count} items");

			// Wire up the Search All Families checkbox
			SearchAllFamiliesCheckBox.Checked   += OnSearchAllFamiliesChanged;
			SearchAllFamiliesCheckBox.Unchecked += OnSearchAllFamiliesChanged;

			// Wire up the Show All Materials checkbox
			ShowAllMaterialsCheckBox.Checked   += ShowAllMaterialsCheckBox_Checked;
			ShowAllMaterialsCheckBox.Unchecked += ShowAllMaterialsCheckBox_Unchecked;

			// Wire up the Made From Another Part checkbox
			MadeFromAnotherPartCheckBox.Checked   += MadeFromAnotherPartCheckBox_Checked;
			MadeFromAnotherPartCheckBox.Unchecked += MadeFromAnotherPartCheckBox_Unchecked;

			// Wire up the search text box
			SearchTextBox.TextChanged += SearchTextBox_TextChanged;
			SearchTextBox.IsEnabled   =  false;

			// If showAllMaterials is requested, check the box to load all materials
			if (showAllMaterials)
				ShowAllMaterialsCheckBox.IsChecked = true;

			// If the part has a Cost Center property, select it in the ComboBox
			if (_dimensions.TryGetValue("Cost Center", out var partCostCenter) &&
			    !string.IsNullOrWhiteSpace(partCostCenter))
			{
				var matchedCostCenter = _costCenters.FirstOrDefault(cc =>
					cc.Equals(partCostCenter, StringComparison.OrdinalIgnoreCase));
				if (matchedCostCenter != null)
				{
					CostCenterComboBox.Text = matchedCostCenter;
					Debug.WriteLine(
						$"StockMaterialSelector: Selected Cost Center '{matchedCostCenter}' from part properties");
				}
			}

			// Try to select a match based on the "RM" custom property from the part
			if (matches is not { Count: > 0 }) return;
			TrySelectMatchByRm(MatchesDataGrid.ItemsSource as IReadOnlyList<Dictionary<string, string>> ?? matches);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error initializing control: {ex.Message}");
		}
	}

	/// <summary>
	///     Occurs when the user selects a material.
	/// </summary>
	public event EventHandler<Dictionary<string, string>> MaterialSelected;

	/// <summary>
	///     Occurs when the user cancels selection.
	/// </summary>
	public event EventHandler SelectionCancelled;

	private List<Dictionary<string, string>> GetCurrentTypeMatches()
	{
		return _allowRoundTypeSelection
			? _activeStockTypeKey == "RoundTube" ? _tubeMatches : _barMatches
			: _initialMatches;
	}

	private void LoadMatchesAndSelect(List<Dictionary<string, string>> matches)
	{
		MatchesDataGrid.ItemsSource = matches.ToList();
		TrySelectMatchByRm(matches);
	}

	private static string InsertSpacesBeforeCaps(string input)
	{
		if (string.IsNullOrEmpty(input))
			return input;

		var result = new StringBuilder();
		for (var i = 0; i < input.Length; i++)
		{
			if (i > 0 && char.IsUpper(input[i]))
				result.Append(' ');
			result.Append(input[i]);
		}

		return result.ToString();
	}

	private static string FormatDimension(string value)
	{
		return decimal.TryParse(value, out var number)
			?
			// Remove trailing zeros by using a format that avoids unnecessary decimal places
			number.ToString("0.####################")
			: value;
	}

	private void UpdateDimensionsDisplay()
	{
		var dimParts     = new List<string>();
		var stockTypeKey = _activeStockTypeKey;

		if (_allowRoundTypeSelection)
		{
			if (_activeStockTypeKey == "RoundTube")
			{
				if (_dimensions.TryGetValue("OD", out var od))
					dimParts.Add($"OD: {FormatDimension(od)}\"");
			}
			else
			{
				if (_dimensions.TryGetValue("Diameter", out var diameter))
					dimParts.Add($"Diameter: {FormatDimension(diameter)}\"");
			}
		}
		else
		{
			switch (stockTypeKey)
			{
				case "BarStock":
				{
					if (_dimensions.TryGetValue("Diameter", out var diameter))
						dimParts.Add($"Diameter: {FormatDimension(diameter)}\"");
					break;
				}
				case "RoundTube":
				{
					if (_dimensions.TryGetValue("OD", out var od))
						dimParts.Add($"OD: {FormatDimension(od)}\"");
					break;
				}
				default:
				{
					if (_dimensions.TryGetValue("Width", out var width))
						dimParts.Add($"Width: {FormatDimension(width)}\"");
					if (_dimensions.TryGetValue("Height", out var height))
						dimParts.Add($"Height: {FormatDimension(height)}\"");
					break;
				}
			}
		}

		if (_dimensions.TryGetValue("Length", out var length))
			dimParts.Add($"Length: {FormatDimension(length)}\"");

		DimensionsText.Text = string.Join(", ", dimParts);
	}

	private void ApplyRoundStockColumnVisibility(string stockTypeKey)
	{
		if (MatchesDataGrid.Columns.Count <= 5) return;

		var isRoundStock = stockTypeKey is "BarStock" or "RoundTube" or "AmbiguousRound";
		MatchesDataGrid.Columns[5].Visibility = isRoundStock ? Visibility.Visible : Visibility.Collapsed;
		MatchesDataGrid.Columns[2].Visibility = isRoundStock ? Visibility.Collapsed : Visibility.Visible;
		MatchesDataGrid.Columns[3].Visibility = isRoundStock ? Visibility.Collapsed : Visibility.Visible;
	}

	private void TrySelectMatchByRm(IReadOnlyList<Dictionary<string, string>> matches)
	{
		if (!_dimensions.TryGetValue("RM", out var rmValue) || string.IsNullOrWhiteSpace(rmValue))
		{
			MatchesDataGrid.SelectedIndex = -1;
			return;
		}

		var rmMatchIndex = matches.ToList().FindIndex(m =>
			m.TryGetValue("RM", out var matchRm) &&
			matchRm.Equals(rmValue, StringComparison.OrdinalIgnoreCase));

		if (rmMatchIndex >= 0)
		{
			MatchesDataGrid.SelectedIndex = rmMatchIndex;
			MatchesDataGrid.ScrollIntoView(MatchesDataGrid.SelectedItem);
			Debug.WriteLine(
				$"StockMaterialSelector: Selected match by RM property '{rmValue}' at index {rmMatchIndex}");
			return;
		}

		MatchesDataGrid.SelectedIndex = -1;
		Debug.WriteLine(
			$"StockMaterialSelector: No match found for RM property '{rmValue}', no item auto-selected");
	}

	private void RoundStockTypeComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (!_allowRoundTypeSelection || _isLoadingRoundTypeMatches) return;
		if (RoundStockTypeComboBox.SelectedItem is not RoundStockTypeOption selectedType) return;
		if (selectedType.Key == _activeStockTypeKey) return;

		try
		{
			_isLoadingRoundTypeMatches = true;
			_activeStockTypeKey        = selectedType.Key;
			UpdateDimensionsDisplay();
			ApplyRoundStockColumnVisibility(_activeStockTypeKey);

			if (ShowAllMaterialsCheckBox.IsChecked == true || MadeFromAnotherPartCheckBox.IsChecked == true)
				return;

			SearchTextBox.Text                    = string.Empty;
			SearchTextBox.Visibility              = Visibility.Collapsed;
			SearchTextBox.IsEnabled               = false;
			_originalMatches                      = null;
			MadeFromAnotherPartCheckBox.IsChecked = false;

			LoadMatchesAndSelect(GetCurrentTypeMatches());

			Debug.WriteLine(
				$"StockMaterialSelector: Showing matches for stock type '{_activeStockTypeKey}'");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error changing round stock type: {ex.Message}");
			MessageBox.Show($"Error loading materials for {selectedType.DisplayName}: {ex.Message}", "Error",
				MessageBoxButton.OK, MessageBoxImage.Error);
		}
		finally
		{
			_isLoadingRoundTypeMatches = false;
		}
	}

	// Apply a list as the current original matches and enable the search UI
	private void ApplyOriginalMatches(List<Dictionary<string, string>> items, bool showMadeFromCheckbox = true)
	{
		_originalMatches            = items != null ? [.. items] : null;
		MatchesDataGrid.ItemsSource = items;
		SearchTextBox.Visibility    = Visibility.Visible;
		SearchTextBox.IsEnabled     = true;
		if (showMadeFromCheckbox)
			MadeFromAnotherPartCheckBox.Visibility = Visibility.Visible;
	}

	// Restore the UI to initial state showing only the initial matches
	private void RestoreInitialMatches()
	{
		SearchTextBox.Text                     = string.Empty;
		SearchTextBox.Visibility               = Visibility.Collapsed;
		SearchTextBox.IsEnabled                = false;
		MadeFromAnotherPartCheckBox.Visibility = Visibility.Collapsed;
		MadeFromAnotherPartCheckBox.IsChecked  = false;
		_allMaterials                          = null;
		_allItems                              = null;
		_originalMatches                       = null;
		LoadMatchesAndSelect(GetCurrentTypeMatches());
	}

	private void OnMatchSelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		try
		{
			if (MatchesDataGrid.SelectedItem is not Dictionary<string, string> selected ||
			    _costCenters is not { Count: > 0 }) return;

			// First, try to match by Cost Center property if it exists
			if (selected.TryGetValue("Cost Center", out var costCenter) && !string.IsNullOrWhiteSpace(costCenter))
			{
				var match = _costCenters.FirstOrDefault(cc =>
					cc.Equals(costCenter, StringComparison.OrdinalIgnoreCase));
				if (match != null)
				{
					CostCenterComboBox.Text = match;
					return;
				}
			}

			// Fall back to matching by Family
			if (!selected.TryGetValue("Family", out var family) || string.IsNullOrWhiteSpace(family)) return;
			{
				var match = _costCenters.FirstOrDefault(cc =>
					cc.Equals(family, StringComparison.OrdinalIgnoreCase));
				if (match != null)
					CostCenterComboBox.Text = match;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error in selection changed: {ex.Message}");
		}
	}


	private void SelectButton_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			if (MatchesDataGrid.SelectedItem is Dictionary<string, string> selected)
			{
				// Apply the user's cost center selection
				var selectedCostCenter = CostCenterComboBox.Text?.Trim();
				if (!string.IsNullOrWhiteSpace(selectedCostCenter))
					selected["Cost Center"] = selectedCostCenter;

				// Mark finalized so window close doesn't trigger duplicate cancel
				_selectionFinalized = true;

				MaterialSelected?.Invoke(this, selected);
			}
			else
			{
				MessageBox.Show("Please select a material from the list.", "No Selection",
					MessageBoxButton.OK, MessageBoxImage.Information);
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error in SelectButton_Click: {ex.Message}");
			MessageBox.Show($"Error selecting material: {ex.Message}", "Error",
				MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	private void OnSearchAllFamiliesChanged(object sender, RoutedEventArgs e)
	{
		try
		{
			if (_sqlDataManager == null) return;

			var searchAll = SearchAllFamiliesCheckBox.IsChecked == true;
			_costCenters = searchAll
				? _sqlDataManager.GetCostCentersAsync(false).GetAwaiter().GetResult()
				: _sqlDataManager.GetCostCentersAsync().GetAwaiter().GetResult();

			CostCenterComboBox.ItemsSource = _costCenters;
			Debug.WriteLine(
				$"StockMaterialSelector: Reloaded {_costCenters.Count} cost centers (Search All Families = {searchAll})");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error reloading cost centers: {ex.Message}");
		}
	}

	private void CancelButton_Click(object sender, RoutedEventArgs e)
	{
		try
		{
			// Mark finalized so window close doesn't also trigger cancellation
			_selectionFinalized = true;
			SelectionCancelled?.Invoke(this, EventArgs.Empty);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error in CancelButton_Click: {ex.Message}");
		}
	}

	private void ParentWindow_Closing(object sender, CancelEventArgs e)
	{
		try
		{
			if (!_selectionFinalized)
				SelectionCancelled?.Invoke(this, EventArgs.Empty);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error in ParentWindow_Closing: {ex.Message}");
		}
	}

	private async void MadeFromAnotherPartCheckBox_Checked(object sender, RoutedEventArgs e)
	{
		try
		{
			if (_sqlDataManager == null) return;

			Debug.WriteLine("StockMaterialSelector: Loading all items...");
			var allItems = await _sqlDataManager.GetAllItemsAsync();
			_allItems = allItems != null
				? [.. allItems]
				: [];
			ApplyOriginalMatches(_allItems, false);

			Debug.WriteLine($"StockMaterialSelector: Loaded {_allItems.Count} total items");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error loading all items: {ex.Message}");
			MessageBox.Show($"Error loading all items: {ex.Message}", "Error",
				MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	private void MadeFromAnotherPartCheckBox_Unchecked(object sender, RoutedEventArgs e)
	{
		try
		{
			SearchTextBox.Text = string.Empty;

			_allItems        = null;
			_originalMatches = null;

			// If Show All Materials is still checked, restore the stock materials list
			if (ShowAllMaterialsCheckBox.IsChecked == true && _allMaterials != null)
			{
				_originalMatches            = [.._allMaterials];
				MatchesDataGrid.ItemsSource = _allMaterials;
			}
			else
			{
				LoadMatchesAndSelect(GetCurrentTypeMatches());
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error unchecking made from another part: {ex.Message}");
		}
	}

	private async void ShowAllMaterialsCheckBox_Checked(object sender, RoutedEventArgs e)
	{
		try
		{
			if (_sqlDataManager == null) return;

			Debug.WriteLine("StockMaterialSelector: Loading all stock materials...");
			_allMaterials = await _sqlDataManager.GetAllStockMaterialsAsync();
			_allMaterials = _allMaterials != null
				? [.. _allMaterials]
				: [];
			ApplyOriginalMatches(_allMaterials);
			SearchTextBox.Text = string.Empty;

			Debug.WriteLine($"StockMaterialSelector: Loaded {_allMaterials.Count} total materials");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error loading all materials: {ex.Message}");
			MessageBox.Show($"Error loading all materials: {ex.Message}", "Error",
				MessageBoxButton.OK, MessageBoxImage.Error);
		}
	}

	private void ShowAllMaterialsCheckBox_Unchecked(object sender, RoutedEventArgs e)
	{
		try
		{
			// Restore search UI and lists to initial state
			RestoreInitialMatches();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error unchecking show all materials: {ex.Message}");
		}
	}

	private void SearchTextBox_GotFocus(object sender, RoutedEventArgs e)
	{
		SearchTextBox.SelectAll();
	}

	private void SearchTextBox_LostFocus(object sender, RoutedEventArgs e)
	{
		// Watermark style handles placeholder display automatically
	}

	private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
	{
		try
		{
			if (_originalMatches == null || _originalMatches.Count == 0) return;

			var searchText = SearchTextBox.Text.Trim();

			if (string.IsNullOrEmpty(searchText))
			{
				MatchesDataGrid.ItemsSource = _originalMatches;
				return;
			}

			var filtered = FuzzySearch(_originalMatches, searchText);
			MatchesDataGrid.ItemsSource = filtered;

			Debug.WriteLine(
				$"StockMaterialSelector: Fuzzy search for '{searchText}' returned {filtered.Count} results");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"StockMaterialSelector: Error in search: {ex.Message}");
		}
	}

	private static List<Dictionary<string, string>> FuzzySearch(
		List<Dictionary<string, string>> materials, string searchText)
	{
		if (string.IsNullOrWhiteSpace(searchText))
			return materials;

		var search = searchText.Trim();

		// Score each item, keep only positive scores, and sort by descending score (best matches first)
		var results = materials
		              .Select(m => new { Item = m, Score = CalculateFuzzyScore(m, search) })
		              .Where(x => x.Score > 0)
		              .OrderByDescending(x => x.Score)
		              .ThenBy(x => x.Item.TryGetValue("RM", out var rm) ? rm : string.Empty,
			              StringComparer.OrdinalIgnoreCase)
		              .Select(x => x.Item)
		              .ToList();

		return results;
	}

	private static int CalculateFuzzyScore(Dictionary<string, string> material, string search)
	{
		if (string.IsNullOrEmpty(search) || material == null) return 0;

		const int rmExact        = 100, rmPrefix   = 80, rmContains   = 60, rmFuzzyMax   = 40;
		const int descExact      = 50,  descPrefix = 40, descContains = 30, descFuzzyMax = 25;
		const int familyContains = 10;

		var score = 0;

		// Item Number (RM) - highest priority
		var itemNumber = GetValue(material, "RM");
		if (!string.IsNullOrWhiteSpace(itemNumber))
		{
			if (string.Equals(itemNumber, search, StringComparison.OrdinalIgnoreCase))
			{
				score += rmExact;
			}
			else if (itemNumber.StartsWith(search, StringComparison.OrdinalIgnoreCase))
			{
				score += rmPrefix;
			}
			else if (itemNumber.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
			{
				score += rmContains;
			}
			else
			{
				// Fuzzy similarity via the Levenshtein ratio (gives a fraction 0..1). Only add if reasonably similar.
				var dist = LevenshteinDistance(itemNumber, search);
				var max  = Math.Max(itemNumber.Length, search.Length);
				if (max > 0)
				{
					var similarity = 1.0 - (double)dist / max; // 0..1
					if (similarity > 0.4)                      // avoid tiny matches
						score += (int)Math.Round(similarity * rmFuzzyMax);
				}
				else if (IsSubsequence(itemNumber, search))
				{
					score += rmFuzzyMax / 4; // subsequence fallback
				}
			}
		}

		// Description - secondary priority
		var description = GetValue(material, "Description1");
		if (!string.IsNullOrWhiteSpace(description))
		{
			if (string.Equals(description, search, StringComparison.OrdinalIgnoreCase))
			{
				score += descExact;
			}
			else if (description.StartsWith(search, StringComparison.OrdinalIgnoreCase))
			{
				score += descPrefix;
			}
			else if (description.IndexOf(search, StringComparison.OrdinalIgnoreCase) >= 0)
			{
				score += descContains;
			}
			else
			{
				var dist = LevenshteinDistance(description, search);
				var max  = Math.Max(description.Length, search.Length);
				if (max > 0)
				{
					var similarity = 1.0 - (double)dist / max;
					if (similarity > 0.35)
						score += (int)Math.Round(similarity * descFuzzyMax);
				}
				else if (IsSubsequence(description, search))
				{
					score += descFuzzyMax / 4;
				}
			}
		}

		// Family - tertiary priority (contains only)
		var family = GetValue(material, "Family");
		if (!string.IsNullOrWhiteSpace(family) && family.Contains(search, StringComparison.OrdinalIgnoreCase))
			score += familyContains;

		return score;

		// Helper to get value safely
		static string GetValue(Dictionary<string, string> d, string key)
		{
			return d != null && d.TryGetValue(key, out var v) ? v ?? string.Empty : string.Empty;
		}
	}

	// Fast subsequence check (e.g., 'rdg' matches 'round gauge') used as a weak fuzzy indicator
	private static bool IsSubsequence(string source, string pattern)
	{
		if (string.IsNullOrEmpty(pattern) || string.IsNullOrEmpty(source))
			return false;
		var si = 0;
		var pi = 0;
		while (si < source.Length && pi < pattern.Length)
		{
			if (char.ToLowerInvariant(source[si]) == char.ToLowerInvariant(pattern[pi])) pi++;
			si++;
		}

		return pi == pattern.Length;
	}

	// Levenshtein distance (iterative, memory-efficient)
	private static int LevenshteinDistance(string a, string b)
	{
		if (string.IsNullOrEmpty(a)) return b?.Length ?? 0;
		if (string.IsNullOrEmpty(b)) return a.Length;

		a = a.ToLowerInvariant();
		b = b.ToLowerInvariant();

		var previous                                    = new int[b.Length + 1];
		for (var j = 0; j <= b.Length; j++) previous[j] = j;

		for (var i = 1; i <= a.Length; i++)
		{
			var current = new int[b.Length + 1];
			current[0] = i;
			for (var j = 1; j <= b.Length; j++)
			{
				var cost = a[i - 1] == b[j - 1] ? 0 : 1;
				current[j] = Math.Min(Math.Min(current[j - 1] + 1, previous[j] + 1), previous[j - 1] + cost);
			}

			previous = current;
		}

		return previous[b.Length];
	}

	private sealed class RoundStockTypeOption(string key, string displayName)
	{
		public string Key { get; } = key;
		public string DisplayName { get; } = displayName;

		public override string ToString()
		{
			return DisplayName;
		}
	}
}