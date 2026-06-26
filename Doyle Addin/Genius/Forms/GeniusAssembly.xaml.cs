namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using TextBox = System.Windows.Controls.TextBox;
using ThemeManager = Options.Themes.ThemeManager;

public partial class GeniusAssembly
{
	private readonly HashSet<string> _memberPartNumbers = new(StringComparer.OrdinalIgnoreCase);

	private readonly Dictionary<string, Dictionary<string, string>> _pendingChanges =
		new(StringComparer.OrdinalIgnoreCase);

	private readonly PropertyComparator _propertyComparator;
	private CancellationTokenSource _calcCts;
	private string _currentMemberName;
	private Document _currentTargetDocument;
	private bool _isIAssembly;
	private ApplicationEventsSink_OnDocumentChangeEventHandler _onDocumentChangeHandler;

	public GeniusAssembly()
	{
		try
		{
			var sqlDataManager = new SqlDataManager(GeniusConstants.DefaultConnectionString);
			_propertyComparator = new PropertyComparator(sqlDataManager);
			CalculateProps.SetSqlDataManager(sqlDataManager);

			InitializeComponent();
			ThemeManager.ApplyTheme(this);

			_ = InitializeAsync();

			Unloaded += (_, _) =>
			{
				try
				{
					ThisApplication.ApplicationEvents.OnDocumentChange -= _onDocumentChangeHandler;
				}
				catch
				{
					// ignored
				}
			};
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Exception in constructor: {ex.Message}");
		}
	}

	public string PanelTitle => _isIAssembly ? "Genius - i-Assembly" : "Genius - Assembly";

	private async Task InitializeAsync()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is AssemblyDocument asmDoc)
			{
				var compDef = asmDoc.ComponentDefinition;
				if (compDef.iAssemblyFactory != null || compDef.ModelStates?.Count > 1)
				{
					compDef.iAssemblyFactory?.MemberEditScope = MemberEditScopeEnum.kEditActiveMember;
					if (compDef.IsModelStateFactory)
						compDef.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditActiveMember;

					_isIAssembly = true;
					SetupIAssemblyTabs();
					_ = LoadMembers();
					_ = LoadAssemblyChildren();
				}
				else
				{
					_isIAssembly = false;
					SetupRegularAssemblyTabs();
					await RefreshData();
				}
			}

			_currentMemberName                                  =  GetCurrentMemberName();
			_onDocumentChangeHandler                            =  OnDocumentChange;
			ThisApplication?.ApplicationEvents.OnDocumentChange += _onDocumentChangeHandler;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: InitializeAsync error: {ex.Message}");
		}
	}

	private void SetupIAssemblyTabs()
	{
		VisualStateManager.GoToState(this, "IAssembly", false);
		CalculateAllPartsButton.ToolTip =
			"Calculate properties for all members and assembly children";
	}

	private void SetupRegularAssemblyTabs()
	{
		VisualStateManager.GoToState(this, "RegularAssembly", false);
		CalculateAllPartsButton.ToolTip =
			"Calculate properties for all parts and subassemblies in this assembly";
	}

	private async Task LoadMembers()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc) return;

			var compDef = assemblyDoc.ComponentDefinition;
			_memberPartNumbers.Clear();
			var members = new List<MemberInfo>();

			if (compDef.iAssemblyFactory != null)
				foreach (iAssemblyTableRow row in compDef.iAssemblyFactory.TableRows)
				{
					members.Add(new MemberInfo
					{
						PartNumber  = row[ColumnNames.PartNumber].Value,
						Description = row[ColumnNames.Description].Value
					});
					_memberPartNumbers.Add(row[ColumnNames.PartNumber].Value);
				}
			else if (compDef.ModelStates?.Count > 1)
				foreach (ModelStateTableRow row in compDef.ModelStates.ModelStateTable.TableRows)
				{
					members.Add(new MemberInfo
					{
						PartNumber  = row[ColumnNames.PartNumber].Value,
						Description = row[ColumnNames.Description].Value
					});
					_memberPartNumbers.Add(row[ColumnNames.PartNumber].Value);
				}

			await Dispatcher.InvokeAsync(() =>
			{
				Members.ItemsSource   = members;
				MembersTab.Visibility = members.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
				UpdateMemberHighlights();

				var activePartNumber = GetActiveMemberPartNumber(assemblyDoc);
				if (activePartNumber != null)
					Members.SelectedItem = members.FirstOrDefault(m =>
						string.Equals(m.PartNumber, activePartNumber, StringComparison.OrdinalIgnoreCase));
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Error loading members: {ex.Message}");
		}
	}

	private async Task LoadAssemblyChildren()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc) return;

			var childrenTable = await Task.Run(() => Geniusinfo.GetAllAssemblyChildren(assemblyDoc));
			var allChildren   = new List<ChildInfo>();
			var purchased     = new List<ChildInfo>();

			foreach (DataRow row in childrenTable.Rows)
			{
				var child = new ChildInfo
				{
					PartNumber   = row["PartNumber"].ToString() ?? "",
					Description  = row["Description"].ToString() ?? "",
					DocumentType = row["DocumentType"].ToString(),
					IsPurchased  = (bool)row["IsPurchased"],
					FullPath     = row["FullPath"].ToString(),
					Level        = (int)row["Level"]
				};

				if (child.IsPurchased)
				{
					purchased.Add(child);
					continue;
				}

				allChildren.Add(child);
			}

			var rootPartNumber = PropertyExtractor.GetPropertiesFromDocumentStatic((Document)assemblyDoc)
			                                      .GetValueOrDefault("Part Number", assemblyDoc.DisplayName);
			var rootDescription = "";
			try
			{
				rootDescription = assemblyDoc.PropertySets["Design Tracking Properties"]?["Description"]?.Value
				                             ?.ToString() ?? "";
			}
			catch
			{
				// ignored
			}

			allChildren.Insert(0, new ChildInfo
			{
				Level        = 0,
				PartNumber   = rootPartNumber,
				Description  = rootDescription,
				DocumentType = "kAssemblyDocumentObject",
				FullPath     = assemblyDoc.FullFileName,
				IsPurchased  = false
			});

			await Dispatcher.InvokeAsync(() =>
			{
				var tree = BuildTree(allChildren);
				MarkPendingOnItems(tree, [], purchased);

				AssembliesTreeView.ItemsSource = tree;
				PurchasedDataGrid.ItemsSource  = purchased;

				var hasChildren = allChildren.Count > 0;
				AssembliesTab.Visibility = hasChildren ? Visibility.Visible : Visibility.Collapsed;
				PurchasedTab.Visibility  = purchased.Count > 0 ? Visibility.Visible : Visibility.Collapsed;
			});

			_currentMemberName = GetCurrentMemberName();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Error loading children: {ex.Message}");
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
				{
					UpdatePendingHighlights();
					return;
				}

				var childrenTable = Geniusinfo.GetAllAssemblyChildren();

				var allChildren   = new List<ChildInfo>();
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

					allChildren.Add(child);

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

				var costCenterCache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
				allChildren =
				[
					.. allChildren.Where(child =>
					{
						if (costCenterCache.TryGetValue(child.PartNumber, out var excluded))
							return !excluded;
						try
						{
							var doc = PropertyExtractor.FindDocumentByPartNumber(child.PartNumber);
							if (doc != null)
							{
								var props = doc.PropertySets["Design Tracking Properties"];
								var cc    = props["Cost Center"].Value?.ToString();
								excluded = string.Equals(cc?.Trim(), "D-HDWR", StringComparison.OrdinalIgnoreCase);
							}
						}
						catch
						{
							// ignored
						}

						costCenterCache[child.PartNumber] = excluded;
						return !excluded;
					})
				];

				if (_currentTargetDocument is AssemblyDocument rootAsmDoc)
				{
					var rootPartNumber = PropertyExtractor.GetPropertiesFromDocumentStatic((Document)rootAsmDoc)
					                                      .GetValueOrDefault("Part Number", rootAsmDoc.DisplayName);
					var rootDescription = "";
					try
					{
						rootDescription = rootAsmDoc.PropertySets["Design Tracking Properties"]?["Description"]?.Value
						                            ?.ToString() ?? "";
					}
					catch
					{
						// ignored
					}

					allChildren.Insert(0, new ChildInfo
					{
						Level        = 0,
						PartNumber   = rootPartNumber,
						Description  = rootDescription,
						DocumentType = "kAssemblyDocumentObject",
						FullPath     = rootAsmDoc.FullFileName,
						IsPurchased  = false
					});
				}

				var treeItems = _isIAssembly
					? allChildren
					: [.. allChildren.Where(c => c.DocumentType == "kAssemblyDocumentObject" || c.Level > 0)];
				var assemblyTree = BuildTree(treeItems);
				MarkPendingOnItems(assemblyTree, parts, purchased);
				AssembliesTreeView.ItemsSource = assemblyTree;
				PurchasedDataGrid.ItemsSource  = purchased;
				PurchasedDataGrid.Items.Refresh();
				var hasPurchased   = purchased.Count > 0;
				var hasAnyChildren = treeItems.Count > 0;

				AssembliesTab.Visibility = hasAnyChildren ? Visibility.Visible : Visibility.Collapsed;
				PurchasedTab.Visibility  = hasPurchased ? Visibility.Visible : Visibility.Collapsed;

				if (tabControl.SelectedItem is not TabItem { Visibility: Visibility.Visible })
					tabControl.SelectedItem = hasAnyChildren ? AssembliesTab :
						hasPurchased                         ? PurchasedTab : tabControl.SelectedItem;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Refresh error: {ex.Message}");
		}
	}

	private void OnDocumentChange(
		_Document documentObject,
		EventTimingEnum beforeOrAfter,
		CommandTypesEnum reasonsForChange,
		NameValueMap context,
		out HandlingCodeEnum handlingCode)
	{
		handlingCode = HandlingCodeEnum.kEventNotHandled;

		try
		{
			if (!_isIAssembly || beforeOrAfter != EventTimingEnum.kAfter) return;

			var newMemberName = GetCurrentMemberName();
			if (newMemberName == null || newMemberName == _currentMemberName) return;

			_currentMemberName = newMemberName;
			_ = Dispatcher.InvokeAsync(async () =>
			{
				await LoadMembers();
				await LoadAssemblyChildren();
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: OnDocumentChange error: {ex.Message}");
		}
	}

	private static string GetCurrentMemberName()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is AssemblyDocument asmDoc)
			{
				var compDef = asmDoc.ComponentDefinition;
				if (compDef.iAssemblyFactory != null)
					return compDef.iAssemblyFactory.DefaultRow.MemberName;
				if (compDef.IsModelStateFactory)
					return compDef.ModelStates.ActiveModelState?.Name;
			}
		}
		catch
		{
			// ignore
		}

		return null;
	}

	private static string GetActiveMemberPartNumber(AssemblyDocument assemblyDoc)
	{
		var compDef = assemblyDoc.ComponentDefinition;
		if (compDef.iAssemblyFactory != null)
			return compDef.iAssemblyFactory.DefaultRow[ColumnNames.PartNumber].Value;
		return compDef.ModelStates?.Count > 1 ? compDef.ModelStates.ActiveModelState?.Name : null;
	}

	private void PopulateDataGrids(DataTable comparisonTable)
	{
		var partNumber = GetPartNumberFromTable(comparisonTable);
		_pendingChanges.TryGetValue(partNumber, out var pendingValues);

		var sqlRows = comparisonTable.Rows.Cast<DataRow>().Select(row =>
		{
			var invVal = row["Inventor Value"].ToString();
			var sqlVal = row["SQL Value"].ToString();
			var prop   = row["Property"].ToString();

			if (prop != null && pendingValues?.TryGetValue(prop, out var pending) == true &&
			    !string.IsNullOrEmpty(pending))
				invVal = pending;

			var areEqual = GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
			return new PropertyRow
			{
				Property   = prop, ["SQL Value"]                            = sqlVal, ["Inventor Value"] = invVal,
				["Status"] = areEqual ? "Match" : "Mismatch", HasDifference = !areEqual
			};
		}).ToList();

		var invRows = comparisonTable.Rows.Cast<DataRow>().Select(row =>
		{
			var invVal = row["Inventor Value"].ToString();
			var sqlVal = row["SQL Value"].ToString();
			var prop   = row["Property"].ToString();

			if (prop != null && pendingValues?.TryGetValue(prop, out var pending) == true &&
			    !string.IsNullOrEmpty(pending))
				invVal = pending;

			var areEqual = GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
			return new PropertyRow
			{
				Property   = prop, ["Inventor Value"]                       = invVal, ["SQL Value"] = sqlVal,
				["Status"] = areEqual ? "Match" : "Mismatch", HasDifference = !areEqual
			};
		}).ToList();

		if (pendingValues != null)
		{
			var existingInInv = new HashSet<string>(invRows.Select(propertyRow => propertyRow.Property),
				StringComparer.OrdinalIgnoreCase);
			foreach (var (prop, val) in pendingValues)
			{
				if (existingInInv.Contains(prop)) continue;
				var areEqual = GeniusFormsHelper.ValuesAreEqual(val, "");
				invRows.Add(new PropertyRow
				{
					Property   = prop, ["Inventor Value"]                       = val, ["SQL Value"] = "",
					["Status"] = areEqual ? "Match" : "Mismatch", HasDifference = !areEqual
				});
				sqlRows.Add(new PropertyRow
				{
					Property   = prop, ["SQL Value"]                            = "", ["Inventor Value"] = val,
					["Status"] = areEqual ? "Match" : "Mismatch", HasDifference = !areEqual
				});
			}
		}

		sqlRows.Sort((propertyRow, row) =>
			string.Compare(propertyRow.Property, row.Property, StringComparison.Ordinal));
		invRows.Sort((propertyRow, row) =>
			string.Compare(propertyRow.Property, row.Property, StringComparison.Ordinal));

		SqlDataGrid.ItemsSource      = sqlRows;
		InventorDataGrid.ItemsSource = invRows;
		SqlDataGrid.Items.Refresh();
		InventorDataGrid.Items.Refresh();
	}

	private static string GetPartNumberFromTable(DataTable table)
	{
		var row = table.Rows.Cast<DataRow>().FirstOrDefault(dataRow =>
			string.Equals(dataRow["Property"].ToString(), "Part Number", StringComparison.OrdinalIgnoreCase));
		return row?["Inventor Value"].ToString() ?? "";
	}

	private static List<ChildInfo> BuildTree(List<ChildInfo> flatList)
	{
		var roots = new List<ChildInfo>();
		var stack = new List<ChildInfo>();

		foreach (var item in flatList)
		{
			while (stack.Count > 0 && stack[^1].Level >= item.Level)
				stack.RemoveAt(stack.Count - 1);

			if (stack.Count > 0)
				stack[^1].Children.Add(item);
			else
				roots.Add(item);

			stack.Add(item);
		}

		return roots;
	}

	private static IEnumerable<ChildInfo> FlattenTree(IEnumerable<ChildInfo> items)
	{
		foreach (var item in items)
		{
			yield return item;
			foreach (var child in FlattenTree(item.Children))
				yield return child;
		}
	}

	private void MarkPendingOnItems(List<ChildInfo> tree, List<ChildInfo> parts, List<ChildInfo> purchased)
	{
		foreach (var item in FlattenTree(tree))
			if (_pendingChanges.ContainsKey(item.PartNumber))
				item.HasDifference = true;
		foreach (var item in parts.Where(item => _pendingChanges.ContainsKey(item.PartNumber)))
			item.HasDifference = true;
		foreach (var item in purchased.Where(item => _pendingChanges.ContainsKey(item.PartNumber)))
			item.HasDifference = true;
	}

	private void UpdatePendingHighlights()
	{
		if (AssembliesTreeView.ItemsSource is IEnumerable<ChildInfo> tree)
			foreach (var item in FlattenTree(tree))
				if (_pendingChanges.ContainsKey(item.PartNumber))
					item.HasDifference = true;
		if (PurchasedDataGrid.ItemsSource is not IEnumerable<ChildInfo> purchased) return;
		{
			foreach (var item in purchased)
				if (_pendingChanges.ContainsKey(item.PartNumber))
					item.HasDifference = true;
		}
	}

	private void UpdateMemberHighlights()
	{
		if (Members.ItemsSource is not IEnumerable<MemberInfo> members) return;
		foreach (var member in members)
			member.HasDifference = _pendingChanges.ContainsKey(member.PartNumber);
	}

	private void UpdateSaveCancelButtons()
	{
		var hasPending = _pendingChanges.Values.Any(d => d.Count > 0);
		SaveButton.IsEnabled   = hasPending;
		CancelButton.IsEnabled = hasPending;
		if (!_isIAssembly) return;
		UpdateMemberHighlights();
	}

	private void SetButtonsEnabled(bool enabled)
	{
		CalculatePropsButton.IsEnabled    = enabled;
		CalculateAllPartsButton.IsEnabled = enabled;
		if (!enabled) return;
		UpdateSaveCancelButtons();
		StopButton.IsEnabled = false;
	}

	private async void CalculatePropsButton_Click(object _sender, RoutedEventArgs _e)
	{
		SetButtonsEnabled(false);
		StopButton.IsEnabled = true;
		_calcCts             = new CancellationTokenSource();
		try
		{
			var target = _currentTargetDocument ?? ThisApplication.ActiveDocument;
			if (target == null) return;

			StatusText.Text = $"Calculating properties for {target.DisplayName}...";
			var calculated = await CalculateProps.CalculateAllPropertiesAsync(target, _calcCts.Token);
			if (calculated.Count <= 0)
			{
				StatusText.Text = "No properties to update.";
				return;
			}

			var current    = PropertyExtractor.GetPropertiesFromDocumentStatic(target);
			var updates    = new Dictionary<string, string>();
			var partNumber = current.GetValueOrDefault("Part Number", target.DisplayName);

			foreach (var calc in calculated)
			{
				var key = calc.Key == "Mass" ? "GeniusMass" : calc.Key;
				if (!GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(key, ""), calc.Value))
					updates[key] = calc.Value;
			}

			if (updates.Count <= 0)
			{
				StatusText.Text = "Properties already up to date.";
				return;
			}

			if (!_pendingChanges.TryGetValue(partNumber, out var partPending))
			{
				partPending                 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
				_pendingChanges[partNumber] = partPending;
			}

			foreach (var kvp in updates)
				partPending[kvp.Key] = kvp.Value;

			UpdateSaveCancelButtons();
			StatusText.Text = $"{updates.Count} property(ies) pending for {target.DisplayName}";

			if (_isIAssembly)
				await RefreshDisplay(partNumber, target);
			else
				await RefreshData(target);
		}
		catch (OperationCanceledException)
		{
			StatusText.Text = "Calculation cancelled.";
			Debug.WriteLine("GeniusAssembly: Calculation cancelled");
		}
		catch (Exception ex)
		{
			StatusText.Text = $"Error: {ex.Message}";
			Debug.WriteLine($"GeniusAssembly: CalcProps error: {ex.Message}");
		}
		finally
		{
			StopButton.IsEnabled = false;
			_calcCts?.Dispose();
			_calcCts = null;
			SetButtonsEnabled(true);
		}
	}

	private async void CalculateAllPartsButton_Click(object _sender, RoutedEventArgs _e)
	{
		SetButtonsEnabled(false);
		StopButton.IsEnabled = true;
		_calcCts             = new CancellationTokenSource();
		try
		{
			if (_isIAssembly)
			{
				await CalculateAllMembersAsync(_calcCts.Token);
				_calcCts.Token.ThrowIfCancellationRequested();
				await CalculateAllChildrenAsync(_calcCts.Token);
				UpdateSaveCancelButtons();
				if (_currentTargetDocument != null)
					await LoadAndDisplayProperties(GetCurrentPartNumber(), _currentTargetDocument);
			}
			else
			{
				var assemblyItems = AssembliesTreeView.ItemsSource as IEnumerable<ChildInfo> ?? [];
				var children      = FlattenTree(assemblyItems).ToList();

				if (children.Count == 0)
				{
					StatusText.Text = "No parts to calculate.";
					return;
				}

				var pending = 0;
				var failed  = 0;
				var skipped = 0;

				for (var index = 0; index < children.Count; index++)
				{
					_calcCts.Token.ThrowIfCancellationRequested();

					var child = children[index];
					StatusText.Text = $"[{index + 1}/{children.Count}] Calculating {child.PartNumber}...";

					await Task.Delay(1);

					var doc = PropertyExtractor.FindDocumentByPartNumber(child.PartNumber);
					if (doc == null)
					{
						skipped++;
						Debug.WriteLine($"GeniusAssembly: CalcAll skipped '{child.PartNumber}' (document not open)");
						continue;
					}

					try
					{
						var calculated = await CalculateProps.CalculateAllPropertiesAsync(doc, _calcCts.Token);
						if (calculated.Count == 0)
						{
							skipped++;
							continue;
						}

						var current = PropertyExtractor.GetPropertiesFromDocumentStatic(doc);
						var updates = new Dictionary<string, string>();

						foreach (var calc in calculated)
						{
							var key = calc.Key == "Mass" ? "GeniusMass" : calc.Key;
							if (!GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(key, ""), calc.Value))
								updates[key] = calc.Value;
						}

						if (updates.Count <= 0) continue;

						if (!_pendingChanges.TryGetValue(child.PartNumber, out var partPending))
						{
							partPending = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
							_pendingChanges[child.PartNumber] = partPending;
						}

						foreach (var kvp in updates)
							partPending[kvp.Key] = kvp.Value;

						pending++;
					}
					catch (OperationCanceledException)
					{
						throw;
					}
					catch (Exception ex)
					{
						failed++;
						Debug.WriteLine($"GeniusAssembly: CalcAll error on '{child.PartNumber}': {ex.Message}");
					}
				}

				UpdateSaveCancelButtons();
				StatusText.Text = $"Done — {pending} pending, {skipped} skipped, {failed} failed";
				await RefreshData(_currentTargetDocument ?? ThisApplication.ActiveDocument);
			}
		}
		catch (OperationCanceledException)
		{
			StatusText.Text = "Calculation cancelled.";
			Debug.WriteLine("GeniusAssembly: CalculateAllParts cancelled");
		}
		catch (Exception ex)
		{
			StatusText.Text = $"Error: {ex.Message}";
			Debug.WriteLine($"GeniusAssembly: CalculateAllParts error: {ex.Message}");
		}
		finally
		{
			StopButton.IsEnabled = false;
			_calcCts?.Dispose();
			_calcCts = null;
			SetButtonsEnabled(true);
		}
	}

	private async void SaveButton_Click(object _sender, RoutedEventArgs _e)
	{
		SetButtonsEnabled(false);
		try
		{
			var              totalSaved = 0;
			AssemblyDocument asmDoc     = null;
			var factory = _isIAssembly && (asmDoc = ThisApplication?.ActiveDocument as AssemblyDocument) != null
				? asmDoc.ComponentDefinition.iAssemblyFactory
				: null;
			object originalRow               = null;
			if (factory != null) originalRow = factory.DefaultRow;

			foreach (var (partNumber, properties) in _pendingChanges)
			{
				if (properties.Count == 0) continue;
				var doc = PropertyExtractor.FindDocumentByPartNumber(partNumber);
				if (doc == null && _isIAssembly && _memberPartNumbers.Contains(partNumber))
				{
					var row = factory?.TableRows.Cast<iAssemblyTableRow>().FirstOrDefault(r =>
						string.Equals(r[ColumnNames.PartNumber].Value, partNumber,
							StringComparison.OrdinalIgnoreCase));
					if (row != null && factory.DefaultRow.MemberName != row.MemberName)
					{
						factory.DefaultRow = row;
						asmDoc.Update();
					}

					doc = ThisApplication?.ActiveDocument;
					Debug.WriteLine($"GeniusAssembly: Save using assembly doc for member '{partNumber}'");
				}

				if (doc == null)
				{
					Debug.WriteLine($"GeniusAssembly: Save skipped '{partNumber}' (document not found)");
					continue;
				}

				GeniusFormsHelper.UpdateInventorProperties(doc, properties, "GeniusAssembly");
				totalSaved += properties.Count;
			}

			if (originalRow != null)
				if (((iAssemblyTableRow)originalRow).MemberName != factory.DefaultRow.MemberName)
				{
					factory.DefaultRow = (iAssemblyTableRow)originalRow;
					asmDoc.Update();
				}

			_pendingChanges.Clear();
			UpdateSaveCancelButtons();
			StatusText.Text = $"{totalSaved} property(ies) saved.";

			if (_isIAssembly)
			{
				if (_currentTargetDocument != null)
					await LoadAndDisplayProperties(GetCurrentPartNumber(), _currentTargetDocument);
			}
			else
			{
				await RefreshData(_currentTargetDocument ?? ThisApplication?.ActiveDocument, true);
			}
		}
		catch (Exception ex)
		{
			StatusText.Text = $"Save error: {ex.Message}";
			Debug.WriteLine($"GeniusAssembly: Save error: {ex.Message}");
		}
		finally
		{
			SetButtonsEnabled(true);
		}
	}

	private async void CancelButton_Click(object _sender, RoutedEventArgs _e)
	{
		_pendingChanges.Clear();
		UpdateSaveCancelButtons();
		StatusText.Text = "Pending changes cancelled.";

		if (_isIAssembly)
		{
			if (_currentTargetDocument != null)
				await LoadAndDisplayProperties(GetCurrentPartNumber(), _currentTargetDocument);
		}
		else
		{
			await RefreshData(_currentTargetDocument ?? ThisApplication.ActiveDocument, true);
		}
	}

	private void StopButton_Click(object _sender, RoutedEventArgs _e)
	{
		_calcCts?.Cancel();
	}

	private async void AssembliesTreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
	{
		if (!AssembliesTab.IsSelected) return;
		await HandleChildSelectionChanged(e.NewValue as ChildInfo);
	}

	private async void PurchasedDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (_isIAssembly)
		{
			if (PurchasedDataGrid.SelectedItem is not ChildInfo selectedChild)
			{
				_currentTargetDocument = null;
				return;
			}

			_currentTargetDocument = PropertyExtractor.FindDocumentByPartNumber(selectedChild.PartNumber);
			await LoadAndDisplayProperties(selectedChild.PartNumber, _currentTargetDocument);
			if (_currentTargetDocument != null) UpdateThumbnail(_currentTargetDocument);

			GeniusFormsHelper.ZoomToOccurrence(_currentTargetDocument);
		}
		else
		{
			await HandleChildSelectionChanged(PurchasedDataGrid.SelectedItem as ChildInfo);
		}
	}

	private async void InventorDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
	{
		if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not PropertyRow invRow) return;

		var propertyName = invRow.Property;
		var newValue     = e.EditingElement is TextBox tb ? tb.Text : "";
		var partNumber   = GetCurrentPartNumber();

		if (string.IsNullOrEmpty(partNumber)) return;

		if (!_pendingChanges.TryGetValue(partNumber, out var partPending))
		{
			partPending                 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			_pendingChanges[partNumber] = partPending;
		}

		if (string.IsNullOrEmpty(newValue))
			partPending.Remove(propertyName);
		else
			partPending[propertyName] = newValue;

		UpdateSaveCancelButtons();

		if (_isIAssembly)
			await RefreshDisplay(partNumber, _currentTargetDocument);
		else
			await RefreshData(_currentTargetDocument);
	}

	private async void MembersDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (Members.SelectedItem is not MemberInfo selectedMember ||
		    ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc) return;

		try
		{
			var compDef          = assemblyDoc.ComponentDefinition;
			var activePartNumber = GetActiveMemberPartNumber(assemblyDoc);
			var isAlreadyActive = string.Equals(selectedMember.PartNumber, activePartNumber,
				StringComparison.OrdinalIgnoreCase);

			Document memberDoc = null;

			if (compDef.iAssemblyFactory != null)
			{
				var row = compDef.iAssemblyFactory.TableRows.Cast<iAssemblyTableRow>().FirstOrDefault(r =>
					string.Equals(r[ColumnNames.PartNumber].Value, selectedMember.PartNumber,
						StringComparison.OrdinalIgnoreCase));
				if (row != null)
				{
					if (!isAlreadyActive && compDef.iAssemblyFactory.DefaultRow.MemberName != row.MemberName)
					{
						compDef.iAssemblyFactory.DefaultRow = row;
						assemblyDoc.Update();
					}

					memberDoc = (Document)assemblyDoc;
				}
			}
			else if (compDef.ModelStates?.Count > 1)
			{
				var state = compDef.ModelStates.Cast<ModelState>()
				                   .FirstOrDefault(s => s.Name == selectedMember.PartNumber);
				if (state != null)
				{
					if (!isAlreadyActive)
					{
						state.Activate();
						memberDoc = ThisApplication.ActiveDocument;
					}
					else
					{
						memberDoc = (Document)assemblyDoc;
					}
				}
			}

			_currentTargetDocument = memberDoc;
			await LoadAndDisplayProperties(selectedMember.PartNumber, memberDoc);
			if (memberDoc != null)
			{
				UpdateThumbnail(memberDoc);
				var camera = ThisApplication.ActiveView.Camera;
				camera.ViewOrientationType = ViewOrientationTypeEnum.kIsoTopRightViewOrientation;
				camera.Fit();
				camera.Apply();
			}

			if (!isAlreadyActive)
				await LoadAssemblyChildren();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: Error in member selection changed: {ex.Message}");
		}
	}

	private async Task HandleChildSelectionChanged(ChildInfo selectedChild)
	{
		if (selectedChild == null) return;
		await UpdateSelectedChildData(selectedChild);
		GeniusFormsHelper.ZoomToOccurrence(_currentTargetDocument);
	}

	private async Task UpdateSelectedChildData(ChildInfo selectedChild)
	{
		try
		{
			var target = ThisApplication.Documents.Cast<Document>().FirstOrDefault(doc =>
				string.Equals(doc.FullFileName, selectedChild.FullPath, StringComparison.OrdinalIgnoreCase));

			target ??= ThisApplication.Documents.Cast<Document>().FirstOrDefault(doc =>
				string.Equals(doc.DisplayName, selectedChild.PartNumber, StringComparison.OrdinalIgnoreCase) ||
				string.Equals(GetFileNameWithoutExtension(doc.FullFileName), selectedChild.PartNumber,
					StringComparison.OrdinalIgnoreCase));

			_currentTargetDocument = target;

			if (target != null)
			{
				UpdateThumbnail(target);
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

	private void UpdateThumbnail(Document document)
	{
		try
		{
			if (document == null)
			{
				image.Source = null;
				return;
			}

			image.Source = ThumbnailHelper.GetThumbnail(document);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: UpdateThumbnail error: {ex.Message}");
			image.Source = null;
		}
	}

	private async Task LoadAndDisplayProperties(string partNumber, Document document)
	{
		var (geniusRows, invRows) = await GeniusFormsHelper.LoadPropertiesForSelectedPart(
			partNumber, _propertyComparator.SqlDataManager, document);
		ApplyPendingOverlay(partNumber, geniusRows, invRows, _pendingChanges);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
	}

	private async Task RefreshDisplay(string partNumber, Document doc)
	{
		var (geniusRows, invRows) = await GeniusFormsHelper.LoadPropertiesForSelectedPart(
			partNumber, _propertyComparator.SqlDataManager, doc);
		ApplyPendingOverlay(partNumber, geniusRows, invRows, _pendingChanges);
		await Dispatcher.InvokeAsync(() =>
		{
			SqlDataGrid.ItemsSource      = geniusRows;
			InventorDataGrid.ItemsSource = invRows;
		});
	}

	private static void ApplyPendingOverlay(string partNumber, List<PropertyRow> geniusRows,
		List<PropertyRow> invRows, Dictionary<string, Dictionary<string, string>> pendingChanges = null)
	{
		if (pendingChanges == null || !pendingChanges.TryGetValue(partNumber, out var pendingValues))
			return;

		foreach (var row in invRows)
			if (pendingValues.TryGetValue(row.Property, out var pendingVal))
			{
				row["Inventor Value"] = pendingVal;
				var sqlVal = row.TryGetValue("SQL Value", out var sql) ? sql?.ToString() : "";
				row.HasDifference = !GeniusFormsHelper.ValuesAreEqual(sqlVal, pendingVal);
			}

		foreach (var row in geniusRows)
			if (pendingValues.TryGetValue(row.Property, out var pendingVal))
			{
				var sqlVal = row.TryGetValue("SQL Value", out var sql) ? sql?.ToString() : "";
				row.HasDifference = !GeniusFormsHelper.ValuesAreEqual(sqlVal, pendingVal);
			}

		var existingInGenius =
			new HashSet<string>(geniusRows.Select(r => r.Property), StringComparer.OrdinalIgnoreCase);
		foreach (var (prop, val) in pendingValues)
		{
			if (existingInGenius.Contains(prop)) continue;
			var areEqual = GeniusFormsHelper.ValuesAreEqual(val, "");
			invRows.Add(new PropertyRow
			{
				Property      = prop, ["Inventor Value"] = val, ["SQL Value"] = "",
				HasDifference = !areEqual
			});
			geniusRows.Add(new PropertyRow
			{
				Property      = prop, ["SQL Value"] = "", ["Inventor Value"] = val,
				HasDifference = !areEqual
			});
		}
	}

	private async Task CalculateAllMembersAsync(CancellationToken cancellationToken = default)
	{
		if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc) return;
		var compDef = assemblyDoc.ComponentDefinition;

		List<(string PartNumber, string Description)> members = null;

		if (compDef.iAssemblyFactory != null)
			members =
			[
				.. compDef.iAssemblyFactory.TableRows.Cast<iAssemblyTableRow>().Select(row =>
					(row[ColumnNames.PartNumber].Value, row[ColumnNames.Description].Value))
			];
		else if (compDef.ModelStates?.Count > 1)
			members =
			[
				.. compDef.ModelStates.ModelStateTable.TableRows.Cast<ModelStateTableRow>().Select(row =>
					(row[ColumnNames.PartNumber].Value, row[ColumnNames.Description].Value))
			];

		if (members == null || members.Count == 0) return;

		var app          = ThisApplication;
		var screenUpdate = app.ScreenUpdating;
		var pending      = 0;
		var failed       = 0;

		try
		{
			app.ScreenUpdating = false;

			var    factory                   = compDef.iAssemblyFactory;
			object originalRow               = null;
			if (factory != null) originalRow = factory.DefaultRow;

			for (var index = 0; index < members.Count; index++)
			{
				cancellationToken.ThrowIfCancellationRequested();

				var (partNumber, _) = members[index];
				StatusText.Text     = $"[{index + 1}/{members.Count}] Calculating member {partNumber}...";
				await Task.Delay(1, cancellationToken);

				try
				{
					if (factory != null)
					{
						var row = factory.TableRows.Cast<iAssemblyTableRow>().FirstOrDefault(r =>
							string.Equals(r[ColumnNames.PartNumber].Value, partNumber,
								StringComparison.OrdinalIgnoreCase));
						if (row == null) continue;

						if (factory.DefaultRow.MemberName != row.MemberName)
						{
							factory.DefaultRow = row;
							assemblyDoc.Update();
						}
					}
					else
					{
						var state = compDef.ModelStates.Cast<ModelState>()
						                   .FirstOrDefault(modelState => modelState.Name == partNumber);
						if (state == null) continue;
						state.Activate();
					}

					var doc = factory != null
						? (Document)assemblyDoc
						: app.ActiveDocument;

					var calculated = await CalculateProps.CalculateAllPropertiesAsync(doc, cancellationToken);
					if (calculated.Count == 0) continue;

					var current = PropertyExtractor.GetPropertiesFromDocumentStatic(doc);
					var updates = new Dictionary<string, string>();
					foreach (var calc in calculated)
					{
						var key = calc.Key == "Mass" ? "GeniusMass" : calc.Key;
						if (!GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(key, ""), calc.Value))
							updates[key] = calc.Value;
					}

					if (updates.Count <= 0) continue;

					if (!_pendingChanges.TryGetValue(partNumber, out var partPending))
					{
						partPending                 = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
						_pendingChanges[partNumber] = partPending;
					}

					foreach (var kvp in updates)
						partPending[kvp.Key] = kvp.Value;

					pending++;
				}
				catch (OperationCanceledException)
				{
					throw;
				}
				catch (Exception ex)
				{
					failed++;
					Debug.WriteLine($"GeniusAssembly: CalcAll member '{partNumber}' error: {ex.Message}");
				}
			}

			if (factory != null && originalRow != null)
				if (((iAssemblyTableRow)originalRow).MemberName != factory.DefaultRow.MemberName)
				{
					factory.DefaultRow = (iAssemblyTableRow)originalRow;
					assemblyDoc.Update();
				}
		}
		finally
		{
			app.ScreenUpdating = screenUpdate;
		}

		StatusText.Text = $"Members: {pending} pending, {failed} failed";
		await LoadAssemblyChildren();
	}

	private async Task CalculateAllChildrenAsync(CancellationToken cancellationToken = default)
	{
		var treeItems = AssembliesTreeView.ItemsSource as IEnumerable<ChildInfo> ?? [];
		var children  = FlattenTree(treeItems).ToList();

		if (children.Count == 0) return;

		var pending = 0;
		var failed  = 0;
		var skipped = 0;

		for (var index = 0; index < children.Count; index++)
		{
			cancellationToken.ThrowIfCancellationRequested();

			var child = children[index];
			StatusText.Text = $"[{index + 1}/{children.Count}] Calculating child {child.PartNumber}...";
			await Task.Delay(1, cancellationToken);

			var doc = PropertyExtractor.FindDocumentByPartNumber(child.PartNumber);
			if (doc == null)
			{
				skipped++;
				continue;
			}

			try
			{
				var calculated = await CalculateProps.CalculateAllPropertiesAsync(doc, cancellationToken);
				if (calculated.Count == 0)
				{
					skipped++;
					continue;
				}

				var current = PropertyExtractor.GetPropertiesFromDocumentStatic(doc);
				var updates = new Dictionary<string, string>();
				foreach (var calc in calculated)
				{
					var key = calc.Key == "Mass" ? "GeniusMass" : calc.Key;
					if (!GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(key, ""), calc.Value))
						updates[key] = calc.Value;
				}

				if (updates.Count <= 0) continue;

				if (!_pendingChanges.TryGetValue(child.PartNumber, out var partPending))
				{
					partPending = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
					_pendingChanges[child.PartNumber] = partPending;
				}

				foreach (var kvp in updates)
					partPending[kvp.Key] = kvp.Value;

				pending++;
			}
			catch (OperationCanceledException)
			{
				throw;
			}
			catch (Exception ex)
			{
				failed++;
				Debug.WriteLine($"GeniusAssembly: CalcAll child '{child.PartNumber}' error: {ex.Message}");
			}
		}

		StatusText.Text = StatusText.Text.Contains("Members:")
			? $"{StatusText.Text} | Children: {pending} pending, {skipped} skipped, {failed} failed"
			: $"Children: {pending} pending, {skipped} skipped, {failed} failed";
	}

	private string GetCurrentPartNumber()
	{
		if (_currentTargetDocument == null) return "";
		var props = PropertyExtractor.GetPropertiesFromDocumentStatic(_currentTargetDocument);
		return props.GetValueOrDefault("Part Number", _currentTargetDocument.DisplayName);
	}

	public void SelectPartByPartNumber(string partNumber)
	{
		try
		{
			if (string.IsNullOrWhiteSpace(partNumber)) return;

			Dispatcher.Invoke(() =>
			{
				// Check Members tab (iAssemblies)
				if (Members.ItemsSource is IEnumerable<MemberInfo> members)
				{
					var member = members.FirstOrDefault(m =>
						string.Equals(m.PartNumber, partNumber, StringComparison.OrdinalIgnoreCase));
					if (member != null)
					{
						Members.SelectedItem = member;
						return;
					}
				}

				// Check PurchasedDataGrid
				if (PurchasedDataGrid.ItemsSource is IEnumerable<ChildInfo> purchased)
				{
					var child = purchased.FirstOrDefault(c =>
						string.Equals(c.PartNumber, partNumber, StringComparison.OrdinalIgnoreCase));
					if (child != null)
					{
						PurchasedDataGrid.SelectedItem = child;
						return;
					}
				}

				// Check AssembliesTreeView
				if (AssembliesTreeView.ItemsSource is not IEnumerable<ChildInfo> treeItems) return;
				{
					var flatItems = FlattenTree(treeItems);
					var child = flatItems.FirstOrDefault(c =>
						string.Equals(c.PartNumber, partNumber, StringComparison.OrdinalIgnoreCase));
					if (child != null)
						SetSelectedTreeViewItem(AssembliesTreeView, child);
				}
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusAssembly: SelectPartByPartNumber error: {ex.Message}");
		}
	}

	private static void SetSelectedTreeViewItem(TreeView treeView, object item)
	{
		var tvi = FindTreeViewItem(treeView, item);
		if (tvi == null) return;
		tvi.IsSelected = true;
		tvi.BringIntoView();
		tvi.Focus();
	}

	private static TreeViewItem FindTreeViewItem(ItemsControl container, object item)
	{
		if (container?.ItemContainerGenerator.Status != GeneratorStatus.ContainersGenerated)
			return null;

		if (container.ItemContainerGenerator.ContainerFromItem(item) is TreeViewItem tvi) return tvi;

		foreach (var child in container.Items)
		{
			if (container.ItemContainerGenerator.ContainerFromItem(child) is not TreeViewItem childContainer)
				continue;
			var result = FindTreeViewItem(childContainer, item);
			if (result != null) return result;
		}

		return null;
	}

	private static class ColumnNames
	{
		public const string PartNumber = "Part Number [Project]";
		public const string Description = "Description [Project]";
	}
}