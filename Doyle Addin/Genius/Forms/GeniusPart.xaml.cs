namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using TextBox = System.Windows.Controls.TextBox;
using ThemeManager = Options.Themes.ThemeManager;

public partial class GeniusPart
{
	private readonly bool _isIPart;
	private readonly Dictionary<string, string> _pendingProperties = [];
	private readonly PropertyComparator _propertyComparator;
	private CancellationTokenSource _calcCts;
	private PartDocument _factoryDocument;

	public GeniusPart()
	{
		try
		{
			var sqlDataManager = new SqlDataManager(GeniusConstants.DefaultConnectionString);
			_propertyComparator = new PropertyComparator(sqlDataManager);
			CalculateProps.SetSqlDataManager(sqlDataManager);

			InitializeComponent();
			ThemeManager.ApplyTheme(this);

			if (ThisApplication?.ActiveDocument is PartDocument partDoc)
			{
				var compDef = partDoc.ComponentDefinition;
				if (compDef.iPartFactory != null || compDef.IsModelStateFactory)
				{
					compDef.iPartFactory?.MemberEditScope = MemberEditScopeEnum.kEditActiveMember;
					if (compDef.IsModelStateFactory)
						compDef.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditActiveMember;

					_isIPart = true;
					SetupIPartUI();
					_ = LoadMembers();
					return;
				}
			}

			_isIPart = false;
			SetupRegularUI();
			_ = RefreshData();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusPart: Exception in constructor: {ex.Message}");
		}
	}

	private void SetupIPartUI()
	{
		VisualStateManager.GoToState(this, "iParts", true);
	}

	private void SetupRegularUI()
	{
		VisualStateManager.GoToState(this, "Regular_Parts", true);
	}

	private async void CalculatePropsButton_Click(object sender, RoutedEventArgs e)
	{
		if (_isIPart)
		{
			if (ThisApplication?.ActiveDocument == null) return;
			_calcCts                   = new CancellationTokenSource();
			IPart_StopButton.IsEnabled = true;
			try
			{
				await CalculatePropertiesForDocument(ThisApplication.ActiveDocument, true, _calcCts.Token);
			}
			catch (OperationCanceledException)
			{
				Debug.WriteLine("GeniusPart: Calculation cancelled");
			}
			finally
			{
				IPart_StopButton.IsEnabled = false;
				_calcCts?.Dispose();
				_calcCts = null;
			}
		}
		else
		{
			await CalculatePropertiesPreview();
		}
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
		UpdateMemberHighlights();
		_ = RefreshData();
	}

	private void StopButton_Click(object _sender, RoutedEventArgs _e)
	{
		_calcCts?.Cancel();
	}

	private void UpdateMemberHighlights()
	{
		if (Members.ItemsSource is not IEnumerable<MemberInfo> members) return;
		var    compDef          = _factoryDocument?.ComponentDefinition;
		string activeMemberName = null;
		if (compDef?.iPartFactory != null)
			activeMemberName = compDef.iPartFactory.DefaultRow?["Part Number [Project]"]?.Value;
		else if (compDef is { IsModelStateFactory: true })
			activeMemberName = compDef.ModelStates?.ActiveModelState?.Name;
		foreach (var member in members)
			member.HasDifference = _pendingProperties.Count > 0 && member.PartNumber == activeMemberName;
	}

	private static bool IsSheetMetalPart()
	{
		return ThisApplication?.ActiveDocument is PartDocument { ComponentDefinition: SheetMetalComponentDefinition };
	}

	private async Task RefreshData()
	{
		try
		{
			if (ThisApplication?.ActiveDocument == null) return;

			var inventorProperties = PropertyExtractor.GetAllProperties();
			var partNumber         = inventorProperties.GetValueOrDefault("Part Number", "");
			var sqlData            = await _propertyComparator.SqlDataManager.GetSqlDataAsync(partNumber);

			var isSheetMetal = IsSheetMetalPart();
			var sheetMetalPropertyNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
			{
				"Thickness", "Extent_Width", "Extent_Length", "Extent_Area"
			};

			var geniusProperties = sqlData
			                       .Select(kvp => new
			                       {
				                       SqlKey   = kvp.Key,
				                       InvName  = GeniusFormsHelper.MapSqlColumnToInventorProperty(kvp.Key),
				                       SqlValue = kvp.Value
			                       })
			                       .Where(x => isSheetMetal || !sheetMetalPropertyNames.Contains(x.InvName))
			                       .Select(x =>
			                       {
				                       var displayValue = _pendingProperties.GetValueOrDefault(x.InvName) ??
				                                          inventorProperties.GetValueOrDefault(x.InvName, "");
				                       return new PropertyRow
				                       {
					                       Property      = x.InvName,
					                       ["SQL Value"] = x.SqlValue,
					                       HasDifference = !GeniusFormsHelper.ValuesAreEqual(x.SqlValue, displayValue)
				                       };
			                       })
			                       .ToList();

			if (geniusProperties.Count == 0)
				geniusProperties.Add(new PropertyRow { Property = "Info", ["SQL Value"] = "No data found" });

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
					Property           = pending.Key,
					["Inventor Value"] = pending.Value,
					HasDifference      = !GeniusFormsHelper.ValuesAreEqual(sqlValue, pending.Value)
				});

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
		IPart_StopButton.IsEnabled     = true;
		_calcCts                       = new CancellationTokenSource();
		try
		{
			var calculated = await CalculateProps.CalculateAllPropertiesAsync(_calcCts.Token);
			if (calculated.Count == 0) return;

			var current = PropertyExtractor.GetAllProperties();
			_pendingProperties.Clear();

			foreach (var calc in calculated.Where(calc =>
				         !GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(calc.Key, ""), calc.Value)))
				_pendingProperties[calc.Key] = calc.Value;

			await RefreshData();
			UpdateMemberHighlights();
			SaveButton.IsEnabled   = _pendingProperties.Count > 0;
			CancelButton.IsEnabled = _pendingProperties.Count > 0;
		}
		catch (OperationCanceledException)
		{
			Debug.WriteLine("GeniusPart: Calculation cancelled");
		}
		finally
		{
			CalculatePropsButton.IsEnabled = true;
			IPart_StopButton.IsEnabled     = false;
			_calcCts?.Dispose();
			_calcCts = null;
		}
	}

	private void SavePendingProperties()
	{
		if (_pendingProperties.Count == 0) return;
		GeniusFormsHelper.UpdateInventorProperties(ThisApplication.ActiveDocument, _pendingProperties, "GeniusPart");
		_pendingProperties.Clear();
		SaveButton.IsEnabled   = false;
		CancelButton.IsEnabled = false;
		UpdateMemberHighlights();
	}

	private async void InventorDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
	{
		if (_isIPart) return;
		if (e.EditAction != DataGridEditAction.Commit || e.Row.Item is not PropertyRow rowData) return;

		var propertyName = rowData.Property;
		var newValue     = e.EditingElement is TextBox tb ? tb.Text : "";

		_pendingProperties[propertyName] = newValue;
		SaveButton.IsEnabled             = true;
		CancelButton.IsEnabled           = true;

		await RefreshData();
	}

	private async Task LoadMembers()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is not PartDocument partDoc) return;

			_factoryDocument = partDoc;
			var compDef = partDoc.ComponentDefinition;
			var members = new List<MemberInfo>();

			if (compDef.iPartFactory != null)
				members.AddRange(compDef.iPartFactory.TableRows.Cast<iPartTableRow>()
				                        .Select(row => new MemberInfo
				                        {
					                        PartNumber  = row["Part Number [Project]"].Value,
					                        Description = row["Description [Project]"].Value
				                        }));
			else if (compDef.IsModelStateFactory)
				members.AddRange(compDef.ModelStates.Cast<ModelState>()
				                        .Select(state => new MemberInfo
				                        {
					                        PartNumber  = state.Name,
					                        Description = ""
				                        }));

			await Dispatcher.InvokeAsync(() =>
			{
				Members.ItemsSource = members;
				if (members.Count > 0) Members.SelectedIndex = 0;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusPart: Error loading members: {ex.Message}");
		}
	}

	private async void ChildrenDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
	{
		if (Members.SelectedItem is not MemberInfo selectedMember ||
		    _factoryDocument?.ComponentDefinition is not { } compDef) return;

		try
		{
			Document memberDoc = null;

			if (compDef.iPartFactory != null)
			{
				var row = compDef.iPartFactory.TableRows.Cast<iPartTableRow>()
				                 .FirstOrDefault(r => r["Part Number [Project]"].Value == selectedMember.PartNumber);
				if (row != null)
				{
					compDef.iPartFactory.DefaultRow = row;
					_factoryDocument.Update();
					memberDoc = (Document)_factoryDocument;
				}
			}
			else if (compDef.IsModelStateFactory)
			{
				var state = compDef.ModelStates.Cast<ModelState>()
				                   .FirstOrDefault(s => s.Name == selectedMember.PartNumber);
				if (state != null)
				{
					state.Activate();
					memberDoc = ThisApplication.ActiveDocument;
				}
			}

			var (geniusRows, invRows) = await GeniusFormsHelper.LoadPropertiesForSelectedPart(
				selectedMember.PartNumber, _propertyComparator.SqlDataManager, memberDoc);
			await Dispatcher.InvokeAsync(() =>
			{
				SqlDataGrid.ItemsSource      = geniusRows;
				InventorDataGrid.ItemsSource = invRows;
			});

			UpdateMemberHighlights();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusPart: Error in selection changed: {ex.Message}");
		}
	}

	private async void CalculateEachMember_Click(object sender, RoutedEventArgs e)
	{
		var compDef = _factoryDocument?.ComponentDefinition;
		if (compDef == null) return;

		var app          = ThisApplication;
		var screenUpdate = app.ScreenUpdating;

		_calcCts = new CancellationTokenSource();
		try
		{
			CalculateEachMember.IsEnabled = CalculatePropsButton.IsEnabled = false;
			IPart_StopButton.IsEnabled    = true;
			app.ScreenUpdating            = false;

			if (compDef.iPartFactory != null)
			{
				var factory  = compDef.iPartFactory;
				var original = factory.DefaultRow;

				foreach (iPartTableRow row in factory.TableRows)
				{
					_calcCts.Token.ThrowIfCancellationRequested();

					if (factory.DefaultRow.MemberName != row.MemberName)
					{
						factory.DefaultRow = row;
						_factoryDocument.Update();
					}

					await CalculatePropertiesForDocument((Document)_factoryDocument, false, _calcCts.Token);
					await Task.Delay(1);
				}

				if (factory.DefaultRow.MemberName != original.MemberName)
				{
					factory.DefaultRow = original;
					_factoryDocument.Update();
				}

				await CalculatePropertiesForDocument((Document)_factoryDocument, true, _calcCts.Token);
			}
			else if (compDef.IsModelStateFactory)
			{
				var originalState = compDef.ModelStates.ActiveModelState;

				foreach (ModelState state in compDef.ModelStates)
				{
					_calcCts.Token.ThrowIfCancellationRequested();

					if (compDef.ModelStates.ActiveModelState?.Name != state.Name)
						state.Activate();

					await CalculatePropertiesForDocument(app.ActiveDocument, false, _calcCts.Token);
					await Task.Delay(1);
				}

				if (originalState != null && compDef.ModelStates.ActiveModelState?.Name != originalState.Name)
					originalState.Activate();

				await CalculatePropertiesForDocument(app.ActiveDocument, true, _calcCts.Token);
			}
		}
		catch (OperationCanceledException)
		{
			Debug.WriteLine("GeniusPart: CalculateEachMember cancelled");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusPart: Error in CalculateEachMember: {ex.Message}");
		}
		finally
		{
			app.ScreenUpdating         = screenUpdate;
			IPart_StopButton.IsEnabled = false;
			_calcCts?.Dispose();
			_calcCts                      = null;
			CalculateEachMember.IsEnabled = CalculatePropsButton.IsEnabled = true;
		}
	}

	private async Task CalculatePropertiesForDocument(Document doc, bool updateUI,
		CancellationToken cancellationToken = default)
	{
		if (updateUI) CalculatePropsButton.IsEnabled = false;
		try
		{
			var calculated = await CalculateProps.CalculateAllPropertiesAsync((_Document)doc, cancellationToken);
			if (calculated.Count > 0)
			{
				var current = PropertyExtractor.GetPropertiesFromDocumentStatic(doc);
				var updates = new Dictionary<string, string>();

				foreach (var calc in calculated)
				{
					var lookupKey = calc.Key == "Mass" ? "GeniusMass" : calc.Key;
					if (!GeniusFormsHelper.ValuesAreEqual(current.GetValueOrDefault(lookupKey, ""), calc.Value))
						updates[lookupKey] = calc.Value;
				}

				if (updates.Count > 0)
				{
					_pendingProperties.Clear();
					foreach (var kvp in updates)
						_pendingProperties[kvp.Key] = kvp.Value;

					if (updateUI)
					{
						await RefreshData();
						UpdateMemberHighlights();
						SaveButton.IsEnabled   = true;
						CancelButton.IsEnabled = true;
					}
				}
			}

			await Task.Delay(1, cancellationToken);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusPart: Calculation error: {ex.Message}");
		}
		finally
		{
			if (updateUI) CalculatePropsButton.IsEnabled = true;
		}
	}
}