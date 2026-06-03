namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Inventor;
using ThemeManager = Options.Themes.ThemeManager;

/// <summary>
///     Represents a WPF control specific to the "IPart" type within the Genius application.
/// </summary>
public partial class GeniusiPart
{
	private readonly PropertyComparator _propertyComparator;
	private PartDocument _factoryDocument;

	/// <summary>
	///     Represents the WPF control associated with the "IPart" type for integration within the Genius application.
	///     This class facilitates working with an iPart factory or model state within an active Inventor document.
	/// </summary>
	public GeniusiPart()
	{
		try
		{
			if (ThisApplication?.ActiveDocument is PartDocument partDoc)
			{
				partDoc.ComponentDefinition.iPartFactory?.MemberEditScope = MemberEditScopeEnum.kEditActiveMember;
				if (partDoc.ComponentDefinition.IsModelStateFactory)
					partDoc.ComponentDefinition.ModelStates.MemberEditScope = MemberEditScopeEnum.kEditActiveMember;
			}

			var sqlDataManager = new SqlDataManager(Geniusinfo.DefaultConnectionString);
			_propertyComparator = new PropertyComparator(sqlDataManager);

			InitializeComponent();
			ThemeManager.ApplyTheme(this);

			_ = LoadMembers();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiPart: Exception in constructor: {ex.Message}");
		}
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
			else if (compDef.ModelStates?.Count > 1)
				members.AddRange(compDef.ModelStates.ModelStateTable.TableRows.Cast<ModelStateTableRow>()
				                        .Select(row => new MemberInfo
				                        {
					                        PartNumber  = row["Part Number [Project]"].Value,
					                        Description = row["Description [Project]"].Value
				                        }));

			await Dispatcher.InvokeAsync(() =>
			{
				Members.ItemsSource = members;
				if (members.Count > 0) Members.SelectedIndex = 0;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiPart: Error loading members: {ex.Message}");
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
			else if (compDef.ModelStates?.Count > 1)
			{
				var row = compDef.ModelStates.ModelStateTable.TableRows.Cast<ModelStateTableRow>()
				                 .FirstOrDefault(r => r["Part Number [Project]"].Value == selectedMember.PartNumber);
				if (row != null)
				{
					compDef.ModelStates[row.MemberName]?.Activate();
					memberDoc = ThisApplication.ActiveDocument;
				}
			}

			await LoadPropertiesForPart(selectedMember.PartNumber, memberDoc);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiPart: Error in selection changed: {ex.Message}");
		}
	}

	private async Task LoadPropertiesForPart(string partNumber, Document document = null)
	{
		try
		{
			var sqlData = await _propertyComparator.SqlDataManager.GetSqlDataAsync(partNumber);
			var invProps = document != null
				? PropertyExtractor.GetPropertiesFromDocumentStatic(document)
				: PropertyExtractor.GetAllProperties();

			var geniusRows = sqlData.Select(kvp =>
			{
				var invName = GeniusFormsHelper.MapSqlColumnToInventorProperty(kvp.Key);
				var invVal  = invProps.GetValueOrDefault(invName, "");
				return new PropertyRow
				{
					Property      = invName, ["SQL Value"] = kvp.Value,
					HasDifference = !GeniusFormsHelper.ValuesAreEqual(kvp.Value, invVal)
				};
			}).ToList();

			if (geniusRows.Count == 0)
				geniusRows.Add(new PropertyRow { Property = "Info", ["SQL Value"] = "No data found" });

			var invRows = invProps.Select(kvp =>
			{
				var sqlVal = sqlData.GetValueOrDefault(Geniusinfo.GetSqlColumnName(kvp.Key), "");
				return new PropertyRow
				{
					Property      = kvp.Key, ["Inventor Value"] = kvp.Value,
					HasDifference = !GeniusFormsHelper.ValuesAreEqual(sqlVal, kvp.Value)
				};
			}).ToList();

			await Dispatcher.InvokeAsync(() =>
			{
				SqlDataGrid.ItemsSource      = geniusRows;
				InventorDataGrid.ItemsSource = invRows;
			});
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiPart: Error loading properties: {ex.Message}");
		}
	}

	private async void CalculatePropsButton_Click(object sender, RoutedEventArgs e)
	{
		if (ThisApplication?.ActiveDocument != null)
			await CalculatePropertiesForDocument(ThisApplication.ActiveDocument, true);
	}

	private async void CalculateEachMember_Click(object sender, RoutedEventArgs e)
	{
		var factory = _factoryDocument?.ComponentDefinition.iPartFactory;
		if (factory == null) return;

		var original     = factory.DefaultRow;
		var app          = ThisApplication;
		var screenUpdate = app.ScreenUpdating;

		try
		{
			CalculateEachMember.IsEnabled = CalculatePropsButton.IsEnabled = false;
			app.ScreenUpdating            = false;

			foreach (iPartTableRow row in factory.TableRows)
			{
				if (factory.DefaultRow.MemberName != row.MemberName)
				{
					factory.DefaultRow = row;
					_factoryDocument.Update();
				}

				await CalculatePropertiesForDocument((Document)_factoryDocument, false);
				await Task.Delay(1);
			}

			if (factory.DefaultRow.MemberName != original.MemberName)
			{
				factory.DefaultRow = original;
				_factoryDocument.Update();
			}

			var pn = PropertyExtractor.GetPropertiesFromDocumentStatic((Document)_factoryDocument)
			                          .GetValueOrDefault("Part Number", _factoryDocument.DisplayName);
			await LoadPropertiesForPart(pn, (Document)_factoryDocument);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiPart: Error in CalculateEachMember: {ex.Message}");
		}
		finally
		{
			app.ScreenUpdating            = screenUpdate;
			CalculateEachMember.IsEnabled = CalculatePropsButton.IsEnabled = true;
		}
	}

	private async Task CalculatePropertiesForDocument(Document doc, bool updateUI)
	{
		if (updateUI) CalculatePropsButton.IsEnabled = false;
		try
		{
			var calculated = await CalculateProps.CalculateAllPropertiesAsync((_Document)doc);
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
					GeniusFormsHelper.UpdateInventorProperties((_Document)doc, updates, "GeniusiPart");
					if (updateUI)
					{
						var pn = current.GetValueOrDefault("Part Number", doc.DisplayName);
						await LoadPropertiesForPart(pn, doc);
					}
				}
			}

			await Task.Delay(1);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiPart: Calculation error: {ex.Message}");
		}
		finally
		{
			if (updateUI) CalculatePropsButton.IsEnabled = true;
		}
	}
}