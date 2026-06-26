namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.ComponentModel;
using System.Threading.Tasks;

/// <summary>
///     Provides shared functionality for all Genius Forms classes.
/// </summary>
public static class GeniusFormsHelper
{
	private const double NumericTolerance = 0.0001;

	public static string MapSqlColumnToInventorProperty(string sqlColumn)
	{
		return Geniusinfo.GetInventorPropertyName(sqlColumn);
	}

	private static string NormalizeNumericValue(string value)
	{
		if (string.IsNullOrEmpty(value)) return value;

		return decimal.TryParse(value, out var decimalValue)
			? decimalValue.ToString("G29").TrimEnd('0').TrimEnd('.')
			: value;
	}

	private static string StripUnits(string value)
	{
		if (string.IsNullOrEmpty(value)) return value;

		var spaceIndex = value.LastIndexOf(' ');
		return spaceIndex > 0 ? value[..spaceIndex].Trim() : value.Trim();
	}

	public static bool ValuesAreEqual(string value1, string value2)
	{
		if (string.Equals(value1, value2)) return true;
		value1 ??= "";
		value2 ??= "";

		var clean1 = StripUnits(value1);
		var clean2 = StripUnits(value2);

		if (double.TryParse(clean1, out var double1) && double.TryParse(clean2, out var double2))
			return Math.Abs(double1 - double2) < NumericTolerance;

		return string.Equals(NormalizeNumericValue(clean1), NormalizeNumericValue(clean2),
			StringComparison.OrdinalIgnoreCase);
	}

	public static void UpdateInventorProperties(Document document, Dictionary<string, string> propertiesToUpdate,
		string logPrefix = "Genius")
	{
		if (document == null || propertiesToUpdate == null || propertiesToUpdate.Count == 0) return;

		try
		{
			var wasAlreadyDirty     = document.Dirty;
			var propertySets        = document.PropertySets;
			var userDefinedProps    = propertySets[GeniusConstants.UserDefinedProperties];
			var designTrackingProps = propertySets[GeniusConstants.DesignTrackingProperties];

			var actualUpdates = 0;

			foreach (var (key, newValue) in propertiesToUpdate)
				try
				{
					var propertySet = GeniusConstants.StandardProps.Contains(key)
						? designTrackingProps
						: userDefinedProps;
					if (propertySet == null) continue;

					Property existingProp;
					var      propertyAlreadyExisted = false;

					try
					{
						existingProp           = propertySet[key];
						propertyAlreadyExisted = true;
					}
					catch
					{
						if (propertySet != userDefinedProps)
						{
							Debug.WriteLine($"{logPrefix}: Standard property '{key}' not found.");
							continue;
						}

						Debug.WriteLine($"{logPrefix}: Creating property '{key}' = '{newValue}'");
						try
						{
							propertySet.Add(newValue, key);
							existingProp = propertySet[key];
						}
						catch (Exception createEx)
						{
							Debug.WriteLine($"{logPrefix}: Failed to create '{key}': {createEx.Message}");
							continue;
						}
					}

					if (existingProp == null) continue;

					var currentValue = existingProp.Value?.ToString();
					if (ValuesAreEqual(currentValue, newValue)) continue;
					existingProp.Value = newValue;
					if (propertyAlreadyExisted) actualUpdates++;
					Debug.WriteLine($"{logPrefix}: Updated '{key}' from '{currentValue}' to '{newValue}'");
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"{logPrefix}: Error updating '{key}': {ex.Message}");
				}

			if (!wasAlreadyDirty && actualUpdates <= 0) return;
			document.Save();
			Debug.WriteLine($"{logPrefix}: Document saved (wasDirty={wasAlreadyDirty}, updates={actualUpdates})");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"{logPrefix}: Error in UpdateInventorProperties: {ex.Message}");
		}
	}

	/// <summary>
	///     Zooms to and selects an occurrence in the active assembly document.
	/// </summary>
	/// <param name="document">The document to zoom to.</param>
	public static void ZoomToOccurrence(Document document)
	{
		try
		{
			if (ThisApplication?.ActiveDocument is not AssemblyDocument assemblyDoc || document == null) return;

			var occurrence = ThisApplication.Documents.ItemByName[document.FullFileName];
			if (occurrence == null) return;

			var asmOcc = assemblyDoc.ComponentDefinition.Occurrences.AllReferencedOccurrences[occurrence];
			if (asmOcc.Count <= 0) return;

			ComponentOccurrence target = null;
			for (var i = 1; i <= asmOcc.Count; i++)
				try
				{
					var occ = asmOcc[i];
					if (occ.Suppressed || occ.Excluded) continue;
					if (HasSuppressedOrExcludedAncestor(occ)) continue;
					target = occ;
					break;
				}
				catch
				{
					// skip occurrences that throw on property access
				}

			if (target == null) return;

			assemblyDoc.SelectSet.Clear();
			assemblyDoc.SelectSet.Select(target);
			ThisApplication.CommandManager.ControlDefinitions["AppZoomSelectCmd"].Execute();
		}
		catch
		{
			/* ignore selection errors */
		}
	}

	private static bool HasSuppressedOrExcludedAncestor(ComponentOccurrence occurrence)
	{
		try
		{
			var parent = occurrence.ParentOccurrence;
			while (parent != null)
			{
				if (parent.Suppressed || parent.Excluded) return true;
				parent = parent.ParentOccurrence;
			}
		}
		catch
		{
			return true;
		}

		return false;
	}

	/// <summary>
	///     Loads properties for a selected part from both SQL and Inventor sources.
	/// </summary>
	/// <param name="partNumber">The part number to load properties for.</param>
	/// <param name="sqlDataManager">The SQL data manager instance.</param>
	/// <param name="document">The Inventor document (optional).</param>
	/// <returns>A tuple containing SQL rows and Inventor rows.</returns>
	public static async Task<(List<PropertyRow> SqlRows, List<PropertyRow> InventorRows)> LoadPropertiesForSelectedPart(
		string partNumber, ISqlDataManager sqlDataManager, Document document = null)
	{
		try
		{
			var sqlData = await sqlDataManager.GetSqlDataAsync(partNumber);
			var invProps = document != null
				? PropertyExtractor.GetPropertiesFromDocumentStatic(document)
				: PropertyExtractor.GetAllProperties();

			var geniusRows = sqlData.Select(kvp =>
			{
				var invName = MapSqlColumnToInventorProperty(kvp.Key);
				var invVal  = invProps.GetValueOrDefault(invName, "");
				return new PropertyRow
				{
					Property      = invName,
					["SQL Value"] = kvp.Value,
					HasDifference = !ValuesAreEqual(kvp.Value, invVal)
				};
			}).ToList();

			if (geniusRows.Count == 0)
				geniusRows.Add(new PropertyRow { Property = "Info", ["SQL Value"] = "No data found" });

			var invRows = invProps.Select(kvp =>
			{
				var sqlVal = sqlData.GetValueOrDefault(Geniusinfo.GetSqlColumnName(kvp.Key), "");
				return new PropertyRow
				{
					Property           = kvp.Key,
					["Inventor Value"] = kvp.Value,
					HasDifference      = !ValuesAreEqual(sqlVal, kvp.Value)
				};
			}).ToList();

			// Sort both lists by Property name to ensure consistent ordering
			geniusRows.Sort((a, b) => string.Compare(a.Property, b.Property, StringComparison.Ordinal));
			invRows.Sort((a, b) => string.Compare(a.Property, b.Property, StringComparison.Ordinal));

			return (geniusRows, invRows);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error loading properties for part {partNumber}: {ex.Message}");
			return ([], []);
		}
	}
}

public class PropertyRow : Dictionary<string, object>, INotifyPropertyChanged
{
	public string Property
	{
		get => TryGetValue("Property", out var value) ? value?.ToString() : "";
		init => this["Property"] = value;
	}

	public bool HasDifference
	{
		get => TryGetValue("HasDifference", out var value) && value is true;
		set
		{
			if (value == (TryGetValue("HasDifference", out var v) && v is true)) return;
			this["HasDifference"] = value;
			OnPropertyChanged(nameof(HasDifference));
		}
	}

	public new object this[string key]
	{
		get => base[key];
		set
		{
			base[key] = value;
			if (key == "Inventor Value") UpdateHasDifference();
		}
	}

	public event PropertyChangedEventHandler PropertyChanged;

	private void UpdateHasDifference()
	{
		var invVal = TryGetValue("Inventor Value", out var inv) ? inv?.ToString() : "";
		var sqlVal = TryGetValue("SQL Value", out var sql) ? sql?.ToString() : "";
		HasDifference  = !GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
		this["Status"] = HasDifference ? "Mismatch" : "Match";
		OnPropertyChanged(nameof(HasDifference));
	}

	private void OnPropertyChanged(string propertyName)
	{
		PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
	}
}

public class MemberInfo : PropertyRow
{
	public string PartNumber
	{
		get => TryGetValue("PartNumber", out var value) ? value?.ToString() : "";
		init => this["PartNumber"] = value;
	}

	public string Description
	{
		get => TryGetValue("Description", out var value) ? value?.ToString() : "";
		set => this["Description"] = value;
	}
}

public class ChildInfo : PropertyRow
{
	public int Level
	{
		get => TryGetValue("Level", out var value) && value is int i ? i : 0;
		init => this["Level"] = value;
	}

	public string PartNumber
	{
		get => TryGetValue("PartNumber", out var value) ? value?.ToString() : "";
		init => this["PartNumber"] = value;
	}

	public string Description
	{
		get => TryGetValue("Description", out var value) ? value?.ToString() : "";
		set => this["Description"] = value;
	}

	public string DocumentType
	{
		get => TryGetValue("DocumentType", out var value) ? value?.ToString() : "";
		init => this["DocumentType"] = value;
	}

	public string FullPath
	{
		get => TryGetValue("FullPath", out var value) ? value?.ToString() : "";
		init => this["FullPath"] = value;
	}

	public bool IsPurchased
	{
		get => TryGetValue("IsPurchased", out var value) && value is true;
		init => this["IsPurchased"] = value;
	}

	public List<ChildInfo> Children { get; set; } = [];
}