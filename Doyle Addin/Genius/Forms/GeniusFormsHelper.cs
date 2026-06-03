namespace DoyleAddin.Genius.Forms;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Inventor;

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
}

public class PropertyRow : Dictionary<string, object>
{
	public string Property
	{
		get => TryGetValue("Property", out var value) ? value?.ToString() : "";
		init => this["Property"] = value;
	}

	public bool HasDifference
	{
		get => TryGetValue("HasDifference", out var value) && value is true;
		set => this["HasDifference"] = value;
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

	private void UpdateHasDifference()
	{
		var invVal = TryGetValue("Inventor Value", out var inv) ? inv?.ToString() : "";
		var sqlVal = TryGetValue("SQL Value", out var sql) ? sql?.ToString() : "";
		HasDifference  = !GeniusFormsHelper.ValuesAreEqual(invVal, sqlVal);
		this["Status"] = HasDifference ? "Mismatch" : "Match";
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
		set => this["Level"] = value;
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
}