namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Forms;
using Inventor;
using Path = Path;

/// <summary>
///     Defines operations for extracting properties from Inventor documents.
/// </summary>
public interface IPropertyExtractor;

/// <summary>
///     Retrieves custom iProperties and standard properties from Inventor documents.
/// </summary>
public abstract class PropertyExtractor : IPropertyExtractor
{
	public static Dictionary<string, string> GetAllProperties()
	{
		return ThisApplication.ActiveDocument is { } doc ? GetPropertiesFromDocumentStatic(doc) : [];
	}

	public static Dictionary<string, string> GetPropertiesFromDocumentStatic(Document document)
	{
		var properties = new Dictionary<string, string>();
		if (document == null) return properties;

		try
		{
			var propertySets = document.PropertySets;

			if (TryGetPropertySet(propertySets, GeniusConstants.DesignTrackingProperties, out var designTrackingProps))
				AddPropertiesFromSet(designTrackingProps, GeniusConstants.StandardProps, properties);

			if (TryGetPropertySet(propertySets, GeniusConstants.UserDefinedProperties, out var customPropertySet))
				AddPropertiesFromSet(customPropertySet, GeniusConstants.CustomProps, properties);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in GetPropertiesFromDocumentStatic: {ex.Message}");
		}

		return properties;
	}

	private static void AddPropertiesFromSet(PropertySet propertySet, string[] propertyNames,
		Dictionary<string, string> properties)
	{
		if (propertySet == null || propertyNames == null || propertyNames.Length == 0) return;

		try
		{
			var namesToFind = new HashSet<string>(propertyNames, StringComparer.OrdinalIgnoreCase);
			foreach (Property prop in propertySet)
			{
				if (!namesToFind.Contains(prop.Name)) continue;

				var value                                               = prop.Value?.ToString();
				if (!string.IsNullOrEmpty(value)) properties[prop.Name] = value;

				namesToFind.Remove(prop.Name);
				if (namesToFind.Count == 0) break;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error iterating properties in set: {ex.Message}");
		}
	}

	public static Document FindDocumentByPartNumber(string partNumber)
	{
		return ThisApplication?.Documents.Cast<Document>().FirstOrDefault(doc =>
			string.Equals(doc.DisplayName, partNumber, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(Path.GetFileNameWithoutExtension(doc.FullFileName), partNumber,
				StringComparison.OrdinalIgnoreCase));
	}

	public static async Task<(List<PropertyRow> geniusRows, List<PropertyRow> invRows)> LoadPropertiesForPart(
		string partNumber, ISqlDataManager sqlDataManager, Document document = null)
	{
		try
		{
			var sqlData    = await sqlDataManager.GetSqlDataAsync(partNumber);
			var invProps   = document != null ? GetPropertiesFromDocumentStatic(document) : GetAllProperties();
			var panelType  = Geniusinfo.GetPanelType(document);
			var isAssembly = panelType is Geniusinfo.PanelType.Assembly or Geniusinfo.PanelType.IAssembly;

			var excludedProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
			if (isAssembly)
			{
				excludedProps.Add("Thickness");
				excludedProps.Add("Extent_Width");
				excludedProps.Add("Extent_Length");
				excludedProps.Add("Extent_Area");
			}

			var geniusRows = sqlData.Select(kvp =>
			{
				var invName = GeniusFormsHelper.MapSqlColumnToInventorProperty(kvp.Key);
				if (excludedProps.Contains(invName)) return null;

				var invVal = invProps.GetValueOrDefault(invName, "");
				return new PropertyRow
				{
					Property           = invName,
					["SQL Value"]      = kvp.Value,
					["Inventor Value"] = invVal,
					HasDifference      = !GeniusFormsHelper.ValuesAreEqual(kvp.Value, invVal)
				};
			}).Where(row => row != null).ToList();

			if (geniusRows.Count == 0)
				geniusRows.Add(new PropertyRow { Property = "Info", ["SQL Value"] = "No data found" });

			var invRows = invProps.Select(kvp =>
			{
				if (excludedProps.Contains(kvp.Key)) return null;

				var sqlVal = sqlData.GetValueOrDefault(Geniusinfo.GetSqlColumnName(kvp.Key), "");
				return new PropertyRow
				{
					Property           = kvp.Key,
					["Inventor Value"] = kvp.Value,
					["SQL Value"]      = sqlVal,
					HasDifference      = !GeniusFormsHelper.ValuesAreEqual(sqlVal, kvp.Value)
				};
			}).Where(row => row != null).ToList();

			return (geniusRows, invRows);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GeniusiAssembly: Load properties error: {ex.Message}");
		}

		return ([], []);
	}

	private static bool TryGetPropertySet(PropertySets sets, string setName, out PropertySet result)
	{
		try
		{
			result = sets[setName];
			return true;
		}
		catch
		{
			result = null;
			return false;
		}
	}
}