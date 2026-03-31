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
	/// <summary>
	///     Gets all custom iProperties and standard properties from the active document.
	/// </summary>
	/// <returns>A dictionary containing property names and values.</returns>
	public static Dictionary<string, string> GetAllProperties()
	{
		var properties = new Dictionary<string, string>();

		try
		{
			var activeDocument = ThisApplication.ActiveDocument;
			if (activeDocument == null) return properties;

			GetStandardProperties(activeDocument, properties);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in GetAllProperties: {ex.Message}");
		}

		return properties;
	}

	/// <summary>
	///     Gets all custom iProperties and standard properties from a specific document.
	/// </summary>
	/// <param name="document">The Inventor document to extract properties from.</param>
	/// <returns>A dictionary containing property names and values.</returns>
	public static Dictionary<string, string> GetPropertiesFromDocumentStatic(Document document)
	{
		var properties = new Dictionary<string, string>();

		try
		{
			if (document == null) return properties;

			GetStandardProperties(document, properties);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in GetPropertiesFromDocumentStatic: {ex.Message}");
		}

		return properties;
	}

	/// <summary>
	///     Gets standard properties like Part Number and Description.
	/// </summary>
	/// <param name="document">The Inventor document.</param>
	/// <param name="properties">Dictionary to populate with properties.</param>
	private static void GetStandardProperties(Document document, Dictionary<string, string> properties)
	{
		try
		{
			var propertySets = document.PropertySets;

			// Design Tracking Properties
			if (TryGetPropertySet(propertySets, "Design Tracking Properties", out var designTrackingProps))
			{
				string[] trackingProps = ["Part Number", "Description", "Cost Center"];
				AddPropertiesFromSet(designTrackingProps, trackingProps, properties);
			}

			// Custom Properties
			if (!TryGetPropertySet(propertySets, "Inventor User Defined Properties", out var customPropertySet)) return;
			string[] customProps =
			[
				"GeniusMass", "Thickness", "Extent_Width", "Extent_Length", "Extent_Area", "RM", "RMUNIT", "RMQTY"
			];
			AddPropertiesFromSet(customPropertySet, customProps, properties);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in GetStandardProperties: {ex.Message}");
		}
	}

	/// <summary>
	///     Adds specified properties from a property set to the dictionary by iterating once.
	///     This avoids multiple COMExceptions that occur when using the indexer for missing properties.
	/// </summary>
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

	/// <summary>
	///     Finds and returns a document based on the given part number.
	/// </summary>
	/// <param name="partNumber">The part number corresponding to the document to be retrieved.</param>
	/// <returns>The document associated with the given part number, or null if no matching document is found.</returns>
	public static Document FindDocumentByPartNumber(string partNumber)
	{
		return ThisApplication?.Documents.Cast<Document>().FirstOrDefault(doc =>
			string.Equals(doc.DisplayName, partNumber, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(Path.GetFileNameWithoutExtension(doc.FullFileName), partNumber,
				StringComparison.OrdinalIgnoreCase));
	}

	/// <summary>
	///     Loads property data for a specified part from both a SQL data source and an Inventor document.
	/// </summary>
	/// <param name="partNumber">The part number of the item for which properties are to be loaded.</param>
	/// <param name="sqlDataManager">The SQL data manager instance used to retrieve data from the SQL database.</param>
	/// <param name="document">
	///     The Inventor document from which to extract properties. If null, a default method is used to retrieve properties.
	/// </param>
	/// <returns>
	///     A tuple containing two lists of property rows:
	///     1. A list of property rows derived from the SQL database (geniusRows).
	///     2. A list of property rows derived from the Inventor document (invRows).
	/// </returns>
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
					Property      = invName, ["SQL Value"] = kvp.Value,
					HasDifference = !GeniusFormsHelper.ValuesAreEqual(kvp.Value, invVal)
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
					Property      = kvp.Key, ["Inventor Value"] = kvp.Value,
					HasDifference = !GeniusFormsHelper.ValuesAreEqual(sqlVal, kvp.Value)
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