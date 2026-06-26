namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Threading.Tasks;
using Forms;

/// <summary>
///     Retrieves custom iProperties and standard properties from Inventor documents.
/// </summary>
public static class PropertyExtractor
{
	public static IReadOnlyDictionary<string, string> GetAllProperties()
	{
		return ThisApplication.ActiveDocument is { } doc
			? GetPropertiesFromDocumentStatic(doc)
			: new Dictionary<string, string>();
	}

	public static IReadOnlyDictionary<string, string> GetPropertiesFromDocumentStatic(Document document)
	{
		ArgumentNullException.ThrowIfNull(document);

		var properties = new Dictionary<string, string>();

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
			Debug.WriteLine($"Error in {nameof(GetPropertiesFromDocumentStatic)}: {ex.Message}");
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

				var value = prop.Value?.ToString();
				if (!string.IsNullOrEmpty(value))
					properties[prop.Name] = value;

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
		var doc = ThisApplication?.Documents.Cast<Document>().FirstOrDefault(d =>
			string.Equals(d.DisplayName, partNumber, StringComparison.OrdinalIgnoreCase) ||
			string.Equals(GetFileNameWithoutExtension(d.FullFileName), partNumber,
				StringComparison.OrdinalIgnoreCase));
		if (doc != null) return doc;

		var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var active  = ThisApplication?.ActiveDocument;
		if (active == null) return null;

		var queue = new Queue<Document>();
		queue.Enqueue(active);
		visited.Add(active.FullFileName ?? string.Empty);

		while (queue.Count > 0)
		{
			var current = queue.Dequeue();
			if (current.FullFileName == null) continue;

			try
			{
				foreach (var referenced in current.ReferencedFiles.Cast<Document>().Where(referenced =>
					         referenced?.FullFileName != null && visited.Add(referenced.FullFileName)))
				{
					if (string.Equals(referenced.DisplayName, partNumber, StringComparison.OrdinalIgnoreCase) ||
					    string.Equals(GetFileNameWithoutExtension(referenced.FullFileName), partNumber,
						    StringComparison.OrdinalIgnoreCase))
						return referenced;

					queue.Enqueue(referenced);
				}
			}
			catch
			{
				// Skip documents whose referenced files can't be enumerated
			}
		}

		return null;
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

			var excludedProps = isAssembly
				? new HashSet<string>(StringComparer.OrdinalIgnoreCase)
				{
					"Thickness", "Extent_Width", "Extent_Length", "Extent_Area"
				}
				: new HashSet<string>(StringComparer.OrdinalIgnoreCase);

			var geniusRows = sqlData
			                 .Where(kvp =>
				                 !excludedProps.Contains(GeniusFormsHelper.MapSqlColumnToInventorProperty(kvp.Key)))
			                 .Select(kvp =>
			                 {
				                 var invName = GeniusFormsHelper.MapSqlColumnToInventorProperty(kvp.Key);
				                 var invVal  = invProps.GetValueOrDefault(invName, "");
				                 return new PropertyRow
				                 {
					                 Property           = invName,
					                 ["SQL Value"]      = kvp.Value,
					                 ["Inventor Value"] = invVal
				                 };
			                 })
			                 .ToList();

			if (geniusRows.Count == 0)
				geniusRows.Add(new PropertyRow { Property = "Info", ["SQL Value"] = "No data found" });

			var invRows = invProps
			              .Where(kvp => !excludedProps.Contains(kvp.Key))
			              .Select(kvp =>
			              {
				              var sqlVal = sqlData.GetValueOrDefault(Geniusinfo.GetSqlColumnName(kvp.Key), "");
				              return new PropertyRow
				              {
					              Property           = kvp.Key,
					              ["Inventor Value"] = kvp.Value,
					              ["SQL Value"]      = sqlVal
				              };
			              })
			              .ToList();

			return (geniusRows, invRows);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"{nameof(PropertyExtractor)}: Load properties error: {ex.Message}");
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
		catch (Exception ex) when (ex is NullReferenceException or IndexOutOfRangeException)
		{
			// The property set does not exist in this document; this is expected for some document types.
			result = null;
			return false;
		}
	}
}