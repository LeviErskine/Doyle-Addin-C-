namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Inventor;
using Path = Path;

/// <summary>
///     Provides a centralized location for constant values used throughout the Genius plugin,
///     including default database connection strings, property metadata, and standard property lists.
///     These constants are used across various components to maintain consistency and reduce duplication.
/// </summary>
public static class GeniusConstants
{
	/// <summary>
	///     Represents the default connection string used to establish a connection
	///     to the DoyleDB database. This constant value is commonly used when
	///     no specific connection string is provided, ensuring consistent database access
	///     across the application.
	/// </summary>
	public const string DefaultConnectionString =
		"Data Source=DOYLE-ERP02;Initial Catalog=DoyleDB;User ID=geniusreporting;Password=geniusreporting;";

	public const string DesignTrackingProperties = "Design Tracking Properties";
	public const string UserDefinedProperties = "Inventor User Defined Properties";
	public const string SummaryInformation = "Inventor Summary Information";

	public static readonly string[] StandardProps = ["Part Number", "Description", "Cost Center"];

	public static readonly string[] CustomProps =
	[
		"GeniusMass", "Thickness", "Extent_Width", "Extent_Length", "Extent_Area", "RM", "RMUNIT", "RMQTY"
	];
}

/// <summary>
///     Determines the type of the currently open document and provides information about which panel to use.
/// </summary>
public static class Geniusinfo
{
	public enum PanelType
	{
		Part,
		Assembly,
		IPart,
		IAssembly
	}

	public const string DefaultConnectionString = GeniusConstants.DefaultConnectionString;

	private static readonly Dictionary<string, string> InventorToSqlColumnMap = new(StringComparer.OrdinalIgnoreCase)
	{
		["Part Number"]   = "Item",
		["Description"]   = "Description1",
		["Cost Center"]   = "Family",
		["GeniusMass"]    = "Weight",
		["Extent_Width"]  = "Width",
		["Extent_Length"] = "Length",
		["Extent_Area"]   = "Diameter",
		["Thickness"]     = "Thickness",
		["RM"]            = "RM",
		["RMUNIT"]        = "RMUNIT",
		["RMQTY"]         = "RMQTY"
	};

	public static string GetSqlColumnName(string inventorPropertyName)
	{
		return string.IsNullOrEmpty(inventorPropertyName)
			? inventorPropertyName
			: InventorToSqlColumnMap.GetValueOrDefault(inventorPropertyName, inventorPropertyName);
	}

	public static string GetInventorPropertyName(string sqlColumnName)
	{
		if (string.IsNullOrEmpty(sqlColumnName)) return sqlColumnName;
		return InventorToSqlColumnMap
		       .FirstOrDefault(kvp => kvp.Value.Equals(sqlColumnName, StringComparison.OrdinalIgnoreCase)).Key ??
		       sqlColumnName;
	}

	public static PanelType GetPanelType(Document doc = null)
	{
		try
		{
			var targetDoc = doc ?? ThisApplication.ActiveDocument;
			if (targetDoc == null) return PanelType.Part;

			return targetDoc switch
			{
				PartDocument partDoc => partDoc.ComponentDefinition.iPartFactory != null ||
				                        partDoc.ComponentDefinition.ModelStates?.Count > 1
					? PanelType.IPart
					: PanelType.Part,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.iAssemblyFactory != null ||
				                            assyDoc.ComponentDefinition.ModelStates?.Count > 1
					? PanelType.IAssembly
					: PanelType.Assembly,
				_ => PanelType.Part
			};
		}
		catch
		{
			return PanelType.Part;
		}
	}

	public static DataTable GetAllAssemblyChildren()
	{
		var childrenTable = new DataTable();
		childrenTable.Columns.Add("Level", typeof(int));
		childrenTable.Columns.Add("PartNumber", typeof(string));
		childrenTable.Columns.Add("Description", typeof(string));
		childrenTable.Columns.Add("DocumentType", typeof(string));
		childrenTable.Columns.Add("HasDifference", typeof(bool));
		childrenTable.Columns.Add("FullPath", typeof(string));
		childrenTable.Columns.Add("IsPurchased", typeof(bool));

		if (ThisApplication.ActiveDocument is not AssemblyDocument assemblyDoc) return childrenTable;

		var processedDocuments = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var parentPath         = assemblyDoc.FullFileName;

		try
		{
			ProcessOccurrences(assemblyDoc.ComponentDefinition.Occurrences, 1);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error extracting children: {ex.Message}");
		}

		return childrenTable;

		void ProcessOccurrences(ComponentOccurrences occurrences, int level)
		{
			foreach (var occ in occurrences.Cast<ComponentOccurrence>().Where(occ =>
				         !occ.Suppressed && !occ.Excluded && (int)occ.BOMStructure is not (2 or 4)))
			{
				string    path;
				_Document doc = null;
				try
				{
					if (occ.Definition.Document is _Document d)
					{
						doc  = d;
						path = doc.FullFileName;
					}
					else
					{
						path = "virtual:" + occ.Name;
					}
				}
				catch
				{
					continue;
				}

				if (path.Equals(parentPath, StringComparison.OrdinalIgnoreCase) ||
				    !processedDocuments.Add(path)) continue;

				string partNumber = "", description = "";
				var    docType    = doc?.DocumentType.ToString() ?? "Virtual";

				if (doc != null)
				{
					try
					{
						var props = doc.PropertySets[GeniusConstants.DesignTrackingProperties];
						partNumber  = props["Part Number"]?.Value?.ToString();
						description = props["Description"]?.Value?.ToString();
					}
					catch
					{
						/* ignore */
					}

					if (string.IsNullOrEmpty(partNumber)) partNumber = Path.GetFileNameWithoutExtension(path);
				}
				else
				{
					partNumber = occ.Name;
				}

				var isPurchased = occ.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure;

				childrenTable.Rows.Add(level, partNumber, description, docType, false, path, isPurchased);

				if (occ.DefinitionDocumentType == kAssemblyDocumentObject &&
				    occ.Definition is AssemblyComponentDefinition subDef &&
				    !isPurchased)
					ProcessOccurrences(subDef.Occurrences, level + 1);
			}
		}
	}
}

public interface IPropertyComparator;

public class PropertyComparator(ISqlDataManager sqlDataManager) : IPropertyComparator
{
	public ISqlDataManager SqlDataManager { get; } = sqlDataManager;

	public async Task<DataTable> ComparePropertiesAsync(Document document = null)
	{
		var table = new DataTable();
		table.Columns.Add("Property", typeof(string));
		table.Columns.Add("Inventor Value", typeof(string));
		table.Columns.Add("SQL Value", typeof(string));
		table.Columns.Add("Status", typeof(string));
		table.Columns.Add("HasDifference", typeof(bool));

		try
		{
			var inventorProps = document == null
				? PropertyExtractor.GetAllProperties()
				: PropertyExtractor.GetPropertiesFromDocumentStatic(document);

			var partNumber = inventorProps.GetValueOrDefault("Part Number", "");
			if (string.IsNullOrEmpty(partNumber))
			{
				table.Rows.Add("Error", "No Part Number", "", "", false);
				return table;
			}

			var sqlData   = await SqlDataManager.GetSqlDataAsync(partNumber);
			var panelType = Geniusinfo.GetPanelType(document);

			var customProps = new List<string> { "GeniusMass" };
			if (panelType is not (Geniusinfo.PanelType.Assembly or Geniusinfo.PanelType.IAssembly))
				customProps.AddRange([
					"Thickness", "Extent_Width", "Extent_Length", "Extent_Area", "RM", "RMUNIT", "RMQTY"
				]);

			foreach (var prop in GeniusConstants.StandardProps.Concat(customProps))
			{
				var invValue = inventorProps.GetValueOrDefault(prop, "");
				var sqlCol   = Geniusinfo.GetSqlColumnName(prop);
				var sqlValue = sqlData.GetValueOrDefault(sqlCol, "");
				var status   = CompareValues(invValue, sqlValue);

				table.Rows.Add(prop, invValue, sqlValue, status, status != "Match");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Comparison error: {ex.Message}");
		}

		return table;
	}

	private static string CompareValues(string inv, string sql)
	{
		if (string.IsNullOrEmpty(inv) && string.IsNullOrEmpty(sql)) return "Match";
		if (string.IsNullOrEmpty(inv)) return "Missing in Inventor";
		if (string.IsNullOrEmpty(sql)) return "Missing in SQL";

		if (double.TryParse(inv.Split(' ')[0], out var dInv) && double.TryParse(sql.Split(' ')[0], out var dSql))
			return Math.Abs(dInv - dSql) < 0.0001 ? "Match" : "Mismatch";

		return string.Equals(inv, sql, StringComparison.OrdinalIgnoreCase) ? "Match" : "Mismatch";
	}
}