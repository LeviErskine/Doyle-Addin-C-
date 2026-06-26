namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

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

	/// <summary>
	///     Reverse mapping from SQL column names to Inventor property names, built once for O(1) lookups.
	/// </summary>
	private static readonly Dictionary<string, string> SqlToInventorColumnMap;

	static Geniusinfo()
	{
		SqlToInventorColumnMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
		foreach (var kvp in InventorToSqlColumnMap)
			// If multiple Inventor properties map to the same SQL column, the last one wins (consistent behavior).
			SqlToInventorColumnMap[kvp.Value] = kvp.Key;
	}

	/// <summary>
	///     Returns true if the given panel type represents a table-driven (iPart/iAssembly) document.
	/// </summary>
	public static bool IsTableDriven(PanelType panelType)
	{
		return panelType is PanelType.IPart or PanelType.IAssembly;
	}

	public static string GetSqlColumnName(string inventorPropertyName)
	{
		return !string.IsNullOrEmpty(inventorPropertyName) &&
		       InventorToSqlColumnMap.TryGetValue(inventorPropertyName, out var sqlColumn)
			? sqlColumn
			: inventorPropertyName;
	}

	public static string GetInventorPropertyName(string sqlColumnName)
	{
		return !string.IsNullOrEmpty(sqlColumnName) &&
		       SqlToInventorColumnMap.TryGetValue(sqlColumnName, out var inventorProperty)
			? inventorProperty
			: sqlColumnName;
	}

	public static PanelType GetPanelType(Document doc = null)
	{
		try
		{
			var targetDoc = doc ?? ThisApplication.ActiveDocument;
			if (targetDoc == null) return PanelType.Part;

			return targetDoc switch
			{
				PartDocument partDoc => HasMultipleModelStatesOrFactory(partDoc.ComponentDefinition)
					? PanelType.IPart
					: PanelType.Part,
				AssemblyDocument assyDoc => HasMultipleModelStatesOrFactory(assyDoc.ComponentDefinition)
					? PanelType.IAssembly
					: PanelType.Assembly,
				_ => PanelType.Part
			};
		}
		catch (Exception ex) when (ex is NullReferenceException or InvalidCastException)
		{
			Debug.WriteLine($"GetPanelType: {ex.Message}");
			return PanelType.Part;
		}
	}

	private static bool HasMultipleModelStatesOrFactory(PartComponentDefinition compDef)
	{
		return compDef.iPartFactory != null || compDef.ModelStates?.Count > 1;
	}

	private static bool HasMultipleModelStatesOrFactory(AssemblyComponentDefinition compDef)
	{
		return compDef.iAssemblyFactory != null || compDef.ModelStates?.Count > 1;
	}

	public static DataTable GetAllAssemblyChildren(AssemblyDocument assemblyDoc = null)
	{
		var childrenTable = new DataTable();
		childrenTable.Columns.Add("Level", typeof(int));
		childrenTable.Columns.Add("PartNumber", typeof(string));
		childrenTable.Columns.Add("Description", typeof(string));
		childrenTable.Columns.Add("DocumentType", typeof(string));
		childrenTable.Columns.Add("HasDifference", typeof(bool));
		childrenTable.Columns.Add("FullPath", typeof(string));
		childrenTable.Columns.Add("IsPurchased", typeof(bool));

		assemblyDoc ??= ThisApplication.ActiveDocument as AssemblyDocument;
		if (assemblyDoc == null) return childrenTable;

		var compDef            = assemblyDoc.ComponentDefinition;
		var processedDocuments = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var parentPath         = assemblyDoc.FullFileName;

		try
		{
			GetOccurrences(compDef.Occurrences, 1);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error extracting children: {ex.Message}");
		}

		return childrenTable;

		void GetOccurrences(ComponentOccurrences occurrences, int level)
		{
			foreach (ComponentOccurrence occ in occurrences)
			{
				try
				{
					if (occ.Suppressed || occ.Excluded) continue;
					if ((int)occ.BOMStructure is 2 or 4) continue;
					if (occ.IsiAssemblyMember || occ.IsiPartMember) continue;
				}
				catch
				{
					continue;
				}

				string    path;
				_Document doc = null;
				try
				{
					if (occ.Definition.Document is _Document document)
					{
						doc  = document;
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

					if (string.IsNullOrEmpty(partNumber)) partNumber = GetFileNameWithoutExtension(path);
				}
				else
				{
					partNumber = occ.Name;
				}

				var isPurchased = occ.BOMStructure == BOMStructureEnum.kPurchasedBOMStructure;

				childrenTable.Rows.Add(level, partNumber, description, docType, false, path, isPurchased);

				if (occ.DefinitionDocumentType == kAssemblyDocumentObject &&
				    occ.Definition is AssemblyComponentDefinition assemblyComponentDefinition &&
				    !isPurchased)
					GetOccurrences(assemblyComponentDefinition.Occurrences, level + 1);
			}
		}
	}
}

public class PropertyComparator(ISqlDataManager sqlDataManager)
{
	/// <summary>
	///     Custom property names to include for non-assembly documents (cached to avoid re-allocation).
	/// </summary>
	private static readonly string[] NonAssemblyCustomProps =
	[
		"Thickness", "Extent_Width", "Extent_Length", "Extent_Area", "RM", "RMUNIT", "RMQTY"
	];

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

			var customProps = new List<string>(1 + NonAssemblyCustomProps.Length) { "GeniusMass" };
			if (!Geniusinfo.IsTableDriven(panelType) && panelType != Geniusinfo.PanelType.Assembly)
				customProps.AddRange(NonAssemblyCustomProps);

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