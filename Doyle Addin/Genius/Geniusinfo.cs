namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Inventor;
using Path = Path;

/// <summary>
///     Determines the type of the currently open document and provides information about which panel to use.
/// </summary>
public static class Geniusinfo
{
	/// <summary>
	///     Specifies the type of document currently open in the application.
	/// </summary>
	/// <remarks>
	///     This enumeration is used to determine the type of the open document and to provide
	///     guidance on which panel to display in the user interface.
	/// </remarks>
	public enum PanelType
	{
		/// <summary>
		///     Represents a panel type used for displaying information about a document categorized as a standard part.
		/// </summary>
		/// <remarks>
		///     This panel type is associated with documents that are identified as regular part files without advanced
		///     configurations such as iParts or multiple model states. It is commonly used as the primary representation
		///     for single-part documents in the application interface.
		/// </remarks>
		Part,

		/// <summary>
		///     Represents a panel type used for displaying information about a document categorized as an assembly.
		/// </summary>
		/// <remarks>
		///     This panel type is associated with documents that are identified as assembly files. It is typically used
		///     to represent collections of multiple components or sub-assemblies structured into a single design.
		///     The associated user interface provides tools and views specific to managing assembly-level content
		///     within the application.
		/// </remarks>
		Assembly,

		/// <summary>
		///     Represents a panel type used for displaying information about an iPart document.
		/// </summary>
		/// <remarks>
		///     The <c>IPart</c> panel type is used when the currently open document is identified as an iPart.
		///     iParts are specialized part files with multiple configurations, typically managed through
		///     a table-driven approach or utilizing model states. This panel type is designed to provide
		///     additional functionality and detail specific to these advanced configurations.
		/// </remarks>
		IPart,

		/// <summary>
		///     Represents a panel type used for displaying information about a document categorized as an iAssembly.
		/// </summary>
		/// <remarks>
		///     This panel type is associated with documents that are identified as an iAssembly, which are assemblies
		///     that have factory configurations or multiple model states. It is utilized to provide specialized
		///     visualization and interaction for documents containing advanced assembly setups in the application interface.
		/// </remarks>
		IAssembly
	}

	/// <summary>
	///     Represents the default connection string used for connecting to the DoyleDB SQL database
	///     within the scope of the DoyleAddin.Genius application.
	/// </summary>
	/// <remarks>
	///     The constant string <c>DefaultConnectionString</c> encapsulates essential connection details,
	///     such as the data source, initial catalog, user credentials, and password, required for database interaction.
	///     This field provides a fallback connection configuration when no specific connection string is supplied.
	/// </remarks>
	/// <value>
	///     A pre-defined connection string enabling seamless access to the DoyleDB database, facilitating
	///     operations like data retrieval, updates, or synchronization within the application.
	/// </value>
	public const string DefaultConnectionString =
		"Data Source=DOYLE-ERP02;Initial Catalog=DoyleDB;User ID=geniusreporting;Password=geniusreporting;";

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
	///     Retrieves the corresponding SQL column name for a given Inventor property name.
	///     If a mapping exists between the Inventor property name and an SQL column name, the SQL column name is returned.
	///     If no mapping exists, the input Inventor property name is returned unchanged.
	/// </summary>
	/// <param name="inventorPropertyName">
	///     The name of the Inventor property for which to find the corresponding SQL column name.
	///     If null or empty, the method returns the input value.
	/// </param>
	/// <returns>
	///     The corresponding SQL column name if a mapping exists; otherwise, the original Inventor property name.
	/// </returns>
	public static string GetSqlColumnName(string inventorPropertyName)
	{
		return string.IsNullOrEmpty(inventorPropertyName)
			? inventorPropertyName
			: InventorToSqlColumnMap.GetValueOrDefault(inventorPropertyName, inventorPropertyName);
	}

	/// <summary>
	///     Retrieves the corresponding Inventor property name for a given SQL column name.
	///     If a mapping exists between the SQL column name and an Inventor property name,
	///     the Inventor property name is returned. If no mapping exists, the input SQL column name is returned.
	/// </summary>
	/// <param name="sqlColumnName">
	///     The name of the SQL column for which to find the corresponding Inventor property name.
	///     If null or empty, the method returns the input value.
	/// </param>
	/// <returns>
	///     The corresponding Inventor property name if a mapping exists; otherwise, the original SQL column name.
	/// </returns>
	public static string GetInventorPropertyName(string sqlColumnName)
	{
		if (string.IsNullOrEmpty(sqlColumnName)) return sqlColumnName;
		return InventorToSqlColumnMap
		       .FirstOrDefault(kvp => kvp.Value.Equals(sqlColumnName, StringComparison.OrdinalIgnoreCase)).Key ??
		       sqlColumnName;
	}

	/// <summary>
	///     Determines the type of panel (Part, Assembly, IPart, or IAssembly) based on the provided Inventor document.
	///     If no document is provided, the active Inventor document is used.
	///     Returns PanelType.Part by default if the document is null or an unsupported type.
	/// </summary>
	/// <param name="doc">The Inventor document for which to determine the panel type. If null, the active document is used.</param>
	/// <returns>
	///     A value of the <c>Geniusinfo.PanelType</c> enumeration representing the type of panel:
	///     Part, Assembly, IPart, or IAssembly.
	/// </returns>
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

	/// <summary>
	///     Retrieves a DataTable containing information about all child components within the active Inventor assembly
	///     document.
	///     The table includes details such as the hierarchy level, part number, description, document type,
	///     whether differences exist, and the full file path of each component.
	/// </summary>
	/// <returns>
	///     A DataTable where each row represents a child component in the assembly and contains information such as
	///     level, part number, description, document type, difference status, and file path.
	///     Returns an empty table if the active document is not an assembly.
	/// </returns>
	public static DataTable GetAllAssemblyChildren()
	{
		var childrenTable = new DataTable();
		childrenTable.Columns.Add("Level", typeof(int));
		childrenTable.Columns.Add("PartNumber", typeof(string));
		childrenTable.Columns.Add("Description", typeof(string));
		childrenTable.Columns.Add("DocumentType", typeof(string));
		childrenTable.Columns.Add("HasDifference", typeof(bool));
		childrenTable.Columns.Add("FullPath", typeof(string));

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

				string partNumber = "", description = "", docType = "Virtual";
				if (doc != null)
				{
					docType = doc.DocumentType.ToString();
					try
					{
						var props = doc.PropertySets["Design Tracking Properties"];
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

				childrenTable.Rows.Add(level, partNumber, description, docType, false, path);

				if (occ.DefinitionDocumentType == kAssemblyDocumentObject &&
				    occ.Definition is AssemblyComponentDefinition subDef)
					ProcessOccurrences(subDef.Occurrences, level + 1);
			}
		}
	}
}

/// <summary>
///     Represents a mechanism for comparing properties between an Inventor document and an external data source.
///     Implementations of this interface define the logic to evaluate, analyze, or reconcile differences in property
///     values.
/// </summary>
public interface IPropertyComparator;

/// <summary>
///     Provides functionality for comparing properties of an Inventor document with data retrieved
///     from a SQL database. The results of the comparison are presented in a tabular format, detailing
///     differences and highlighting inconsistencies.
/// </summary>
public class PropertyComparator(ISqlDataManager sqlDataManager) : IPropertyComparator
{
	/// <summary>
	///     Represents the SQL data manager used for retrieving and managing data from a SQL database
	///     within the context of the DoyleAddin.Genius application.
	/// </summary>
	/// <remarks>
	///     The <c>SqlDataManager</c> property provides access to an implementation of the
	///     <c>ISqlDataManager</c> interface, enabling querying and retrieval of SQL data.
	///     This property is typically utilized within property comparison, data synchronization, or data extraction workflows
	///     in the application.
	/// </remarks>
	/// <value>
	///     An object implementing the <c>ISqlDataManager</c> interface used to perform SQL data operations.
	/// </value>
	public ISqlDataManager SqlDataManager { get; } = sqlDataManager;

	/// <summary>
	///     Compares properties from an Inventor document with data retrieved from an SQL database
	///     and returns a DataTable containing the comparison results.
	/// </summary>
	/// <param name="document">
	///     An Inventor document containing the properties to compare. If null, all properties will be
	///     extracted.
	/// </param>
	/// <returns>
	///     A DataTable containing the comparison results. The table includes columns for the property name, Inventor value,
	///     SQL value, comparison status, and a boolean indicating if there is a difference.
	/// </returns>
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

			var standardProps = new[] { "Part Number", "Description", "Cost Center" };
			var customProps   = new List<string> { "GeniusMass" };

			if (panelType is not (Geniusinfo.PanelType.Assembly or Geniusinfo.PanelType.IAssembly))
				customProps.AddRange([
					"Thickness", "Extent_Width", "Extent_Length", "Extent_Area", "RM", "RMUNIT", "RMQTY"
				]);

			foreach (var prop in standardProps.Concat(customProps))
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