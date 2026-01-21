#region

using System.Data.SqlClient;
using System.Threading.Tasks;
using Dapper;

#endregion

namespace Doyle_Addin.Genius;

public class DatabaseService(string connectionString = null)
{
	// Default connection string for Genius database

	private const string DefaultConnectionString =
		"Data Source=DOYLE-ERP02;Initial Catalog=DoyleDB;User ID=GeniusReporting;Password=geniusreporting";

	private readonly string connectionString = connectionString ?? DefaultConnectionString;

	public async Task<PartInfo> GetPartByNumberAsync(string partNumber)
	{
		if (string.IsNullOrWhiteSpace(partNumber))
			return null;

		try
		{
			await using var connection = new SqlConnection(connectionString);
			await connection.OpenAsync();

			const string sql = """
			                                   SELECT 
			                                       m.Family AS CostCenter,
			                                       m.Item AS PartNumber,
			                                       m.Description1 AS Description,
			                                       m.Diameter AS Extent_Area,
			                                       m.Length AS Extent_Length,
			                                       m.Width AS Extent_Width,
			                                       m.Weight AS GeniusMass,
			                                       m.Thickness,
			                                       b.Item AS Stock,
			                                       b.QuantityInConversionUnit AS RMQTY,
			                                       b.ConversionUnit AS RMUNIT
			                                   FROM vgMfiItems m
			                                   LEFT JOIN vgIcoBillOfMaterials b ON m.Item = b.Product
			                                   WHERE m.Item = @ItemNumber
			                   """;

			var parameters = new { ItemNumber = partNumber };

			var result = await connection.QueryFirstOrDefaultAsync<PartInfo>(sql, parameters);
			return result;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error retrieving part {partNumber}: {ex.Message}");
			return null;
		}
	}

	/// <summary>
	///     Finds the correct raw stock based on thickness and material for parts not in database
	/// </summary>
	/// <param name="thickness">Thickness in inches</param>
	/// <param name="material">Material type (MS, SS, etc.)</param>
	/// <param name="partLength">Part length in inches</param>
	/// <param name="partWidth">Part width in inches</param>
	/// <returns>Raw stock information or null if not found</returns>
	public async Task<RawStockInfo> GetRawStockByThicknessAndMaterialAsync(double thickness, string material,
		double partLength = 0, double partWidth = 0)
	{
		if (thickness <= 0 || string.IsNullOrWhiteSpace(material))
			return null;

		try
		{
			await using var connection = new SqlConnection(connectionString);
			await connection.OpenAsync();

			// Determine item filter based on material
			var itemFilter = material switch
			{
				"SS" => "FS%",
				"MS" => "FM%",
				_    => "%"
			};

			const string sql = """
			                   SELECT
			                       Stock,
			                       CostCenter,
			                       Length,
			                       Width,
			                       Thickness,
			                       Material,
			                       RMUNIT
			                   FROM (
			                            SELECT
			                                m.Item AS Stock,
			                                m.Family AS CostCenter,
			                                m.Length AS Length,
			                                m.Width AS Width,
			                                m.Thickness,
			                                m.Specification6 AS Material,
			                                IIF(m.Family = 'DSHEET', 'FT2', b.ConversionUnit) AS RMUNIT,
			                                -- Assign a unique rank to handle duplicate items
			                                -- We PARTITION BY m.Item to handle duplicates but keep one per item
			                                ROW_NUMBER() OVER (PARTITION BY m.Item ORDER BY m.Length DESC, m.Width DESC) AS rn
			                            FROM vgMfiItems m
			                                     LEFT JOIN vgIcoBillOfMaterials b ON b.ItemLink = m.Item
			                            WHERE m.Thickness = @Thickness
			                              AND m.Specification6 = @Material
			                              AND m.Item LIKE @ItemFilter
			                              AND m.Length > 0 AND m.Width > 0
			                        ) AS RankedResults
			                   -- Filter to only keep the row that was ranked #1 for each Item
			                   WHERE rn = 1
			                   ORDER BY Stock;
			                   """;

			var parameters = new
			{
				Thickness  = thickness,
				Material   = material,
				ItemFilter = itemFilter
			};

			var results = await connection.QueryAsync<RawStockInfo>(sql, parameters);

			// If part dimensions are provided, select the smallest sheet that fits
			if (partLength <= 0 || partWidth <= 0)
				return results.FirstOrDefault();

			var rawStockInfos = results as RawStockInfo[] ?? results.ToArray();
			var fittingSheets = GetFittingSheets(rawStockInfos, partLength, partWidth);

			return fittingSheets.Count != 0
				? GetOptimalSheet(fittingSheets)
				:
				// Fallback: return the first result if no part dimensions or no fitting sheets
				rawStockInfos.FirstOrDefault();
		}
		catch (Exception ex)
		{
			Debug.WriteLine(
				$"Error retrieving raw stock for thickness {thickness} and material {material}: {ex.Message}");
			return null;
		}
	}

	/// <summary>
	///     Parses sheet dimensions from string values to doubles with fallback
	/// </summary>
	/// <param name="sheet">Raw stock info containing length and width as strings</param>
	/// <returns>Tuple of parsed length and width (double.MaxValue if parsing fails)</returns>
	private static (double length, double width) ParseSheetDimensions(RawStockInfo sheet)
	{
		var length = double.TryParse(sheet.Length, out var l) ? l : double.MaxValue;
		var width  = double.TryParse(sheet.Width, out var w) ? w : double.MaxValue;
		return (length, width);
	}

	/// <summary>
	///     Checks if a part fits within a sheet in either orientation
	/// </summary>
	/// <param name="sheetLength">Sheet length</param>
	/// <param name="sheetWidth">Sheet width</param>
	/// <param name="partLength">Part length</param>
	/// <param name="partWidth">Part width</param>
	/// <returns>True if part fits in either orientation</returns>
	private static bool DoesPartFit(double sheetLength, double sheetWidth, double partLength, double partWidth)
	{
		return (sheetLength >= partLength && sheetWidth >= partWidth) ||
		       (sheetLength >= partWidth && sheetWidth >= partLength);
	}

	/// <summary>
	///     Calculates the area of a sheet
	/// </summary>
	/// <param name="length">Sheet length</param>
	/// <param name="width">Sheet width</param>
	/// <returns>Sheet area</returns>
	private static double GetSheetArea(double length, double width)
	{
		return length * width;
	}

	/// <summary>
	///     Filters sheets that can accommodate the specified part dimensions
	/// </summary>
	/// <param name="sheets">Collection of raw stock sheets</param>
	/// <param name="partLength">Part length in inches</param>
	/// <param name="partWidth">Part width in inches</param>
	/// <returns>List of sheets that can fit the part</returns>
	private static List<RawStockInfo> GetFittingSheets(IEnumerable<RawStockInfo> sheets, double partLength,
		double partWidth)
	{
		return sheets.Where(sheet =>
		{
			var (sheetLength, sheetWidth) = ParseSheetDimensions(sheet);
			return DoesPartFit(sheetLength, sheetWidth, partLength, partWidth);
		}).ToList();
	}

	/// <summary>
	///     Selects the optimal sheet from fitting sheets based on area efficiency
	/// </summary>
	/// <param name="fittingSheets">List of sheets that can fit the part</param>
	/// <returns>The most efficient sheet (smallest area, then length, then width)</returns>
	private static RawStockInfo GetOptimalSheet(List<RawStockInfo> fittingSheets)
	{
		return fittingSheets
		       .OrderBy(s =>
		       {
			       var (length, width) = ParseSheetDimensions(s);
			       return GetSheetArea(length, width);
		       })
		       .ThenBy(s =>
		       {
			       var (length, _) = ParseSheetDimensions(s);
			       return length;
		       })
		       .ThenBy(s =>
		       {
			       var (_, width) = ParseSheetDimensions(s);
			       return width;
		       })
		       .First();
	}
}