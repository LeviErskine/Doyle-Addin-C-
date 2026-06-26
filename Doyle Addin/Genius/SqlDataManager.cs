namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Threading.Tasks;

public interface ISqlDataManager
{
	Task<Dictionary<string, string>> GetSqlDataAsync(string partNumber, CancellationToken cancellationToken = default);
	Task<List<string>> GetCostCentersAsync(bool partsOnly = true, CancellationToken cancellationToken = default);

	Task<List<Dictionary<string, string>>> GetStockMaterialAsync(string stockType,
		Dictionary<string, string> dimensions,
		string material, CancellationToken cancellationToken = default);

	Task<List<Dictionary<string, string>>> GetAllStockMaterialsAsync(CancellationToken cancellationToken = default);
	Task<List<Dictionary<string, string>>> GetAllItemsAsync(CancellationToken cancellationToken = default);
}

public class SqlDataManager(string connectionString) : ISqlDataManager
{
	private const string Query = """
	                             SELECT
	                                 i.Item,
	                                 i.Description1,
	                                 i.Weight,
	                                 i.Thickness,
	                                 i.Width,
	                                 i.Length,
	                                 i.Diameter,
	                                 i.Family,
	                                 b.Item as RM,
	                                 b.ConversionUnit as RMUNIT,
	                                 b.QuantityInConversionUnit as RMQTY
	                             FROM vgMfiItems i
	                             LEFT JOIN vgIcoBillOfMaterials b ON i.Item = b.Product
	                             WHERE i.Item = @PartNumber
	                             """;

	private const string CostCenterQuery = """
	                                       SELECT Family
	                                       FROM vgMfiFamilies
	                                       WHERE FamilyGroup = 'PARTS'
	                                         AND Active = 'TRUE'
	                                       ORDER BY Family
	                                       """;

	private const string AllCostCenterQuery = """
	                                          SELECT Family
	                                          FROM vgMfiFamilies
	                                          WHERE Active = 'TRUE'
	                                          ORDER BY Family
	                                          """;

	private const string StockMaterialQuery = """
	                                          SELECT
	                                              i.Item as RM,
	                                              i.Description1,
	                                              i.Weight,
	                                              i.Thickness,
	                                              i.Width,
	                                              i.Length,
	                                              i.Diameter,
	                                              i.Family,
	                                              b.ConversionUnit as RMUNIT,
	                                              b.QuantityInConversionUnit as RMQTY
	                                          FROM vgMfiItems i
	                                          LEFT JOIN vgIcoBillOfMaterials b ON i.Item = b.Product
	                                          WHERE i.Item LIKE @ItemPattern OR i.Item LIKE @AltPattern
	                                          ORDER BY i.Item
	                                          """;

	private const string AllStockMaterialsQuery = """
	                                              SELECT
	                                                  i.Item as RM,
	                                                  i.Description1,
	                                                  i.Weight,
	                                                  i.Thickness,
	                                                  i.Width,
	                                                  i.Length,
	                                                  i.Diameter,
	                                                  i.Family,
	                                                  b.ConversionUnit as RMUNIT,
	                                                  b.QuantityInConversionUnit as RMQTY
	                                              FROM vgMfiItems i
	                                              LEFT JOIN vgIcoBillOfMaterials b ON i.Item = b.Product
	                                              WHERE i.Family LIKE 'D-BAR' OR i.Family LIKE 'D-PCUT' OR i.Family LIKE 'DSHEET' OR i.Family LIKE 'R-BAR' OR i.Family LIKE 'R-PCUT' OR i.Family LIKE 'RSHEET'
	                                              ORDER BY i.Item
	                                              """;

	private const string AllItemsQuery = """
	                                     SELECT
	                                         i.Item as RM,
	                                         i.Description1,
	                                         i.Weight,
	                                         i.Thickness,
	                                         i.Width,
	                                         i.Length,
	                                         i.Diameter,
	                                         i.Family
	                                     FROM vgMfiItems i
	                                     ORDER BY i.Item
	                                     """;

	private readonly string _connectionString = !string.IsNullOrWhiteSpace(connectionString)
		? connectionString
		: ConnectionStringProvider.GetConnectionString();

	public async Task<Dictionary<string, string>> GetSqlDataAsync(string partNumber,
		CancellationToken cancellationToken = default)
	{
		if (string.IsNullOrWhiteSpace(partNumber))
			throw new ArgumentNullException(nameof(partNumber), "Part number cannot be null or empty.");

		var results = await ExecuteQueryAsync(Query, $"properties for part {partNumber}",
			cmd => cmd.Parameters.Add("@PartNumber", SqlDbType.NVarChar, -1).Value = partNumber,
			cancellationToken).ConfigureAwait(false);

		return results.Count > 0 ? results[0] : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
	}

	public async Task<List<string>> GetCostCentersAsync(bool partsOnly = true,
		CancellationToken cancellationToken = default)
	{
		var query = partsOnly ? CostCenterQuery : AllCostCenterQuery;
		var results = await ExecuteQueryAsync(query, "cost centers", null, cancellationToken)
			.ConfigureAwait(false);

		return
		[
			.. results
			   .Select(row => row.GetValueOrDefault("Family", string.Empty))
			   .Where(family => !string.IsNullOrWhiteSpace(family))
		];
	}

	public async Task<List<Dictionary<string, string>>> GetStockMaterialAsync(string stockType,
		Dictionary<string, string> dimensions, string material, CancellationToken cancellationToken = default)
	{
		if (string.IsNullOrWhiteSpace(stockType))
			throw new ArgumentNullException(nameof(stockType), "Stock type cannot be null or empty.");

		if (dimensions == null || dimensions.Count == 0)
			throw new ArgumentNullException(nameof(dimensions), "Dimensions cannot be null or empty.");

		var results = await QueryStockMaterialAsync(stockType, dimensions, material, cancellationToken)
			.ConfigureAwait(false);

		if (results.Count > 0)
			return results;

		if (stockType.Equals("BarStock", StringComparison.OrdinalIgnoreCase) &&
		    dimensions.TryGetValue("Diameter", out var diameterStr) &&
		    double.TryParse(diameterStr, out var finishedDiameter))
			return await QueryBarStockWithDimensionFallbacksAsync(dimensions, material, finishedDiameter, diameterStr,
				"Diameter", "BarStock", cancellationToken).ConfigureAwait(false);

		if ((stockType.Equals("SquareBar", StringComparison.OrdinalIgnoreCase) ||
		     stockType.Equals("RectangleBar", StringComparison.OrdinalIgnoreCase)) &&
		    dimensions.TryGetValue("Width", out var widthStr) &&
		    double.TryParse(widthStr, out var finishedWidth))
			return await QueryBarStockWithDimensionFallbacksAsync(dimensions, material, finishedWidth, widthStr,
				"Width", stockType, cancellationToken).ConfigureAwait(false);

		return results;
	}

	public async Task<List<Dictionary<string, string>>> GetAllStockMaterialsAsync(
		CancellationToken cancellationToken = default)
	{
		return await ExecuteQueryAsync(AllStockMaterialsQuery, "total stock materials", null, cancellationToken)
			.ConfigureAwait(false);
	}

	public async Task<List<Dictionary<string, string>>> GetAllItemsAsync(CancellationToken cancellationToken = default)
	{
		return await ExecuteQueryAsync(AllItemsQuery, "total items", null, cancellationToken)
			.ConfigureAwait(false);
	}

	private async Task<List<Dictionary<string, string>>> ExecuteQueryAsync(
		string query,
		string operationName,
		Action<SqlCommand> configureCommand,
		CancellationToken cancellationToken)
	{
		var results = new List<Dictionary<string, string>>();
		try
		{
			await using var connection = new SqlConnection(_connectionString);
			await using var command    = new SqlCommand(query, connection);
			configureCommand?.Invoke(command);

			await connection.OpenAsync(cancellationToken).ConfigureAwait(false);
			await using var reader = await command.ExecuteReaderAsync(cancellationToken).ConfigureAwait(false);

			while (await reader.ReadAsync(cancellationToken).ConfigureAwait(false))
			{
				var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
				for (var i = 0; i < reader.FieldCount; i++)
					row[reader.GetName(i)] = reader.IsDBNull(i) ? string.Empty : reader.GetValue(i).ToString();
				results.Add(row);
			}

			Debug.WriteLine($"SqlDataManager: Retrieved {results.Count} {operationName}");
		}
		catch (OperationCanceledException)
		{
			Debug.WriteLine($"SqlDataManager: {operationName} cancelled");
			throw;
		}
		catch (SqlException ex)
		{
			Debug.WriteLine($"SqlDataManager: SQL error {operationName}: {ex.Message}");
			throw;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"SqlDataManager: Unexpected error {operationName}: {ex.Message}");
			throw;
		}

		return results;
	}

	private async Task<List<Dictionary<string, string>>> QueryStockMaterialAsync(string stockType,
		Dictionary<string, string> dimensions, string material, CancellationToken cancellationToken)
	{
		var (itemPattern, altPattern) = BuildItemPattern(stockType, dimensions, material);
		if (string.IsNullOrEmpty(itemPattern))
		{
			Debug.WriteLine($"SqlDataManager: Could not build item pattern for stock type {stockType}");
			return [];
		}

		var results = await ExecuteQueryAsync(StockMaterialQuery, $"stock material for {stockType}",
			cmd =>
			{
				cmd.Parameters.Add("@ItemPattern", SqlDbType.NVarChar, -1).Value = itemPattern;
				cmd.Parameters.Add("@AltPattern", SqlDbType.NVarChar, -1).Value  = altPattern ?? string.Empty;
			},
			cancellationToken).ConfigureAwait(false);

		if (results.Count == 0)
			Debug.WriteLine($"SqlDataManager: No stock material found for pattern {itemPattern}");

		return results;
	}

	private async Task<List<Dictionary<string, string>>> QueryBarStockWithDimensionFallbacksAsync(
		Dictionary<string, string> dimensions, string material, double finishedDimension, string exactDimensionText,
		string dimensionKey, string stockType, CancellationToken cancellationToken)
	{
		const double tolerance        = 0.001;
		var          merged           = new List<Dictionary<string, string>>();
		var          seenRm           = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var          fallbacks        = GetBarStockDimensionFallbacks(finishedDimension, exactDimensionText);
		var          truncatedCatalog = Math.Truncate(finishedDimension * 100.0) / 100.0;

		for (var index = 0; index < fallbacks.Count; index++)
		{
			var stockDimension = fallbacks[index];
			var fallbackResults = await QueryBarStockByDimensionAsync(dimensions, material, stockDimension,
				finishedDimension, dimensionKey, stockType, cancellationToken).ConfigureAwait(false);
			if (fallbackResults.Count == 0) continue;

			MergeStockMaterialResults(merged, seenRm, fallbackResults);

			var isTruncatedCatalogTier = truncatedCatalog > 0 &&
			                             Math.Abs(stockDimension - truncatedCatalog) < tolerance;
			if (!isTruncatedCatalogTier)
				break;

			for (var j = index + 1; j < fallbacks.Count; j++)
			{
				if (fallbacks[j] <= finishedDimension + tolerance) continue;

				var nextSizeResults = await QueryBarStockByDimensionAsync(dimensions, material, fallbacks[j],
					finishedDimension, dimensionKey, stockType, cancellationToken).ConfigureAwait(false);
				MergeStockMaterialResults(merged, seenRm, nextSizeResults);
				break;
			}

			break;
		}

		return merged;
	}

	private async Task<List<Dictionary<string, string>>> QueryBarStockByDimensionAsync(
		Dictionary<string, string> dimensions, string material, double stockDimension, double finishedDimension,
		string dimensionKey, string stockType, CancellationToken cancellationToken)
	{
		var fallbackDimensions = new Dictionary<string, string>(dimensions, StringComparer.OrdinalIgnoreCase)
		{
			[dimensionKey] = stockDimension.ToString("F4")
		};

		var fallbackResults = await QueryStockMaterialAsync(stockType, fallbackDimensions, material, cancellationToken)
			.ConfigureAwait(false);
		if (fallbackResults.Count == 0) return fallbackResults;

		foreach (var row in fallbackResults)
			row["IsFallbackMatch"] = "true";

		Debug.WriteLine(
			$"SqlDataManager: Found {fallbackResults.Count} {stockType} match(es) using purchase {dimensionKey} {stockDimension:F4}\" (finished {finishedDimension:F4}\")");
		return fallbackResults;
	}

	private static void MergeStockMaterialResults(List<Dictionary<string, string>> merged,
		HashSet<string> seenRm, List<Dictionary<string, string>> batch)
	{
		merged.AddRange(from row in batch
			let rm = row.GetValueOrDefault("RM", string.Empty)
			where string.IsNullOrWhiteSpace(rm) || seenRm.Add(rm)
			select row);
	}

	private static (string Primary, string Alternate) BuildItemPattern(string stockType,
		Dictionary<string, string> dimensions, string material)
	{
		try
		{
			var isStainless = !string.IsNullOrEmpty(material) &&
			                  material.Contains("stainless", StringComparison.OrdinalIgnoreCase);

			return stockType switch
			{
				"BarStock" when dimensions.TryGetValue("Diameter", out var diameter) =>
					BuildBarStockPattern(diameter, isStainless),
				"RoundTube" when dimensions.TryGetValue("OD", out var od) =>
					BuildRoundTubePattern(od, isStainless),
				"SquareBar" when dimensions.TryGetValue("Width", out var sqBarWidth) =>
					BuildRectangleBarPattern(sqBarWidth, sqBarWidth, isStainless),
				"RectangleBar" when dimensions.TryGetValue("Width", out var rectBarWidth) &&
				                    dimensions.TryGetValue("Height", out var rectBarHeight) =>
					BuildRectangleBarPattern(rectBarWidth, rectBarHeight, isStainless),
				"SquareTube" when dimensions.TryGetValue("Width", out var width) =>
					BuildSquareTubePattern(width, isStainless),
				"RectangleTube" when dimensions.TryGetValue("Width", out var rectWidth) &&
				                     dimensions.TryGetValue("Height", out var rectHeight) =>
					BuildRectangleTubePattern(rectWidth, rectHeight, isStainless),
				_ => (null, null)
			};
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"SqlDataManager: Error building item pattern: {ex}");
			return (null, null);
		}
	}

	private static (string Primary, string Alternate) BuildBarStockPattern(string diameter,
		bool isStainless = false)
	{
		var normalizedDiameter = NormalizeDimension(diameter);
		var prefix             = isStainless ? "RS" : "RM";
		var primary            = $"{prefix}%-{normalizedDiameter}%D-%";

		var    altDiameter = GetAlternateForm(normalizedDiameter);
		string alternate   = null;
		if (altDiameter != null)
			alternate = $"{prefix}%-{altDiameter}%D-%";

		return (primary, alternate);
	}

	private static (string Primary, string Alternate) BuildRoundTubePattern(string od,
		bool isStainless = false)
	{
		var normalizedOd = NormalizeDimension(od);
		var prefix       = isStainless ? "TSD" : "TMD";
		var altPrefix    = isStainless ? "PS" : "PM";
		var primary      = $"{prefix}%{normalizedOd}%DX%";
		var alternate    = $"{altPrefix}%{normalizedOd}%X%";
		return (primary, alternate);
	}

	private static (string Primary, string Alternate) BuildRectangleBarPattern(string width, string height,
		bool isStainless = false)
	{
		var normalizedWidth  = NormalizeDimension(width);
		var normalizedHeight = NormalizeDimension(height);
		var prefix           = isStainless ? "BS" : "BM";
		var primary          = $"{prefix}%{normalizedWidth}%X{normalizedHeight}%-%";

		var    altWidth  = GetAlternateForm(normalizedWidth);
		var    altHeight = GetAlternateForm(normalizedHeight);
		string alternate = null;
		if (altWidth != null || altHeight != null)
			alternate = $"{prefix}%{altWidth ?? normalizedWidth}%X{altHeight ?? normalizedHeight}%-%";

		return (primary, alternate);
	}

	private static (string Primary, string Alternate) BuildSquareTubePattern(string width,
		bool isStainless = false)
	{
		var normalizedWidth = NormalizeDimension(width);
		var prefix          = isStainless ? "TS" : "TM";
		var primary         = $"{prefix}%{normalizedWidth}%X{normalizedWidth}%X%";

		var    altWidth  = GetAlternateForm(normalizedWidth);
		string alternate = null;
		if (altWidth != null)
			alternate = $"{prefix}%{altWidth}%X{altWidth}%X%";

		return (primary, alternate);
	}

	private static (string Primary, string Alternate) BuildRectangleTubePattern(string width, string height,
		bool isStainless = false)
	{
		var normalizedWidth  = NormalizeDimension(width);
		var normalizedHeight = NormalizeDimension(height);
		var prefix           = isStainless ? "TS" : "TM";
		var primary          = $"{prefix}%{normalizedWidth}%X{normalizedHeight}%X%";

		var    altWidth  = GetAlternateForm(normalizedWidth);
		var    altHeight = GetAlternateForm(normalizedHeight);
		string alternate = null;
		if (altWidth != null || altHeight != null)
			alternate = $"{prefix}%{altWidth ?? normalizedWidth}%X{altHeight ?? normalizedHeight}%X%";

		return (primary, alternate);
	}

	private static string NormalizeDimension(string dimension)
	{
		if (string.IsNullOrWhiteSpace(dimension))
			return dimension;

		var value     = double.Parse(dimension);
		var formatted = value.ToString("F2");
		if (!formatted.Contains('.')) return formatted;
		formatted = formatted.TrimEnd('0');
		if (formatted.EndsWith('.'))
			formatted = formatted.TrimEnd('.');

		return formatted;
	}

	private static string GetAlternateForm(string normalized)
	{
		if (string.IsNullOrWhiteSpace(normalized))
			return null;

		if (!normalized.Contains('.'))
			return normalized + ".0";

		return null;
	}

	private static List<double> GetBarStockDimensionFallbacks(double finishedDimensionInches, string exactDimensionText)
	{
		const double tolerance  = 0.001;
		var          tried      = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var          candidates = new List<double>();

		if (!string.IsNullOrWhiteSpace(exactDimensionText))
			tried.Add(NormalizeDimension(exactDimensionText));

		var truncatedCatalog = Math.Truncate(finishedDimensionInches * 100.0) / 100.0;
		if (truncatedCatalog > 0)
			TryAdd(truncatedCatalog);

		foreach (var purchaseSize in GetBarStockPurchaseSizesInches(finishedDimensionInches)
			         .Where(purchaseSize => !(purchaseSize <= finishedDimensionInches + tolerance)))
			TryAdd(purchaseSize);

		return candidates;

		void TryAdd(double candidate)
		{
			var key = NormalizeDimension(candidate.ToString("F4"));
			if (!tried.Add(key)) return;
			candidates.Add(candidate);
		}
	}

	private static List<double> GetBarStockPurchaseSizesInches(double finishedDimensionInches)
	{
		const double tolerance  = 0.001;
		var          tried      = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		var          candidates = new List<double>();

		TryAdd(Math.Ceiling(finishedDimensionInches * 16.0) / 16.0);
		TryAdd(Math.Ceiling(finishedDimensionInches * 8.0) / 8.0);
		TryAdd(Math.Ceiling(finishedDimensionInches * 4.0) / 4.0);
		TryAdd(Math.Ceiling(finishedDimensionInches));

		return candidates;

		void TryAdd(double candidate)
		{
			if (candidate <= finishedDimensionInches + tolerance) return;

			var key = NormalizeDimension(candidate.ToString("F4"));
			if (!tried.Add(key)) return;
			candidates.Add(candidate);
		}
	}
}