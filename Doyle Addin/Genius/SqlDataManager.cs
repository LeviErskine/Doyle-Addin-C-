namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading.Tasks;

/// <summary>
///     Interface for managing SQL Server connections and data retrieval.
/// </summary>
public interface ISqlDataManager
{
	/// <summary>
	///     Gets matching data from SQL Server based on part number.
	/// </summary>
	/// <param name="partNumber">The part number to search for.</param>
	/// <returns>Dictionary of property names and SQL values.</returns>
	Task<Dictionary<string, string>> GetSqlDataAsync(string partNumber);
}

/// <summary>
///     Manages SQL Server connections and data retrieval for property comparison.
/// </summary>
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

	private readonly string _connectionString = string.IsNullOrWhiteSpace(connectionString)
		? Geniusinfo.DefaultConnectionString
		: connectionString;

	/// <summary>
	///     Gets matching data from SQL Server based on part number.
	/// </summary>
	/// <param name="partNumber">The part number to search for.</param>
	/// <returns>Dictionary of property names and SQL values.</returns>
	public async Task<Dictionary<string, string>> GetSqlDataAsync(string partNumber)
	{
		var sqlData = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

		if (string.IsNullOrWhiteSpace(partNumber))
		{
			Debug.WriteLine("SqlDataManager: Part number is null or empty");
			return sqlData;
		}

		try
		{
			await using var connection = new SqlConnection(_connectionString);
			await using var command    = new SqlCommand(Query, connection);
			command.Parameters.Add("@PartNumber", SqlDbType.NVarChar).Value = partNumber;

			await connection.OpenAsync();
			await using var reader = await command.ExecuteReaderAsync();

			if (await reader.ReadAsync())
			{
				for (var i = 0; i < reader.FieldCount; i++)
				{
					var columnName = reader.GetName(i);
					var value      = reader.GetValue(i).ToString() ?? string.Empty;
					sqlData[columnName] = value;
				}

				Debug.WriteLine($"SqlDataManager: Retrieved {sqlData.Count} properties for part {partNumber}");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"SqlDataManager: Error for part {partNumber}: {ex.Message}");
		}

		return sqlData;
	}
}