namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading.Tasks;

public interface ISqlDataManager
{
	Task<Dictionary<string, string>> GetSqlDataAsync(string partNumber);
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

	private readonly string _connectionString = string.IsNullOrWhiteSpace(connectionString)
		? GeniusConstants.DefaultConnectionString
		: connectionString;

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
					sqlData[reader.GetName(i)] = reader.GetValue(i).ToString() ?? string.Empty;

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