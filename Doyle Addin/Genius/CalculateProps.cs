namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading.Tasks;
using Inventor;

/// <summary>
///     Calculates properties for parts/assemblies that are new or have been updated.
/// </summary>
public static class CalculateProps
{
	/// <summary>
	///     Calculates all required properties for the active document.
	/// </summary>
	/// <returns>A dictionary containing calculated property names and values.</returns>
	public static async Task<Dictionary<string, string>> CalculateAllPropertiesAsync()
	{
		return ThisApplication?.ActiveDocument is Document document ? await CalculateAllPropertiesAsync(document) : [];
	}

	/// <summary>
	///     Calculates all required properties for the specified document.
	/// </summary>
	/// <param name="document">The Inventor document to calculate properties for.</param>
	/// <returns>A dictionary containing calculated property names and values.</returns>
	public static async Task<Dictionary<string, string>> CalculateAllPropertiesAsync(Document document)
	{
		var properties = new Dictionary<string, string>();

		try
		{
			if (document == null) return properties;

			var unitsOfMeasure = document.UnitsOfMeasure;

			// Calculate GeniusMass from Inventor's mass
			CalculateGeniusMass(document, properties);

			// Handle sheet metal specific properties
			if (document is PartDocument { ComponentDefinition: SheetMetalComponentDefinition smDef } partDoc)
			{
				unitsOfMeasure.LengthDisplayPrecision = 4;

				// Calculate Thickness for sheet metal parts
				var thickness = CalculateThickness(smDef, unitsOfMeasure, properties);

				// Calculate extent properties from flat pattern
				CalculateExtentProperties(smDef, unitsOfMeasure, properties);

				// Calculate raw material properties
				await CalculateRawMaterialPropertiesAsync(partDoc, thickness, properties);
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating properties: {ex.Message}");
		}

		return properties;
	}

	private static async Task CalculateRawMaterialPropertiesAsync(PartDocument partDoc, double thickness,
		Dictionary<string, string> properties)
	{
		try
		{
			var material = GetMaterialFromPart(partDoc);
			properties["Cost Center"] = "D-RMTO";

			if (string.IsNullOrEmpty(material) || thickness <= 0)
			{
				Debug.WriteLine("CalculateRawMaterialProperties: Missing material or thickness information");
				return;
			}

			var rawMaterialData = await FindMatchingRawMaterialAsync(material, thickness);
			if (rawMaterialData is { Count: > 0 })
			{
				if (rawMaterialData.TryGetValue("RM", out var rmValue))
					properties["RM"] = rmValue;
				if (rawMaterialData.TryGetValue("RMUNIT", out var rmUnitValue))
					properties["RMUNIT"] = rmUnitValue;

				if (properties.TryGetValue("Extent_Area", out var extentAreaStr) &&
				    double.TryParse(extentAreaStr, out var extentAreaIn2))
					properties["RMQTY"] = (extentAreaIn2 / 144.0).ToString("F4");

				Debug.WriteLine(
					$"CalculateRawMaterialProperties: Found matching raw material: {rawMaterialData.GetValueOrDefault("RM", "N/A")}");
			}
			else
			{
				Debug.WriteLine(
					$"CalculateRawMaterialProperties: No matching raw material found for {material}, thickness {thickness}");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in CalculateRawMaterialProperties: {ex.Message}");
		}
	}

	private static void CalculateGeniusMass(Document document, Dictionary<string, string> properties)
	{
		try
		{
			var massKg = document switch
			{
				PartDocument partDoc     => partDoc.ComponentDefinition.MassProperties.Mass,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.MassProperties.Mass,
				_                        => 0
			};

			if (massKg <= 0) return;

			var unitsOfMeasure = document.UnitsOfMeasure;
			var massLbs        = unitsOfMeasure.ConvertUnits(massKg, "kg", "lb");
			properties["GeniusMass"] = massLbs.ToString("F4");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating GeniusMass: {ex.Message}");
		}
	}

	private static double CalculateThickness(SheetMetalComponentDefinition smDef, UnitsOfMeasure unitsOfMeasure,
		Dictionary<string, string> properties)
	{
		try
		{
			var parameters = smDef.Parameters;
			if (parameters == null) return 0;

			// Try common thickness parameter names without multiple try-catches
			var paramNames = new[] { "Thickness", "SheetMetalThickness", "MaterialThickness" };
			foreach (var name in paramNames)
				try
				{
					var param = parameters[name];
					if (param == null) continue;
					var thicknessValue    = Convert.ToDouble(param.Value);
					var thicknessInInches = unitsOfMeasure.ConvertUnits(thicknessValue, "cm", "in");
					properties["Thickness"] = thicknessInInches.ToString("F4") + " in";
					return thicknessInInches;
				}
				catch
				{
					/* skip to next name */
				}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating Thickness: {ex.Message}");
		}

		return 0;
	}

	private static void CalculateExtentProperties(SheetMetalComponentDefinition smDef, UnitsOfMeasure unitsOfMeasure,
		Dictionary<string, string> properties)
	{
		try
		{
			var flatPattern = smDef.FlatPattern;
			if (flatPattern == null) return;

			var width  = flatPattern.Width;
			var length = flatPattern.Length;
			var area   = width * length;

			properties["Extent_Width"]  = unitsOfMeasure.ConvertUnits(width, "cm", "in").ToString("F4");
			properties["Extent_Length"] = unitsOfMeasure.ConvertUnits(length, "cm", "in").ToString("F4");
			properties["Extent_Area"]   = unitsOfMeasure.ConvertUnits(area, "cm^2", "in^2").ToString("F4");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating extent properties: {ex.Message}");
		}
	}

	private static string GetMaterialFromPart(PartDocument partDoc)
	{
		try
		{
			return partDoc.ComponentDefinition.Material?.Name ?? string.Empty;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error getting material from part: {ex.Message}");
			return string.Empty;
		}
	}

	private static async Task<Dictionary<string, string>> FindMatchingRawMaterialAsync(string material,
		double thickness)
	{
		const double thicknessTolerance = 0.010;

		try
		{
			const string query = """
			                     SELECT TOP 1
			                         Item, 
			                         Unit, 
			                         Item as RM,
			                         'FT2' as RMUNIT
			                     FROM vgMfiItems
			                     WHERE (Item LIKE 'FS%' OR Item LIKE 'FM%')
			                     AND Specification6 LIKE @Material 
			                     AND Thickness BETWEEN @ThicknessMin AND @ThicknessMax
			                     AND (OnHand > 0 OR OnOrder > 0)
			                     ORDER BY ABS(Thickness - @Thickness), Width * Length 
			                     """;

			await using var connection = new SqlConnection(Geniusinfo.DefaultConnectionString);
			await using var command    = new SqlCommand(query, connection);

			var materialParam = material.ToLowerInvariant() switch
			{
				var m when m.Contains("stainless") => "%SS%",
				var m when m.Contains("mild")      => "%MS%",
				_                                  => $"%{material}%"
			};

			command.Parameters.Add("@Material", SqlDbType.NVarChar).Value  = materialParam;
			command.Parameters.Add("@ThicknessMin", SqlDbType.Float).Value = thickness - thicknessTolerance;
			command.Parameters.Add("@ThicknessMax", SqlDbType.Float).Value = thickness + thicknessTolerance;
			command.Parameters.Add("@Thickness", SqlDbType.Float).Value    = thickness;

			await connection.OpenAsync();
			await using var reader = await command.ExecuteReaderAsync();

			if (!await reader.ReadAsync()) return null;

			var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			for (var i = 0; i < reader.FieldCount; i++)
				result[reader.GetName(i)] = reader.GetValue(i).ToString() ?? string.Empty;
			return result;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in FindMatchingRawMaterial: {ex.Message}");
			return null;
		}
	}
}