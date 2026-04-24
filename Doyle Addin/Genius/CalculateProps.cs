namespace DoyleAddin.Genius;

using System.Collections.Generic;
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
	public static Task<Dictionary<string, string>> CalculateAllPropertiesAsync(Document document)
	{
		try
		{
			var properties = new Dictionary<string, string>();

			try
			{
				if (document == null) return Task.FromResult(properties);

				var unitsOfMeasure = document.UnitsOfMeasure;

				// Calculate GeniusMass from Inventor's mass
				CalculateGeniusMass(document, properties);

				// Handle sheet metal-specific properties
				if (document is PartDocument { ComponentDefinition: SheetMetalComponentDefinition smDef } partDoc)
				{
					unitsOfMeasure.LengthDisplayPrecision = 4;

					// Calculate Thickness for sheet metal parts
					var thickness = CalculateThickness(smDef, unitsOfMeasure, properties);

					// Calculate extent properties from flat pattern
					CalculateExtentProperties(smDef, unitsOfMeasure, properties);

					// Calculate raw material properties
					CalculateRawMaterialProperties(partDoc, thickness, smDef, properties);
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Error calculating properties: {ex.Message}");
			}

			return Task.FromResult(properties);
		}
		catch (Exception exception)
		{
			return Task.FromException<Dictionary<string, string>>(exception);
		}
	}

	private static void CalculateRawMaterialProperties(PartDocument partDoc, double thickness,
		SheetMetalComponentDefinition smDef, Dictionary<string, string> properties)
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

			var gaugeName       = GetSheetMetalGaugeName(smDef);
			var rawMaterialData = FindMatchingRawMaterial(material, gaugeName, thickness);

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
					$"CalculateRawMaterialProperties: No matching raw material found for {material}, gauge {gaugeName}, thickness {thickness}");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in CalculateRawMaterialProperties: {ex.Message}");
		}
	}

	private static string GetSheetMetalGaugeName(SheetMetalComponentDefinition smDef)
	{
		try
		{
			// Try to get the active sheet metal style name which typically contains gauge info
			var style = smDef.SheetMetalStyles[smDef.ActiveSheetMetalStyle];
			return style?.Name ?? string.Empty;
		}
		catch
		{
			return string.Empty;
		}
	}

	private static void CalculateGeniusMass(Document document, Dictionary<string, string> properties)
	{
		try
		{
			var massProperties = document switch
			{
				PartDocument partDoc     => partDoc.ComponentDefinition.MassProperties,
				AssemblyDocument assyDoc => assyDoc.ComponentDefinition.MassProperties,
				_                        => null
			};

			if (massProperties == null) return;
			var massKg = massProperties.Mass;
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

			// Try common thickness parameter names
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

	private static Dictionary<string, string> FindMatchingRawMaterial(string material, string gaugeName,
		double thickness)
	{
		// Hardcoded mappings based on VBA pnShtMetalHardCoded function
		var    materialKey = material.ToLowerInvariant();
		string rmValue     = null;

		if (materialKey.Contains("stainless"))
			rmValue = gaugeName switch
			{
				_ when gaugeName.Contains("18")   => "FS-48x96x0.048",
				_ when gaugeName.Contains("14")   => "FS-60x120x0.075",
				_ when gaugeName.Contains("13")   => "FS-60x97x0.09",
				_ when gaugeName.Contains("12")   => "FS-60x120x0.105",
				_ when gaugeName.Contains("10")   => "FS-60x144x0.135",
				_ when gaugeName.Contains("3/16") => "FS-60x144x0.188",
				_ when gaugeName.Contains("1/4")  => "FS-60x144x0.25",
				_ when gaugeName.Contains("5/16") => "FS-60x144x0.313",
				_ when gaugeName.Contains("3/8")  => "FS-60x144x0.375",
				_ when gaugeName.Contains("1/2")  => "FS-60x144x0.5",
				_                                 => MatchStainlessSteelByThickness(thickness)
			};
		else if (materialKey.Contains("mild"))
			rmValue = gaugeName switch
			{
				_ when gaugeName.Contains("14")   => "FM-60x144x0.075",
				_ when gaugeName.Contains("12")   => "FM-60x144x0.105",
				_ when gaugeName.Contains("10")   => "FM-60x144x0.135",
				_ when gaugeName.Contains("3/16") => "FM-60x144x0.188",
				_ when gaugeName.Contains("1/4")  => "FM-60x144x0.25",
				_ when gaugeName.Contains("5/16") => "FM-60x144x0.313",
				_ when gaugeName.Contains("3/8")  => "FM-60x144x0.375",
				_ when gaugeName.Contains("1/2")  => "FM-60x144x0.5",
				_ when gaugeName.Contains("5/8")  => "FM-60x144x0.625",
				_ when gaugeName.Contains("3/4")  => "FM-60x120x0.75",
				_ when gaugeName.Contains("1") && !gaugeName.Contains("1/2") && !gaugeName.Contains("1/4")
					=> "FM-48x120x1",
				_ => MatchMildSteelByThickness(thickness)
			};
		else if (materialKey.Contains("rubber")) rmValue = "LG";

		if (string.IsNullOrEmpty(rmValue))
			return null;

		return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			["RM"]     = rmValue,
			["RMUNIT"] = "FT2"
		};
	}

	private static string MatchStainlessSteelByThickness(double thickness)
	{
		const double tolerance = 0.005;
		return thickness switch
		{
			_ when Math.Abs(thickness - 0.048) <= tolerance => "FS-48x96x0.048",
			_ when Math.Abs(thickness - 0.075) <= tolerance => "FS-60x120x0.075",
			_ when Math.Abs(thickness - 0.090) <= tolerance => "FS-60x97x0.09",
			_ when Math.Abs(thickness - 0.105) <= tolerance => "FS-60x120x0.105",
			_ when Math.Abs(thickness - 0.135) <= tolerance => "FS-60x144x0.135",
			_ when Math.Abs(thickness - 0.188) <= tolerance => "FS-60x144x0.188",
			_ when Math.Abs(thickness - 0.250) <= tolerance => "FS-60x144x0.25",
			_ when Math.Abs(thickness - 0.313) <= tolerance => "FS-60x144x0.313",
			_ when Math.Abs(thickness - 0.375) <= tolerance => "FS-60x144x0.375",
			_ when Math.Abs(thickness - 0.500) <= tolerance => "FS-60x144x0.5",
			_                                               => null
		};
	}

	private static string MatchMildSteelByThickness(double thickness)
	{
		const double tolerance = 0.005;
		return thickness switch
		{
			_ when Math.Abs(thickness - 0.075) <= tolerance => "FM-60x144x0.075",
			_ when Math.Abs(thickness - 0.105) <= tolerance => "FM-60x144x0.105",
			_ when Math.Abs(thickness - 0.135) <= tolerance => "FM-60x144x0.135",
			_ when Math.Abs(thickness - 0.188) <= tolerance => "FM-60x144x0.188",
			_ when Math.Abs(thickness - 0.250) <= tolerance => "FM-60x144x0.25",
			_ when Math.Abs(thickness - 0.313) <= tolerance => "FM-60x144x0.313",
			_ when Math.Abs(thickness - 0.375) <= tolerance => "FM-60x144x0.375",
			_ when Math.Abs(thickness - 0.500) <= tolerance => "FM-60x144x0.5",
			_ when Math.Abs(thickness - 0.625) <= tolerance => "FM-60x144x0.625",
			_ when Math.Abs(thickness - 0.750) <= tolerance => "FM-60x120x0.75",
			_ when Math.Abs(thickness - 1.000) <= tolerance => "FM-48x120x1",
			_                                               => null
		};
	}
}