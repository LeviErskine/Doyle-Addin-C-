namespace DoyleAddin.Genius;

using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Forms;
using My_Project;

/// <summary>
///     Calculates properties for parts/assemblies that are new or have been updated.
/// </summary>
internal static class CalculateProps
{
	private const double ThicknessTolerance = 0.005;
	private const double SquareInchesPerSquareFoot = 144.0;
	private const string DefaultCostCenter = "D-RMTO";
	private const string DefaultRmUnit = "FT2";
	private static ISqlDataManager _sqlDataManager;

	private static readonly string[] ThicknessParameterNames =
		["Thickness", "SheetMetalThickness", "MaterialThickness"];

	private static readonly Dictionary<string, string> StainlessSteelGaugeMap = new(StringComparer.OrdinalIgnoreCase)
	{
		["18"]   = "FS-48X96X0.048",
		["14"]   = "FS-60X120X0.075",
		["13"]   = "FS-60X97X0.090",
		["12"]   = "FS-60X120X0.105",
		["10"]   = "FS-60X144X0.135",
		["3/16"] = "FS-60X144X0.188",
		["1/4"]  = "FS-60X144X0.250",
		["5/16"] = "FS-60X144X0.313",
		["3/8"]  = "FS-60X144X0.375",
		["1/2"]  = "FS-60X144X0.500"
	};

	private static readonly Dictionary<string, string> MildSteelGaugeMap = new(StringComparer.OrdinalIgnoreCase)
	{
		["14"]   = "FM-60X144X0.075",
		["12"]   = "FM-60X144X0.105",
		["10"]   = "FM-60X144X0.135",
		["3/16"] = "FM-60X144X0.188",
		["1/4"]  = "FM-60X144X0.250",
		["5/16"] = "FM-60X144X0.313",
		["3/8"]  = "FM-60X144X0.375",
		["1/2"]  = "FM-60X144X0.500",
		["5/8"]  = "FM-60X144X0.625",
		["3/4"]  = "FM-60X120X0.750",
		["1"]    = "FM-48X120X1.000"
	};

	public static void SetSqlDataManager(ISqlDataManager sqlDataManager)
	{
		_sqlDataManager = sqlDataManager;
	}

	/// <summary>
	///     Calculates all required properties for the active document.
	/// </summary>
	/// <returns>A dictionary containing calculated property names and values.</returns>
	public static Task<Dictionary<string, string>> CalculateAllPropertiesAsync(
		CancellationToken cancellationToken = default)
	{
		return ThisApplication?.ActiveDocument is Document document
			? CalculateAllPropertiesAsync(document, cancellationToken)
			: Task.FromResult(new Dictionary<string, string>());
	}

	/// <summary>
	///     Calculates all required properties for the specified document.
	/// </summary>
	/// <param name="document">The Inventor document to calculate properties for.</param>
	/// <param name="cancellationToken"></param>
	/// <returns>A dictionary containing calculated property names and values.</returns>
	public static async Task<Dictionary<string, string>> CalculateAllPropertiesAsync(Document document,
		CancellationToken cancellationToken = default)
	{
		var properties = new Dictionary<string, string>();

		if (document == null)
			return properties;

		if (IsComponentSuppressed(document))
		{
			Debug.WriteLine($"CalculateAllPropertiesAsync: Skipping suppressed component: {document.DisplayName}");
			return properties;
		}

		cancellationToken.ThrowIfCancellationRequested();

		try
		{
			var unitsOfMeasure = document.UnitsOfMeasure;

			// Calculate GeniusMass from Inventor's mass
			CalculateGeniusMass(document, properties);

			if (IsPurchasedPart(document))
				return properties;

			cancellationToken.ThrowIfCancellationRequested();

			switch (document)
			{
				// Handle sheet metal-specific properties
				case PartDocument { ComponentDefinition: SheetMetalComponentDefinition smDef } partDoc:
				{
					unitsOfMeasure.LengthDisplayPrecision = 4;

					// Calculate Thickness for sheet metal parts
					var thickness = CalculateThickness(smDef, unitsOfMeasure, properties);

					// Ensure flat pattern exists before calculating extent properties
					EnsureFlatPattern(smDef);

					// Calculate extent properties from flat pattern
					CalculateExtentProperties(smDef, unitsOfMeasure, properties);

					// Calculate raw material properties
					CalculateRawMaterialProperties(partDoc, thickness, smDef, properties);
					break;
				}
				// Handle non-sheet metal parts that may be stock materials
				case PartDocument { ComponentDefinition: { } partDef } nonSheetPartDoc:
					unitsOfMeasure.LengthDisplayPrecision = 4;

					cancellationToken.ThrowIfCancellationRequested();

					// Detect and calculate stock material properties
					await CalculateStockMaterialProperties(nonSheetPartDoc, partDef, unitsOfMeasure, properties,
						cancellationToken);
					break;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating properties: {ex}");
		}

		return properties;
	}

	private static void CalculateRawMaterialProperties(PartDocument partDoc, double thickness,
		SheetMetalComponentDefinition smDef, Dictionary<string, string> properties)
	{
		try
		{
			var material = GetMaterialFromPart(partDoc);
			properties["Cost Center"] = DefaultCostCenter;

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
					properties["RMQTY"] = (extentAreaIn2 / SquareInchesPerSquareFoot).ToString("F4");

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
			Debug.WriteLine($"Error in CalculateRawMaterialProperties: {ex}");
		}
	}

	private static void EnsureFlatPattern(SheetMetalComponentDefinition smDef)
	{
		try
		{
			if (smDef.HasFlatPattern) return;
			smDef.Unfold();
			smDef.FlatPattern.ExitEdit();
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error creating flat pattern: {ex}");
		}
	}

	private static string GetSheetMetalGaugeName(SheetMetalComponentDefinition smDef)
	{
		try
		{
			// Try to get the active sheet metal style name which typically contains gauge info
			var style = smDef.ActiveSheetMetalStyle;
			return style?.Name ?? string.Empty;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error getting sheet metal gauge name: {ex}");
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
			Debug.WriteLine($"Error calculating GeniusMass: {ex}");
		}
	}

	private static bool IsComponentSuppressed(Document document)
	{
		try
		{
			return document switch
			{
				PartDocument { ComponentDefinition: { } partDef } => partDef.SurfaceBodies.Count == 0,
				_                                                 => false
			};
		}
		catch
		{
			return true;
		}
	}

	private static bool IsPurchasedPart(Document document)
	{
		try
		{
			return document switch
			{
				PartDocument { ComponentDefinition: { } partDef } => partDef.BOMStructure ==
				                                                     BOMStructureEnum.kPurchasedBOMStructure,
				AssemblyDocument { ComponentDefinition: { } assyDef } => assyDef.BOMStructure ==
				                                                         BOMStructureEnum.kPurchasedBOMStructure,
				_ => false
			};
		}
		catch
		{
			return false;
		}
	}

	private static void SetBomStructureToPurchased(PartDocument partDoc)
	{
		try
		{
			partDoc?.ComponentDefinition?.BOMStructure = BOMStructureEnum.kPurchasedBOMStructure;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"SetBomStructureToPurchased: Error setting BOM structure: {ex.Message}");
		}
	}

	private static double CalculateThickness(SheetMetalComponentDefinition smDef, UnitsOfMeasure unitsOfMeasure,
		Dictionary<string, string> properties)
	{
		try
		{
			var parameters = smDef.Parameters;
			if (parameters == null) return 0;

			foreach (var name in ThicknessParameterNames)
				try
				{
					var param = GetParameterSafely(parameters, name);
					if (param == null) continue;
					var thicknessValue    = Convert.ToDouble(param.Value);
					var thicknessInInches = unitsOfMeasure.ConvertUnits(thicknessValue, "cm", "in");
					properties["Thickness"] = thicknessInInches.ToString("F4") + " in";

					if (param.ExposedAsProperty) return thicknessInInches;
					param.ExposedAsProperty = true;
					param.CustomPropertyFormat.PropertyType = CustomPropertyTypeEnum.kTextPropertyType;
					param.CustomPropertyFormat.Precision = CustomPropertyPrecisionEnum.kThreeDecimalPlacesPrecision;
					param.CustomPropertyFormat.ShowLeadingZeros = false;
					param.CustomPropertyFormat.ShowTrailingZeros = true;
					param.CustomPropertyFormat.ShowUnitsString = true;
					param.CustomPropertyFormat.set_Units("in");

					return thicknessInInches;
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"Error calculating thickness for parameter '{name}': {ex}");
				}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error calculating Thickness: {ex}");
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
			Debug.WriteLine($"Error calculating extent properties: {ex}");
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
			Debug.WriteLine($"Error getting material from part: {ex}");
			return string.Empty;
		}
	}

	private static Dictionary<string, string> FindMatchingRawMaterial(string material, string gaugeName,
		double thickness)
	{
		var    materialKey = material.ToLowerInvariant();
		string rmValue     = null;

		if (materialKey.Contains("stainless"))
		{
			rmValue = TryMatchGauge(gaugeName, StainlessSteelGaugeMap) ??
			          MatchByThickness(thickness, MatchStainlessSteelByThickness);
		}
		else if (materialKey.Contains("mild"))
		{
			// Special case: "1" gauge should not match "1/2" or "1/4"
			if (gaugeName.Contains('1') && !gaugeName.Contains("1/2") && !gaugeName.Contains("1/4") &&
			    MildSteelGaugeMap.TryGetValue("1", out var oneGaugeValue))
				rmValue = oneGaugeValue;
			else
				rmValue = TryMatchGauge(gaugeName, MildSteelGaugeMap) ??
				          MatchByThickness(thickness, MatchMildSteelByThickness);
		}
		else if (materialKey.Contains("rubber"))
		{
			rmValue = "LG";
		}

		if (string.IsNullOrEmpty(rmValue))
			return null;

		return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
		{
			["RM"]     = rmValue,
			["RMUNIT"] = DefaultRmUnit
		};
	}

	private static string TryMatchGauge(string gaugeName, Dictionary<string, string> gaugeMap)
	{
		foreach (var (gauge, materialCode) in gaugeMap)
			if (gaugeName.Contains(gauge, StringComparison.OrdinalIgnoreCase))
				return materialCode;

		return null;
	}

	private static string MatchByThickness(double thickness, Func<double, string> thicknessMatcher)
	{
		return thickness > 0 ? thicknessMatcher(thickness) : null;
	}

	private static string MatchStainlessSteelByThickness(double thickness)
	{
		return thickness switch
		{
			_ when Math.Abs(thickness - 0.048) <= ThicknessTolerance => "FS-48X96X0.048",
			_ when Math.Abs(thickness - 0.075) <= ThicknessTolerance => "FS-60X120X0.075",
			_ when Math.Abs(thickness - 0.090) <= ThicknessTolerance => "FS-60X97X0.090",
			_ when Math.Abs(thickness - 0.105) <= ThicknessTolerance => "FS-60X120X0.105",
			_ when Math.Abs(thickness - 0.135) <= ThicknessTolerance => "FS-60X144X0.135",
			_ when Math.Abs(thickness - 0.188) <= ThicknessTolerance => "FS-60X144X0.188",
			_ when Math.Abs(thickness - 0.250) <= ThicknessTolerance => "FS-60X144X0.250",
			_ when Math.Abs(thickness - 0.313) <= ThicknessTolerance => "FS-60X144X0.313",
			_ when Math.Abs(thickness - 0.375) <= ThicknessTolerance => "FS-60X144X0.375",
			_ when Math.Abs(thickness - 0.500) <= ThicknessTolerance => "FS-60X144X0.500",
			_                                                        => null
		};
	}

	private static string MatchMildSteelByThickness(double thickness)
	{
		return thickness switch
		{
			_ when Math.Abs(thickness - 0.075) <= ThicknessTolerance => "FM-60X144X0.075",
			_ when Math.Abs(thickness - 0.105) <= ThicknessTolerance => "FM-60X144X0.105",
			_ when Math.Abs(thickness - 0.135) <= ThicknessTolerance => "FM-60X144X0.135",
			_ when Math.Abs(thickness - 0.188) <= ThicknessTolerance => "FM-60X144X0.188",
			_ when Math.Abs(thickness - 0.250) <= ThicknessTolerance => "FM-60X144X0.250",
			_ when Math.Abs(thickness - 0.313) <= ThicknessTolerance => "FM-60X144X0.313",
			_ when Math.Abs(thickness - 0.375) <= ThicknessTolerance => "FM-60X144X0.375",
			_ when Math.Abs(thickness - 0.500) <= ThicknessTolerance => "FM-60X144X0.500",
			_ when Math.Abs(thickness - 0.625) <= ThicknessTolerance => "FM-60X144X0.625",
			_ when Math.Abs(thickness - 0.750) <= ThicknessTolerance => "FM-60X120X0.750",
			_ when Math.Abs(thickness - 1.000) <= ThicknessTolerance => "FM-48X120X1.000",
			_                                                        => null
		};
	}

	private static async Task CalculateStockMaterialProperties(PartDocument partDoc, PartComponentDefinition partDef,
		UnitsOfMeasure unitsOfMeasure, Dictionary<string, string> properties,
		CancellationToken cancellationToken = default)
	{
		try
		{
			var material = GetMaterialFromPart(partDoc);

			if (string.IsNullOrEmpty(material))
			{
				Debug.WriteLine("CalculateStockMaterialProperties: Missing material information");
				return;
			}

			cancellationToken.ThrowIfCancellationRequested();

			var stockType       = DetectStockMaterialType(partDef);
			var stockDimensions = GetStockDimensions(partDoc, partDef, unitsOfMeasure, stockType);

			List<Dictionary<string, string>> matches;
			List<Dictionary<string, string>> barMatches  = null;
			List<Dictionary<string, string>> tubeMatches = null;

			var showAllMaterials = stockType == StockMaterialType.None;

			if (stockType != StockMaterialType.None && stockDimensions is { Count: > 0 })
			{
				cancellationToken.ThrowIfCancellationRequested();

				if (stockType == StockMaterialType.AmbiguousRound)
				{
					barMatches =
						await FindMatchingStockRawMaterialAsync(material, StockMaterialType.BarStock, stockDimensions,
							cancellationToken);
					tubeMatches =
						await FindMatchingStockRawMaterialAsync(material, StockMaterialType.RoundTube, stockDimensions,
							cancellationToken);
					matches = MergeStockMaterialMatches(barMatches, tubeMatches);
				}
				else
				{
					matches = await FindMatchingStockRawMaterialAsync(material, stockType, stockDimensions,
						cancellationToken);
				}

				if (matches is { Count: 0 })
				{
					Debug.WriteLine(
						$"CalculateStockMaterialProperties: No matching stock material found for {material}, type {stockType}");
					showAllMaterials = true;
				}
			}
			else
			{
				matches = [];
			}

			var partNumber = PropertyExtractor.GetPropertiesFromDocumentStatic((Document)partDoc)
			                                  .GetValueOrDefault("Part Number", partDoc.DisplayName);

			// If the part has no Cost Center, or no stock matches were found, ask if it should be purchased
			var hasCostCenter = stockDimensions.TryGetValue("Cost Center", out var existingCc) &&
			                    !string.IsNullOrWhiteSpace(existingCc);
			if (_sqlDataManager != null && (!hasCostCenter || showAllMaterials))
			{
				cancellationToken.ThrowIfCancellationRequested();

				try
				{
					var askCcList =
						await _sqlDataManager.GetCostCentersAsync(cancellationToken: cancellationToken);
					if (askCcList is { Count: > 0 })
					{
						var askPurchased = new AskPurchased(partNumber, askCcList, _sqlDataManager, (Document)partDoc,
							existingCc);
						if (askPurchased.ShowDialog() == true)
						{
							SetBomStructureToPurchased(partDoc);
							if (!string.IsNullOrWhiteSpace(askPurchased.SelectedCostCenter))
								properties["Cost Center"] = askPurchased.SelectedCostCenter;
							Debug.WriteLine(
								$"CalculateStockMaterialProperties: Part {partNumber} set to purchased by user");
							return;
						}
					}
				}
				catch (Exception ex)
				{
					Debug.WriteLine(
						$"CalculateStockMaterialProperties: Error in AskPurchased dialog: {ex.Message}");
				}
			}

			// Check if the Cost Center should be updated to match Genius
			try
			{
				if (_sqlDataManager != null && !string.IsNullOrWhiteSpace(partNumber) &&
				    stockDimensions.TryGetValue("Cost Center", out var inventorCc) &&
				    !string.IsNullOrWhiteSpace(inventorCc))
				{
					cancellationToken.ThrowIfCancellationRequested();
					var geniusData   = await _sqlDataManager.GetSqlDataAsync(partNumber, cancellationToken);
					var geniusFamily = geniusData.GetValueOrDefault("Family", "");
					if (!string.IsNullOrWhiteSpace(geniusFamily) &&
					    !inventorCc.Equals(geniusFamily, StringComparison.OrdinalIgnoreCase))
					{
						var costCenters =
							await _sqlDataManager.GetCostCentersAsync(cancellationToken: cancellationToken);
						cancellationToken.ThrowIfCancellationRequested();
						var askFamily = new AskFamily(inventorCc, geniusFamily, costCenters, _sqlDataManager,
							(Document)partDoc);
						if (askFamily.ShowDialog() == true &&
						    !string.IsNullOrWhiteSpace(askFamily.SelectedCostCenter))
						{
							properties["Cost Center"]      = askFamily.SelectedCostCenter;
							stockDimensions["Cost Center"] = askFamily.SelectedCostCenter;
						}
					}
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"CalculateStockMaterialProperties: Error checking Cost Center match: {ex}");
			}

			cancellationToken.ThrowIfCancellationRequested();

			GeniusFormsHelper.ZoomToOccurrence((Document)partDoc);

			var selectedMatch = await ShowStockMaterialSelector(stockType, stockDimensions, matches,
				barMatches, tubeMatches, showAllMaterials, partNumber,
				cancellationToken);
			if (selectedMatch == null)
			{
				Debug.WriteLine("CalculateStockMaterialProperties: User cancelled stock material selection");
				return;
			}

			if (selectedMatch.TryGetValue("RM", out var rmValue))
				properties["RM"] = rmValue;
			properties["RMUNIT"] = "IN";

			// Apply cost center from the selected match if present; otherwise use the default only when
			// a selection was actually made (prevents cancel from overwriting existing part value).
			if (selectedMatch.TryGetValue("Cost Center", out var costCenter) &&
			    !string.IsNullOrWhiteSpace(costCenter))
				properties["Cost Center"] = costCenter;
			else
				properties["Cost Center"] = DefaultCostCenter;
			if (stockDimensions != null && stockDimensions.TryGetValue("Length", out var lengthStr) &&
			    double.TryParse(lengthStr, out var lengthInInches))
				properties["RMQTY"] = lengthInInches.ToString("F4");

			Debug.WriteLine(
				$"CalculateStockMaterialProperties: Selected stock material: {selectedMatch.GetValueOrDefault("RM", "N/A")}");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error in CalculateStockMaterialProperties: {ex}");
		}
	}

	private static async Task<Dictionary<string, string>> ShowStockMaterialSelector(StockMaterialType stockType,
		Dictionary<string, string> dimensions, List<Dictionary<string, string>> matches,
		List<Dictionary<string, string>> barMatches = null, List<Dictionary<string, string>> tubeMatches = null,
		bool showAllMaterials = false, string partNumber = null,
		CancellationToken cancellationToken = default)
	{
		try
		{
			List<string> costCenters;
			try
			{
				if (_sqlDataManager == null)
				{
					Debug.WriteLine("ShowStockMaterialSelector: SQL data manager is null");
					costCenters = [];
				}
				else
				{
					costCenters = await _sqlDataManager.GetCostCentersAsync(cancellationToken: cancellationToken);
					Debug.WriteLine($"ShowStockMaterialSelector: Retrieved {costCenters.Count} cost centers");
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"ShowStockMaterialSelector: Error retrieving cost centers: {ex.Message}");
				costCenters = [];
			}

			var allowRoundTypeSelection = stockType == StockMaterialType.AmbiguousRound;
			var selector = new StockMaterialSelector(stockType.ToString(), dimensions, matches, costCenters,
				_sqlDataManager, allowRoundTypeSelection: allowRoundTypeSelection,
				barMatches: barMatches, tubeMatches: tubeMatches, showAllMaterials: showAllMaterials,
				partNumber: partNumber);
			var tcs = new TaskCompletionSource<Dictionary<string, string>>();

			selector.MaterialSelected   += (_, match) => tcs.TrySetResult(match);
			selector.SelectionCancelled += (_, _) => tcs.TrySetResult(null);

			var panelWrapper = new PanelWrapper(selector, "Select Stock Material");
			selector.MaterialSelected   += (_, _) => panelWrapper.Close();
			selector.SelectionCancelled += (_, _) => panelWrapper.Close();

			return await tcs.Task;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error showing stock material selector: {ex}");
			return null;
		}
	}

	private static StockMaterialType DetectStockMaterialType(PartComponentDefinition partDef)
	{
		try
		{
			return DetectStockMaterialTypeFromRangebox(partDef);
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error detecting stock material type: {ex}");
			return StockMaterialType.None;
		}
	}

	private static StockMaterialType DetectStockMaterialTypeFromRangebox(PartComponentDefinition partDef)
	{
		try
		{
			var rangeBox = partDef.RangeBox;
			if (rangeBox == null) return StockMaterialType.None;

			var sorted = GetSortedRangeboxDimensionsInches(rangeBox);
			if (sorted == null) return StockMaterialType.None;

			var smallest = sorted[0];
			var middle   = sorted[1];
			var largest  = sorted[2];

			Debug.WriteLine(
				$"DetectStockMaterialTypeFromRangebox: smallest={smallest:F4}, middle={middle:F4}, largest={largest:F4}");

			// If largest is significantly larger than middle, it's the length
			// Use smallest and middle to determine cross-section type
			const double relativeEqualityTolerance   = 0.10; // 10% relative equality threshold
			const double crossSectionDominanceFactor = 2.0;  // original dominance factor
			const double crossSectionToLengthFactor  = 1.1;  // used for short-length special case

			// Special case: two largest dimensions nearly equal and noticeably larger than the smallest
			// -> the smallest may actually be the length (short part). Treat the two larger dims as the cross-section.
			if (Math.Abs(middle - largest) / Math.Max(1.0, largest) < relativeEqualityTolerance &&
			    middle > smallest * crossSectionToLengthFactor)
			{
				// cross-section dims are middle and largest (nearly equal)
				if (Math.Abs(middle - largest) < middle * 0.05)
				{
					if (HasSignificantCylindricalFaces(partDef, middle))
						return ClassifyRoundCrossSection(
							"cylindrical faces, square bounding box cross-section - short length");

					if (!HasInternalVoid(partDef, smallest, middle, largest))
					{
						Debug.WriteLine(
							"DetectStockMaterialTypeFromRangebox: Detected Bar Stock (square cross-section, solid - short length)");
						return StockMaterialType.BarStock;
					}

					// Check if this is a diagonally-modeled square bar (rotated 45° about the length axis)
					// The rangebox measures the diagonal, so the cross-section appears larger, and the
					// internal void (hole) makes it look like a tube — but it's actually a solid bar with a hole.
					if (IsDiagonalCrossSection(partDef))
					{
						Debug.WriteLine(
							"DetectStockMaterialTypeFromRangebox: Detected Square Bar (diagonal cross-section, solid with hole - short length)");
						return StockMaterialType.SquareBar;
					}

					Debug.WriteLine(
						"DetectStockMaterialTypeFromRangebox: Detected Square Tube (cross-section is square - short length)");
					return StockMaterialType.SquareTube;
				}

				if (!HasInternalVoid(partDef, smallest, middle, largest))
				{
					Debug.WriteLine(
						"DetectStockMaterialTypeFromRangebox: Detected Bar Stock (rectangular cross-section, solid - short length)");
					return StockMaterialType.BarStock;
				}

				Debug.WriteLine(
					"DetectStockMaterialTypeFromRangebox: Detected Rectangle Tube (cross-section is rectangular - short length)");
				return StockMaterialType.RectangleTube;
			}

			// Square cross-section: two equal dims are the profile, largest is length.
			// Does not require length to be 2x the cross-section (handles short vertical bar, etc.).
			if (Math.Abs(smallest - middle) < ThicknessTolerance && largest > middle + ThicknessTolerance)
			{
				if (HasSignificantCylindricalFaces(partDef, middle))
					return ClassifyRoundCrossSection(
						"cylindrical faces, square bounding box cross-section");

				if (!HasInternalVoid(partDef, smallest, middle, largest))
				{
					Debug.WriteLine(
						"DetectStockMaterialTypeFromRangebox: Detected Square Bar (square cross-section, solid)");
					return StockMaterialType.SquareBar;
				}

				// Check if this is a diagonally-modeled square bar (rotated 45° about the length axis)
				if (IsDiagonalCrossSection(partDef))
				{
					Debug.WriteLine(
						"DetectStockMaterialTypeFromRangebox: Detected Square Bar (diagonal cross-section, solid with hole)");
					return StockMaterialType.SquareBar;
				}

				Debug.WriteLine(
					"DetectStockMaterialTypeFromRangebox: Detected Square Tube (cross-section is square)");
				return StockMaterialType.SquareTube;
			}

			// Rectangular cross-section: length must clearly dominate the larger cross-section dimension.
			if (largest > middle * crossSectionDominanceFactor)
			{
				if (!HasInternalVoid(partDef, smallest, middle, largest))
				{
					Debug.WriteLine(
						"DetectStockMaterialTypeFromRangebox: Detected Rectangle Bar (rectangular cross-section, solid)");
					return StockMaterialType.RectangleBar;
				}

				Debug.WriteLine(
					"DetectStockMaterialTypeFromRangebox: Detected Rectangle Tube (cross-section is rectangular)");
				return StockMaterialType.RectangleTube;
			}

			// If all dimensions are similar, it could be a solid cube/block or a short cylinder
			if (Math.Abs(smallest - middle) < ThicknessTolerance && Math.Abs(middle - largest) < ThicknessTolerance)
			{
				if (HasSignificantCylindricalFaces(partDef, middle))
					return ClassifyRoundCrossSection("cylindrical faces, all dimensions similar");

				Debug.WriteLine("DetectedStockMaterialTypeFromRangebox: Detected Bar Stock (all dimensions similar)");
				return StockMaterialType.BarStock;
			}

			Debug.WriteLine("DetectStockMaterialTypeFromRangebox: Could not determine type from Rangebox");
			return StockMaterialType.None;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error detecting stock material type from Rangebox: {ex}");
			return StockMaterialType.None;
		}
	}

	/// <summary>
	///     Round cross-sections are always ambiguous — bar vs tube/pipe is chosen in the material selector.
	/// </summary>
	private static StockMaterialType ClassifyRoundCrossSection(string context)
	{
		Debug.WriteLine($"DetectStockMaterialTypeFromRangebox: Ambiguous round profile ({context})");
		return StockMaterialType.AmbiguousRound;
	}

	/// <summary>
	///     Returns true if the part body contains a cylindrical face whose radius is
	///     a significant fraction of the expected cross-section size.
	/// </summary>
	private static bool HasSignificantCylindricalFaces(PartComponentDefinition partDef, double crossSectionSizeInches)
	{
		try
		{
			return GetDistinctSignificantCylinderRadiiInches(partDef, crossSectionSizeInches).Count > 0;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"HasSignificantCylindricalFaces: Could not inspect faces: {ex.Message}");
			return false;
		}
	}

	private static List<double> GetDistinctSignificantCylinderRadiiInches(PartComponentDefinition partDef,
		double crossSectionSizeInches)
	{
		var body = partDef.SurfaceBodies.Count > 0 ? partDef.SurfaceBodies[1] : null;
		if (body == null) return [];

		const double radiusFraction  = 0.35;
		var          minRadiusInches = crossSectionSizeInches * radiusFraction;
		var          radii           = new List<double>();

		foreach (var face in body.Faces.Cast<Face>()
		                         .Where(face => face.SurfaceType == SurfaceTypeEnum.kCylinderSurface))
		{
			if (face.Geometry is not Cylinder cylinder) continue;
			var unitsOfMeasure = ThisApplication.ActiveDocument.UnitsOfMeasure;
			var radiusInches   = unitsOfMeasure.ConvertUnits(cylinder.Radius, "cm", "in");
			if (radiusInches >= minRadiusInches)
				radii.Add(radiusInches);
		}

		return DistinctCylinderRadii(radii);
	}

	private static List<double> DistinctCylinderRadii(List<double> radii)
	{
		if (radii.Count == 0) return [];

		const double clusterTolerance = 0.05;
		var          distinct         = new List<double>();

		foreach (var radius in radii.OrderByDescending(r => r))
		{
			if (distinct.Any(d => Math.Abs(d - radius) / Math.Max(d, radius) < clusterTolerance))
				continue;
			distinct.Add(radius);
		}

		return [.. distinct.OrderByDescending(r => r)];
	}

	/// <summary>
	///     Detects if the part has a square cross-section that is rotated approximately 45°
	///     about the length axis (making the rangebox measure the diagonal instead of the side).
	///     Checks for planar side faces whose normals are at 45° to the principal axes.
	/// </summary>
	private static bool IsDiagonalCrossSection(PartComponentDefinition partDef)
	{
		try
		{
			var body = partDef.SurfaceBodies.Count > 0 ? partDef.SurfaceBodies[1] : null;
			if (body == null) return false;

			const double diagonalNormal    = 0.7071067811865476; // 1/√2
			const double normalTolerance   = 0.15;
			var          diagonalFaceCount = 0;

			foreach (var face in body.Faces.Cast<Face>()
			                         .Where(face => face.SurfaceType == SurfaceTypeEnum.kPlaneSurface))
			{
				if (face.Geometry is not Plane plane) continue;

				var normal    = plane.Normal;
				var magnitude = Math.Sqrt(normal.X * normal.X + normal.Y * normal.Y + normal.Z * normal.Z);
				if (magnitude < 0.001) continue;

				var nx = Math.Abs(normal.X / magnitude);
				var ny = Math.Abs(normal.Y / magnitude);
				var nz = Math.Abs(normal.Z / magnitude);

				// Check for planar side faces at 45°: one component near 0, two near 0.707.
				// This is the characteristic pattern of a square cross-section rotated 45°:
				// the length axis has near-zero contribution to the normal,
				// and the two cross-section axes each contribute ~0.707 (1/√2).
				var hasNearZeroComponent   = nx < 0.1 || ny < 0.1 || nz < 0.1;
				var diagonalComponentCount = 0;
				if (Math.Abs(nx - diagonalNormal) < normalTolerance) diagonalComponentCount++;
				if (Math.Abs(ny - diagonalNormal) < normalTolerance) diagonalComponentCount++;
				if (Math.Abs(nz - diagonalNormal) < normalTolerance) diagonalComponentCount++;

				if (!hasNearZeroComponent || diagonalComponentCount != 2) continue;

				diagonalFaceCount++;
			}

			// At least 2 diagonal side faces strongly suggests a 45° rotated cross-section.
			// A typical square nut or bar has 4 such faces (one on each side of the square).
			// Small chamfers/fillets rarely produce 2+ faces matching this exact pattern.
			return diagonalFaceCount >= 2;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"IsDiagonalCrossSection: {ex.Message}");
			return false;
		}
	}

	private static bool HasInternalVoid(PartComponentDefinition partDef, double smallest, double middle, double largest)
	{
		try
		{
			var body = partDef.SurfaceBodies.Count > 0 ? partDef.SurfaceBodies[1] : null;
			if (body is not { IsSolid: true }) return false;

			var actualVolumeCm3 = body.Volume[0.01];
			if (actualVolumeCm3 <= 0) return false;

			var theoreticalVolumeIn3 = smallest * middle * largest;
			var unitsOfMeasure       = ThisApplication.ActiveDocument.UnitsOfMeasure;
			var theoreticalVolumeCm3 = unitsOfMeasure.ConvertUnits(theoreticalVolumeIn3, "in^3", "cm^3");
			if (theoreticalVolumeCm3 <= 0) return false;

			var volumeRatio = actualVolumeCm3 / theoreticalVolumeCm3;

			Debug.WriteLine($"HasInternalVoid: volume ratio = {volumeRatio:F4}");

			const double hollowThreshold = 0.85;
			return volumeRatio < hollowThreshold;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"HasInternalVoid: Could not check: {ex.Message}");
			return false;
		}
	}

	private static List<Dictionary<string, string>> MergeStockMaterialMatches(
		List<Dictionary<string, string>> primary, List<Dictionary<string, string>> secondary)
	{
		var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

		return
		[
			.. from match in (primary ?? []).Concat(secondary ?? [])
			let rm = match.GetValueOrDefault("RM", string.Empty)
			where string.IsNullOrWhiteSpace(rm) || seen.Add(rm)
			select match
		];
	}

	private static double[] GetSortedRangeboxDimensionsInches(Box rangeBox)
	{
		try
		{
			if (rangeBox == null) return null;

			// Get raw extents (Inventor uses cm)
			var xLength = rangeBox.MaxPoint.X - rangeBox.MinPoint.X;
			var yLength = rangeBox.MaxPoint.Y - rangeBox.MinPoint.Y;
			var zLength = rangeBox.MaxPoint.Z - rangeBox.MinPoint.Z;

			var unitsOfMeasure = ThisApplication.ActiveDocument.UnitsOfMeasure;
			var xInches        = unitsOfMeasure.ConvertUnits(xLength, "cm", "in");
			var yInches        = unitsOfMeasure.ConvertUnits(yLength, "cm", "in");
			var zInches        = unitsOfMeasure.ConvertUnits(zLength, "cm", "in");

			return [.. new[] { xInches, yInches, zInches }.OrderBy(d => d)];
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GetSortedRangeboxDimensionsInches: {ex}");
			return null;
		}
	}

	private static Dictionary<string, string> GetStockDimensions(PartDocument partDoc, PartComponentDefinition partDef,
		UnitsOfMeasure unitsOfMeasure, StockMaterialType stockType)
	{
		var dimensions = new Dictionary<string, string>();
		try
		{
			// Get the part's existing Cost Center iProperty if it exists
			GetPartCostCenter(partDoc, dimensions);

			// Get the part's existing RM iProperty if it exists
			GetPartRM(partDoc, dimensions);

			var parameters = partDef.Parameters;
			if (parameters == null) return dimensions;

			switch (stockType)
			{
				case StockMaterialType.None:
					GetDimensionsFromRangebox(partDoc, StockMaterialType.BarStock, dimensions);
					break;

				case StockMaterialType.BarStock:
					if (TryGetDimension(parameters, unitsOfMeasure, "Diameter", "in", out var diameter))
						dimensions["Diameter"] = diameter;
					else
						GetDimensionsFromRangebox(partDoc, stockType, dimensions);
					if (TryGetDimension(parameters, unitsOfMeasure, "Length", "in", out var length))
						dimensions["Length"] = length;
					break;

				case StockMaterialType.RoundTube:
					if (TryGetDimension(parameters, unitsOfMeasure, "Diameter", "in", out var od))
						dimensions["OD"] = od;
					else
						GetDimensionsFromRangebox(partDoc, stockType, dimensions);
					if (TryGetDimension(parameters, unitsOfMeasure, "Length", "in", out var rtLength))
						dimensions["Length"] = rtLength;
					break;

				case StockMaterialType.AmbiguousRound:
					if (TryGetDimension(parameters, unitsOfMeasure, "Diameter", "in", out var roundSize))
					{
						dimensions["Diameter"] = roundSize;
						dimensions["OD"]       = roundSize;
					}
					else
					{
						GetDimensionsFromRangebox(partDoc, StockMaterialType.BarStock, dimensions);
						if (dimensions.TryGetValue("Diameter", out var derivedDiameter))
							dimensions["OD"] = derivedDiameter;
						else
							GetDimensionsFromRangebox(partDoc, StockMaterialType.RoundTube, dimensions);
					}

					if (TryGetDimension(parameters, unitsOfMeasure, "Length", "in", out var ambiguousLength))
						dimensions["Length"] = ambiguousLength;
					else if (!dimensions.ContainsKey("Length"))
						GetDimensionsFromRangebox(partDoc, StockMaterialType.BarStock, dimensions);
					break;

				case StockMaterialType.SquareBar:
				case StockMaterialType.RectangleBar:
				case StockMaterialType.SquareTube:
				case StockMaterialType.RectangleTube:
					if (TryGetDimension(parameters, unitsOfMeasure, "Width", "in", out var width))
						dimensions["Width"] = width;
					if (TryGetDimension(parameters, unitsOfMeasure, "Height", "in", out var height))
						dimensions["Height"] = height;
					else
						GetDimensionsFromRangebox(partDoc, stockType, dimensions);
					if (TryGetDimension(parameters, unitsOfMeasure, "Length", "in", out var stLength))
						dimensions["Length"] = stLength;
					else if (!dimensions.ContainsKey("Length"))
						GetDimensionsFromRangebox(partDoc, stockType, dimensions);
					break;
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error getting stock dimensions: {ex}");
		}

		return dimensions;
	}

	/// <summary>
	///     Gets the Cost Center iProperty from the Inventor part if it exists.
	/// </summary>
	/// <param name="partDoc">The part document to read the property from.</param>
	/// <param name="dimensions">The dictionary to add the Cost Center to.</param>
	private static void GetPartCostCenter(PartDocument partDoc, Dictionary<string, string> dimensions)
	{
		try
		{
			var propertySets = partDoc.PropertySets;
			if (propertySets == null) return;

			// Look for the Cost Center in the Summary Information or Document Information properties
			foreach (PropertySet propertySet in propertySets)
				try
				{
					// Check for Cost Center in the property set
					if (!propertySet.Name.Equals("Design Tracking Properties", StringComparison.OrdinalIgnoreCase) &&
					    !propertySet.Name.Equals("Summary Information", StringComparison.OrdinalIgnoreCase)) continue;
					foreach (Property property in propertySet)
					{
						if (!property.Name.Equals("Cost Center", StringComparison.OrdinalIgnoreCase) ||
						    property.Value is not string costCenterValue ||
						    string.IsNullOrWhiteSpace(costCenterValue)) continue;
						dimensions["Cost Center"] = costCenterValue;
						Debug.WriteLine($"GetPartCostCenter: Found Cost Center '{costCenterValue}' in part properties");
						return;
					}
				}
				catch (Exception ex)
				{
					Debug.WriteLine($"GetPartCostCenter: Error reading property set: {ex.Message}");
				}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GetPartCostCenter: Error getting part cost center: {ex.Message}");
		}
	}

	/// <summary>
	///     Gets the RM iProperty from the Inventor part if it exists.
	/// </summary>
	/// <param name="partDoc">The part document to read the property from.</param>
	/// <param name="dimensions">The dictionary to add the RM to.</param>
	private static void GetPartRM(PartDocument partDoc, Dictionary<string, string> dimensions)
	{
		try
		{
			var propertySets = partDoc.PropertySets;
			if (propertySets == null)
			{
				Debug.WriteLine("GetPartRM: PropertySets is null");
				return;
			}

			// The RM property is always in 'Inventor User Defined Properties'
			var userDefinedSet = GetPropertySetSafely(propertySets, "Inventor User Defined Properties");
			if (userDefinedSet == null)
			{
				Debug.WriteLine("GetPartRM: 'Inventor User Defined Properties' property set not found");
				return;
			}

			Debug.WriteLine("GetPartRM: Checking property set 'Inventor User Defined Properties'");

			try
			{
				var rmProperty = GetPropertySafely(userDefinedSet, "RM");
				if (rmProperty?.Value is string rmValue && !string.IsNullOrWhiteSpace(rmValue))
				{
					dimensions["RM"] = rmValue;
					Debug.WriteLine(
						$"GetPartRM: SUCCESS - Found RM '{rmValue}' in property set 'Inventor User Defined Properties'");
				}
				else
				{
					Debug.WriteLine("GetPartRM: RM property not found or empty in 'Inventor User Defined Properties'");
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine(
					$"GetPartRM: Error accessing RM in set 'Inventor User Defined Properties': {ex.GetType().Name} - {ex.Message}");
			}
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"GetPartRM: Error processing property sets: {ex.GetType().Name} - {ex.Message}");
		}
	}

	private static void GetDimensionsFromRangebox(PartDocument partDoc, StockMaterialType stockType,
		Dictionary<string, string> dimensions)
	{
		try
		{
			var rangeBox = partDoc.ComponentDefinition.RangeBox;
			if (rangeBox == null) return;

			var sorted = GetSortedRangeboxDimensionsInches(rangeBox);
			if (sorted == null) return;

			var smallest = sorted[0];
			var middle   = sorted[1];
			var largest  = sorted[2];

			// Detect short-length case: middle and largest nearly equal and significantly larger than smallest.
			// In that case the smallest is the length and the other two are the cross-section.
			const double relativeEqualityTolerance  = 0.10; // 10%
			const double crossSectionToLengthFactor = 1.1;

			var treatSmallestAsLength = Math.Abs(middle - largest) / Math.Max(1.0, largest) < relativeEqualityTolerance
			                            && middle > smallest * crossSectionToLengthFactor;

			if (treatSmallestAsLength)
				// Short part: smallest is length
				dimensions["Length"] = smallest.ToString("F4");
			else
				// The largest dimension is typically the length
				dimensions["Length"] = largest.ToString("F4");

			switch (stockType)
			{
				case StockMaterialType.BarStock:
					// Diameter is one of the cross-section dimensions. If smallest was treated as the length,
					// use middle (smaller cross-section) as the diameter; otherwise use middle as before.
					dimensions["Diameter"] = middle.ToString("F4");
					break;

				case StockMaterialType.RoundTube:
					// OD is one of the cross-section dimensions. If smallest is the length, OD is the larger of the two remaining dims.
					if (treatSmallestAsLength)
						dimensions["OD"] = Math.Max(middle, largest).ToString("F4");
					else
						dimensions["OD"] = middle.ToString("F4");
					break;

				case StockMaterialType.SquareBar:
				case StockMaterialType.RectangleBar:
				case StockMaterialType.SquareTube:
				case StockMaterialType.RectangleTube:
					// Smallest and middle are width and height (cross-section) when largest is length.
					// If smallest is the length instead, cross-section is middle and largest.
					if (treatSmallestAsLength)
					{
						dimensions["Width"]  = largest.ToString("F4");
						dimensions["Height"] = middle.ToString("F4");
					}
					else
					{
						dimensions["Width"]  = middle.ToString("F4");
						dimensions["Height"] = smallest.ToString("F4");
					}

					// Correct for diagonal square cross-section (rotated 45° about the length axis).
					// The rangebox gives the diagonal measurement; divide by √2 to get the true side length.
					if (stockType is StockMaterialType.SquareBar or StockMaterialType.SquareTube &&
					    IsDiagonalCrossSection(partDoc.ComponentDefinition))
					{
						var currentWidth  = double.Parse(dimensions["Width"]);
						var currentHeight = double.Parse(dimensions["Height"]);
						if (Math.Abs(currentWidth - currentHeight) / Math.Max(1.0, currentWidth) < 0.05)
						{
							var correctedSide = currentWidth / Math.Sqrt(2);
							dimensions["Width"]  = correctedSide.ToString("F4");
							dimensions["Height"] = correctedSide.ToString("F4");
						}
					}

					break;
			}

			Debug.WriteLine(
				$"GetDimensionsFromRangebox: Extracted dimensions for {stockType}: {string.Join(", ", dimensions)}");
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"Error getting dimensions from Rangebox: {ex}");
		}
	}


	private static bool TryGetDimension(Parameters parameters, UnitsOfMeasure unitsOfMeasure, string paramName,
		string targetUnit, out string value)
	{
		value = string.Empty;
		try
		{
			var param = GetParameterSafely(parameters, paramName);
			if (param == null) return false;

			var paramValue     = Convert.ToDouble(param.Value);
			var convertedValue = unitsOfMeasure.ConvertUnits(paramValue, "cm", targetUnit);
			value = convertedValue.ToString("F4");
			return true;
		}
		catch
		{
			return false;
		}
	}

	private static async Task<List<Dictionary<string, string>>> FindMatchingStockRawMaterialAsync(string material,
		StockMaterialType stockType, Dictionary<string, string> dimensions,
		CancellationToken cancellationToken = default)
	{
		if (_sqlDataManager == null)
		{
			Debug.WriteLine("FindMatchingStockRawMaterialAsync: SQL data manager not initialized");
			return [];
		}

		try
		{
			var stockTypeString = stockType.ToString();
			var results =
				await _sqlDataManager.GetStockMaterialAsync(stockTypeString, dimensions, material, cancellationToken);

			Debug.WriteLine(
				results is { Count: > 0 }
					? $"FindMatchingStockRawMaterialAsync: Found {results.Count} SQL match(es) for {stockType}"
					: $"FindMatchingStockRawMaterialAsync: No SQL match found for {stockType}");

			return results;
		}
		catch (Exception ex)
		{
			Debug.WriteLine($"FindMatchingStockRawMaterialAsync: SQL lookup failed: {ex}");
			return [];
		}
	}

	private static Parameter GetParameterSafely(Parameters parameters, string name)
	{
		try
		{
			foreach (var param in parameters.Cast<Parameter>()
			                                .Where(param =>
				                                param.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
				return param;
		}
		catch
		{
			// Ignore
		}

		return null;
	}

	private static PropertySet GetPropertySetSafely(PropertySets propertySets, string name)
	{
		try
		{
			foreach (var getPropertySetSafely in propertySets.Cast<PropertySet>()
			                                                 .Where(propertySet =>
				                                                 propertySet.Name.Equals(name,
					                                                 StringComparison.OrdinalIgnoreCase)))
				return getPropertySetSafely;
		}
		catch
		{
			// Ignore
		}

		return null;
	}

	private static Property GetPropertySafely(PropertySet propertySet, string name)
	{
		try
		{
			foreach (var prop in propertySet.Cast<Property>()
			                                .Where(prop => prop.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
				return prop;
		}
		catch
		{
			// Ignore
		}

		return null;
	}

	private enum StockMaterialType
	{
		None,
		BarStock,
		SquareBar,
		RectangleBar,
		RoundTube,
		AmbiguousRound,
		SquareTube,
		RectangleTube
	}
}