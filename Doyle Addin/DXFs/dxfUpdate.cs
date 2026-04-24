namespace DoyleAddin.DXFs;

using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;
using Options;
using Services;
using Environment = Environment;
using Path = Path;

internal static class DxfUpdate
{
	private static bool CreateFlatPattern(SheetMetalComponentDefinition def, string partName,
		List<string> failedExports)
	{
		if (def.HasFlatPattern) return true;
		try
		{
			def.Unfold();
			def.FlatPattern.ExitEdit();
			return true;
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			failedExports.Add("Failed to create flat pattern for: " + partName);
			return false;
		}
	}

	private static bool ValidateFlatPattern(SheetMetalComponentDefinition memberDef, string partNumber,
		List<string> failedExports)
	{
		if (memberDef.FlatPattern.FlatBendResults.Count != 0 ||
		    memberDef.FlatPattern.RangeBox.MaxPoint.Z - memberDef.FlatPattern.RangeBox.MinPoint.Z <=
		    Convert.ToDouble(memberDef.Thickness.Value) + 0.003d) return true;
		failedExports.Add("Invalid flat pattern for: " + partNumber);
		return false;
	}

	private static void ExportDxf(SheetMetalComponentDefinition def, string fileName, string partNumber,
		List<string> failedExports, UserOptions userOptions)
	{
		const string oFormat =
				"FLAT PATTERN DXF?AcadVersion=2018"
				+ "&TangentLayer=FlatPattern_Tangent Lines"
				+ "&OuterProfileLayer=FlatPattern_Outer Profile"
				+ "&ArcCentersLayer=FlatPattern_Arc Centers"
				+ "&InteriorProfilesLayer=FlatPattern_Interior Profile"
				+ "&BendLayer=FlatPattern_Bend Lines"
				+ "&BendUpLayer=FlatPattern_Bend Lines"
				+ "&BendDownLayer=FlatPattern_Bend Lines backside"
				+ "&ToolCenterLayer=FlatPattern_Tool Centers"
				+ "&ToolCenterUpLayer=FlatPattern_Tool Centers"
				+ "&ToolCenterDownLayer=FlatPattern_Tool Centers backside"
				+ "&FeatureProfilesLayer=FlatPattern_Feature Profile"
				+ "&FeatureProfilesUpLayer=FlatPattern_Feature Profile"
				+ "&FeatureProfilesDownLayer=FlatPattern_Feature Profile backside"
				+ "&AltRepFrontLayer=FlatPattern_AltRep Front"
				+ "&AltRepBackLayer=FlatPattern_AltRep Back"
				+ "&UnconsumedSketchesLayer=FlatPattern_Unconsumed Sketches"
				+ "&TangentRollLinesLayer=FlatPattern_Roll Tangent"
				+ "&RollLinesLayer=FlatPattern_Roll Lines"
				+ "&OuterProfileLayerColor=0;0;255"
				+ "&InteriorProfilesLayerColor=0;0;0"
				+ "&ArcCentersLayerColor=255;0;255"
				+ "&TangentLayerColor=255;255;0"
				+ "&BendLayerColor=0;255;0"
				+ "&BendUpLayerColor=0;255;0"
				+ "&BendDownLayerColor=255;0;0"
				+ "&FeatureProfilesUpLayerColor=0;128;255"
				+ "&FeatureProfilesDownLayerColor=255;0;0"
				+ "&ToolCenterUpLayerColor=0;0;0"
				+ "&ToolCenterDownLayerColor=0;0;0"
			;
		try
		{
			def.DataIO.WriteDataToFile(oFormat, fileName);

			// Create ClickUp task after successful DXF export if enabled
			if (userOptions.EnableClickUpIntegration)
				_ = Task.Run(async () => await ClickUpService.CreateDxfExportTaskAsync(partNumber));
		}
		catch (Exception ex)
		{
			Console.WriteLine(ex);
			failedExports.Add("DXF failed to generate for: " + partNumber);
		}
	}

	private static void GenerateMissingMembers(iPartFactory oFactory, Documents oDoc)
	{
		oDoc.CloseAll(true);
		foreach (var iPartTableRow in oFactory.TableRows)
			oFactory.CreateMember(iPartTableRow);
	}

	private static void ProcessIPartMember(string filepath, Documents oDoc, List<string> failedExports,
		UserOptions userOptions)
	{
		if (oDoc.Open(filepath) is not PartDocument openedDoc) return;

		var memberDef  = openedDoc.ComponentDefinition as SheetMetalComponentDefinition;
		var partNumber = openedDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();
		var oiFileName = Path.Combine(userOptions.DxfExportLocation, partNumber + ".dxf");

		if (!ValidateFlatPattern(memberDef, partNumber, failedExports))
		{
			openedDoc.Close(true);
			return;
		}

		ExportDxf(memberDef, oiFileName, partNumber, failedExports, userOptions);
		openedDoc.Close(true);
	}

	private static void ProcessIPartFactory(PartDocument oPartDoc,
		SheetMetalComponentDefinition oDef, string pn, List<string> failedExports, UserOptions userOptions)
	{
		var oFactory = oDef.iPartFactory;
		var oDoc     = ThisApplication.Documents;
		var total    = oFactory.TableRows.Count;
		oDoc.CloseAll(true);
		oPartDoc.ReleaseReference();

		if (!CreateFlatPattern(oDef, pn, failedExports)) return;

		if (!Directory.Exists(oFactory.MemberCacheDir)) Directory.CreateDirectory(oFactory.MemberCacheDir);
		var partFiles = Directory.GetFiles(oFactory.MemberCacheDir);

		if (total > partFiles.Length)
		{
			var result = MessageBox.Show(
				"Warning: The factory has " + total + " members, but " + partFiles.Length +
				" files were found in the folder. Generate files?.", "Missing Members",
				MessageBoxButtons.YesNo);
			if (result == DialogResult.Yes)
				GenerateMissingMembers(oFactory, oDoc);
			else
				return;
		}

		foreach (var filepath in Directory.GetFiles(oFactory.MemberCacheDir))
			ProcessIPartMember(filepath, oDoc, failedExports, userOptions);

		if (failedExports.Count > 0)
			MessageBox.Show(failedExports.Count + " Members have errors and were skipped." +
			                Environment.NewLine + string.Join(Environment.NewLine, failedExports));
		else
			MessageBox.Show("Created " + total + " DXFs. All exports succeeded.");
	}

	private static void ProcessNonIPart(PartDocument oPartDoc, SheetMetalComponentDefinition oDef, string oFileName,
		List<string> failedExports, UserOptions userOptions)
	{
		if (oPartDoc._ComatoseNodesCount > 0 || oPartDoc._SickNodesCount > 0)
		{
			MessageBox.Show(oPartDoc.DisplayName + " has errors, fix before export." + Environment.NewLine +
			                string.Join(Environment.NewLine, failedExports));
			return;
		}

		var modelStates = oDef.ModelStates;
		var exportCount = 0;
		var totalCount  = modelStates.Count;
		var partNumber  = oPartDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();

		foreach (ModelState modelState in modelStates)
			try
			{
				// Activate the model state
				modelState.Activate();
				var currentStatePartNumber =
					oPartDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();
				var fileName = Path.Combine(Path.GetDirectoryName(oFileName) ?? "", $"{currentStatePartNumber}.dxf");

				if (!CreateFlatPattern(oDef, currentStatePartNumber, failedExports))
				{
					failedExports.Add(
						$"Failed to create flat pattern for model state '{modelState.Name}' in: {currentStatePartNumber}");
					continue;
				}

				ExportDxf(oDef, fileName, currentStatePartNumber, failedExports, userOptions);
				exportCount++;
			}
			catch (Exception ex)
			{
				failedExports.Add($"Error processing model state '{modelState.Name}': {ex.Message}");
			}

		// Show results
		if (failedExports.Count > 0)
		{
			MessageBox.Show($"{exportCount} of {totalCount} model states exported successfully." + Environment.NewLine +
			                $"{failedExports.Count} errors:" + Environment.NewLine +
			                string.Join(Environment.NewLine, failedExports),
				"Export Results", MessageBoxButtons.OK, MessageBoxIcon.Warning);
		}
		else
		{
			var message = totalCount > 1
				? $"All {totalCount} model states exported successfully."
				: $"{partNumber} exported successfully.";
			MessageBox.Show(message, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}

	public static void RunDxfUpdate()
	{
		if (ThisApplication.ActiveDocument is not PartDocument oPartDoc)
		{
			MessageBox.Show("ActiveDocument is not a PartDocument");
			MessageBox.Show("DXF Export can only be used on Part documents.", "Error", MessageBoxButtons.OK,
				MessageBoxIcon.Warning);
			return;
		}

		if (oPartDoc.ComponentDefinition is not SheetMetalComponentDefinition oDef)
		{
			MessageBox.Show("ComponentDefinition is null or not SheetMetalComponentDefinition");
			MessageBox.Show("This is not a sheet metal part.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			return;
		}

		var failedExports = new List<string>();
		var propertySets  = oPartDoc.PropertySets["Design Tracking Properties"];
		var pn            = propertySets["Part Number"].Value.ToString();
		var userOptions   = UserOptions.Load();
		var oFileName     = Path.Combine(userOptions.DxfExportLocation, pn + ".dxf");

		// Check if a part is a factory and delegate to the appropriate processor
		if (oDef.IsiPartFactory)
			ProcessIPartFactory(oPartDoc, oDef, pn, failedExports, userOptions);
		else
			ProcessNonIPart(oPartDoc, oDef, oFileName, failedExports, userOptions);
	}
}