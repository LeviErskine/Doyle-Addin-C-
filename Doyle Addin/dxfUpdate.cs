using Doyle_Addin.Options;
using Inventor;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;
using Application = Inventor.Application;

namespace Doyle_Addin
{
    internal static class DxfUpdate
    {
        public static void RunDxfUpdate(object? unknown, Application thisApplication)
        {
            var oPartDoc = (PartDocument)thisApplication.ActiveDocument;
            var oDef = (SheetMetalComponentDefinition)oPartDoc.ComponentDefinition;
            var oFactory = oDef.iPartFactory;
            var failedExports = new List<string>();
            var oDoc = thisApplication.Documents;
            var pn = oPartDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();
            var oFileName = UserOptions.Load().DxfExportLocation + pn + ".dxf";


            // Check if a part is a factory
            if (oDef.IsiPartFactory)
            {
                var total = oFactory.TableRows.Count;
                oDoc.CloseAll(true);
                oPartDoc.ReleaseReference();

                // Create a flat pattern if missing
                if (!oDef.HasFlatPattern)
                {
                    try
                    {
                        oDef.Unfold();
                        oDef.FlatPattern.ExitEdit();
                    }
                    catch (Exception)
                    {
                        failedExports.Add("Failed to create flat pattern for: " + pn);
                    }
                }

                if (Directory.Exists(oFactory.MemberCacheDir) == false)
                {
                    Directory.CreateDirectory(oFactory.MemberCacheDir);
                }

                var partFiles = Directory.GetFiles(oFactory.MemberCacheDir);

                if (total > partFiles.Length)
                {
                    var result =
                        Interaction.MsgBox(
                            "Warning: The factory has " + total + " members, but " + partFiles.Length +
                            " files were found in the folder. Generate files?.", MsgBoxStyle.YesNo, "Missing Members");
                    if (result == MsgBoxResult.Yes)
                    {
                        oDoc.CloseAll(true);
                        oPartDoc.ReleaseReference();
                        foreach (var iPartTableRow in oFactory.TableRows)
                            oFactory.CreateMember(iPartTableRow);
                    }
                    else
                    {
                        return;
                    }
                }

                foreach (var filepath in Directory.GetFiles(oFactory.MemberCacheDir))
                {
                    var openedDoc = (PartDocument)oDoc.Open(filepath);
                    var memberDef = openedDoc.ComponentDefinition as SheetMetalComponentDefinition;
                    var partnumber = openedDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value
                        .ToString();
                    var oiFileName = UserOptions.Load().DxfExportLocation + partnumber + ".dxf";


                    // Validate flat pattern
                    if (memberDef != null && Conversions.ToBoolean(Operators.AndObject(
                            memberDef.FlatPattern.FlatBendResults.Count == 0,
                            Operators.ConditionalCompareObjectGreater(
                                memberDef.FlatPattern.RangeBox.MaxPoint.Z - memberDef.FlatPattern.RangeBox.MinPoint.Z,
                                Operators.AddObject(memberDef.Thickness.Value, 0.003d), false))))
                    {
                        failedExports.Add("Invalid flat pattern for: " + partnumber);
                        continue;
                    }

                    // Prepare output path and format
                    const string oFormat = "FLAT PATTERN DXF?AcadVersion=2018" +
                                           "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" +
                                           "&BendUpLayer=UP&BendUpLayerColor=255;0;0" +
                                           "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" +
                                           "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" +
                                           "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" +
                                           "&TangentLayer=RADIUS&TangentLayerColor=255;255;0";

                    // Export DXF
                    try
                    {
                        memberDef?.DataIO.WriteDataToFile(oFormat, oiFileName);
                        openedDoc.Close(true);
                    }
                    catch (Exception)
                    {
                        failedExports.Add("DXF failed to generate for: " + partnumber);
                    }
                }

                if (failedExports.Count > 0)
                {
                    Interaction.MsgBox(failedExports.Count + " Members have errors and were skipped." +
                                       Constants.vbCrLf + string.Join(Constants.vbCrLf, failedExports));
                }
                else
                {
                    Interaction.MsgBox("Created " + total + " DXFs. All exports succeeded.");
                }
            }

            else
            {
                // Debug
                // MsgBox(ThisApplication.ActiveDocument._SickNodesCount.ToString, MsgBoxStyle.OkOnly, "Sick Nodes")
                // MsgBox(ThisApplication.ActiveDocument._ComatoseNodesCount.ToString, MsgBoxStyle.OkOnly, "Comatose Nodes")
                // All the same as above, for non-iparts
                const string oFormat = "FLAT PATTERN DXF?AcadVersion=2018" +
                                       "&BendDownLayer=DOWN&BendDownLayerColor=255;0;0" +
                                       "&BendUpLayer=UP&BendUpLayerColor=255;0;0" +
                                       "&OuterProfileLayer=OUTER&OuterProfileLayerColor=0;0;255" +
                                       "&InteriorProfilesLayer=INNER&InteriorProfilesLayerColor=0;0;0" +
                                       "&ArcCentersLayer=POINT&ArcCentersLayerColor=255;0;255" +
                                       "&TangentLayer=RADIUS&TangentLayerColor=255;255;0";

                // Skip export if comatose nodes exist
                if (oPartDoc._ComatoseNodesCount > 0 | oPartDoc._SickNodesCount > 0)
                {
                    Interaction.MsgBox(oPartDoc.DisplayName + " has errors, fix before export." + Constants.vbCrLf +
                                       string.Join(Constants.vbCrLf, failedExports));
                    return;
                }

                if (!oDef.HasFlatPattern)
                {
                    try
                    {
                        oDef.Unfold();
                        oDef.FlatPattern.ExitEdit();
                    }
                    catch (Exception)
                    {
                        Interaction.MsgBox("Failed to create flat pattern", MsgBoxStyle.OkOnly, "Error");
                        return;
                    }
                }

                try
                {
                    oDef.DataIO.WriteDataToFile(oFormat, oFileName);
                    Interaction.MsgBox(oPartDoc.DisplayName + " exported successfully.", MsgBoxStyle.Information,
                        "Success");
                }
                catch (Exception ex)
                {
                    Interaction.MsgBox(
                        "DXF failed to generate. Check connection to X drive" + Constants.vbCrLf + "Error: " +
                        ex.Message, MsgBoxStyle.Critical, "Error");
                }
            }
        }
    }
}