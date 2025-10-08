using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Doyle_Addin.Options;
using Inventor;
using Application = Inventor.Application;
using Environment = System.Environment;

namespace Doyle_Addin.DXFs;

internal static class DxfUpdate
{
    public static void RunDxfUpdate(Application thisApplication)
    {
        var oPartDoc = (PartDocument)thisApplication.ActiveDocument;
        var oDef = (SheetMetalComponentDefinition)oPartDoc.ComponentDefinition;
        var oFactory = oDef.iPartFactory;
        var failedExports = new List<string>();
        var oDoc = thisApplication.Documents;
        var pn = oPartDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value.ToString();
        var oFileName = UserOptions.Load().DxfExportLocation + pn + ".dxf";
        
        // Check if a part is a factory
        if (!oDef.IsiPartFactory) goto NonIPart;
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
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                failedExports.Add("Failed to create flat pattern for: " + pn);
            }
        }

        if (!Directory.Exists(oFactory.MemberCacheDir))
        {
            Directory.CreateDirectory(oFactory.MemberCacheDir);
        }

        var partFiles = Directory.GetFiles(oFactory.MemberCacheDir);

        if (total > partFiles.Length)
        {
            var result =
                MessageBox.Show(
                    @"Warning: The factory has " + total + @" members, but " + partFiles.Length +
                    @" files were found in the folder. Generate files?.", @"Missing Members",
                    MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
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
            Debug.Assert(memberDef != null, nameof(memberDef) + " != null");
            if (memberDef.FlatPattern.FlatBendResults.Count == 0 &&
                memberDef.FlatPattern.RangeBox.MaxPoint.Z - memberDef.FlatPattern.RangeBox.MinPoint.Z >
                Convert.ToDouble(memberDef.Thickness.Value) + 0.003d)
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
                memberDef.DataIO.WriteDataToFile(oFormat, oiFileName);
                if (true)
                {
                    openedDoc.Close(true);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                failedExports.Add("DXF failed to generate for: " + partnumber);
            }
        }
        
        if (failedExports.Count > 0)
        {
            MessageBox.Show(failedExports.Count + @" Members have errors and were skipped." +
                            Environment.NewLine + string.Join(Environment.NewLine, failedExports));
        }
        else
        {
            MessageBox.Show(@"Created " + total + @" DXFs. All exports succeeded.");
        }
        return;

NonIPart:
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
                MessageBox.Show(oPartDoc.DisplayName + @" has errors, fix before export." + Environment.NewLine +
                                string.Join(Environment.NewLine, failedExports));
                return;
            }

            if (!oDef.HasFlatPattern)
            {
                try
                {
                    oDef.Unfold();
                    oDef.FlatPattern.ExitEdit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    MessageBox.Show(@"Failed to create flat pattern", @"Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }
            }
            
            try
            {
                oDef.DataIO.WriteDataToFile(oFormat, oFileName);
                MessageBox.Show(oPartDoc.DisplayName + @" exported successfully.", @"Success", MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    @"DXF failed to generate. Check connection to X drive" + Environment.NewLine + @"Error: " +
                    ex.Message, @"Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}