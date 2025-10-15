using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

/// <summary>
/// 
/// </summary>
public class mod1
{
    public static Dictionary m1g0f0()
    {
        var rt = new Dictionary();
        {
            var withBlock = dcAiSheetMetal(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument)));
            long gt0 = 0;
            long eq0 = 0;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocPart(withBlock.get_Item(ky));
                    {
                        var withBlock2 = vcChkFlatPat(aiCompDefShtMetal(withBlock1.ComponentDefinition));
                        if (Abs(withBlock2.X * withBlock2.Y * withBlock2.Z) > 0)
                        {
                            Debug.Print.X(*withBlock2.Y * withBlock2.Z, ky);
                            // Stop
                            gt0 = 1 + gt0;
                        }
                        else
                            // Stop
                            eq0 = 1 + eq0;
                    }
                }
            }

            Debug.Print(gt0, eq0);
        }
        return rt;
    }

    public static Dictionary m1g0f1()
    {
        var rt = new Dictionary();
        {
            var withBlock = dcAssyDocComponents(ThisApplication.ActiveDocument, null, 1);
            const long gt0 = 0;
            const long eq0 = 0;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocument(withBlock.get_Item(ky));
                    if (withBlock1.DocumentInterests.HasInterest(guidDesignAccl))
                        Debugger.Break();
                }
            }

            Debug.Print(gt0, eq0);
        }
        return rt;
    }

    public static Vector vcDiaBox(Box bx)
    {
        //Given a Box dynamic
        //containing diametrically opposed
        //MinPoint and MaxPoint Objects,
        //return the Vector from Min to Max
        {
            return bx.MinPoint.VectorTo(bx.MaxPoint);
        }
    }

    public static Vector vc2flatPat(SheetMetalComponentDefinition df)
    {
        //From an Inventor Sheet Metal Component Definition,
        //obtain a Vector representing the translation
        //of the Folded Model's bounding box diagonal
        //to that of the Flat Pattern.
        // '
        Vector rt;

        {
            rt = vcDiaBox(df.RangeBox);
            rt.SubtractVector(df.HasFlatPattern ? vcDiaBox(df.FlatPattern.RangeBox) : rt);
        }
        // '
        //If they're the same point, this vector should
        //have zero length, however, this is NOT proof
        //positive of an invalid Flat Pattern. A valid
        //flat piece, with no folds, should produce
        //the same result.
        // '
        //A good follow-up check would probably be to
        //compare the Flat Pattern diagonal vector's
        //Z component to the model's Thickness

        return rt;
    }
    // Debug.Print vc2flatPat(aiCompDefShtMetal(aiDocPart(ThisApplication.Documents(2)).ComponentDefinition)).Length

    public static Vector vcCubicThickness(SheetMetalComponentDefinition df)
    {
        double hk = df.Thickness.Value;
        {
            var withBlock = ThisApplication.TransientGeometry;
            return withBlock.CreateVector(hk, hk, hk);
        }
    }

    public static Vector vcChkFlatPat(SheetMetalComponentDefinition df)
    {
        //From an Inventor Sheet Metal Component Definition,
        //subtract a vector of cubic thickness from the
        //diagonal vector of either its Flat Pattern,
        //if available, or otherwise, the Folded Model.
        // '
        Vector rt;

        {
            rt = vcDiaBox(df.HasFlatPattern ? df.FlatPattern.RangeBox : df.RangeBox);
        }
        rt.SubtractVector(vcCubicThickness(df));
        // '
        //If the model is a valid sheet metal part,
        //one of the dimensions of its flat pattern's
        //bounding box diagonal should either equal
        //the defined sheet metal thickness,
        //or fall very close. At least, in theory.
        // '
        //Plan at this point is to try to determine
        //just how often this bears out.
        //While this HAS failed one pretest, the Flat
        //Pattern of that model includes features;
        //a relatively infrequent occurrence, and
        //quite possibly one that can throw off
        //the boundaries.

        return rt;
    }
    // Debug.Print vcChkFlatPat(aiCompDefShtMetal(aiDocPart(ThisApplication.Documents(2)).ComponentDefinition)).Length

    public static dynamic m1tst0()
    {
        Debug.Print(iFacAssy(aiCompDefAssy(
            aiDocAssy(aiCompDefAssy(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition).Occurrences(1)
                .Definition.Document).ComponentDefinition)));
    }

    public static dynamic m1tst1()
    {
        // Dim bm As Inventor.BOMStructureEnum

        foreach (var ky in Filter(dcAssyCompAndSub(aiDocDefAssy(ThisApplication.ActiveDocument).Occurrences).Keys,
                     "(DC)"))
        {
            {
                var withBlock = aiDocPart(ThisApplication.Documents.ItemByName((ky)));
                Property pr = withBlock.PropertySets("Design Tracking Properties").get_Item("Cost Center");
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    if (withBlock1.BOMStructure == kPurchasedBOMStructure) continue;
                    withBlock1.BOMStructure = kPurchasedBOMStructure;
                    pr.Value = "D-PTS";
                    Debug.Print(pr.Value, withBlock1.BOMStructure, ky);
                }
            }
        }
    }

    public static iAssemblyFactory iFacAssy(AssemblyComponentDefinition ob)
    {
        {
            return ob.iAssemblyFactory ?? ob.iAssemblyMember?.ParentFactory;
        }
    }

    public static Document aiOccDoc(ComponentOccurrence ob)
    {
        return ob.Definition.Document;
    }

    public static AssemblyComponentDefinition aiDocDefAssy(AssemblyDocument ob)
    {
        return ob.ComponentDefinition;
    }

    public static Dictionary dcAssyComponentsImmediate(AssemblyDocument ob)
    {
        var rt = new Dictionary();
        foreach (ComponentOccurrence oc in ob.ComponentDefinition.Occurrences)
        {
            // aiDocument(oc.Definition.Document).FullFileName
            Document cp = oc.Definition.Document;
            var fn = cp.FullFileName;
            {
                if (!rt.Exists(fn))
                    rt.Add(fn, cp);
            }
        }

        return rt;
    }
    // For Each pn In new [] {pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnArea, pnLength, pnWidth, pnThickness): Debug.Print pn & "=" & aiDocument(ThisApplication.ActiveDocument).PropertySets.get_Item(gnCustom).Item((pn)).Value: Next

    public static long iSyncPartFactory(PartDocument pd)
    {
        // '
        //Backport iPart Member Properties to parent Factory
        // '

        long rt = 0;
        {
            var withBlock = pd.ComponentDefinition;
            var ps = aiDocument(withBlock.Document).PropertySets.get_Item(gnCustom);
            if (withBlock.iPartMember == null)
            {
            }
            else
            {
                var withBlock1 = withBlock.iPartMember;
                var dcRows = dcIPartTbRows(withBlock1.ParentFactory.TableRows);
                VbMsgBoxResult ck;
                if (dcRows.Exists(pd.DisplayName))
                {
                    var dcCols = dcIPartTbCols(withBlock1.ParentFactory.TableColumns);

                    Information.Err().Clear();
                    var tr = withBlock1.Row;

                    // REV[2022.03.23.1624]
                    // adding "pickup" attempt to capture
                    // and recover from errors encountered
                    // in trying to retrieve .Row directly
                    if (Information.Err().Number == 0)
                    {
                    }
                    else
                    {
                        tr = dcRows.get_Item(pd.DisplayName);
                        if (tr.MemberName == pd.DisplayName || tr.PartName == pd.DisplayName)
                            Information.Err().Clear();
                        else
                            Debugger.Break();
                    }

                    if (Information.Err().Number == 0)
                    {
                        foreach (Property pr in ps)
                        {
                            if (!dcCols.Exists(pr.Name)) continue;
                            Debug.Print(pr.Name);
                            {
                                var withBlock2 = tr.get_Item(dcCols.get_Item(pr.Name));
                                Debug.Print(" ");
                                if (pr.Value == withBlock2.Value)
                                    // Stop 'No change necessary
                                    Debug.Print(" (NO CHANGE)");
                                else
                                {
                                    withBlock2.Value = pr.Value;
                                    if (Information.Err().Number == 0)
                                    {
                                        rt = 1 + rt;

                                        //The update invalidated the dynamic
                                        //We'll have to grab it again
                                        {
                                            var withBlock3 = tr.get_Item(dcCols.get_Item(pr.Name));
                                            Debug.Print(" -> ");
                                        }
                                    }
                                    else
                                        Debug.Print(" <!!ERROR!!> Couldn't Change");
                                }
                            }
                        }

                        Debug.Print(""); // Breakpoint Landing
                    }
                    else
                    {
                        Debug.Print("=== CAN'T SYNC IFACTORY ===");
                        Debug.Print(" Failed to access Row");
                        Debug.Print(" for Member " + pd.DisplayName);
                        Debug.Print(" of Factory " + aiDocument(withBlock1.ParentFactory.Parent).DisplayName);
                        Debug.Print("=== PLEASE CHECK PARENT ===");
                        Debug.Print("====== FACTORY TABLE ======");
                        Debug.Print("Error 0x" + Hex(Information.Err()Number) + "(");
                        Debug.Print(" " + Information.Err()Description);
                        Debug.Print("===========================");
                        ck = MessageBox.Show(Join(new[]
                            {
                                "" + "iPart Member " + pd.DisplayName,
                                "in Factory" + aiDocument(withBlock1.ParentFactory.Parent).DisplayName,
                                "could not be accessed for updates.", "", "Its Row might still be present",
                                "but somehow unavailable.", "", "Please review iPart Factory."
                            }, Constants.vbCrLf),
                            Constants.vbOKCancel, "ERROR ACCESSING MEMBER ROW!");
                        if (ck == Constants.vbCancel)
                            Debugger.Break();
                    }

                    Debug.Print(""); // Breakpoint Landing
                }
                else
                {
                    Debug.Print("==== CAN'T FIND MEMBER ====");
                    Debug.Print(" Failed to locate Row");
                    Debug.Print(" for Member " + pd.DisplayName);
                    Debug.Print(" of Factory " + aiDocument(withBlock1.ParentFactory.Parent).DisplayName);
                    Debug.Print("=== PLEASE CHECK PARENT ===");
                    Debug.Print("====== FACTORY TABLE ======");
                    // Debug.Print "Error 0x" & Hex$(Err()Number) & "("; CStr(Err()Number) & ")"
                    // Debug.Print " " & Err()Description
                    // Debug.Print "==========================="
                    ck = MessageBox.Show(Join(new[]
                        {
                            "" + "iPart Member " + pd.DisplayName, "could not be located in Factory",
                            aiDocument(withBlock1.ParentFactory.Parent).DisplayName, "",
                            "Its Row might have been removed",
                            "or separated from the main table.", "", "Please review iPart Factory Table."
                        },
                        Constants.vbCrLf), Constants.vbOKCancel, "WARNING!! MEMBER NOT FOUND!");
                    if (ck == Constants.vbCancel)
                        Debugger.Break();
                }
            }
        }
    }

    public static long iSyncAssyFactory(AssemblyDocument pd)
    {
        // '
        //Backport iPart Member Properties to parent Factory
        // '

        long rt = 0;
        {
            var withBlock = pd.ComponentDefinition;
            var ps = aiDocument(withBlock.Document).PropertySets.get_Item(gnCustom);
            if (withBlock.iAssemblyMember == null)
            {
            }
            else
            {
                var withBlock1 = withBlock.iAssemblyMember;
                var dcRows = dcIAssyTbRows(withBlock1.ParentFactory.TableRows);
                if (dcRows.Exists(pd.DisplayName))
                {
                    var dcCols = dcIAssyTbCols(withBlock1.ParentFactory.TableColumns);
                    var tr = withBlock1.Row;
                    foreach (Property pr in ps)
                    {
                        if (!dcCols.Exists(pr.Name)) continue;
                        Debug.Print(pr.Name);
                        {
                            var withBlock2 = tr.get_Item(dcCols.get_Item(pr.Name));
                            Debug.Print(" ");
                            if (pr.Value == withBlock2.Value)
                                // Stop 'No change necessary
                                Debug.Print(" (NO CHANGE)");
                            else
                            {
                                withBlock2.Value = pr.Value;
                                if (Information.Err().Number == 0)
                                {
                                    rt = 1 + rt;

                                    //The update invalidated the dynamic
                                    //We'll have to grab it again
                                    {
                                        var withBlock3 = tr.get_Item(dcCols.get_Item(pr.Name));
                                        Debug.Print(" -> ");
                                    }
                                }
                                else
                                    Debug.Print(" <!!ERROR!!> Couldn't Change");
                            }
                        }
                    }
                }
                else
                    Debugger.Break();
            }
        }
    }

    public static Dictionary dcColumnsIPart(PartDocument pd)
    {
        //Retrieve Dictionary of iPart Factory Table Columns
        //If supplied Part Document is NOT an iPart Factory
        //OR Member, returned Dictionary will be null
        // '
        {
            var withBlock = pd.ComponentDefinition;
            if (withBlock.iPartMember == null)
            {
                return withBlock.iPartFactory == null
                    ? new Dictionary()
                    : dcIPartTbCols(withBlock.iPartFactory.TableColumns);
            }

            return dcIPartTbCols(withBlock.iPartMember.ParentFactory.TableColumns);
        }
    }

    public static Dictionary dcColumnsIAssy(AssemblyDocument pd)
    {
        //Retrieve Dictionary of iAssembly Factory Table Columns
        //If supplied Assembly Document is NOT an iAssembly Factory
        //OR Member, returned Dictionary will be null
        // '
        {
            var withBlock = pd.ComponentDefinition;
            if (withBlock.iAssemblyMember == null)
            {
                return withBlock.iAssemblyFactory == null
                    ? new Dictionary()
                    : dcIAssyTbCols(withBlock.iAssemblyFactory.TableColumns);
            }

            return dcIAssyTbCols(withBlock.iAssemblyMember.ParentFactory.TableColumns);
        }
    }

    public static Dictionary dcIPartTbCols(iPartTableColumns ls)
    {
        var rt = new Dictionary();

        foreach (iPartTableColumn it in ls)
        {
            rt.Add(it.DisplayHeading, it.Index);
            if (!Information.Err().Number) continue;
            rt.Add(it.FormattedHeading, it.Index);
            Information.Err().Clear();
        }

        return rt;
    }

    public static Dictionary dcIPartTbRows(iPartTableRows ls)
    {
        var rt = new Dictionary();

        // REV[2022.08.05.1003]
        // added error trap code to more gracefully handle
        // errors accessing iPart/iAssembly factory table
        // (replicated changes also to dcIAssyTbRows)
        long ck = ls.Count; // to trigger potential error

        if (Information.Err().Number == 0)
        {
            foreach (iPartTableRow it in ls)
            {
                // REV[2022.03.23.1618]
                // replacing Index of iPartTableRow
                // with iPartTableRow itself, so it
                // can just be pulled directly out
                // of the Dictionary by the client
                // process. if it needs the Index,
                // it can just get it itself, right?
                rt.Add(it.MemberName, it);
                rt.Add(it.PartName, it);
                Debug.Print(""); // debug landing
            }
        }
        else
        {
            Debug.Print(
                "ERROR " + Convert.ToHexString(Information.Err().Number) + " (" + Hex(Information.Err().Number) + ")" +
                Constants.vbCrLf + Information.Err().Description);
            Debug.Print(Join(new[]
            {
                "Could not access Table Rows", "for member of iPart factory."
            }, Constants.vbCrLf));
            // Stop
            Debug.Print(""); // Breakpoint Landing
        }

        return rt;
    }

    public static Dictionary dcIAssyTbCols(iAssemblyTableColumns ls)
    {
        var rt = new Dictionary();

        foreach (iAssemblyTableColumn it in ls)
        {
            rt.Add(it.DisplayHeading, it.Index);
            if (!Information.Err().Number) continue;
            rt.Add(it.FormattedHeading, it.Index);
            Information.Err().Clear();
        }

        return rt;
    }

    public static Dictionary dcIAssyTbRows(iAssemblyTableRows ls)
    {
        var rt = new Dictionary();

        // REV[2022.08.05.1003]
        // added error trap code to more gracefully handle
        // errors accessing iPart/iAssembly factory table
        // (replicated from dcIPartTbRows)
        long ck = ls.Count; // to trigger potential error

        if (Information.Err().Number == 0)
        {
            foreach (iAssemblyTableRow it in ls)
            {
                rt.Add(it.MemberName, it.Index);
                rt.Add(it.DocumentName, it.Index);
                Debug.Print(""); // debug landing
            }
        }
        else
        {
            Debug.Print(
                "ERROR " + Convert.ToHexString(Information.Err().Number) + " (" + Hex(Information.Err().Number) + ")" +
                Constants.vbCrLf + Information.Err().Description);
            Debug.Print(Join(new[]
            {
                "Could not access Table Rows", "for member of iAssembly factory."
            }, Constants.vbCrLf));
            Debugger.Break();
        }

        return rt;
    }

    public static iPartTableColumn m1g1f2(dynamic vr)
    {
        return vr;
    }

    public static long m1g1f3(AssemblyDocument ad)
    {
        long rt = 0;
        foreach (ComponentOccurrence oc in ad.ComponentDefinition.Occurrences) // (1)
        {
            {
                var withBlock = oc.Definition;
                // Debug.Print Hex$(aiDocument(.Document).Type)
                // Stop
                if (aiDocument(withBlock.Document).DocumentType == kPartDocumentObject)
                    rt = rt + iSyncPartFactory(aiDocPart(withBlock.Document));
            }
        }

        return rt;
    }

    public static long m1g1f4()
    {
        return m1g1f3(aiDocAssy(ThisApplication.ActiveDocument));
    }

    public static Dictionary m1g1f5(PartDocument pd)
    {
        //Retrieve Dictionary of Custom Properties
        // '
        // Dim dcMember As Scripting.Dictionary

        var rt = new Dictionary();
        {
            var psMember = pd.PropertySets.get_Item(gnCustom);
            {
                var withBlock1 = pd.ComponentDefinition;
                if (!withBlock1.iPartMember == null)
                {
                    {
                        var withBlock2 = withBlock1.iPartMember;
                        Dictionary dcFactry;
                        PropertySet psFactry;
                        {
                            var withBlock3 = aiDocument(withBlock2.ParentFactory.Parent);
                            psFactry = withBlock3.PropertySets.get_Item(gnCustom);
                            dcFactry = dcAiPropsInSet(psFactry);
                        }

                        foreach (Property pr in psMember)
                        {
                            if (!dcFactry.Exists(pr.Name))
                                // rt.Add pr.Name, pr
                                rt = dcWithProp(psFactry, pr.Name, pr.Value, rt);
                        }
                    }
                }
            }
        }
        return rt;
    }

    public static string m1g1f5t0()
    {
        return Join(
            m1g1f5(aiDocPart(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences.get_Item(1)
                .Definition.Document)).Keys, Constants.vbCrLf);
    }
}