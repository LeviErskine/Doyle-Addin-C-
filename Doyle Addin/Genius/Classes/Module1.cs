using Microsoft.VisualBasic;
using stdole;

namespace Doyle_Addin.Genius.Classes;

class Module1

{
    public static Dictionary dcCutTimePerimeter(Document ad, Dictionary dc = null, bool incTop = false)
    {
        Document ActiveDoc;

        // If dc Is Nothing Then
        // dcCutTimePerimeter = dcCutTimePerimeter( ad, New Scripting.Dictionary, incTop )
        // Else
        var rt = new Dictionary();

        {
            var withBlock = dcAiDocComponents(ad, dc, incTop);
            foreach (var ky in withBlock.Keys)
            {
                Document pt = aiDocument(withBlock.get_Item(ky));
                rt.Addpt.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value(null, fpPerimeterInch(pt));
            }
        }

        return rt;
    }
    // Debug.Print dumpLsKeyVal(dcCutTimePerimeter(ThisApplication.ActiveDocument))

    public static long mdl1g0f0()
    {
        // Dim ps As Inventor.PropertySet
        var dc =
            dcAssyDocComponents(
                ThisApplication.Documents.ItemByName(@"C:\Doyle_Vault\Designs\Misc\andrewT\02\02-weldmentStd-01.iam"));
        {
            foreach (var ky in dc.Keys)
            {
                Document ad = aiDocument(dc.get_Item(ky));
                {
                    var withBlock1 = dcGeniusProps(ad);
                    if (!withBlock1.Exists(pnRawMaterial)) continue;
                    Property pr = ad.PropertySets(gnCustom).get_Item(pnRawMaterial);
                    {
                        var withBlock2 = pr;
                        Debug.Print.Value();

                        if (.Value Like "FM-*");
                        {
                            with new fmTest1;

                            if (.AskAbout(ad) = vbYes);
                            {
                                with.ItemData;
                            }

                            pr.Value =  .get_Item(pnRawMaterial);
                            end with
                        }
                        end With
                    }
                }
            }
        }
    }
}
// Debug.Print cnGnsDoyle.Execute("select I.ItemID, I.Thickness, I.Item, I.Description1 from vgMfiItems as I where I.Family='DSHEET'").GetString

public static long mdl1g1f0()
{
    {
        var withBlock = new fmTest1();
        withBlock.AskAbout(ThisApplication.ActiveDocument);
    }
}

public static float mdl1g1f2(Label lb) // , txt As String
{
    long x0;
    long y0;

    Control ct = lb;
    {
        // x0 = .Left
        // y0 = .Top
        float x1 = ct.Width;
        float y1 = ct.Height;
        {
            // .Caption = txt
            lb.AutoSize = true;
            lb.AutoSize = false;
        }
        ct.Width = (int)x1;
        return ct.Height - y1;
    }
}

public static float mdl1g1f3(Control ct, float byX, float byY)
{
    {
        ct.Left = (int)(ct.Left + byX);
        ct.Top = (int)(ct.Top + byY);
    }
    return Sqr(byX * byX + byY * byY);
}

// For lack of a better place to put it, creating this node

// The following is a basic example of accessing Parameters,

// such as dimensions, from an Inventor Part Document.
public static long mdl1g1f1()
{
    {
        var withBlock = aiDocPart(ThisApplication.ActiveDocument);
        // aiDocPart casts an Inventor Part Document
        // from its general Document reference, if valid.
        {
            var withBlock1 = withBlock.ComponentDefinition.Parameters.get_Item("Thickness");
            Debug.Print.ExposedAsProperty();
            Debug.Print.Value();
        }
    }
}
// This example was written as a quick one-off to see how

// an Inventor Parameter like the Thickness setting for

// Sheet Metal Parts might have its "Export" status

// modified programmatically.

public static dynamic aiPropVal(Property pr, dynamic ifNot = "")
{
    return pr == null ? ifNot : aiPropValAux(pr.Value, ifNot);
}
// This example was written as a quick one-off to see how

// an Inventor Parameter like the Thickness setting for

// Sheet Metal Parts might have its "Export" status

// modified programmatically.

public static dynamic aiPropValAux(dynamic vl, string ifNot = "")
{
    if (vl is null) return null;
    switch (vl)
    {
        case null:
            return ifNot;
        case StdPicture:
            return "<stdole.StdPicture>";
            Debug.Print(""); // Breakpoint Landing
            break;
        default:
            Debugger.Break(); // and see what we need to do
            return "<dynamic:" + TypeName(vl) + ">";
    }

    return vl;
}

public static Property aiPropGnsItmFamily(Document AiDoc)
{
    if (AiDoc == null)
        return null;
    return AiDoc.PropertySets(gnDesign).get_Item(pnFamily);
}

public static Property aiPropShtMetalThickness(PartDocument adPart)
{
    if (adPart == null)
        return null;
    if (adPart.SubType != guidSheetMetal) return null;
    if (smThicknessExposed(adPart.ComponentDefinition))
        return adPart.PropertySets(gnCustom).get_Item(pnThickness);
    return null;
}

public static long smThicknessExposed(SheetMetalComponentDefinition smDef)
{
    if (smDef.Parameters.IsExpressionValid(pnThickness, "in"))
        return parExposed(smDef.Parameters(pnThickness), true);
    Debugger.Break();
}

public static long parExposed(Parameter par, bool tryTo = false)
{
    // ' Check Inventor Parameter for exposure as Property.
    // ' Return 0 if not, unless caller requests exposure
    // ' (tryTo <> 0). Nonzero return indicates exposed
    // ' Parameter, with sign indicating initial status.
    // ' -1 indicates Parameter already exposed
    // ' 1 indicates status change to expose it.
    // ' No provision is made for failure to expose,
    // ' nor to reverse exposure status.
    {
        if (par.ExposedAsProperty)
            return -1;
        if (!tryTo) return 0;
        par.ExposedAsProperty = true;
        return 1 & parExposed(par);
    }
}

public static Dictionary dcGnsPropsListed(Document ad, dynamic ls, Dictionary dc = null, bool ifNone = true)
{
    while (true)
    {
        // dcGnsPropsListed --
        // Return a Dictionary of any
        // Properties in the supplied list
        // from the "custom" PropertySet.
        // 
        // Missing Property names are addressed
        // in one of three (present) ways,
        // based on optional argument ifNone:
        // 0 - do not add to Dictionary
        // missing name is missing
        // 1 - attempt to create. failure
        // returns Nothing, which is
        // still not added
        // 2 - add Nothing to Dictionary
        // under missing name
        // 3 - attempt to create, adding
        // Nothing for any failures
        // (combines options 1 and 2)
        // 

        // rt = New Scripting.Dictionary
        var ps = ad.PropertySets.get_Item(gnCustom);

        if (dc == null)
        {
            dc = new Dictionary();
            continue;
        }

        if (ls is Array)
        {
            var mkNf = ifNone & 1; // try to make any not found
            var rtNf = ifNone & 2; // return Nothing for not found
            // originally used IIf construct
            // to force mapping of exact values
            // of ifNone to corresponding behaviors.
            // 
            // changed to bitcode matching once clear
            // that each bit would map exclusively
            // to a particular behavior, and could
            // be combined with the other, if desired.
            // 

            foreach (var ky in
                     ls) // new string[] {pnMass, pnArea, pnWidth, pnLength, pnRawMaterial, pnRmQty, pnRmUnit) ', "SPEC01", "SPEC02", "SPEC03", "SPEC04", "SPEC05", "SPEC06", "SPEC07", "SPEC08", "SPEC16"'
            {
                Property pr = aiGetProp(ps, Convert.ToHexString(ky), mkNf);
                dynamic wk = new[] { pr };

                if (pr == null)
                {
                    // if supposed to return Nothings

                    if (rtNf == 0) wk = Array.Empty<string>();
                }

                if (UBound(wk) < LBound(wk))
                {
                }
                else
                {
                    if (dc.Exists(ky))
                    {
                        dc.Remove(ky);
                        // WARNING[2021.11.19]
                        // This was added to permit
                        // replacement of elements
                        // already present under a
                        // supplied key. It might
                        // NOT be the best way to
                        // address this situation.
                        // Be prepared to correct
                        // this with a more robust
                        // solution in future.
                        // Meanwhile, have a
                        Debug.Print(""); // Breakpoint Landing
                    }

                    dc.Add(ky, pr);
                }
            }

            return dc; // rt
        }

        if (VarType(ls) == Constants.vbString)
            return dcGnsPropsListed(ad, new[] { ls }, dc, ifNone);
        Debugger.Break();

        break;
    }
}

public static Dictionary dcGnsPropsPart(Document ad, Dictionary dc = null, bool ifNone = true)
{
    // dcGnsPropsPart
    // 
    // REV[2021.11.18]:
    // Added pnThickness to list
    // of Properties to return.
    // 
    return dcGnsPropsListed(ad, new string[]
    {
        pnMass, pnArea, pnWidth, pnLength, pnThickness, pnRawMaterial, pnRmQty, pnRmUnit
    }, dc, ifNone);
}

public static Dictionary dcGnsPropsAssy(Document ad, Dictionary dc = null, long ifNone = 1)
{
    // Dim rt As Scripting.Dictionary
    // Dim ps As Inventor.PropertySet
    // Dim pr As Inventor.Property
    // Dim ky As dynamic

    return dcGnsPropsListed(ad,
        new[] { pnMass, "SPEC01", "SPEC02", "SPEC03", "SPEC04", "SPEC05", "SPEC06", "SPEC07", "SPEC08", "SPEC16" },
        dc, ifNone);
}

public static Dictionary dcProps4genius(Document ad, Dictionary dc = null, bool Create = true)
{
    while (true)
    {
        // Dim rt As Scripting.Dictionary
        PropertySet ps;
        Property pr;
        dynamic ky;

        if (dc == null)
        {
            dc = new Dictionary();
        }
        else
        {
            return ad.DocumentType switch
            {
                kAssemblyDocumentObject => dcGnsPropsAssy(ad, dc, Create),
                kPartDocumentObject => dcGnsPropsPart(ad, dc, Create),
                _ => dc
            };
        }
    }
}

public static WorkPlanes mdl1g2f1(Document ad)
{
    return aiDocPart(ad).ComponentDefinition.WorkPlanes;
}

public static double mdl1g3f0(Document ad)
{
    var rt = ad.DocumentType switch
    {
        dynamic _ when kPartDocumentObject => aiDocPart(ad).ComponentDefinition.MassProperties.Mass,
        dynamic _ when kAssemblyDocumentObject => aiDocAssy(ad).ComponentDefinition.MassProperties.Mass,
        _ => 0
    };

    {
        var withBlock = ad.UnitsOfMeasure;
        return withBlock.ConvertUnits(rt, kKilogramMassUnits, kLbMassMassUnits); // .MassUnits)
    }
}

public static long mdl1g4f0()
{
    {
        var withBlock = ThisApplication.CommandManager.ControlDefinitions;
        long mx = withBlock.Count;
        for (long dx = 1; dx <= mx; dx++)
        {
            {
                var withBlock1 = withBlock.get_Item(dx);

                if (.InternalNameLike "*ault*")
                {
                    Debug.Print(CStr(dx) & ": " &.InternalName & "/" & .DisplayName);
                }
            }
        }
    }
}

public static Dictionary mdl1g5f0(Document ad)
{
    // The purpose of this function is to return a Dictionary
    // of Genius Family Inventor Properties
    // for each component Document of an assembly
    // or a single part Document.

    var rt = new Dictionary();
    {
        var withBlock = dcAiDocComponents(ad, new Dictionary(), 1) // sc
            ;
        foreach (var ky in withBlock.Keys)
        {
            {
                var withBlock1 = aiDocument(withBlock.get_Item(ky));
                rt.Add.FullFileName(null,
                    aiPropGnsItmFamily(withBlock1.PropertySets.Parent));
            }
        }
    }
    return rt;
}

public static Dictionary mdl1g5f1(Document ad)
{
    // This function calls mdl1g5f0 to retrieve a Dictionary
    // of Genius Family Inventor Properties, and then
    // transforms it into a Dictionary of Dictionaries
    // grouped by Family Property Value
    Dictionary gp;

    var rt = new Dictionary();
    {
        var withBlock = mdl1g5f0(ad);
        foreach (var ky in withBlock.Keys)
        {
            Property pr = aiProperty(withBlock.get_Item(ky));
            string fm = pr.Value;
            {
                if (!rt.Exists(fm))
                    rt.Add(fm, new Dictionary());
                dcOb(rt.get_Item(fm)).Add(ky, pr);
            }
        }
    }
    return rt;
}
// Debug.Print txDumpLs(mdl1g5f1(ThisApplication.ActiveDocument).Keys)
// Debug.Print txDumpLs(dcOb(mdl1g5f1(ThisApplication.ActiveDocument).get_Item("")).Keys)

public static Dictionary mdl1g5f2(Document ad)
{
    // The purpose of this function is to return a Dictionary
    // of Genius Family Inventor Properties
    // for each component Document of an assembly
    // or a single part Document.

    var fm = new fmTest1();
    var rt = new Dictionary();
    {
        var withBlock = mdl1g5f1(ad);
        if (!withBlock.Exists("")) return rt;
        {
            var withBlock1 = dcOb(withBlock.get_Item(""));
            foreach (var ky in withBlock1.Keys)
            {
                {
                    var withBlock2 = aiProperty(withBlock1.get_Item(ky));
                    if (fm.AskAbout(withBlock2.Parent) == Constants.vbOK)
                        Debugger.Break();
                    else
                        Debugger.Break();
                    Debugger.Break();
                }
            }
        }
    }
    return rt;
}

public static Dictionary mdl1g5f3(AssemblyDocument ad)
{
    // Scan immediate members of Assembly document
    // and group in Dictionary by declared Part Number
    // and sub-grouped by Full Document Name.
    // 
    // (I wonder if an ADO Recordset wouldn't be a better choice?)
    // 
    // Dim fm As fmTest1
    // Dim ky As dynamic

    // fm = New fmTest1
    var rt = new Dictionary();
    foreach (Document sd in from ComponentOccurrence oc in ad.ComponentDefinition.Occurrences
             select oc.Definition.Document)
    {
        string pn = sd.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;
        {
            if (rt.Exists(pn))
                dcAiDocsByFullDocName(sd, rt.get_Item(pn));
            else
                rt.Add(pn, dcAiDocsByFullDocName(sd, new Dictionary()));
        }
    }

    return rt;
}
// Debug.Print txDumpLs(mdl1g5f3(ThisApplication.ActiveDocument).Keys)

public static Dictionary mdl1g5f4(Dictionary dc)
{
    // Transform keys from supplied Dictionary
    // (expected from mdl1g5f3)
    // into header/member indented form.

    var dl = Constants.vbCrLf + Constants.vbTab;
    var rt = new Dictionary();
    {
        foreach (var ky in dc.Keys)
            rt.Add(ky + Constants.vbTab + Join(dcOb(dc.get_Item(ky)).Keys, Constants.vbCrLf + ky + Constants.vbTab),
                dc.get_Item(ky));
    }
    return rt;
}
// Debug.Print txDumpLs(mdl1g5f4(mdl1g5f3(ThisApplication.ActiveDocument)).Keys)

public static Dictionary dcAiDocsByFullDocName(Document ad, Dictionary dc)
{
    while (true)
    {
        // Add supplied Inventor Document
        // to supplied Dictionary
        // under its Full Document Name
        // (supports mdl1g5f3)

        var ky = ad.FullDocumentName;
        if (dc == null)
        {
            dc = new Dictionary();
        }
        else
        {
            {
                if (dc.Exists(ky))
                    dc.get_Item(ky) = 1 + dc.get_Item(ky);
                else
                    dc.Add(ky, 1);
            }
            return dc;
        }
    }
}

public static Dictionary dcAssyDocsByPtNum(AssemblyDocument ad)
{
    // Derived from mdl1g5f3
    // 
    // Scan immediate members of Assembly Document
    // and collect source Documents in Dictionary,
    // grouped by declared Part Number.
    // 

    var rt = new Dictionary();
    foreach (ComponentOccurrence oc in ad.ComponentDefinition.Occurrences)
    {
        Document sd = oc.Definition.Document;
        string pn = sd.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;
        {
            if (rt.Exists(pn))
            {
                if (sd == rt.get_Item(pn))
                {
                }
                else
                    Debugger.Break(); // and check it out
            }
            else
                rt.Add(pn, sd);
        }
    }

    return rt;
}
// Debug.Print txDumpLs(dcAssyDocsByPtNum(ThisApplication.ActiveDocument).Keys)

public static Dictionary dcAiDocsByCompList(Dictionary dc)
{
    // Derived from mdl1g5f4
    // Transform keys from supplied Dictionary
    // (expected from dcAssyDocsByPtNum)
    // into tab-delimited list form.

    var rt = new Dictionary();
    {
        foreach (var ky in dc.Keys)
        {
            Document sd = aiDocument(dc.get_Item(ky));
            {
                switch (sd.DocumentType)
                {
                    case kAssemblyDocumentObject:
                        rt.Add(
                            ky + Constants.vbTab +
                            Join(Split(txDumpLs(mdl1g5f4(mdl1g5f3(sd)).Keys), Constants.vbCrLf),
                                Constants.vbCrLf + ky + Constants.vbTab), sd);
                        break;
                    case kPartDocumentObject:
                    {
                        // Stop
                        var dl = Constants.vbCrLf + Constants.vbTab;
                        {
                            var withBlock2 = dcAiPropsInSet(sd.PropertySets.get_Item(gnCustom));
                            if (withBlock2.Exists(pnRawMaterial))
                            {
                                dl = Trim(aiProperty(withBlock2.get_Item(pnRawMaterial)).Value);
                                if (Strings.Len(dl) == 0)
                                    dl = "NO_RAW_STOCK" + Constants.vbTab + "<No Raw Stock Declared>";
                                else
                                {
                                    Debugger.Break();
                                    {
                                        var withBlock3 = cnGnsDoyle.Execute(Join(new[]
                                            {
                                                "select Description1", "from vgMfiItems",
                                                "where Item = '" + dl + "';"
                                            },
                                            Constants.vbCrLf));
                                        if (withBlock3.BOF | withBlock3.EOF)
                                            dl = dl + Constants.vbTab + "<Stock Number Not Found>";
                                        else
                                            dl = dl + Constants.vbTab + withBlock3.Fields(0).Value;
                                    }
                                }
                            }

                            else
                                dl = "NO_RAW_STOCK" + Constants.vbTab + "<No Raw Stock Declared>";
                        }

                        rt.Add(ky + Constants.vbTab + dl, sd);
                        break;
                    }
                    case kUnknownDocumentObject:
                    case kDrawingDocumentObject:
                    case kPresentationDocumentObject:
                    case kDesignElementDocumentObject:
                    case kForeignModelDocumentObject:
                    case kSATFileDocumentObject:
                    case kNoDocument:
                    case kNestingDocument:
                    default:
                        rt.Add(ky + Constants.vbTab + "(UNSUPPORTED DOCUMENT TYPE)", sd);
                        break;
                }
            }
        }
    }

    return rt;
}
// Debug.Print txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)
// send2clipBd txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)

public static ADODB.Recordset rsWinUpdHist()
{
    // Windows Update History
    WUApiLib.IUpdateHistoryEntry it;

    ADODB.Recordset rt = rsNewWinUpdHist;
    dynamic ls = new[] { "ResultCode", "Operation", "Title", "Description", "Date" };
    {
        var withBlock = new WUApiLib.UpdateSession() // .CreateUpdateSearcher
            ;
        {
            var withBlock1 = withBlock.CreateUpdateSearcher;
            foreach (var it in withBlock1.QueryHistory(0, withBlock1.GetTotalHistoryCount))
            {
                {
                    var withBlock2 = it;
                    // Debug.Print .ResultCode, .Operation, .Title, .Description, .Date
                    rt.AddNew(ls, new[]
                    {
                        withBlock2.ResultCode, withBlock2.Operation, withBlock2.Title, withBlock2.Description,
                        withBlock2.Date
                    });
                }
            }

            rt.Filter = "";
        }
    }
    return rt;
}

public static ADODB.Recordset rsNewWinUpdHist()
{
    var rt = new ADODB.Recordset();
    {
        {
            var withBlock1 = rt.Fields;
            // .Append "", adBigInt
            // .Append "", adVarChar, 1024
            withBlock1.Append("ResultCode", adBigInt);
            withBlock1.Append("Operation", adBigInt);
            withBlock1.Append("Title", adVarChar, 256);
            withBlock1.Append("Description", adVarChar, 1024);
            withBlock1.Append("Date", adDBDate);
        }
        rt.Open();
    }
    return rt;
}

public static ADODB.Recordset rsShtMtlCutPars(Document ad, long incTop = 0)
{
    // Windows Update History
    Document ActiveDoc;

    ADODB.Recordset rt = rsNewShtMtlCutPars;
    dynamic ls = new[]
    {
        "Item", "Description", "Thickness", "Perimeter"
    };
    {
        var withBlock = dcAiDocComponents(ad, null,
            incTop);
        foreach (var ky in withBlock.Keys)
        {
            Document pt = aiDocument(withBlock.get_Item(ky));
            {
                var withBlock1 = pt.PropertySets.get_Item(gnDesign);
                rt.AddNew(ls, new[]
                {
                    withBlock1.get_Item(pnPartNum).Value, withBlock1.get_Item(pnDesc).Value,
                    aiPropVal(aiPropShtMetalThickness(aiDocPart(pt)), -1), fpPerimeterInch(pt)
                });
            }
        }

        rt.Filter = "";
    }

    return rt;
}
// send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , "|")
// send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , vbTab)

public static ADODB.Recordset rsNewShtMtlCutPars()
{
    var rt = new ADODB.Recordset();
    {
        {
            var withBlock1 = rt.Fields;
            // .Append "", adBigInt
            // .Append "", adVarChar, 1024
            // .Append "Date", adDBDate
            // 
            withBlock1.Append("Item", adVarChar, 32, adFldKeyColumn);
            withBlock1.Append("Description", adVarChar, 128);
            withBlock1.Append("Thickness", adDouble);
            withBlock1.Append("Perimeter", adDouble);
        }
        rt.Open();
    }

    return rt;
}

}