using Microsoft.VisualBasic;
using static Doyle_Addin.Genius.Classes.lib0;

namespace Doyle_Addin.Genius.Classes;

 public class dvlDict0
{
    public static string aiDocPartNum(Document AiDoc, string ifNone = "")
    {
        return AiDoc == null ? ifNone : AiDoc.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;
    }

    public static Dictionary dc0g1f0(Document AiDoc, string prName = "", Dictionary dc = null)
    {
        while (true)
        {
            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            var rt = dc;

            {
                string pn = AiDoc.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;
                var ky = prName + Constants.vbTab + pn;
                {
                    if (rt.Exists(ky))
                    {
                        if (rt.get_Item(ky) == AiDoc)
                        {
                        }
                    }
                    else
                        rt.Add(ky, AiDoc);
                }

                if (AiDoc.DocumentType == kAssemblyDocumentObject)
                    rt = dc0g1f1(AiDoc, rt);
                else if (AiDoc.DocumentType != kPartDocumentObject) Debugger.Break();
            }

            return rt;
        }
    }

    public static Dictionary dc0g1f1(AssemblyDocument AiDoc, Dictionary dc = null)
    {
        while (true)
        {
            if (dc == null)
            {
                dc = new Dictionary();
            }
            else
            {
                {
                    string pn = AiDoc.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;
                    {
                        var withBlock1 = AiDoc.ComponentDefinition;
                    }
                }

                return Enumerable.Cast<ComponentOccurrence>(withBlock1.Occurrences).Aggregate(dc,
                    (current, aiOcc) => dc0g1f0(aiOcc.Definition.Document, pn, current));
            }
        }
    }

    public static Dictionary dc0g2f0(Document AiDoc = null)
    {
        while (true)
        {
            if (AiDoc == null)
            {
                AiDoc = ThisApplication.ActiveDocument;
                continue;
            }

            var wk = dcAiDocsByPtNum(dcAiDocComponents(AiDoc));
            var rt = new Dictionary();
            {
                foreach (var ky in wk.Keys) rt = dc0g2f2(aiDocument(obOf(wk.get_Item(ky))), rt);
            }
            return rt;
        }
    }

    public static Dictionary dc0g2f1(AssemblyDocument AiDoc, Dictionary dc = null)
    {
        while (true)
        {
            // Dim rt As Scripting.Dictionary

            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            var prName = aiDocPartNum(AiDoc);
            {
                var withBlock = AiDoc.ComponentDefinition;
                foreach (ComponentOccurrence aiOcc in withBlock.Occurrences)
                {
                    var ptName = aiDocPartNum(aiOcc.Definition.Document);
                    var ky = prName + Constants.vbTab + ptName;

                    {
                        if (dc.Exists(ky))
                        {
                            long ct = dc.get_Item(ky);
                            dc.get_Item(ky) = 1 + ct;
                        }
                        else
                            dc.Add(ky, 1);
                    }
                }
            }
            return dc;
        }
    }

    public static Dictionary dc0g2f2(Document AiDoc, Dictionary dc = null)
    {
        return AiDoc.DocumentType == kAssemblyDocumentObject ? dc0g2f1(AiDoc, dc) : dc;
    }

    public static Dictionary dc0g3f0(Dictionary dc)
    {
        // (just so we don't forget what this is for)
        // This function accepts a Dictionary
        // of Inventor Documents, and generates
        // a new Dictionary of Dictionaries,
        // keyed on Genius Family names, and
        // containing all Documents in its Family,
        // themselves keyed on Part Number.
        // 
        // Function dc0g3f1 makes use of this below.
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document ad = aiDocument(obOf(dc.get_Item(ky)));
                if (ad == null)
                {
                }
                else
                {
                    string nm;
                    string pn;
                    {
                        var withBlock1 = ad.PropertySets.get_Item(gnDesign);
                        nm = withBlock1.get_Item(pnFamily).Value;
                        pn = withBlock1.get_Item(pnPartNum).Value;
                    }

                    Dictionary fm;
                    {
                        if (rt.Exists(nm))
                            fm = rt.get_Item(nm);
                        else
                        {
                            fm = new Dictionary();
                            rt.Add(nm, fm);
                        }
                    }

                    {
                        if (fm.Exists(pn))
                            Debugger.Break();
                        else
                            fm.Add(pn, ad);
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary dc0g3f1()
    {
        // This function calls dc0g3f0 above
        // against a Dictionary of Inventor Documents
        // generated from the components of the active
        // Inventor Document. It then dumps a list of
        // the Genius Family names encountered, and if
        // any were blank, the list of part numbers
        // in the blank Family group is also revealed.
        // 
        {
            var withBlock = dc0g3f0(dcAssyDocComponents(aiDocAssy(aiDocActive())));
            Debug.Print(txDumpLs(withBlock.Keys));
            Debugger.Break();
            if (withBlock.Exists(""))
            {
                Debug.Print("NO FAMILY");
                {
                    var withBlock1 = dcOb(withBlock.get_Item(""));
                    Debug.Print(txDumpLs(withBlock1.Keys));
                }
                Debugger.Break();
            }
            else
                Debugger.Break();
        }
    }

    public static Dictionary dc0g4f0(AssemblyDocument AiDoc)
    {
        var rt = new Dictionary();
        {
            var withBlock = dcAssyComponentsImmediate(AiDoc);
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcAssyComponentsImmediate(aiDocAssy(withBlock.get_Item(ky)));
                    foreach (var ki in withBlock1.Keys)
                    {
                        if (!rt.Exists(ki))
                            rt.Add(ki, withBlock1.get_Item(ki));
                    }
                }
            }
        }
        return rt;
    }
    // Debug.Print txDumpLs(dcAssyComponentsImmediate(ThisApplication.ActiveDocument).Keys)
    // Debug.Print txDumpLs(dc0g4f0(ThisApplication.ActiveDocument).Keys)
    // Debug.Print txDumpLs(dcAiDocPartNumbers(dc0g4f0(ThisApplication.ActiveDocument)).Keys)

    public static Dictionary dc0g4f1(AssemblyDocument adIn, AssemblyDocument adOut)
    {
        var ps = ThisApplication.TransientGeometry.CreateMatrix();

        var tg = adOut.ComponentDefinition.Occurrences;
        {
            var withBlock = dc0g4f0(adIn);
            foreach (var ky in withBlock.Keys)
            {
                var oc = tg.Add(ky, ps);
            }
        }
    }

    public static Dictionary dcBoxDims(Box bx)
    {
        Inventor.Point mx;
        Inventor.Point mn;

        var rt = new Dictionary();

        {
            mx = bx.MaxPoint;
            mn = bx.MinPoint;
        }

        {
            rt.Add("X", (mx.X - mn.X));
            rt.Add("Y", (mx.Y - mn.Y));
            rt.Add("Z", (mx.Z - mn.Z));
        }

        return rt;
    }

    public static Dictionary dcBoxDimsCm2in(Box bx)
    {
        var rt = new Dictionary();

        {
            var withBlock = dcBoxDims(bx);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, Convert.ToDouble(withBlock.get_Item(ky)) / cvLenIn2cm);
        }

        return rt;
    }

    public static Dictionary dcAiPropsInSet(PropertySet ps)
    {
        var rt = new Dictionary();
        foreach (Property pr in ps)
        {
            if (rt.Exists(pr.Name))
                Debugger.Break();
            else
                rt.Add(pr.Name, pr);
        }

        return rt;
    }
    // Debug.Print Join(Filter(dcAiPropsInSet(ThisApplication.ActiveDocument.PropertySets(gnDesign)).Keys, "web", , vbTextCompare), vbCrLf)

    public static Dictionary dcAiDocParVals(Document AiDoc)
    {
        return dcAiParValues(dcAiDocParams(AiDoc));
    }

    public static Dictionary dcAiParValues(Dictionary dc)
    {
        var rt = new Dictionary();

        if (dc == null)
        {
        }
        else
        {
            foreach (var ky in dc.Keys)
            {
                Parameter pr = obAiParam(obOf(dc.get_Item(ky)));
                if (pr == null)
                {
                }
                else
                    rt.Add(ky, new[] { pr.Value, pr.Units });
            }
        }

        return rt;
    }

    public static Dictionary dcAiDocParams(Document AiDoc)
    {
        return dcCompDefParams(compDefOf(AiDoc));
    }
    // Debug.Print Join(Filter(dcAiDocParams(ThisApplication.ActiveDocument.PropertySets(gnDesign)).Keys, "web", , vbTextCompare), vbCrLf)

    public static Dictionary dcCompDefParams(ComponentDefinition CpDef, Dictionary dc = null)
    {
        while (true)
        {
            switch (CpDef)
            {
                case null:
                    return new Dictionary();
                case AssemblyComponentDefinition:
                    return dcCompDefParamsAssy(CpDef);
                case PartComponentDefinition:
                    return dcCompDefParamsPart(CpDef);
                default:
                    CpDef = null;
                    dc = null;
                    break;
            }
        }
    }

    public static Dictionary dcCompDefParamsPart(PartComponentDefinition CpDef, Dictionary dc = null)
    {
        var pr = CpDef?.Parameters;

        return dcOfAiParameters(pr, dc);
    }

    public  static Dictionary dcCompDefParamsAssy(AssemblyComponentDefinition CpDef, Dictionary dc = null)
    {
        var pr = CpDef?.Parameters;

        return dcOfAiParameters(pr, dc);
    }

    public static Dictionary dcOfAiParameters(Parameters AiPars, Dictionary dc = null)
    {
        Dictionary rt;

        if (dc == null)
            rt = dcOfAiParameters(AiPars, new Dictionary());
        else
        {
            rt = dc;

            if (AiPars == null)
            {
            }
            else
                foreach (Parameter pr in AiPars)
                    rt.Add(pr.Name, pr);
        }

        return rt;
    }

    public static Dictionary dcOfPropsInAiDoc(Document AiDoc)
    {
        var rt = new Dictionary();

        if (AiDoc == null)
        {
        }
        else
        {
            foreach (var wk in from PropertySet ps in AiDoc.PropertySets select dcAiPropsInSet(ps))
            {
                {
                    var withBlock1 = dcKeysMissing(wk, rt);
                    foreach (var ky in withBlock1.Keys)
                    {
                        rt.Add(ky, withBlock1.get_Item(ky));
                        wk.Remove(ky);
                    }
                }

                {
                    if (wk.Count <= 0) continue;
                    Debug.Print("=== DUPLICATE PROPERTY NAMES ===");
                    Debug.Print(" Item " + aiProperty(rt.get_Item(pnPartNum)).Value + " (" + AiDoc.FullDocumentName +
                                ")");
                    Debug.Print(dumpLsKeyVal(dcPropVals(wk), ": "));
                    Debug.Print("--- previously found");
                    Debug.Print(dumpLsKeyVal(dcPropVals(dcKeysInCommon(wk, rt, 2)), ": "));
                    Debug.Print(""); // Breakpoint Landing
                }
            }
        }

        return rt;
    }

    public static Dictionary dcAiPropValsFromDc(Dictionary dc, bool Flags = false)
    {
        var rt = new Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
            {
                Property pr = aiProperty(obOf(withBlock.get_Item(ky)));
                if (pr == null)
                {
                    if (Flags & true)
                        // Keep non-Property Items
                        rt.Add(ky, withBlock.get_Item(ky));
                }
                else
                    rt.Add(ky, aiPropVal(pr, null));
            }
        }

        Debug.Print(""); // Breakpoint Landing
        return rt;
    }

    public static Dictionary dcForAiDocIType(Dictionary dc, Document AiDoc)
    {
        Dictionary wk;
        string ky;

        switch (AiDoc)
        {
            case PartDocument:
            {
                ky = "Part";
                {
                    var withBlock = aiDocPart(AiDoc).ComponentDefinition;
                    if (withBlock.IsContentMember)
                        ky = "c" + ky;
                    if (withBlock.IsiPartFactory)
                        ky = "f" + ky;
                    if (withBlock.IsiPartMember)
                        ky = "i" + ky;
                    if (withBlock.IsModelStateFactory)
                    {
                    }

                    if (withBlock.IsModelStateMember)
                        ky = "s" + ky;
                }
                break;
            }
            case AssemblyDocument:
            {
                ky = "Assy";
                {
                    var withBlock = aiDocAssy(AiDoc).ComponentDefinition;
                    if (withBlock.IsiAssemblyFactory)
                        ky = "f" + ky;
                    if (withBlock.IsiAssemblyMember)
                        ky = "i" + ky;
                    if (withBlock.IsModelStateFactory)
                        ky = "msf" + ky;
                    if (withBlock.IsModelStateMember)
                        ky = "s" + ky;
                }
                break;
            }
            default:
                ky = "";
                break;
        }

        {
            if (!dc.Exists(ky))
                dc.Add(ky, new Dictionary());
            return dc.get_Item(ky);
        }
    }

    public static Dictionary dcAiDocsByIType(Dictionary dc, bool Flags = false)
    {
        // Dim pr As Inventor.Property

        var rt = new Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
                // pr = aiProperty(obOf(.get_Item(ky)))
                // If pr Is Nothing Then
                // If Flags And 1 Then
                // Keep non-Property Items
                rt.Add(ky, withBlock.get_Item(ky));
        }

        Debug.Print(""); // Breakpoint Landing
        return rt;
    }

    public static dynamic nvmTest01()
    {
        ApplicationAddIn ad = ThisApplication.ApplicationAddIns.ItemById(guidILogicAdIn);
        if (!ad.Activated)
            ad.Activate();
        var il = ad.Automation; // Inventor.ApplicationAddIn '
        Document md =
            ThisApplication.Documents.ItemByName(
                @"C:\Doyle_Vault\Designs\Misc\andrewT\dvl\iLogVltSrch_2022-0622_01.ipt");
        NameValueMap mp = dc2aiNameValMap(nuDcPopulator().Setting("PartNumber", "60-").Dictionary); // IN 60- 04-

        il.RunRuleWithArguments(md, "vlt02", mp); // il.RunRule md, "tst01" ', mp

        Debug.Print(mp.Value("OUT"));
        Debug.Print(mp.Count);
    }

    // 

    // 
    public string dvlDict0()
    {
        return "dvlDict0";
    }
}