class dvlDict0
{
    public string aiDocPartNum(Inventor.Document AiDoc, string ifNone = "")
    {
        if (AiDoc == null)
            aiDocPartNum = ifNone;
        else
            aiDocPartNum = AiDoc.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
    }

    public Scripting.Dictionary dc0g1f0(Inventor.Document AiDoc, string prName = "", Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        string pn;
        string ky;

        if (dc == null)
            dc0g1f0 = dc0g1f0(AiDoc, prName, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = AiDoc;
                pn = withBlock.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
                ky = prName + Constants.vbTab + pn;
                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(ky))
                    {
                        if (withBlock1.Item(ky) == AiDoc)
                        {
                        }
                        else
                        {
                        }
                    }
                    else
                        withBlock1.Add(ky, AiDoc);
                }

                if (withBlock.DocumentType == kAssemblyDocumentObject)
                    rt = dc0g1f1(AiDoc, rt);
                else if (withBlock.DocumentType != kPartDocumentObject)
                    System.Diagnostics.Debugger.Break();
                else
                {
                }
            }

            dc0g1f0 = rt;
        }
    }

    public Scripting.Dictionary dc0g1f1(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence aiOcc;
        string pn;

        if (dc == null)
            dc0g1f1 = dc0g1f1(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = AiDoc;
                pn = withBlock.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    foreach (var aiOcc in withBlock1.Occurrences)
                        rt = dc0g1f0(aiOcc.Definition.Document, pn, rt);
                }
            }

            dc0g1f1 = rt;
        }
    }

    public Scripting.Dictionary dc0g2f0(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary wk;
        Scripting.Dictionary rt;
        Variant ky;

        if (AiDoc == null)
            dc0g2f0 = dc0g2f0(ThisApplication.ActiveDocument);
        else
        {
            wk = dcAiDocsByPtNum(dcAiDocComponents(AiDoc)); // dcAiDocPartNumbers
            rt = new Scripting.Dictionary();
            {
                var withBlock = wk;
                foreach (var ky in withBlock.Keys)
                    rt = dc0g2f2(aiDocument(obOf(withBlock.Item(ky))), rt);
            }
            dc0g2f0 = rt;
        }
    }

    public Scripting.Dictionary dc0g2f1(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        // Dim rt As Scripting.Dictionary
        Inventor.ComponentOccurrence aiOcc;
        string prName;
        string ptName;
        string ky;
        long ct;

        if (dc == null)
            dc0g2f1 = dc0g2f1(AiDoc, new Scripting.Dictionary());
        else
        {
            prName = aiDocPartNum(AiDoc);
            {
                var withBlock = AiDoc.ComponentDefinition;
                foreach (var aiOcc in withBlock.Occurrences)
                {
                    ptName = aiDocPartNum(aiOcc.Definition.Document);
                    ky = prName + Constants.vbTab + ptName;

                    {
                        var withBlock1 = dc;
                        if (withBlock1.Exists(ky))
                        {
                            ct = withBlock1.Item(ky);
                            withBlock1.Item(ky) = 1 + ct;
                        }
                        else
                            withBlock1.Add(ky, 1);
                    }
                }
            }
            dc0g2f1 = dc;
        }
    }

    public Scripting.Dictionary dc0g2f2(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (AiDoc.DocumentType == kAssemblyDocumentObject)
            dc0g2f2 = dc0g2f1(AiDoc, dc);
        else
            dc0g2f2 = dc;
    }

    public Scripting.Dictionary dc0g3f0(Scripting.Dictionary dc)
    {
        /// (just so we don't forget what this is for)
        /// This function accepts a Dictionary
        /// of Inventor Documents, and generates
        /// a new Dictionary of Dictionaries,
        /// keyed on Genius Family names, and
        /// containing all Documents in its Family,
        /// themselves keyed on Part Number.
        /// 
        /// Function dc0g3f1 makes use of this below.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary fm;
        Variant ky;
        Inventor.Document ad;
        string nm;
        string pn;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocument(obOf(withBlock.Item(ky)));
                if (ad == null)
                {
                }
                else
                {
                    {
                        var withBlock1 = ad.PropertySets.Item(gnDesign);
                        nm = withBlock1.Item(pnFamily).Value;
                        pn = withBlock1.Item(pnPartNum).Value;
                    }

                    {
                        var withBlock1 = rt;
                        if (withBlock1.Exists(nm))
                            fm = withBlock1.Item(nm);
                        else
                        {
                            fm = new Scripting.Dictionary();
                            withBlock1.Add(nm, fm);
                        }
                    }

                    {
                        var withBlock1 = fm;
                        if (withBlock1.Exists(pn))
                            System.Diagnostics.Debugger.Break();
                        else
                            withBlock1.Add(pn, ad);
                    }
                }
            }
        }

        dc0g3f0 = rt;
    }

    public Scripting.Dictionary dc0g3f1()
    {
        /// This function calls dc0g3f0 above
        /// against a Dictionary of Inventor Documents
        /// generated from the components of the active
        /// Inventor Document. It then dumps a list of
        /// the Genius Family names encountered, and if
        /// any were blank, the list of part numbers
        /// in the blank Family group is also revealed.
        /// 
        {
            var withBlock = dc0g3f0(dcAssyDocComponents(aiDocAssy(aiDocActive())));
            Debug.Print(txDumpLs(withBlock.Keys));
            System.Diagnostics.Debugger.Break();
            if (withBlock.Exists(""))
            {
                Debug.Print("NO FAMILY");
                {
                    var withBlock1 = dcOb(withBlock.Item(""));
                    Debug.Print(txDumpLs(withBlock1.Keys));
                }
                System.Diagnostics.Debugger.Break();
            }
            else
                System.Diagnostics.Debugger.Break();
        }
    }

    public Scripting.Dictionary dc0g4f0(Inventor.AssemblyDocument AiDoc)
    {
        Scripting.Dictionary rt;
        Variant ky;
        Variant ki;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAssyComponentsImmediate(AiDoc);
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcAssyComponentsImmediate(aiDocAssy(withBlock.Item(ky)));
                    foreach (var ki in withBlock1.Keys)
                    {
                        if (!rt.Exists(ki))
                            rt.Add(ki, withBlock1.Item(ki));
                    }
                }
            }
        }
        dc0g4f0 = rt;
    }
    // Debug.Print txDumpLs(dcAssyComponentsImmediate(ThisApplication.ActiveDocument).Keys)
    // Debug.Print txDumpLs(dc0g4f0(ThisApplication.ActiveDocument).Keys)
    // Debug.Print txDumpLs(dcAiDocPartNumbers(dc0g4f0(ThisApplication.ActiveDocument)).Keys)

    public Scripting.Dictionary dc0g4f1(Inventor.AssemblyDocument adIn, Inventor.AssemblyDocument adOut)
    {
        Variant ky;
        Inventor.ComponentOccurrences tg;
        Inventor.ComponentOccurrence oc;
        Inventor.Matrix ps;

        ps = ThisApplication.TransientGeometry.CreateMatrix();

        tg = adOut.ComponentDefinition.Occurrences;
        {
            var withBlock = dc0g4f0(adIn);
            foreach (var ky in withBlock.Keys)
                oc = tg.Add(ky, ps);
        }
    }

    public Scripting.Dictionary dcBoxDims(Inventor.Box bx)
    {
        Scripting.Dictionary rt;
        Inventor.Point mx;
        Inventor.Point mn;

        rt = new Scripting.Dictionary();

        {
            var withBlock = bx;
            mx = withBlock.MaxPoint;
            mn = withBlock.MinPoint;
        }

        {
            var withBlock = rt;
            withBlock.Add("X", (mx.X - mn.X)); withBlock.Add("Y", (mx.Y - mn.Y)); withBlock.Add("Z", (mx.Z - mn.Z));
        }

        dcBoxDims = rt;
    }

    public Scripting.Dictionary dcBoxDimsCm2in(Inventor.Box bx)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcBoxDims(bx);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, System.Convert.ToDouble(withBlock.Item(ky)) / cvLenIn2cm);
        }

        dcBoxDimsCm2in = rt;
    }

    public Scripting.Dictionary dcAiPropsInSet(Inventor.PropertySet ps)
    {
        Scripting.Dictionary rt;
        Inventor.Property pr;

        rt = new Scripting.Dictionary();
        foreach (var pr in ps)
        {
            if (rt.Exists(pr.Name))
                System.Diagnostics.Debugger.Break();
            else
                rt.Add(pr.Name, pr);
        }
        dcAiPropsInSet = rt;
    }
    // Debug.Print Join(Filter(dcAiPropsInSet(ThisApplication.ActiveDocument.PropertySets(gnDesign)).Keys, "web", , vbTextCompare), vbNewLine)

    public Scripting.Dictionary dcAiDocParVals(Inventor.Document AiDoc)
    {
        dcAiDocParVals = dcAiParValues(dcAiDocParams(AiDoc));
    }

    public Scripting.Dictionary dcAiParValues(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;
        Inventor.Parameter pr;

        rt = new Scripting.Dictionary();

        if (dc == null)
        {
        }
        else
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pr = obAiParam(obOf(withBlock.Item(ky)));
                if (pr == null)
                {
                }
                else
                    rt.Add(ky, Array(pr.Value, pr.Units));
            }
        }

        dcAiParValues = rt;
    }

    public Scripting.Dictionary dcAiDocParams(Inventor.Document AiDoc)
    {
        dcAiDocParams = dcCompDefParams(compDefOf(AiDoc));
    }
    // Debug.Print Join(Filter(dcAiDocParams(ThisApplication.ActiveDocument.PropertySets(gnDesign)).Keys, "web", , vbTextCompare), vbNewLine)

    public Scripting.Dictionary dcCompDefParams(Inventor.ComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (CpDef == null)
            dcCompDefParams = new Scripting.Dictionary();
        else if (CpDef is Inventor.AssemblyComponentDefinition)
            dcCompDefParams = dcCompDefParamsAssy(CpDef);
        else if (CpDef is Inventor.PartComponentDefinition)
            dcCompDefParams = dcCompDefParamsPart(CpDef);
        else
            dcCompDefParams = dcCompDefParams(null/* TODO Change to default(_) if this is not a reference type */);
    }

    public Scripting.Dictionary dcCompDefParamsPart(Inventor.PartComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.Parameters pr;

        if (CpDef == null)
            pr = null/* TODO Change to default(_) if this is not a reference type */;
        else
            pr = CpDef.Parameters;

        dcCompDefParamsPart = dcOfAiParameters(pr, dc);
    }

    public Scripting.Dictionary dcCompDefParamsAssy(Inventor.AssemblyComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.Parameters pr;

        if (CpDef == null)
            pr = null/* TODO Change to default(_) if this is not a reference type */;
        else
            pr = CpDef.Parameters;

        dcCompDefParamsAssy = dcOfAiParameters(pr, dc);
    }

    public Scripting.Dictionary dcOfAiParameters(Inventor.Parameters AiPars, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Inventor.Parameter pr;

        if (dc == null)
            rt = dcOfAiParameters(AiPars, new Scripting.Dictionary());
        else
        {
            rt = dc;

            if (AiPars == null)
            {
            }
            else
                foreach (var pr in AiPars)
                    rt.Add(pr.Name, pr);
        }

        dcOfAiParameters = rt;
    }

    public Scripting.Dictionary dcOfPropsInAiDoc(Inventor.Document AiDoc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Inventor.PropertySet ps;
        Variant ky;

        rt = new Scripting.Dictionary();

        if (AiDoc == null)
        {
        }
        else
        {
            var withBlock = AiDoc;
            foreach (var ps in withBlock.PropertySets)
            {
                wk = dcAiPropsInSet(ps);

                {
                    var withBlock1 = dcKeysMissing(wk, rt);
                    foreach (var ky in withBlock1.Keys)
                    {
                        rt.Add(ky, withBlock1.Item(ky));
                        wk.Remove(ky);
                    }
                }

                {
                    var withBlock1 = wk // dcKeysInCommon(wk, rt)
       ;
                    if (withBlock1.Count > 0)
                    {
                        Debug.Print("=== DUPLICATE PROPERTY NAMES ===");
                        Debug.Print("  Item " + aiProperty(rt.Item(pnPartNum)).Value + " (" + AiDoc.FullDocumentName + ")");
                        Debug.Print(dumpLsKeyVal(dcPropVals(wk), ": "));
                        Debug.Print("--- previously found");
                        Debug.Print(dumpLsKeyVal(dcPropVals(dcKeysInCommon(wk, rt, 2)), ": "));
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                }
            }
        }

        dcOfPropsInAiDoc = rt;
    }

    public Scripting.Dictionary dcAiPropValsFromDc(Scripting.Dictionary dc, long Flags = 0)
    {
        Scripting.Dictionary rt;
        Inventor.Property pr;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
            {
                pr = aiProperty(obOf(withBlock.Item(ky)));
                if (pr == null)
                {
                    if (Flags & 1)
                        // Keep non-Property Items
                        rt.Add(ky, withBlock.Item(ky));
                }
                else
                    rt.Add(ky, aiPropVal(pr, Empty));
            }
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        dcAiPropValsFromDc = rt;
    }

    public Scripting.Dictionary dcForAiDocIType(Scripting.Dictionary dc, Inventor.Document AiDoc)
    {
        Scripting.Dictionary wk;
        string ky;

        if (AiDoc is Inventor.PartDocument)
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
        }
        else if (AiDoc is Inventor.AssemblyDocument)
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
        }
        else
            ky = "";

        {
            var withBlock = dc;
            if (!withBlock.Exists(ky))
                withBlock.Add(ky, new Scripting.Dictionary());
            dcForAiDocIType = withBlock.Item(ky);
        }
    }

    public Scripting.Dictionary dcAiDocsByIType(Scripting.Dictionary dc, long Flags = 0)
    {
        Scripting.Dictionary rt;
        // Dim pr As Inventor.Property
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
                // pr = aiProperty(obOf(.Item(ky)))
                // If pr Is Nothing Then
                // If Flags And 1 Then
                // Keep non-Property Items
                rt.Add(ky, withBlock.Item(ky));
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        dcAiDocsByIType = rt;
    }

    public Variant nvmTest01()
    {
        Inventor.ApplicationAddIn ad;
        object il; // Inventor.ApplicationAddIn '
        Inventor.NameValueMap mp;
        Inventor.Document md;

        ad = ThisApplication.ApplicationAddIns.ItemById(guidILogicAdIn);
        if (!ad.Activated)
            ad.Activate();
        il = ad.Automation;
        md = ThisApplication.Documents.ItemByName(@"C:\Doyle_Vault\Designs\Misc\andrewT\dvl\iLogVltSrch_2022-0622_01.ipt");
        mp = dc2aiNameValMap(nuDcPopulator().Setting("PartNumber", "60-").Dictionary);  // IN 60- 04-

        il.RunRuleWithArguments(md, "vlt02", mp);     // il.RunRule md, "tst01" ', mp

        Debug.Print(mp.Value("OUT"));
        Debug.Print(mp.Count);
    }

    /// 

    /// 
    public string dvlDict0()
    {
        dvlDict0 = "dvlDict0";
    }
}