class dvlBomView
{
    public Inventor.BOMView bomViewStruct(Inventor.AssemblyDocument pd)
    {
        /// bomViewStruct --  Structured BOM View
        /// for supplied Assembly, if available
        /// 
        Inventor.BOMView bv;
        // Dim br As Inventor.BOMRow

        if (pd == null)
            bv = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            var withBlock = pd // aiDocAssy(aiDocActive())
;
            Information.Err.Clear();
            bv = withBlock.ComponentDefinition.BOM.BOMViews.Item("Structured");

            if (Information.Err.Number == 0)
            {
            }
            else
                bv = null/* TODO Change to default(_) if this is not a reference type */;
        }
        bomViewStruct = bv;
    }

    public string[] dBVg1f1(string itmPath)
    {
        string[] rt = new string[2];
        long bk;

        bk = InStrRev(itmPath, ".");

        if (bk > 0)
        {
            rt[0] = Left(itmPath, bk - 1);
            rt[1] = Mid(itmPath, bk + 1);
        }
        else
        {
            rt[0] = "";
            rt[1] = itmPath;
        }

        dBVg1f1 = rt;
    }

    public Scripting.Dictionary bomLnumBkDn(Scripting.Dictionary dc)
    {
        string[] ls;

        {
            var withBlock = dc;
            if (withBlock.Exists("path"))
            {
                ls = dBVg1f1(dc.Item("path"));
                withBlock.Item("base") = ls[0];
                withBlock.Item("seq") = ls[1];
            }
        }
        bomLnumBkDn = dc;
    }

    public Scripting.Dictionary bomLineInfo(Inventor.BOMRow brw)
    {
        Scripting.Dictionary rt;

        rt = new Scripting.Dictionary();
        {
            var withBlock = brw;
            rt.Add("bomStruct", withBlock.BOMStructure);
            rt.Add("path", withBlock.ItemNumber);
            // rt.Add "seq", .ItemNumber

            rt.Add("qty", withBlock.ItemQuantity);
            rt.Add("qtTotal", withBlock.TotalQuantity);
            rt.Add("qtUnit", "EA");
            rt.Add("mrg", withBlock.Merged);
            rt.Add("pro", withBlock.Promoted);
            rt.Add("rol", withBlock.RolledUp);
        }
        bomLineInfo = bomLnumBkDn(rt);
    }

    public Scripting.Dictionary dBVg1f2(Inventor.Document AiDoc, Scripting.Dictionary wk, Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary ck;
        string bs;
        string pn;
        string rm;
        string k0;
        Variant k1;

        {
            var withBlock = wk;
            bs = withBlock.Item("path");
            pn = withBlock.Item("ptNum");
        }

        {
            var withBlock = dcOfPropsInAiDoc(AiDoc);
            if (withBlock.Exists("RM"))
            {
                rm = aiProperty(withBlock.Item("RM")).Value;
                k0 = pn + "|" + rm;

                rt = new Scripting.Dictionary();
                rt.Add("bomStruct", kPurchasedBOMStructure);
                rt.Add("path", bs + ".1");
                if (withBlock.Exists("RMQTY"))
                    rt.Add("qty", aiProperty(withBlock.Item("RMQTY")).Value);
                else
                    rt.Add("qty", -1);
                rt.Add("qtTotal", rt.Item("qty"));
                if (withBlock.Exists("RMUNIT"))
                    rt.Add("qtUnit", aiProperty(withBlock.Item("RMUNIT")).Value);
                else
                    rt.Add("qtUnit", "EA");

                rt.Add("mrg", false);
                rt.Add("pro", false);
                rt.Add("rol", false);
                rt.Add("base", bs);
                rt.Add("seq", "1");
                rt.Add("ptNum", rm);
                // rt.Add "aiDoc", ""

                if (dc.Exists(k0))
                {
                    ck = dcOb(dc.Item(k0));
                    // send2clipBd ConvertToJson(dcWBQbyCmpResult(dcCmpTextOf2dc(ck, rt)), vbTab)
                    {
                        var withBlock1 = dcWBQbyCmpResult(dcCmpTextOf2dc(ck, rt));
                        if (withBlock1.Exists("!="))
                        {
                            {
                                var withBlock2 = dcOb(withBlock1.Item("!="));
                                foreach (var k1 in withBlock2.Keys)
                                    ck.Item(k1) = ck.Item(k1) + Constants.vbTab + rt.Item(k1);
                            }
                        }
                    }
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                }
                else
                    dc.Add(pn + "|" + rm, rt);
            }
            else
            {
            }
        }

        // Stop
        dBVg1f2 = dc;
    }

    public Scripting.Dictionary bomItemInfo(Inventor.BOMRow rw, Scripting.Dictionary dc = )
    {
        Scripting.Dictionary rt;
        // Dim wk As Scripting.Dictionary
        Inventor.ComponentDefinition df;
        Inventor.Document pt;
        string pn;
        string ck;
        string fn;
        long ct;

        if (dc == null)
            rt = bomLineInfo(rw);
        else
            rt = dc;

        {
            var withBlock = rw;
            {
                var withBlock1 = withBlock.ComponentDefinitions;
                ct = withBlock1.Count;
                if (ct > 0)
                {
                    {
                        var withBlock2 = withBlock1.Item(1);
                        pt = aiDocument(withBlock2.Document);
                        pn = aiDocPartNum(pt);
                        fn = pt.FullDocumentName;
                    }
                }
                else
                {
                    System.Diagnostics.Debugger.Break();
                    pn = "";
                }

                rt.Add("ptNum", pn);
            }

            if (ct > 1)
            {
                System.Diagnostics.Debugger.Break();
                fn = "";
                foreach (var df in withBlock.ComponentDefinitions)
                {
                    pt = aiDocument(withBlock.Document);
                    ck = aiDocPartNum(pt);
                    if (ck == pn)
                        fn = fn + Constants.vbNewLine + pt.FullDocumentName;
                    else
                        System.Diagnostics.Debugger.Break();
                }
            }

            rt.Add("aiDoc", fn);
        }

        bomItemInfo = rt;
    }

    public Scripting.Dictionary dBVg7f4(Inventor.AssemblyDocument aiAssy, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// not sure what doing with this one
        /// further developemt on hold
        /// 
        // Dim pd As Inventor.AssemblyDocument
        Scripting.Dictionary rt;

        Inventor.BOMView bv;
        Inventor.BOMRow br;

        if (dc == null)
            rt = dBVg7f4(aiAssy, new Scripting.Dictionary());
        else if (aiAssy == null)
            rt = dc;
        else
        {
            var withBlock = aiAssy;
            bv = withBlock.ComponentDefinition.BOM.BOMViews.Item("Structured");
            {
                var withBlock1 = bv;
                System.Diagnostics.Debugger.Break();
            }
        }
        dBVg7f4 = rt;
    }

    public Scripting.Dictionary bomInfoBkDn(Inventor.BOMRowsEnumerator rwSet, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, string fn = "")
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Inventor.BOMRow rw;
        string pn;
        string ck;

        if (dc == null)
            rt = bomInfoBkDn(rwSet, new Scripting.Dictionary(), fn);
        else
        {
            rt = dc;
            if (rwSet == null)
            {
            }
            else
                foreach (var rw in rwSet)
                {
                    // Stop
                    DoEvents();
                    wk = bomItemInfo(rw);
                    wk.Add("ptOf", fn);
                    pn = System.Convert.ToHexString(wk.Item("ptNum"));
                    ck = fn + "|" + pn;
                    {
                        var withBlock = rt;
                        if (withBlock.Exists(ck))
                        {
                            // Stop
                            // debug.Print ConvertToJson(dcCmpTextOf2dc(wk,dcOb(.Item(ck))),vbTab)
                            // debug.Print ConvertToJson(
                            {
                                var withBlock1 = dcWBQbyCmpResult(dcCmpTextOf2dc(wk, dcOb(withBlock.Item(ck)))) // ,vbTab)
       ;
                                {
                                    var withBlock2 = dcOb(withBlock1.Item("!="));
                                    withBlock2.Remove("path");
                                    withBlock2.Remove("base");
                                    if (withBlock2.Count > 0)
                                    {
                                        Debug.Print("MISMATCH: ", ", ");
                                        System.Diagnostics.Debugger.Break();
                                    }
                                }
                            }
                        }
                        else
                            withBlock.Add(ck, wk);
                    }

                    {
                        var withBlock = rw;
                        if (withBlock.ChildRows == null)
                            dc = dBVg1f2(ThisApplication.Documents.ItemByName(wk.Item("aiDoc")), wk, dc);
                        else
                        {
                            DoEvents();
                            dc = bomInfoBkDn(withBlock.ChildRows, dc, pn);
                        }
                    }
                    DoEvents();
                }
        }
        bomInfoBkDn = rt;
    }

    public Scripting.Dictionary dcOfBomsFromAiStructured(Inventor.BOMRowsEnumerator rwSet, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, string fn = "")
    {
        /// dcOfBomsFromAiStructured --
        /// generate Dictionary of BOMs:
        /// one for each distinct Assembly in
        /// supplied Inventor BOM (structured)
        /// 
        /// returned as Dictionary of Assembly
        /// sub Dictionaries, each keyed to its
        /// Part Number and containing a set of
        /// Item sub Dictionaries, again keyed
        /// to Item P/N. Each Item sub Dictionary
        /// represents a BOM line item
        /// 
        Scripting.Dictionary rt;

        dcOfBomsFromAiStructured = dBV0g0f4(dBV0g0f3(dBV0g0f1(rwSet, dc, fn)));
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
    }

    public Scripting.Dictionary dBV0g0f1(Inventor.BOMRowsEnumerator rwSet, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, string fn = "")
    {
        /// dBV0g0f1 -- retrieve BOM data
        /// from a BOMRowsEnumerator
        /// and its child row enumerators
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary pd;
        Scripting.Dictionary it;
        Scripting.Dictionary dt;
        // Dim fd As Scripting.Dictionary
        /// 
        // Dim kyIt As Variant
        /// 
        Inventor.BOMRow rw;
        string pn;
        string th;

        if (dc == null)
            rt = dBV0g0f1(rwSet, new Scripting.Dictionary(), fn);
        else
        {
            rt = dc;
            if (rwSet == null)
                // Stop
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            else
            {
                {
                    var withBlock = rt;
                    if (!withBlock.Exists(fn))
                        withBlock.Add(fn, new Scripting.Dictionary());
                    pd = withBlock.Item(fn);
                }

                foreach (var rw in rwSet)
                {
                    DoEvents();

                    dt = bomItemInfo(rw);
                    {
                        var withBlock = dt;
                        pn = System.Convert.ToHexString(withBlock.Item("ptNum"));
                        th = System.Convert.ToHexString(withBlock.Item("path"));
                    }

                    {
                        var withBlock = pd;
                        if (!withBlock.Exists(pn))
                            withBlock.Add(pn, new Scripting.Dictionary());
                        it = withBlock.Item(pn);
                    }

                    {
                        var withBlock = it;
                        if (withBlock.Exists(th))
                            System.Diagnostics.Debugger.Break();
                        else
                            withBlock.Add(th, dt);
                    }

                    rt = dBV0g0f1(rw.ChildRows, rt, pn);
                    DoEvents();
                }
            }
        }

        // If Len(fn) = 0 Then Stop
        dBV0g0f1 = rt;
    }

    public Scripting.Dictionary dBV0g0f2(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary dcField;
        Variant k0path;
        Variant k1Field;
        Variant itValue;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var k0path in withBlock.Keys)
            {
                {
                    var withBlock1 = dcOb(withBlock.Item(k0path));
                    foreach (var k1Field in withBlock1.Keys)
                    {
                        itValue = withBlock1.Item(k1Field);

                        {
                            var withBlock2 = rt;
                            if (!withBlock2.Exists(k1Field))
                                withBlock2.Add(k1Field, new Scripting.Dictionary());
                            dcField = withBlock2.Item(k1Field);
                        }

                        {
                            var withBlock2 = dcField;
                            if (!withBlock2.Exists(itValue))
                                withBlock2.Add(itValue, new Scripting.Dictionary());

                            {
                                var withBlock3 = dcOb(withBlock2.Item(itValue));
                                if (withBlock3.Exists(k0path))
                                {
                                }
                                else
                                    withBlock3.Add(k0path, 1);
                            }
                        }
                    }
                }
            }
        }

        dBV0g0f2 = rt;
    }

    public Scripting.Dictionary dBV0g0f3(Scripting.Dictionary dc)
    {
        /// dBV0g0f3 -- summarize BOM line item fields
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcProd;
        Variant k0Prod;
        Variant k1Item;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var k0Prod in withBlock.Keys)
            {
                {
                    var withBlock1 = rt;
                    withBlock1.Add(k0Prod, new Scripting.Dictionary());
                    dcProd = withBlock1.Item(k0Prod);
                }

                {
                    var withBlock1 = dcOb(withBlock.Item(k0Prod));
                    foreach (var k1Item in withBlock1.Keys)
                        dcProd.Add(k1Item, dBV0g0f2(withBlock1.Item(k1Item)));
                }
            }
        }

        dBV0g0f3 = rt;
    }

    public Scripting.Dictionary dBV0g0f4(Scripting.Dictionary dc)
    {
        /// dBV0g0f4 -- reduce results of dBV0g0f3
        /// to single values per field
        /// for each Item under each Product
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcProd;
        Scripting.Dictionary dcItem;
        Variant k0Prod;
        Variant k1Item;
        Variant k2Feld;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var k0Prod in withBlock.Keys) // Products
            {
                {
                    var withBlock1 = rt;
                    withBlock1.Add(k0Prod, new Scripting.Dictionary());
                    dcProd = withBlock1.Item(k0Prod);
                }

                {
                    var withBlock1 = dcOb(withBlock.Item(k0Prod)) // Items
           ;
                    foreach (var k1Item in withBlock1.Keys)
                    {
                        {
                            var withBlock2 = dcProd;
                            withBlock2.Add(k1Item, new Scripting.Dictionary());
                            dcItem = withBlock2.Item(k1Item);
                        }

                        {
                            var withBlock2 = dcOb(withBlock1.Item(k1Item)) // Fields
               ;
                            withBlock2.Remove("path");
                            withBlock2.Remove("base");
                            foreach (var k2Feld in withBlock2.Keys)
                            {
                                {
                                    var withBlock3 = dcOb(withBlock2.Item(k2Feld)) // Value(s)
              ;
                                    if (withBlock3.Count > 1)
                                        System.Diagnostics.Debugger.Break();
                                    else
                                        dcItem.Add(k2Feld, withBlock3.Keys(0));
                                }
                            }
                        }
                    }
                }
            }
        }

        dBV0g0f4 = rt;
    }

    public Scripting.Dictionary dBV0g0f5(Scripting.Dictionary dc, string dlm = "|")
    {
        /// dBV0g0f5 -- reduce results of dBV0g0f3
        /// to single values per field
        /// for each Item under each Product
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcItem;
        Variant k0Prod;
        Variant k1Item;
        Variant k2Feld;
        string rw;
        string co;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var k0Prod in withBlock.Keys) // Products
            {
                if (Len(k0Prod) > 0)
                {
                    {
                        var withBlock1 = dcOb(withBlock.Item(k0Prod)) // Items
;
                        foreach (var k1Item in withBlock1.Keys)
                        {
                            dcItem = withBlock1.Item(k1Item);

                            {
                                var withBlock2 = dcItem;
                                rw = k0Prod + dlm + k1Item;
                                foreach (var k2Feld in Array("seq", "qty", "qtUnit"))
                                {
                                    if (withBlock2.Exists(k2Feld))
                                        co = System.Convert.ToHexString(withBlock2.Item(k2Feld));
                                    else
                                        co = "";

                                    rw = rw + dlm + co;
                                }
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }

                            rt.Add(rw, dcItem);
                        }
                    }
                }
                else
                {
                }
            }
        }

        dBV0g0f5 = rt;
    }

    public string csvOfBomsFromDc(Scripting.Dictionary dc, string dlm = "|")
    {
        // Product|Item|ItemOrder|QuantityInConversionUnit|ConversionUnit
        // NOTE[2021.08.20]: want to change 'Item'
        // to 'ItemCode' for compatibility with
        // current Genius BOM import format.
        // Will hold off for now.
        csvOfBomsFromDc = Join(Array("Product", "Item", "ItemOrder", "QuantityInConversionUnit", "ConversionUnit"), dlm) + Constants.vbNewLine + txDumpLs(dBV0g0f5(dc, dlm).Keys);
    }

    public string csvOfBomsFromAiStructured(Inventor.Document AiDoc, string dlm = "|")
    {
        csvOfBomsFromAiStructured = csvOfBomsFromDc(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows), dlm);
    }

    /// 

    /// 
    private string dvlBomView()
    {
        dvlBomView = "dvlBomView";
    }
}