using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class dvlBomView
{
    public BOMView bomViewStruct(AssemblyDocument pd)
    {
        // bomViewStruct -- Structured BOM View
        // for supplied Assembly, if available
        // 
        BOMView bv;
        // Dim br As Inventor.BOMRow

        if (pd == null)
            bv = null;
        else
        {
            Information.Err().Clear();
            bv = pd.ComponentDefinition.BOM.BOMViews.Add("Structured");

            if (Information.Err().Number == 0)
            {
            }
            else
                bv = null;
        }

        return bv;
    }

    public string[] dBVg1f1(string itmPath)
    {
        var rt = new string[2];

        long bk = InStrRev(itmPath, ".");

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

        return rt;
    }

    public Dictionary bomLnumBkDn(Dictionary dc)
    {
        {
            if (!dc.Exists("path")) return dc;
            string[] ls = dBVg1f1(dc.get_Item("path"));
            dc.get_Item("base") = ls[0];
            dc.get_Item("seq") = ls[1];
        }
        return dc;
    }

    public Dictionary bomLineInfo(BOMRow brw)
    {
        var rt = new Dictionary();
        {
            rt.Add("bomStruct", brw.BOMStructure);
            rt.Add("path", brw.ItemNumber);
            // rt.Add "seq", .ItemNumber

            rt.Add("qty", brw.ItemQuantity);
            rt.Add("qtTotal", brw.TotalQuantity);
            rt.Add("qtUnit", "EA");
            rt.Add("mrg", brw.Merged);
            rt.Add("pro", brw.Promoted);
            rt.Add("rol", brw.RolledUp);
        }
        return bomLnumBkDn(rt);
    }

    public Dictionary dBVg1f2(Document AiDoc, Dictionary wk, Dictionary dc)
    {
        string bs;
        string pn;

        {
            bs = wk.get_Item("path");
            pn = wk.get_Item("ptNum");
        }

        {
            var withBlock = dcOfPropsInAiDoc(AiDoc);
            if (!withBlock.Exists("RM")) return dc;
            string rm = aiProperty(withBlock.get_Item("RM")).Value;
            var k0 = pn + "|" + rm;

            var rt = new Dictionary();
            rt.Add("bomStruct", kPurchasedBOMStructure);
            rt.Add("path", bs + ".1");
            if (withBlock.Exists("RMQTY"))
                rt.Add("qty", aiProperty(withBlock.get_Item("RMQTY")).Value);
            else
                rt.Add("qty", -1);
            rt.Add("qtTotal", rt.get_Item("qty"));
            if (withBlock.Exists("RMUNIT"))
                rt.Add("qtUnit", aiProperty(withBlock.get_Item("RMUNIT")).Value);
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
                Dictionary ck = dcOb(dc.get_Item(k0));
                // send2clipBd ConvertToJson(dcWBQbyCmpResult(dcCmpTextOf2dc(ck, rt)), vbTab)
                {
                    var withBlock1 = dcWBQbyCmpResult(dcCmpTextOf2dc(ck, rt));
                    if (withBlock1.Exists("!="))
                    {
                        {
                            var withBlock2 = dcOb(withBlock1.get_Item("!="));
                            foreach (var k1 in withBlock2.Keys)
                                ck.get_Item(k1) = ck.get_Item(k1) + Constants.vbTab + rt.get_Item(k1);
                        }
                    }
                }
                Debug.Print(""); // Breakpoint Landing
            }
            else
                dc.Add(pn + "|" + rm, rt);
        }

        // Stop
        return dc;
    }

    public Dictionary bomItemInfo(BOMRow rw, Dictionary dc = )
    {
        // Dim wk As Scripting.Dictionary

        var rt = dc ?? bomLineInfo(rw);

        {
            Document pt;
            string pn;
            string fn;
            long ct;
            {
                var withBlock1 = rw.ComponentDefinitions;
                ct = withBlock1.Count;
                if (ct > 0)
                {
                    {
                        var withBlock2 = withBlock1.get_Item(1);
                        pt = aiDocument(withBlock2.Document);
                        pn = aiDocPartNum(pt);
                        fn = pt.FullDocumentName;
                    }
                }
                else
                {
                    Debugger.Break();
                    pn = "";
                }

                rt.Add("ptNum", pn);
            }

            if (ct > 1)
            {
                Debugger.Break();
                fn = "";
                foreach (ComponentDefinition df in rw.ComponentDefinitions)
                {
                    pt = aiDocument(rw.Document);
                    var ck = aiDocPartNum(pt);
                    if (ck == pn)
                        fn = fn + Constants.vbCrLf + pt.FullDocumentName;
                    else
                        Debugger.Break();
                }
            }

            rt.Add("aiDoc", fn);
        }

        return rt;
    }

    public Dictionary dBVg7f4(AssemblyDocument aiAssy, Dictionary dc = null)
    {
        // not sure what doing with this one
        // further developemt on hold
        // 
        // Dim pd As Inventor.AssemblyDocument
        Dictionary rt;

        BOMRow br;

        if (dc == null)
            rt = dBVg7f4(aiAssy, new Dictionary());
        else if (aiAssy == null)
            rt = dc;
        else
        {
            var bv = aiAssy.ComponentDefinition.BOM.BOMViews.get_Item("Structured");
            {
                var withBlock1 = bv;
                Debugger.Break();
            }
        }

        return rt;
    }

    public Dictionary bomInfoBkDn(BOMRowsEnumerator rwSet, Dictionary dc = null, string fn = "")
    {
        Dictionary rt;

        if (dc == null)
            rt = bomInfoBkDn(rwSet, new Dictionary(), fn);
        else
        {
            rt = dc;
            if (rwSet == null)
            {
            }
            else
                foreach (BOMRow rw in rwSet)
                {
                    // Stop
                    DoEvents();
                    var wk = bomItemInfo(rw);
                    wk.Add("ptOf", fn);
                    string pn = Convert.ToHexString(wk.get_Item("ptNum"));
                    var ck = fn + "|" + pn;
                    {
                        if (rt.Exists(ck))
                        {
                            // Stop
                            // debug.Print ConvertToJson(dcCmpTextOf2dc(wk,dcOb(.get_Item(ck))),vbTab)
                            // debug.Print ConvertToJson(
                            {
                                var withBlock1 =
                                        dcWBQbyCmpResult(dcCmpTextOf2dc(wk, dcOb(rt.get_Item(ck)))) // ,vbTab)
                                    ;
                                {
                                    var withBlock2 = dcOb(withBlock1.get_Item("!="));
                                    withBlock2.Remove("path");
                                    withBlock2.Remove("base");
                                    if (withBlock2.Count > 0)
                                    {
                                        Debug.Print("MISMATCH: ", ", ");
                                        Debugger.Break();
                                    }
                                }
                            }
                        }
                        else
                            rt.Add(ck, wk);
                    }

                    {
                        if (rw.ChildRows == null)
                            dc = dBVg1f2(ThisApplication.Documents.ItemByName(wk.get_Item("aiDoc")), wk, dc);
                        else
                        {
                            DoEvents();
                            dc = bomInfoBkDn(rw.ChildRows, dc, pn);
                        }
                    }
                    DoEvents();
                }
        }

        return rt;
    }

    public Dictionary dcOfBomsFromAiStructured(BOMRowsEnumerator rwSet, Dictionary dc = null, string fn = "")
    {
        // dcOfBomsFromAiStructured --
        // generate Dictionary of BOMs:
        // one for each distinct Assembly in
        // supplied Inventor BOM (structured)
        // 
        // returned as Dictionary of Assembly
        // sub Dictionaries, each keyed to its
        // Part Number and containing a set of
        // Item sub Dictionaries, again keyed
        // to Item P/N. Each Item sub Dictionary
        // represents a BOM line item
        // 
        Dictionary rt;

        return dBV0g0f4(dBV0g0f3(dBV0g0f1(rwSet, dc, fn)));
        Debug.Print(""); // Breakpoint Landing
    }

    public Dictionary dBV0g0f1(BOMRowsEnumerator rwSet, Dictionary dc = null, string fn = "")
    {
        // dBV0g0f1 -- retrieve BOM data
        // from a BOMRowsEnumerator
        // and its child row enumerators
        // 
        Dictionary rt;
        // Dim fd As Scripting.Dictionary
        // 
        // Dim kyIt As dynamic
        // 

        if (dc == null)
            rt = dBV0g0f1(rwSet, new Dictionary(), fn);
        else
        {
            rt = dc;
            if (rwSet == null)
                // Stop
                Debug.Print(""); // Breakpoint Landing
            else
            {
                Dictionary pd;
                {
                    if (!rt.Exists(fn))
                        rt.Add(fn, new Dictionary());
                    pd = rt.get_Item(fn);
                }

                foreach (BOMRow rw in rwSet)
                {
                    DoEvents();

                    var dt = bomItemInfo(rw);
                    string pn;
                    string th;
                    {
                        pn = Convert.ToHexString(dt.get_Item("ptNum"));
                        th = Convert.ToHexString(dt.get_Item("path"));
                    }

                    Dictionary it;
                    {
                        if (!pd.Exists(pn))
                            pd.Add(pn, new Dictionary());
                        it = pd.get_Item(pn);
                    }

                    {
                        if (it.Exists(th))
                            Debugger.Break();
                        else
                            it.Add(th, dt);
                    }

                    rt = dBV0g0f1(rw.ChildRows, rt, pn);
                    DoEvents();
                }
            }
        }

        // If Len(fn) = 0 Then Stop
        return rt;
    }

    public Dictionary dBV0g0f2(Dictionary dc)
    {
        var rt = new Dictionary();
        {
            foreach (var k0path in dc.Keys)
            {
                {
                    var withBlock1 = dcOb(dc.get_Item(k0path));
                    foreach (var k1Field in withBlock1.Keys)
                    {
                        var itValue = withBlock1.get_Item(k1Field);

                        Dictionary dcField;
                        {
                            if (!rt.Exists(k1Field))
                                rt.Add(k1Field, new Dictionary());
                            dcField = rt.get_Item(k1Field);
                        }

                        {
                            if (!dcField.Exists(itValue))
                                dcField.Add(itValue, new Dictionary());

                            {
                                var withBlock3 = dcOb(dcField.get_Item(itValue));
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

        return rt;
    }

    public Dictionary dBV0g0f3(Dictionary dc)
    {
        // dBV0g0f3 -- summarize BOM line item fields
        // 

        var rt = new Dictionary();

        {
            foreach (var k0Prod in dc.Keys)
            {
                Dictionary dcProd;
                {
                    rt.Add(k0Prod, new Dictionary());
                    dcProd = rt.get_Item(k0Prod);
                }

                {
                    var withBlock1 = dcOb(dc.get_Item(k0Prod));
                    foreach (var k1Item in withBlock1.Keys)
                        dcProd.Add(k1Item, dBV0g0f2(withBlock1.get_Item(k1Item)));
                }
            }
        }

        return rt;
    }

    public Dictionary dBV0g0f4(Dictionary dc)
    {
        // dBV0g0f4 -- reduce results of dBV0g0f3
        // to single values per field
        // for each Item under each Product
        // 

        var rt = new Dictionary();

        {
            foreach (var k0Prod in dc.Keys) // Products
            {
                Dictionary dcProd;
                {
                    rt.Add(k0Prod, new Dictionary());
                    dcProd = rt.get_Item(k0Prod);
                }

                {
                    var withBlock1 = dcOb(dc.get_Item(k0Prod)) // Items
                        ;
                    foreach (var k1Item in withBlock1.Keys)
                    {
                        Dictionary dcItem;
                        {
                            dcProd.Add(k1Item, new Dictionary());
                            dcItem = dcProd.get_Item(k1Item);
                        }

                        {
                            var withBlock2 = dcOb(withBlock1.get_Item(k1Item)) // Fields
                                ;
                            withBlock2.Remove("path");
                            withBlock2.Remove("base");
                            foreach (var k2Feld in withBlock2.Keys)
                            {
                                {
                                    var withBlock3 = dcOb(withBlock2.get_Item(k2Feld)) // Value(s)
                                        ;
                                    if (withBlock3.Count > 1)
                                        Debugger.Break();
                                    else
                                        dcItem.Add(k2Feld, withBlock3.Keys(0));
                                }
                            }
                        }
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary dBV0g0f5(Dictionary dc, string dlm = "|")
    {
        // dBV0g0f5 -- reduce results of dBV0g0f3
        // to single values per field
        // for each Item under each Product
        // 

        var rt = new Dictionary();

        {
            foreach (var k0Prod in dc.Keys) // Products
            {
                if (Len(k0Prod) <= 0) continue;
                {
                    var withBlock1 = dcOb(dc.get_Item(k0Prod)) // Items
                        ;
                    foreach (var k1Item in withBlock1.Keys)
                    {
                        Dictionary dcItem = withBlock1.get_Item(k1Item);

                        string rw;
                        {
                            rw = k0Prod + dlm + k1Item;
                            foreach (var k2Feld in new[]
                                     {
                                         "seq", "qty", "qtUnit"
                                     })
                            {
                                var co = dcItem.Exists(k2Feld)
                                    ? (string)Convert.ToHexString(dcItem.get_Item(k2Feld))
                                    : "";
                                rw = rw + dlm + co;
                            }

                            Debug.Print(""); // Breakpoint Landing
                        }

                        rt.Add(rw, dcItem);
                    }
                }
            }
        }

        return rt;
    }

    public string csvOfBomsFromDc(Dictionary dc, string dlm = "|")
    {
        // Product|Item|ItemOrder|QuantityInConversionUnit|ConversionUnit
        // NOTE[2021.08.20]: want to change 'Item'
        // to 'ItemCode' for compatibility with
        // current Genius BOM import format.
        // Will hold off for now.
        return Join(new[]
            {
                "Product", "Item", "ItemOrder", "QuantityInConversionUnit", "ConversionUnit"
            },
            dlm) + Constants.vbCrLf + txDumpLs(dBV0g0f5(dc, dlm).Keys);
    }

    public string csvOfBomsFromAiStructured(Document AiDoc, string dlm = "|")
    {
        return csvOfBomsFromDc(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows), dlm);
    }

    // 

    // 
    private string dvlBomView()
    {
        return "dvlBomView";
    }
}