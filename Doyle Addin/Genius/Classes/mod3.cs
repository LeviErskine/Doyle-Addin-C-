using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class mod3
{
    public Dictionary d0g2f1b(Dictionary dc)
    {
        var rt = new Dictionary();

        {
            foreach (var k0 in dc.Keys)
            {
                Dictionary i0 = dc.get_Item(k0);
                string fx = i0.get_Item("ext");
                Dictionary xt;
                {
                    if (!rt.Exists(fx))
                        rt.Add(fx, new Dictionary());

                    xt = rt.get_Item(fx); // i0(1)
                }

                xt.Add(k0, i0);
            }
        }

        return rt;
    }

    public Dictionary d0g2f1c(Dictionary dc)
    {
        // d0g2f1c --
        // derived from d0g2f1b
        // 

        string pn;

        var rt = new Dictionary();

        {
            foreach (var k0 in dc.Keys)
            {
                Dictionary i0 = dc.get_Item(k0);
                var xt = new Dictionary();

                foreach (var k1 in new[] { "Part Number", "Description", "ext", "fullname" })
                {
                    var ds = "";
                    if (i0.Exists(k1))
                    {
                        if ((i0.get_Item(k1)) is null)
                        {
                        }
                        else
                            ds = i0.get_Item(k1); // "Description"
                    }

                    xt.Add(k1, ds);
                }

                string fx = xt.get_Item("Part Number");

                {
                    if (!rt.Exists(fx))
                        rt.Add(fx, new Dictionary());

                    // xt =
                    dcOb(rt.get_Item(fx)).Add(xt.get_Item("fullname"), xt);
                }
            }
        }

        return rt;
    }

    public void m3g0f0()
    {
        {
            var withBlock = dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences);
            foreach (var ky in withBlock.Keys)
            {
                Document ad = aiDocument(withBlock.get_Item(ky));
                var dt = ad.DocumentType;
                if (ad.NeedsMigrating)
                    Debug.Print(ky);
            }
        }
    }

    public Dictionary m3g0f1migrate(Dictionary dc)
    {
        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                if (aiDocument(dc.get_Item(ky)).NeedsMigrating)
                    rt.Add(ky, dc.get_Item(ky));
            }
        }
        return rt;
    }
    // Debug.Print Join(m3g0f1migrate(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, vbCrLf)

    public Dictionary m3g0f1factories(Dictionary dc)
    {
        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
                rt = m3g0f3(m3g0f2(dc.get_Item(ky)), rt);
        }
        return rt;
    }

    public Dictionary m3g0f3(Document ad, Dictionary dc = null)
    {
        while (true)
        {
            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            if (ad == null)
            {
            }
            else if (dc.Exists(ad.FullFileName))
            {
            }
            else
                dc.Add(ad.FullDocumentName, ad);

            return dc;
        }
    }

    public Document m3g0f2(Document ad)
    {
        var dt = ad.DocumentType;
        return dt switch
        {
            kAssemblyDocumentObject => m3g0f2a(ad),
            kPartDocumentObject => m3g0f2p(ad),
            _ => null
        };
    }

    public Document m3g0f2a(AssemblyDocument ad)
    {
        if (ad.ComponentDefinition.IsiAssemblyFactory)
            return ad;
        return ad.ComponentDefinition.IsiAssemblyMember
            ? m3g0f2a(ad.ComponentDefinition.iAssemblyMember.ParentFactory.Parent.Document)
            : null;
    }

    public Document m3g0f2p(PartDocument ad)
    {
        if (ad.ComponentDefinition.IsiPartFactory)
            return ad;
        return ad.ComponentDefinition.IsiPartMember
            ? m3g0f2p(ad.ComponentDefinition.iPartMember.ParentFactory.Parent)
            : null;
    }

    public Dictionary m3g1f1()
    {
        // '
        // ' Test time taken for several operations
        // ' involving collection of Item data from Genius
        // ' and correlation with Inventor Model/Assembly
        // '
        // Dim cn As ADODB.Connection

        Document ad = ThisApplication.ActiveDocument;
        float tm = DateTime.Timer;
        var rs = cnGnsDoyle().Execute("select Item, Family from vgMfiItems");
        var ms = DateTime.Timer - tm;
        Debug.Print("Query Genius for Items: " + Convert.ToHexString(ms) + "sec");
        Debugger.Break();

        tm = DateTime.Timer;
        var dcGns = dcFrom2Fields(rs, "Item", "Family");
        ms = DateTime.Timer - tm;
        Debug.Print("Generate Dictionary from Result: " + Convert.ToHexString(ms) + "sec");
        Debugger.Break();

        tm = DateTime.Timer;
        var dcInv = m3g1f2(ad);
        ms = DateTime.Timer - tm;
        Debug.Print("Generate Dictionary from Assembly: " + Convert.ToHexString(ms) + "sec");
        Debugger.Break();

        tm = DateTime.Timer;
        {
            var withBlock = dcKeysInCommon(dcGns, dcInv);
            ms = DateTime.Timer - tm;
            Debug.Print("Join Dictionaries: " + Convert.ToHexString(ms) + "sec");
            Debugger.Break();

            Debugger.Break();
        }
        Debug.Print("");
    }

    public Dictionary m3g1f2(AssemblyDocument ad, long ct = 0)
    {
        // 
        return dcRemapByPtNum(dcAiDocComponents(ad, null, ct));
    }

    public Dictionary m3g1f3(ADODB.Recordset rs)
    {
        dynamic dt;
        long mxRw;
        long dxRw;

        var rt = new Dictionary();
        {
            if (rs.State == adStateClosed)
            {
            }
            else
            {
                {
                    var withBlock1 = rs.Fields;
                    var lsFd = new Dictionary();
                    var tx = "";
                    long mxCo = withBlock1.Count - 1;
                    for (long dxCo = 0; dxCo <= mxCo; dxCo++)
                    {
                        tx = tx + Constants.vbTab + withBlock1.get_Item(dxCo).Name;
                        lsFd.Add.get_Item(dxCo)
                            .Name(null /* Conversion error: Set to default value for this argument */, dxCo);
                    }

                    dynamic lsNm = Split(Mid(tx, 2), Constants.vbTab);
                }

                if (rs.BOF & rs.EOF)
                {
                }
                else
                    // dt = .GetRows
                    // dt = split(left$(.GetRows
                    // mxRw = UBound(dt, 2)
                    // For dxRw = 0 To mxRw
                    // Stop
                    // Next
                {
                    var withBlock1 = m3g1f4(rs.GetString(adClipString, null, Constants.vbTab, Constants.vbVerticalTab));
                }
            }
        }
    }

    public Dictionary m3g1f4(string txData)
    {
        var rt = new Dictionary();

        var lsDt = string.Split(Left(txData, InStrRev(txData, Constants.vbVerticalTab) - 1), Constants.vbVerticalTab);

        long mxCo = 0;
        long mxRw = UBound(lsDt);
        for (long dxRw = 0; dxRw <= mxRw; dxRw++)
        {
            string[] lsRw = Split(lsDt[dxRw], Constants.vbTab);
            Dictionary[] lsDc;
            if (mxCo == 0)
            {
                mxCo = UBound(lsRw);
                lsDc = new Dictionary[mxCo + 1];
                rt.Add("COLIDX", lsDc);
            }

            if (mxCo == UBound(lsRw))
            {
                for (long dxCo = 0; dxCo <= mxCo; dxCo++)
                {
                    var dcCo = lsDc[dxCo];
                    if (dcCo == null)
                    {
                        dcCo = new Dictionary();
                        lsDc[dxCo] = dcCo;
                    }

                    {
                        var ck = lsRw[dxCo];
                        Dictionary dcKy;
                        if (dcCo.Exists(ck))
                            dcKy = dcCo.get_Item(ck);
                        else
                        {
                            dcKy = new Dictionary();
                            dcCo.Add(ck, dcKy);
                        }

                        dcKy.Add(dxRw, dxRw);
                    }
                }
            }
            else
                Debugger.Break();
        }

        return rt;
    }
}