class SurroundingClass
{
    public Scripting.Dictionary d0g2f1b(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary xt;
        Variant k0;
        Scripting.Dictionary i0;
        string fx;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var k0 in withBlock.Keys)
            {
                i0 = withBlock.Item(k0);
                fx = i0.Item("ext");
                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(fx))
                        withBlock1.Add(fx, new Scripting.Dictionary());

                    xt = withBlock1.Item(fx); // i0(1)
                }

                xt.Add(k0, i0);
            }
        }

        d0g2f1b = rt;
    }

    public Scripting.Dictionary d0g2f1c(Scripting.Dictionary dc)
    {
        /// d0g2f1c --
        /// derived from d0g2f1b
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary xt;
        Variant k0;
        Variant k1;
        Scripting.Dictionary i0;
        string fx;
        string ds;
        string pn;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var k0 in withBlock.Keys)
            {
                i0 = withBlock.Item(k0);
                xt = new Scripting.Dictionary();

                foreach (var k1 in Array("Part Number", "Description", "ext", "fullname"))
                {
                    ds = "";
                    if (i0.Exists(k1))
                    {
                        if (IsEmpty(i0.Item(k1)))
                        {
                        }
                        else
                            ds = i0.Item(k1);// "Description"
                    }
                    xt.Add(k1, ds);
                }
                fx = xt.Item("Part Number"); // i0.Item("ext")

                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(fx))
                        withBlock1.Add(fx, new Scripting.Dictionary());

                    // xt =
                    dcOb(withBlock1.Item(fx)).Add(xt.Item("fullname"), xt);
                }
            }
        }

        d0g2f1c = rt;
    }

    public void m3g0f0()
    {
        Variant ky;
        Inventor.DocumentTypeEnum dt;
        Inventor.Document ad;

        {
            var withBlock = dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences);
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocument(withBlock.Item(ky));
                dt = ad.DocumentType;
                if (ad.NeedsMigrating)
                    Debug.Print(ky);
            }
        }
    }

    public Scripting.Dictionary m3g0f1migrate(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                if (aiDocument(withBlock.Item(ky)).NeedsMigrating)
                    rt.Add(ky, withBlock.Item(ky));
            }
        }
        m3g0f1migrate = rt;
    }
    // Debug.Print Join(m3g0f1migrate(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, vbNewLine)

    public Scripting.Dictionary m3g0f1factories(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt = m3g0f3(m3g0f2(withBlock.Item(ky)), rt);
        }
        m3g0f1factories = rt;
    }

    public Scripting.Dictionary m3g0f3(Inventor.Document ad, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;

        if (dc == null)
            m3g0f3 = m3g0f3(ad, new Scripting.Dictionary());
        else
        {
            rt = dc;
            if (ad == null)
            {
            }
            else if (rt.Exists(ad.FullFileName))
            {
            }
            else
                rt.Add(ad.FullDocumentName, ad);
        }

        m3g0f3 = rt;
    }

    public Inventor.Document m3g0f2(Inventor.Document ad)
    {
        Inventor.DocumentTypeEnum dt;

        dt = ad.DocumentType;
        if (dt == kAssemblyDocumentObject)
            m3g0f2 = m3g0f2a(ad);
        else if (dt == kPartDocumentObject)
            m3g0f2 = m3g0f2p(ad);
        else
            m3g0f2 = null/* TODO Change to default(_) if this is not a reference type */; // m3g0f2
    }

    public Inventor.Document m3g0f2a(Inventor.AssemblyDocument ad)
    {
        if (ad.ComponentDefinition.IsiAssemblyFactory)
            m3g0f2a = ad;
        else if (ad.ComponentDefinition.IsiAssemblyMember)
            m3g0f2a = m3g0f2a(ad.ComponentDefinition.iAssemblyMember.ParentFactory.Parent.Document);
        else
            m3g0f2a = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Document m3g0f2p(Inventor.PartDocument ad)
    {
        if (ad.ComponentDefinition.IsiPartFactory)
            m3g0f2p = ad;
        else if (ad.ComponentDefinition.IsiPartMember)
            m3g0f2p = m3g0f2p(ad.ComponentDefinition.iPartMember.ParentFactory.Parent);
        else
            m3g0f2p = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Scripting.Dictionary m3g1f1()
    {
        // '
        // '  Test time taken for several operations
        // '  involving collection of Item data from Genius
        // '  and correlation with Inventor Model/Assembly
        // '
        Inventor.Document ad;
        // Dim cn As ADODB.Connection
        ADODB.Recordset rs;
        Scripting.Dictionary dcGns;
        Scripting.Dictionary dcInv;
        float tm;
        float ms;

        ad = ThisApplication.ActiveDocument;
        tm = DateTime.Timer;
        rs = cnGnsDoyle().Execute("select Item, Family from vgMfiItems");
        ms = DateTime.Timer - tm;
        Debug.Print("Query Genius for Items: " + System.Convert.ToHexString(ms) + "sec");
        System.Diagnostics.Debugger.Break();

        tm = DateTime.Timer;
        dcGns = dcFrom2Fields(rs, "Item", "Family");
        ms = DateTime.Timer - tm;
        Debug.Print("Generate Dictionary from Result: " + System.Convert.ToHexString(ms) + "sec");
        System.Diagnostics.Debugger.Break();

        tm = DateTime.Timer;
        dcInv = m3g1f2(ad);
        ms = DateTime.Timer - tm;
        Debug.Print("Generate Dictionary from Assembly: " + System.Convert.ToHexString(ms) + "sec");
        System.Diagnostics.Debugger.Break();

        tm = DateTime.Timer;
        {
            var withBlock = dcKeysInCommon(dcGns, dcInv);
            ms = DateTime.Timer - tm;
            Debug.Print("Join Dictionaries: " + System.Convert.ToHexString(ms) + "sec");
            System.Diagnostics.Debugger.Break();

            System.Diagnostics.Debugger.Break();
        }
        Debug.Print();
    }

    public Scripting.Dictionary m3g1f2(Inventor.AssemblyDocument ad, long ct = 0)
    {
        /// 
        m3g1f2 = dcRemapByPtNum(dcAiDocComponents(ad, null/* Conversion error: Set to default value for this argument */, ct));
    }

    public Scripting.Dictionary m3g1f3(ADODB.Recordset rs)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary lsFd;
        Variant lsNm;
        Variant dt;
        string tx;
        long mxCo;
        long dxCo;
        long mxRw;
        long dxRw;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rs;
            if (withBlock.State == adStateClosed)
            {
            }
            else
            {
                {
                    var withBlock1 = withBlock.Fields;
                    lsFd = new Scripting.Dictionary();
                    tx = "";
                    mxCo = withBlock1.Count - 1;
                    for (dxCo = 0; dxCo <= mxCo; dxCo++)
                    {
                        tx = tx + Constants.vbTab + withBlock1.Item(dxCo).Name;
                        lsFd.Add.Item(dxCo).Name(null/* Conversion error: Set to default value for this argument */, dxCo);
                    }
                    lsNm = Split(Mid(tx, 2), Constants.vbTab);
                }

                if (withBlock.BOF & withBlock.EOF)
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
                    var withBlock1 = m3g1f4(withBlock.GetString(adClipString, null/* Conversion error: Set to default value for this argument */, Constants.vbTab, Constants.vbVerticalTab));
                }
            }
        }
    }

    public Scripting.Dictionary m3g1f4(string txData)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary dcCo;
        Scripting.Dictionary dcKy;
        Scripting.Dictionary[] lsDc;
        string[] lsDt;
        string[] lsRw;
        string ck;
        long mxRw;
        long dxRw;
        long mxCo;
        long dxCo;

        rt = new Scripting.Dictionary();

        lsDt = Split(Left(txData, InStrRev(txData, Constants.vbVerticalTab) - 1), Constants.vbVerticalTab);

        mxCo = 0;
        mxRw = UBound(lsDt);
        for (dxRw = 0; dxRw <= mxRw; dxRw++)
        {
            lsRw = Split(lsDt[dxRw], Constants.vbTab);
            if (mxCo == 0)
            {
                mxCo = UBound(lsRw);
                lsDc = new Scripting.Dictionary[mxCo + 1];
                rt.Add("COLIDX", lsDc);
            }

            if (mxCo == UBound(lsRw))
            {
                for (dxCo = 0; dxCo <= mxCo; dxCo++)
                {
                    dcCo = lsDc[dxCo];
                    if (dcCo == null)
                    {
                        dcCo = new Scripting.Dictionary();
                        lsDc[dxCo] = dcCo;
                    }

                    {
                        var withBlock = dcCo;
                        ck = lsRw[dxCo];
                        if (withBlock.Exists(ck))
                            dcKy = withBlock.Item(ck);
                        else
                        {
                            dcKy = new Scripting.Dictionary();
                            withBlock.Add(ck, dcKy);
                        }

                        dcKy.Add(dxRw, dxRw);
                    }
                }
            }
            else
                System.Diagnostics.Debugger.Break();
        }

        m3g1f4 = rt;
    }
}