class dvl1
{
    // 
    /// Development Module dvl1 -- (generic until renamed)

    /// begun 2019.08.21

    /// by Andrew Thompson ()

    /// 

    /// Initial Purpose: Begin design of new Genius Properties Generator/Populator

    /// 

    public Variant d1g0f0()
    {
        d1g0f0 = 0;
    }

    public Scripting.Dictionary d1g4f0(Inventor.Document AiDoc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcRemapByPtNum(dcAiDocComponents(AiDoc));
            foreach (var ky in withBlock.Keys)
                // rt.Add ky, dcProps4genius(                aiDocument(.Item(ky)), , 0)
                rt.Add(ky, dcAiPropValsFromDc(dcOfPropsInAiDoc(aiDocument(withBlock.Item(ky)))));
        }

        d1g4f0 = rt;
    }

    public Scripting.Dictionary d1g4f1(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcOb(withBlock.Item(ky)).Count);
        }
        d1g4f1 = rt;
    }

    public string d1g4f2(string md, string pr)
    {
        VBIDE.VBComponent vbc;
        Variant rt;
        Variant ls;
        long mx;
        long dx;
        string ck;

        vbc = ThisDocument.VBAProject.InventorVBAComponents.Item(md).VBComponent;
        {
            var withBlock = vbc.CodeModule;
            ls = Split(withBlock.Lines(withBlock.ProcBodyLine(pr, vbext_pk_Proc), withBlock.ProcCountLines(pr, vbext_pk_Proc)), Constants.vbNewLine);
            mx = UBound(ls);
            for (dx = LBound(ls); dx <= mx; dx++)
            {
                ck = Trim(ls(dx));
                if (Left(ck, 1) == "'")
                    rt = rt + Mid(ck, 2) + Constants.vbNewLine;
            }
        }
        d1g4f2 = rt;
    }

    public Scripting.Dictionary d1g4f3(string hdr, string dlm)
    {
        Scripting.Dictionary rt;
        Variant ls;
        long mx;
        long dx;

        rt = new Scripting.Dictionary();
        ls = Split(hdr, dlm);
        mx = UBound(ls);
        for (dx = LBound(ls); dx <= mx; dx++)
        {
            rt.Add(dx, ls(dx));
            rt.Add(ls(dx), dx);
        }
        rt.Add("", dlm);
        d1g4f3 = rt;
    }

    public Scripting.Dictionary d1g4f4(Scripting.Dictionary dc, string tx)
    {
        Scripting.Dictionary hd;
        Scripting.Dictionary rt;
        Variant ls;
        long mx;
        long dx;
        string dlm;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            hd = dcOb(withBlock.Item(""));
            {
                var withBlock1 = hd;
                dlm = withBlock1.Item("");

                ls = Split(tx, dlm);
                mx = UBound(ls);
                for (dx = LBound(ls); dx <= mx; dx++)
                {
                    if (withBlock1.Exists(dx))
                        rt.Add.Item(dx);/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                    rt.Add(dx, ls(dx));
                }
            }

            dx = withBlock.Count;
            while (withBlock.Exists(dx))
                dx = 1 + dx;
            withBlock.Add(dx, rt);
        }

        d1g4f4 = dc;
    }

    public Scripting.Dictionary d1g4f5(string tx, Scripting.Dictionary dc, string bk = Constants.vbNewLine)
    {
        long ck;

        if (Strings.Len(tx) > 0)
        {
            ck = InStr(tx, bk);
            if (ck > 0)
                d1g4f5 = d1g4f5(Mid(tx, ck + Strings.Len(bk)), d1g4f4(dc, Left(tx, ck - 1)), bk);
            else
                d1g4f5 = d1g4f4(dc, tx);
        }
        else
            d1g4f5 = dc;
    }

    public Scripting.Dictionary d1g4f6(string tx, string dlm = ",", string bk = Constants.vbNewLine)
    {
        Scripting.Dictionary rt;
        string hdr;
        long ck;

        rt = new Scripting.Dictionary();
        ck = InStr(1, tx, bk);
        if (ck > 0)
        {
            hdr = Left(tx, ck - 1);
            rt.Add("", d1g4f3(hdr, dlm));
            rt = d1g4f5(Mid(tx, ck + Strings.Len(bk)), rt, bk);
        }
        else
        {
        }

        d1g4f6 = rt;
    }

    public Variant d1g1f0()
    {
        Scripting.Dictionary dc;

        dc = dcRemapByPtNum(dcAiDocComponents(aiDocActive()));
        Debug.Print(txDumpLs(dc.Keys));
        d1g1f0 = dc.Keys;
    }

    public string d1g1f2(long pd, long fc = 2)
    {
        long ct;
        long dv;
        string rt;

        if (fc > pd)
            d1g1f2 = "";
        else
        {
            ct = 0;
            dv = pd;
            while (!dv % fc > 0)
            {
                ct = 1 + ct;
                dv = dv / fc;
            }
            if (ct > 0)
                rt = System.Convert.ToHexString(fc) + "," + System.Convert.ToHexString(ct) + Constants.vbNewLine;
            else
                rt = "";
            d1g1f2 = rt + d1g1f2(dv, 1 + fc);
        }
    }

    public string d1g1f3(long pd, string tHdr = "Factor,Power", string fSep = ",", string lSep = Constants.vbNewLine)
    {
        long fc;
        long ct;
        long dv;
        string wk;
        string rt;

        rt = tHdr;
        ct = d1g1f3b2(pd);
        dv = pd / Math.Pow(2, ct);
        if (ct > 0)
            rt = rt + lSep; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */

        fc = 3;
        while (!fc > dv)
        {
            ct = 0;
            while (!dv % fc > 0)
            {
                ct = 1 + ct;
                dv = dv / fc;
            }

            if (ct > 0)
                rt = rt + lSep; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */

            fc = fc + 2;
        }
        d1g1f3 = rt;
    }

    public long d1g1f3b2(long pd)
    {
        long fc;
        long ct;
        long dv;
        string wk;
        string rt;

        ct = 0;
        dv = pd;
        while (!1 & dv)
        {
            ct = 1 + ct;
            dv = dv / 2;
        }

        d1g1f3b2 = ct;
    }

    public string fcPrime(long n, string rt = "", string ls = "BCEGKMQSW")
    {
        long nx;
        long md;
        long fc;
        long ct;

        if (n > 0)
        {
            if (n == 1)
                fcPrime = rt;
            else if (Strings.Len(ls) > 0)
            {
                nx = n;
                fc = 31 & Strings.Asc(ls);
                md = nx % fc;
                ct = 0;
                while (!md > 0)
                {
                    ct = 1 + ct;
                    nx = nx / fc;
                    md = nx % fc;
                }
                fcPrime = fcPrime(nx, rt + Chr(48 + ct), Mid(ls, 2));
            }
            else
                fcPrime = rt + "|" + System.Convert.ToHexString(n);
        }
        else
            fcPrime = "";
    }

    public string fcCommon(string s0, string s1)
    {
        long n0;
        long n1;

        n0 = Strings.Len(s0) * Strings.Len(s1);
        if (n0 > 0)
        {
            n0 = Strings.Asc(s0);
            n1 = Strings.Asc(s1);
            fcCommon = Chr(Interaction.IIf(n0 > n1, n1, n0)) + fcCommon(Mid(s0, 2), Mid(s1, 2));
        }
        else
            fcCommon = "";
    }

    public long fcProduct(string s, string ls = "BCEGKMQSW")
    {
        if (Strings.Len(ls) > 0)
        {
            if (Strings.Len(s) > 0)
                fcProduct = Math.Pow((31 & Strings.Asc(ls)), (15 & Strings.Asc(s))) * fcProduct(Mid(s, 2), Mid(ls, 2));
            else
                fcProduct = 1;
        }
        else
            fcProduct = -1;
    }

    public long fcMaxComm(long n0, long n1)
    {
        /// fcMaxComm -- Return Greatest Common Factor
        /// 
        fcMaxComm = fcProduct(fcCommon(fcPrime(n0), fcPrime(n1)));
    }

    public long gcfTest()
    {
        /// gcfTest -- Test GCF Function fcMaxComm
        /// 
        long rt;
        long n0;
        long n1;
        long nd;
        long gf;

        rt = 0;
        for (n0 = 4; n0 <= 49; n0++)
        {
            nd = n0 - 1;
            for (n1 = 2; n1 <= nd; n1++)
            {
                gf = fcMaxComm(n0, n1);
                if ((n0 % gf) + (n1 % gf) > 0)
                {
                    rt = 0;
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                }
            }
        }
    }

    public LongPtr[] tbPrimesWithSquare(long ct = 100000)
    {
        /// tbPrimesWithSquare
        /// 
        /// Generate a table of primes
        /// and their corresponding squares
        /// 
        long dbg;
        LongPtr[] p;
        LongPtr n;
        LongPtr mp;
        long dx;
        long mx;
        long nx;

        double d0;
        double d1;

        p = new LongPtr[2, ct + 1];
        mx = UBound(p, 2);
        p[0, 0] = 2; p[1, 0] = 4;
        nx = 1;
        n = 3;

        d0 = DateTime.Timer;
        do
        {
            dx = 0;
            mp = 1;
            do
            {
                mp = n % p[0, dx];
                if (n > p[1, dx])
                    dx = dx + 1;
                else
                    dx = nx;
            }
            while (mp * p[0, dx] > 0);
            if (mp > 0)
            {
                if (p[0, dx] == 0)
                {
                    p[0, dx] = n;
                    p[1, dx] = n * n;
                    if (Information.Err.Number == 0)
                        nx = dx + 1;
                    else
                        nx = mx + 1;
                }
                else
                    System.Diagnostics.Debugger.Break();
            }

            n = 1 + n;
        }
        while (!nx > mx);
        d1 = DateTime.Timer - d0;

        dbg = 0; // Change to 1 for debug mode
        if (dbg)
        {
            Debug.Print(1000 * d1); System.Diagnostics.Debugger.Break();
        }

        tbPrimesWithSquare = p;
    }

    public long d1g1f7()
    {
        double d0;
        double d1;
        VbMsgBoxResult ur;

        d0 = DateTime.Timer;
        ur = MsgBox("", Constants.vbOKOnly, "");
        d1 = DateTime.Timer - d0;

        System.Diagnostics.Debugger.Break();
    }

    public long bcCtCommFac(Scripting.Dictionary dc)
    {
        Variant ls;
        long rt;
        long mx;
        long dx;

        {
            var withBlock = dc;
            if (withBlock.Count > 0)
            {
                ls = withBlock.Keys;
                mx = UBound(ls);
                rt = System.Convert.ToInt64(withBlock.Item(ls(0)));
                dx = 1;

                do
                {
                    rt = fcMaxComm(rt, System.Convert.ToInt64(withBlock.Item(ls(dx))));
                    if (rt == 1)
                        dx = 1 + mx;
                    else
                        dx = 1 + dx;
                }
                while (!dx > mx);
            }
            else
                rt = 1;
        }

        bcCtCommFac = rt;
    }

    public Scripting.Dictionary dcBoltConn1byGCF(Scripting.Dictionary dc, long fc = 0)
    {
        Scripting.Dictionary rt;
        Variant ky;
        long ct;

        if (fc > 0)
        {
            rt = new Scripting.Dictionary();
            {
                var withBlock = dc;
                foreach (var ky in withBlock.Keys)
                {
                    ct = System.Convert.ToInt64(withBlock.Item(ky));
                    rt.Add(ky, ct / fc);
                }
            }
        }
        else
            rt = dcBoltConn1byGCF(dc, bcCtCommFac(dc));

        dcBoltConn1byGCF = rt;
    }

    public Inventor.Property aiDocProp(Inventor.Document AiDoc, string propName, string propSet = gnCustom)
    {
        /// Proposed Name: aiDocProp
        /// 
        Inventor.Property rt;

        if (AiDoc == null)
            aiDocProp = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            var withBlock = AiDoc.PropertySets;
            if (withBlock.PropertySetExists(propSet))
                // .Item(propSet).GetPropertyInfo()
                aiDocProp = aiGetProp(withBlock.Item(propSet), propName, 0);
            else
                aiDocProp = null/* TODO Change to default(_) if this is not a reference type */;
        }
    }

    public Variant aiDocPropVal(Inventor.Document AiDoc, string propName, string propSet = gnCustom)
    {
        /// Proposed Name: aiDocPropVal
        /// 
        aiDocPropVal = aiPropVal(aiDocProp(AiDoc, propName, propSet));
    }

    public Variant d1g2f0()
    {
        // 
        d1g2f0 = "";
    }

    public Variant d1g2f1(Inventor.Document AiDoc)
    {
        /// 
        /// 
        long pt;
        long sc;

        pt = 0; sc = 0;

        if (InStr(1, AiDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
        {
            pt = pt | 1; sc = sc + 1;
        }

        if (InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + System.Convert.ToHexString(aiDocPropVal(AiDoc, pnFamily, gnDesign)) + "|") > 0)
        {
            pt = pt | 2; sc = sc + 1;
        }

        d1g2f1 = "";
    }

    public Scripting.Dictionary d1g2f2(Inventor.AssemblyDocument AiDoc)
    {
        /// 
        /// 
        // Dim dc As Scripting.Dictionary
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence oc;
        // Dim rs As ADODB.Recordset
        Inventor.Document ob;
        // Dim ky As Variant
        // Dim pn As String
        // Dim pt As Long
        // Dim sc As Long
        long dx;
        Inventor.BOMStructureEnum bs;

        rt = new Scripting.Dictionary();
        dx = rt.Count;
        // pt = 0: sc = 0

        {
            var withBlock = AiDoc.ComponentDefinition;
            foreach (var oc in withBlock.Occurrences)
            {
                {
                    var withBlock1 = oc;
                    ob = aiDocument(withBlock1.Definition.Document);
                    if (withBlock1.BOMStructure == kPhantomBOMStructureThen)
                    {
                        if (withBlock1.DefinitionDocumentType == kAssemblyDocumentObjectThen)
                        {
                            {
                                var withBlock2 = aiDocAssy(ob);
                                // .ComponentDefinition.BOMStructure
                                if (withBlock2.DocumentInterests.HasInterest(guidDesignAccl))
                                {
                                }
                                else
                                {
                                }
                            }
                        }
                        else
                        {
                        }
                    }
                    else
                    {
                    }
                    {
                        var withBlock2 = withBlock1.Definition;
                        {
                            var withBlock3 = ob;
                        }
                    }
                }
            }
        }
        // If InStr(1, aiDoc.FullFileName,"\Doyle_Vault\Designs\purchased\") > 0 Then pt = pt Or 1: sc = sc + 1

        // If InStr(1,"|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|","|" & CStr(aiDocPropVal(aiDoc, pnFamily, gnDesign)) & "|") > 0 Then pt = pt Or 2: sc = sc + 1

        d1g2f2 = rt;
    }

    public Scripting.Dictionary d1g2f3(Inventor.Document AiDoc)
    {
        Scripting.Dictionary rt;

        rt = new Scripting.Dictionary();
        rt.Add(AiDoc.DocumentType, AiDoc);
        rt.Add(AiDoc.DocumentSubType.DocumentSubTypeID, AiDoc);
        d1g2f3 = rt;
    }

    public Variant d1g3f0()
    {
        /// 
        /// 
        d1g3f0 = "";
    }

    public Scripting.Dictionary d1g3f1(Inventor.AssemblyDocument ad)
    {
        /// d1g3f1 --
        /// 
        /// Generate counts of components
        /// in supplied Assembly, adding
        /// a sub-Dictionary for any
        /// "phantom" component recognized
        /// as either a Bolted Connection,
        /// or an Assembly of entirely
        /// Content Center components.
        /// 
        /// (the latter case addresses
        /// an issue encountered with
        /// just such an Assembly)
        /// 
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence oc;
        Inventor.Document sd;
        string nm;
        Scripting.Dictionary bc;
        Variant ar;

        rt = new Scripting.Dictionary();
        if (ad == null)
        {
        }
        else
            foreach (var oc in ad.ComponentDefinition.Occurrences)
            {
                sd = aiDocument(oc.Definition.Document);
                nm = sd.FullDocumentName;

                {
                    var withBlock = rt;
                    if (withBlock.Exists(nm))
                    {
                        ar = withBlock.Item(nm);
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                        ar(1) = ar(1) + 1;
                        withBlock.Item(nm) = ar;
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                    }
                    else
                    {
                        bc = null/* TODO Change to default(_) if this is not a reference type */;
                        if (oc.BOMStructure == kPhantomBOMStructure)
                        {
                            {
                                var withBlock1 = sd.DocumentInterests;
                                if (withBlock1.HasInterest(guidDesignAccl))
                                {
                                    Debug.Print("FOUND Design Accelerator");
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                                    bc = d1g3f1(sd); // New Scripting.Dictionary
                                }
                                else
                                {
                                    Debug.Print("FOUND Phantom Assembly");
                                    Debug.Print(Constants.vbTab + "NOT Design Accelerator");
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                                    bc = dcIfDesignAccel(d1g3f1(sd));
                                    if (bc == null)
                                    {
                                    }
                                    else
                                    {
                                        Debug.Print(Constants.vbTab + "but ALL Members ARE Content Center");
                                        Debug.Print(Constants.vbTab + "so WILL Process as Such");
                                    }
                                }
                            }

                            Debug.Print(Constants.vbTab + sd.FullDocumentName);
                            Debug.Print(Constants.vbTab + aiDocPropVal(sd, pnPartNum, gnDesign));
                        }
                        rt.Add(nm, Array(sd, 1, bc));
                    }
                }
            }

        d1g3f1 = rt;
    }

    public Scripting.Dictionary d1g3f2(Inventor.AssemblyDocument ad)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant sk;
        Variant ar;
        Variant sa;
        long ct;

        rt = new Scripting.Dictionary();
        {
            var withBlock = d1g3f1(ad);
            foreach (var ky in withBlock.Keys)
            {
                ar = withBlock.Item(ky);
                if (ar(2) == null)
                    rt.Add(ky, ar);
                else
                {
                    var withBlock1 = dcOb(ar(2));
                    foreach (var sk in withBlock1.Keys)
                    {
                        sa = withBlock1.Item(sk);
                        ct = ar(1) * sa(1);
                        {
                            var withBlock2 = rt;
                            if (withBlock2.Exists(sk))
                            {
                                // so need to add to existing total

                                // ct = ct + sa(1) '.Item(sk)
                                sa(1) = ct + withBlock2.Item(sk)(1);
                                // got type mismatch here, and fixed
                                // but not sure fix is correct

                                withBlock2.Item(sk) = sa; // ct
                            }
                            else
                            {
                                // so just add its count to the list

                                sa(1) = ct;
                                withBlock2.Add(sk, sa);
                            }
                        }
                    }
                }
            }
        }

        d1g3f2 = rt;
    }

    public Scripting.Dictionary d1g3f3(Inventor.AssemblyDocument ad)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant ar;

        rt = new Scripting.Dictionary();
        {
            var withBlock = d1g3f2(ad);
            foreach (var ky in withBlock.Keys)
            {
                ar = withBlock.Item(ky);
                {
                    var withBlock1 = aiDocument(obOf(ar(0))).PropertySets;
                    {
                        var withBlock2 = withBlock1.Item(gnDesign).Item(pnPartNum);
                        rt.Add.Value(null/* Conversion error: Set to default value for this argument */, ar);
                    }
                }
            }
        }

        d1g3f3 = rt;
    }

    public Scripting.Dictionary d1g3f4(Inventor.AssemblyDocument ad)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant ar;

        rt = new Scripting.Dictionary();
        {
            var withBlock = d1g3f3(ad);
            foreach (var ky in withBlock.Keys)
            {
                ar = withBlock.Item(ky);
                rt.Add(ky, ar(1));
            }
        }

        d1g3f4 = rt;
    }

    public Scripting.Dictionary d1g3f5(Inventor.AssemblyDocument ad, long incTop = 0)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Inventor.AssemblyDocument sd;
        Variant ky;
        // Dim ar As Variant

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcRemapByPtNum(dcAiDocComponents(ad, null/* Conversion error: Set to default value for this argument */, incTop, 1));
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocAssy(obOf(withBlock.Item(ky)));
                if (ad == null)
                {
                }
                else
                // '  Previous test, just for Bolted Connection
                // With ad.DocumentInterests
                // If .HasInterest(guidDesignAccl) Then
                // '  Replaced with test for ALL Phantom (below)

                {
                    var withBlock1 = ad.ComponentDefinition;
                    if (withBlock1.BOMStructure == kPhantomBOMStructure)
                        // Phantom -- don't add to Dictionary
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                    else
                        rt.Add(ky, d1g3f4(ad));
                }
            }
        }

        d1g3f5 = rt;
    }

    public Scripting.Dictionary d1g3f6(Inventor.AssemblyDocument ad, long incTop = 0)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant ar;

        rt = new Scripting.Dictionary();
        {
            var withBlock = d1g3f5(ad, incTop);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky + "|" + ky + "|1", Constants.vbNewLine + ky + "|" + dumpLsKeyVal(dcOb(withBlock.Item(ky)), "|", Constants.vbNewLine + ky + "|"));
        }

        d1g3f6 = rt;
    }

    public string d1g3f7(Inventor.AssemblyDocument ad, long incTop = 0)
    {
        d1g3f7 = "Product|ItemCode|Qty" + vbNewLine; /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
    }

    public Scripting.Dictionary dcOfBoltConnReLabeled(Scripting.Dictionary dc)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary rd;
        Scripting.Dictionary wd;
        Variant ky;
        string pn;
        Variant fn;
        Inventor.Document ad;
        Inventor.Property pr;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                wd = dcOb(withBlock.Item(ky));
                rd = new Scripting.Dictionary();

                pn = InputBox(Join(Array("Part Number proposed", "  for subassemblies", Join(wd.Keys, Constants.vbNewLine + "    "), "", "Modify as necessary,", "then click OK to confirm.", "")), "Verify BC Part Number", System.Convert.ToHexString(ky));
                {
                    var withBlock1 = wd;
                    foreach (var fn in withBlock1.Keys)
                    {
                        ad = aiDocument(obOf(withBlock1.Item(fn)));
                        pr = aiDocProp(ad, pnPartNum, gnDesign);
                        if (pr == null)
                        {
                        }
                        else
                        {
                            pr.Value = pn;
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */

                            rd.Add(fn, ad);
                        }
                    }
                }

                if (rd.Count > 0)
                    rt.Add(pn, rd);
            }
        }

        dcOfBoltConnReLabeled = rt;
    }

    public Scripting.Dictionary dcOfBoltConnIn(Inventor.AssemblyDocument ad, long incTop = 0)
    {
        {
            var withBlock = dcAiDocComponents(ad, null/* Conversion error: Set to default value for this argument */, incTop, 1);
        }
        /// 
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wd;
        Inventor.AssemblyDocument sd;
        Variant ky;
        string pn;
        string dn;
        // Dim ar As Variant

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAiDocComponents(ad, null/* Conversion error: Set to default value for this argument */, incTop, 1);
            foreach (var ky in withBlock.Keys)
            {
                sd = aiDocAssy(obOf(withBlock.Item(ky)));
                pn = pnOfBoltConn(sd);

                if (Strings.Len(pn) > 0)
                {
                    dn = sd.FullDocumentName;

                    {
                        var withBlock1 = rt;
                        if (withBlock1.Exists(pn))
                            wd = dcOb(withBlock1.Item(pn));
                        else
                        {
                            wd = new Scripting.Dictionary();
                            withBlock1.Add(pn, wd);
                        }
                    }

                    {
                        var withBlock1 = wd;
                        if (withBlock1.Exists(dn))
                        {
                            if (obOf(withBlock1.Item(dn)) == sd)
                            {
                            }
                            else
                                System.Diagnostics.Debugger.Break();
                        }
                        else
                            withBlock1.Add(dn, sd);
                    }
                }
            }
        }

        dcOfBoltConnIn = rt;
    }

    public Inventor.PartDocument aiDocContentMember(Inventor.PartDocument ad)
    {
        if (ad == null)
            aiDocContentMember = ad;
        else if (ad.ComponentDefinition.IsContentMember)
            aiDocContentMember = ad;
        else
            aiDocContentMember = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Scripting.Dictionary dcIfDesignAccel(Scripting.Dictionary dc)
    {
        /// dcIfDesignAccel
        /// 
        /// Accepting a Dictionary of form
        /// generated by d1g3f1, verify
        /// that all Items represent Content
        /// Center components, and return
        /// same Dictionary if so.
        /// 
        /// If any Items are NOT Content Center
        /// components, return Nothing
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant ar;

        rt = dc;

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ar = withBlock.Item(ky);
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                if (aiDocContentMember(aiDocPart(aiDocument(obOf(ar(0))))) == null)
                    rt = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }

        dcIfDesignAccel = rt;
    }

    public ADODB.Recordset rsOfBoltConn(Inventor.AssemblyDocument ad)
    {
        /// rsOfBoltConn -- rsOfBoltConn
        /// 
        /// Return Recordset of Components
        /// of one supplied Assembly Document,
        /// provided it's a Bolted Connection.
        /// 
        /// Call rsOfBoltConnRedux against this function's
        /// resulting Recordset rt to condense it
        /// to definition of a single instance,
        /// with a count of each member indicating
        /// the number of instances.
        /// 
        /// (was going to call rsOfBoltConnRedux here and
        /// return THAT result, but realized
        /// this function's preprocessed result
        /// might prove useful in itself, so
        /// decided to return it directly
        /// after all)
        /// 
        ADODB.Recordset rt; // Scripting.Dictionary
        ADODB.Field pNum;
        ADODB.Field fNam;
        ADODB.Field zPos;
        ADODB.Field xCen;
        ADODB.Field yCen;

        Inventor.ComponentOccurrence oc;
        Inventor.PartDocument sd;
        Scripting.Dictionary bc;
        double[] p0 = new double[3];
        double[] p1 = new double[3];

        rt = rsForBoltConn(); // New Scripting.Dictionary
        {
            var withBlock = rt.Fields;
            pNum = withBlock.Item("pNum");
            fNam = withBlock.Item("fNam");
            zPos = withBlock.Item("zPos");
            xCen = withBlock.Item("xCen");
            yCen = withBlock.Item("yCen");
        }

        if (ad == null)
        {
        }
        else
        {
            bc = null/* TODO Change to default(_) if this is not a reference type */;
            if (ad.ComponentDefinition.BOMStructure == kPhantomBOMStructure)
            {
                {
                    var withBlock = ad.DocumentInterests;
                    if (withBlock.HasInterest(guidDesignAccl))
                        bc = d1g3f1(ad);
                    else
                        bc = dcIfDesignAccel(d1g3f1(ad));
                }

                if (bc == null)
                {
                }
                else
                {
                    foreach (var oc in ad.ComponentDefinition.Occurrences)
                    {
                        {
                            var withBlock = oc;
                            sd = aiDocument(withBlock.Definition.Document);
                            {
                                var withBlock1 = withBlock.RangeBox;
                                withBlock1.MinPoint.GetPointData(p0);
                                withBlock1.MaxPoint.GetPointData(p1);
                            }
                        }


                        rt.AddNew();
                        pNum.Value = aiDocPropVal(sd, pnPartNum, gnDesign);
                        fNam.Value = sd.FullDocumentName;

                        zPos.Value = Round(p0[2], 3);
                        // Debug.Print FormatNumber(p0(2), 3); " ";
                        xCen.Value = Round((p0[0] + p1[0]) / 2, 3);
                        // Debug.Print FormatNumber((p0(0) + p1(0)) / 2, 3); " ";
                        yCen.Value = Round((p0[1] + p1[1]) / 2, 3);
                        // Debug.Print FormatNumber((p0(1) + p1(1)) / 2, 3); " ";
                        // Debug.Print

                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                    }

                    {
                        var withBlock = rt;
                        withBlock.Filter = "";
                        if (withBlock.BOF)
                        {
                            withBlock.AddNew();
                            pNum.Value = "NONE";
                            fNam.Value = "No Hardware, or Not Bolted Connection!";
                            zPos.Value = 0;
                            xCen.Value = 0;
                            yCen.Value = 0;
                        }
                        withBlock.Sort = "zPos, pNum, yCen, xCen";
                    }
                }
            }
        }

        rsOfBoltConn = rt; // rsOfBoltConnRedux()
    }

    public Scripting.Dictionary dcOfBoltConn(Inventor.AssemblyDocument ad) // ADODB.Recordset 'Scripting.Dictionary
    {
        /// dcOfBoltConn
        /// 
        /// Alternate implementation of rsOfBoltConn
        /// returning a Dictionary instead of
        /// a Recordset. However, this loses
        /// the benefit of a Recordset's Sort
        /// capability, and so is unlikely
        /// to prove as useful.
        /// 
        Scripting.Dictionary dc;
        string k0;
        long ct;

        ADODB.Recordset rt; // Scripting.Dictionary
        ADODB.Field pNum;
        ADODB.Field fNam;
        ADODB.Field zPos;
        ADODB.Field xCen;
        ADODB.Field yCen;

        Inventor.ComponentOccurrence oc;
        Inventor.PartDocument sd;
        Scripting.Dictionary bc;
        double[] p0 = new double[3];
        double[] p1 = new double[3];

        dc = new Scripting.Dictionary();
        rt = rsForBoltConn();
        {
            var withBlock = rt.Fields;
            pNum = withBlock.Item("pNum");
            fNam = withBlock.Item("fNam");
            zPos = withBlock.Item("zPos");
            xCen = withBlock.Item("xCen");
            yCen = withBlock.Item("yCen");
        }

        if (ad == null)
        {
        }
        else
        {
            bc = null/* TODO Change to default(_) if this is not a reference type */;
            if (ad.ComponentDefinition.BOMStructure == kPhantomBOMStructure)
            {
                {
                    var withBlock = ad.DocumentInterests;
                    if (withBlock.HasInterest(guidDesignAccl))
                        bc = d1g3f1(ad);
                    else
                        bc = dcIfDesignAccel(d1g3f1(ad));
                }

                if (bc == null)
                {
                }
                else
                {
                    foreach (var oc in ad.ComponentDefinition.Occurrences)
                    {
                        {
                            var withBlock = oc;
                            sd = aiDocument(withBlock.Definition.Document);
                            {
                                var withBlock1 = withBlock.RangeBox;
                                withBlock1.MinPoint.GetPointData(p0);
                                withBlock1.MaxPoint.GetPointData(p1);
                            }
                        }

                        k0 = FormatNumber(p0[2], 3) + "|" + aiDocPropVal(sd, pnPartNum, gnDesign);
                        {
                            var withBlock = dc;
                            if (withBlock.Exists(k0))
                            {
                                ct = 1 + withBlock.Item(k0);
                                withBlock.Item(k0) = ct;
                            }
                            else
                                withBlock.Add(k0, 1);
                        }

                        rt.AddNew();
                        pNum.Value = aiDocPropVal(sd, pnPartNum, gnDesign);
                        fNam.Value = sd.FullDocumentName;

                        zPos.Value = Round(p0[2], 3);
                        // Debug.Print FormatNumber(p0(2), 3); " ";
                        xCen.Value = Round((p0[0] + p1[0]) / 2, 3);
                        // Debug.Print FormatNumber((p0(0) + p1(0)) / 2, 3); " ";
                        yCen.Value = Round((p0[1] + p1[1]) / 2, 3);
                        // Debug.Print FormatNumber((p0(1) + p1(1)) / 2, 3); " ";
                        // Debug.Print

                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                    }

                    {
                        var withBlock = rt;
                        withBlock.Filter = "";
                        if (withBlock.BOF)
                        {
                            withBlock.AddNew();
                            pNum.Value = "NONE";
                            fNam.Value = "No Hardware, or Not Bolted Connection!";
                            zPos.Value = 0;
                            xCen.Value = 0;
                            yCen.Value = 0;
                        }
                        withBlock.Sort = "zPos, pNum, yCen, xCen";
                    }
                }
            }
        }

        dcOfBoltConn = dc; // rt
    }

    public ADODB.Recordset rsForBoltConn()
    {
        /// rsForBoltConn -- rsForBoltConn
        /// 
        /// Generate an new, empty Recordset
        /// to gather data on Bolted Connection
        /// 
        ADODB.Recordset rt;

        rt = new ADODB.Recordset();
        {
            var withBlock = rt;
            {
                var withBlock1 = withBlock.Fields;
                withBlock1.Append("zPos", adDouble);
                withBlock1.Append("pNum", adVarChar, 63);
                withBlock1.Append("fNam", adVarChar, 255);
                // .Append "", adVarChar, 63
                withBlock1.Append("xCen", adDouble);
                withBlock1.Append("yCen", adDouble);
            }
            withBlock.Open();
        }
        rsForBoltConn = rt;
    }

    public ADODB.Recordset rsOfBoltConnRedux(ADODB.Recordset rs)
    {
        /// rsOfBoltConnRedux
        /// 
        /// Condense supplied Recordset
        /// of Bolted Connection Assembly
        /// to summary of Components of
        /// ONE instance.
        /// 
        /// Include count of each member
        /// Component in Assembly, which
        /// should be the same for ALL
        /// Components, and reflect the
        /// total number of instances
        /// in the Assembly.
        /// 
        /// In most cases, this count
        /// should be just one, given
        /// the way Bolted Connections
        /// are generated and used here.
        /// However, some models might
        /// be found which use patterns
        /// or multiple holes, thus
        /// producing one BC Assembly
        /// defining multiple instances.
        /// A means to address this might
        /// therefore be required in future.
        /// 
        ADODB.Recordset rt;
        ADODB.Field pNumIn;
        ADODB.Field zPosIn;
        ADODB.Field pNumOut;
        ADODB.Field zPosOut;
        ADODB.Field xCenOut;

        Scripting.Dictionary dc;
        Scripting.Dictionary wk;
        Variant ky;

        double zp;
        string pn;

        dc = new Scripting.Dictionary();

        {
            var withBlock = rs;
            {
                var withBlock1 = withBlock.Fields;
                pNumIn = withBlock1.Item("pNum");
                zPosIn = withBlock1.Item("zPos");
            }

            withBlock.Sort = "zPos";
            if (!withBlock.BOF)
            {
                while (!withBlock.EOF)
                {
                    {
                        var withBlock1 = dc;
                        zp = zPosIn.Value;
                        if (withBlock1.Exists(zp))
                            wk = withBlock1.Item(zp);
                        else
                        {
                            wk = new Scripting.Dictionary();
                            withBlock1.Add(zp, wk);
                        }
                    }

                    {
                        var withBlock1 = wk;
                        pn = pNumIn.Value;
                        if (withBlock1.Exists(pn))
                            withBlock1.Item(pn) = 1 + withBlock1.Item(pn);
                        else
                            withBlock1.Add(pn, 1);
                    }
                    withBlock.MoveNext();
                }
            }
        }

        rt = rsForBoltConn();
        {
            var withBlock = rt;
            {
                var withBlock1 = withBlock.Fields;
                pNumOut = withBlock1.Item("pNum");
                zPosOut = withBlock1.Item("zPos");
                xCenOut = withBlock1.Item("xCen");
            }

            {
                var withBlock1 = dc;
                foreach (var ky in withBlock1.Keys)
                {
                    wk = withBlock1.Item(ky);

                    {
                        var withBlock2 = wk;
                        if (withBlock2.Count > 1)
                            System.Diagnostics.Debugger.Break();
                        else
                        {
                            rt.AddNew();
                            zPosOut.Value = System.Convert.ToDouble(ky);
                            pNumOut.Value = withBlock2.Keys(0);
                            xCenOut.Value = System.Convert.ToDouble(withBlock2.Items(0));
                        }
                    }
                }
            }

            withBlock.Filter = "";
            withBlock.Sort = "zPos, xCen";
        }

        rsOfBoltConnRedux = rt;
    }

    public ADODB.Recordset rsOfBoltConnRedux02(ADODB.Recordset rs)
    {
        /// rsOfBoltConnRedux02
        /// 
        /// Condense supplied Recordset
        /// of Bolted Connection Assembly
        /// to summary of Components of
        /// ONE instance.
        /// 
        /// Include count of each member
        /// Component in Assembly, which
        /// should be the same for ALL
        /// Components, and reflect the
        /// total number of instances
        /// in the Assembly.
        /// 
        /// In most cases, this count
        /// should be just one, given
        /// the way Bolted Connections
        /// are generated and used here.
        /// However, some models might
        /// be found which use patterns
        /// or multiple holes, thus
        /// producing one BC Assembly
        /// defining multiple instances.
        /// A means to address this might
        /// therefore be required in future.
        /// 
        ADODB.Recordset rt;
        ADODB.Field pNumIn;
        ADODB.Field zPosIn;
        ADODB.Field pNumOut;
        ADODB.Field zPosOut;
        ADODB.Field xCenOut;

        Scripting.Dictionary dc;
        Scripting.Dictionary wk;
        Variant ky;

        double zp;
        string pn;

        dc = new Scripting.Dictionary();

        {
            var withBlock = rs;
            {
                var withBlock1 = withBlock.Fields;
                pNumIn = withBlock1.Item("pNum");
                zPosIn = withBlock1.Item("zPos");
            }

            withBlock.Sort = "zPos";
            if (!withBlock.BOF)
            {
                while (!withBlock.EOF)
                {
                    {
                        var withBlock1 = dc;
                        zp = zPosIn.Value;
                        if (withBlock1.Exists(zp))
                            wk = withBlock1.Item(zp);
                        else
                        {
                            wk = new Scripting.Dictionary();
                            withBlock1.Add(zp, wk);
                        }
                    }

                    {
                        var withBlock1 = wk;
                        pn = pNumIn.Value;
                        if (withBlock1.Exists(pn))
                            withBlock1.Item(pn) = 1 + withBlock1.Item(pn);
                        else
                            withBlock1.Add(pn, 1);
                    }
                    withBlock.MoveNext();
                }
            }
        }

        rt = rsForBoltConn();
        {
            var withBlock = rt;
            {
                var withBlock1 = withBlock.Fields;
                pNumOut = withBlock1.Item("pNum");
                zPosOut = withBlock1.Item("zPos");
                xCenOut = withBlock1.Item("xCen");
            }

            {
                var withBlock1 = dc;
                foreach (var ky in withBlock1.Keys)
                {
                    wk = withBlock1.Item(ky);

                    {
                        var withBlock2 = wk;
                        if (withBlock2.Count > 1)
                            System.Diagnostics.Debugger.Break();
                        else
                        {
                            rt.AddNew();
                            zPosOut.Value = System.Convert.ToDouble(ky);
                            pNumOut.Value = withBlock2.Keys(0);
                            xCenOut.Value = System.Convert.ToDouble(withBlock2.Items(0));
                        }
                    }
                }
            }

            withBlock.Filter = "";
            withBlock.Sort = "zPos, xCen";
        }

        rsOfBoltConnRedux02 = rt;
    }

    public string bcPtNumFromRS(ADODB.Recordset rs)
    {
        bcPtNumFromRS = bcPtNumFromRSv2(rs);
    }

    public string bcPtNumFromRSv1(ADODB.Recordset rs)
    {
        /// bcPtNumFromRSv1
        /// 
        /// Generate a uniquely identifying
        /// Part Number from supplied Recordset
        /// Given a "Bolted Connection",
        /// 
        ADODB.Field pNumIn;
        ADODB.Field xCenIn;
        string rt;
        string pn;
        Variant ft;
        long ct;

        {
            var withBlock = rs;
            {
                var withBlock1 = withBlock.Fields;
                pNumIn = withBlock1.Item("pNum");
                xCenIn = withBlock1.Item("xCen");
            }

            withBlock.Sort = "zPos";
            if (withBlock.BOF | withBlock.EOF)
                rt = "";
            else
            {
                withBlock.Sort = "zPos";
                pn = pNumIn.Value;
                rt = "BC" + Mid(pn, 3, Strings.Len(pn) - 4) + Right(pn, 2);
                ct = xCenIn.Value;

                foreach (var ft in Array("zPos <= 0", "zPos > 0"))
                {
                    rt = rt + "-";
                    withBlock.Filter = ft;
                    if (!withBlock.BOF)
                    {
                        withBlock.Sort = "zPos";
                        while (!withBlock.EOF)
                        {
                            if (ct != xCenIn.Value)
                                System.Diagnostics.Debugger.Break();
                            pn = pNumIn.Value;
                            rt = rt + Left(pn, 2);
                            withBlock.MoveNext();
                        }
                    }
                }

                if (ct > 1)
                    rt = rt + Format(ct, "-X00");
            }
        }

        bcPtNumFromRSv1 = rt;
    }

    public string bcPtNumFromRSv2(ADODB.Recordset rs)
    {
        /// bcPtNumFromRSv2
        /// 
        /// Generate a uniquely identifying
        /// Part Number from supplied Recordset
        /// Given a "Bolted Connection",
        /// 
        ADODB.Field pNumIn;
        ADODB.Field xCenIn;
        string rt;
        string pn;
        Variant ft;
        long ct;

        {
            var withBlock = rs;
            {
                var withBlock1 = withBlock.Fields;
                pNumIn = withBlock1.Item("pNum");
                xCenIn = withBlock1.Item("xCen");
            }

            // .Sort = "zPos"
            withBlock.Filter = "";
            if (withBlock.BOF | withBlock.EOF)
                rt = "";
            else
            {
                withBlock.Sort = "zPos";
                pn = pNumIn.Value;
                rt = "BC" + Right(pn, 1) + Mid(pn, 3, Strings.Len(pn) - 4); // & Right$(pn, 2)
                ct = xCenIn.Value;

                foreach (var ft in Array("zPos <= 0|zPos", "zPos > 0|zPos desc"))
                {
                    rt = rt + "-";
                    withBlock.Filter = Left(ft, InStr(ft, "|") - 1);
                    if (!withBlock.BOF)
                    {
                        withBlock.Sort = Mid(ft, InStr(ft, "|") + 1); // "zPos"
                        rt = rt + Left(pNumIn.Value, 2);
                        withBlock.MoveNext();
                        while (!withBlock.EOF)
                        {
                            if (ct != xCenIn.Value)
                                System.Diagnostics.Debugger.Break();
                            // pn = pNumIn.Value
                            // rt = rt & Left$(pn, 2)
                            rt = rt + Left(pNumIn.Value, 1);
                            withBlock.MoveNext();
                        }
                    }
                }

                if (ct > 1)
                    rt = rt + Format(ct, "-X00");
            }
        }

        if (Strings.Len(rt) > 23)
            System.Diagnostics.Debugger.Break();
        bcPtNumFromRSv2 = rt;
    }

    public string pnOfBoltConn(Inventor.AssemblyDocument ad)
    {
        pnOfBoltConn = bcPtNumFromRSv1(rsOfBoltConnRedux(rsOfBoltConn(ad)));
    }

    public Scripting.Dictionary dcOfBoltConn02(Inventor.AssemblyDocument ad)
    {
        /// dcOfBoltConn02
        /// 
        /// Second variation on dcOfBoltConn
        /// returning a Dictionary of Component
        /// quantities, keyed on Item Number.
        /// 
        Scripting.Dictionary rt;
        string pNum;
        long ct;

        Inventor.ComponentOccurrence oc;
        Inventor.PartDocument sd;
        Scripting.Dictionary bc;

        rt = new Scripting.Dictionary();

        if (ad == null)
        {
        }
        else
        {
            bc = null/* TODO Change to default(_) if this is not a reference type */;
            if (ad.ComponentDefinition.BOMStructure == kPhantomBOMStructure)
            {
                {
                    var withBlock = ad.DocumentInterests;
                    if (withBlock.HasInterest(guidDesignAccl))
                        bc = d1g3f1(ad);
                    else
                        bc = dcIfDesignAccel(d1g3f1(ad));
                }

                if (bc == null)
                {
                }
                else
                    foreach (var oc in ad.ComponentDefinition.Occurrences)
                    {
                        sd = aiDocument(oc.Definition.Document);

                        pNum = aiDocPropVal(sd, pnPartNum, gnDesign);
                        {
                            var withBlock = rt;
                            if (withBlock.Exists(pNum))
                            {
                                ct = 1 + withBlock.Item(pNum);
                                withBlock.Item(pNum) = ct;
                            }
                            else
                                withBlock.Add(pNum, 1);
                        }

                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                    }
            }
        }

        dcOfBoltConn02 = rt;
    }

    public ADODB.Recordset rsFiltered(ADODB.Recordset rs, string flText = "")
    {
        rs.Filter = flText;
        rsFiltered = rs;
    }

    public ADODB.Recordset rsFromGnsSql(string sqlText)
    {
        /// 
        /// 
        ADODB.Recordset rt;

        {
            var withBlock = cnGnsDoyle();
            rt = withBlock.Execute(sqlText);
            if (rt == null)
                System.Diagnostics.Debugger.Break();
            rsFromGnsSql = rt;
        }
    }

    public ADODB.Recordset rsAiPurch01fromDict(Scripting.Dictionary dc)
    {
        /// 
        /// 
        rsAiPurch01fromDict = rsFromGnsSql(sqlSelAiPurch01fromDict(dc));
    }

    public ADODB.Recordset rsAiPurch01fromAssy(Inventor.Document AiDoc)
    {
        /// 
        /// 
        rsAiPurch01fromAssy = rsFromGnsSql(sqlSelAiPurch01fromAssy(AiDoc));
    }

    public ADODB.Recordset rsAiPdParts01fromAssy(Inventor.Document AiDoc)
    {
        /// 
        /// 
        rsAiPdParts01fromAssy = rsFromGnsSql(sqlSelAiPdParts01fromAssy(AiDoc));
    }

    public Scripting.Dictionary dcAiPurch01fromAdoRs(ADODB.Recordset rs)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        ADODB.Field fdItem;
        ADODB.Field fdType;
        ADODB.Field fdFmly;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rs;
            if (!withBlock.BOF)
            {
                withBlock.Filter = "";

                {
                    var withBlock1 = withBlock.Fields;
                    fdItem = withBlock1.Item("Item");
                    fdType = withBlock1.Item("Type");
                    fdFmly = withBlock1.Item("Family");
                }

                while (!withBlock.EOF)
                {
                    rt.Add(fdItem.Value, Array(fdType.Value, fdFmly.Value));
                    withBlock.MoveNext();
                }

                withBlock.Close();
            }
            dcAiPurch01fromAdoRs = rt;
        }
    }

    public Scripting.Dictionary dcAiPurch01fromDict(Scripting.Dictionary dc)
    {
        /// 
        /// 
        dcAiPurch01fromDict = dcAiPurch01fromAdoRs(rsAiPurch01fromDict(dc));
    }

    public Scripting.Dictionary dcAiPurch01fromAssy(Inventor.Document AiDoc)
    {
        /// 
        /// 
        dcAiPurch01fromAssy = dcAiPurch01fromAdoRs(rsAiPurch01fromAssy(AiDoc));
    }
}