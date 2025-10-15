using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class dvl1
{
    // 
    // Development Module dvl1 -- (generic until renamed)

    // begun 2019.08.21

    // by Andrew Thompson ()

    // 

    // Initial Purpose: Begin design of new Genius Properties Generator/Populator

    // 

    public static dynamic d1g0f0()
    {
        return 0;
    }

    public static Dictionary d1g4f0(Document AiDoc)
    {
        var rt = new Dictionary();

        {
            var withBlock = dcRemapByPtNum(dcAiDocComponents(AiDoc));
            foreach (var ky in withBlock.Keys)
                // rt.Add ky, dcProps4genius( aiDocument(.get_Item(ky)), , 0)
                rt.Add(ky, dcAiPropValsFromDc(dcOfPropsInAiDoc(aiDocument(withBlock.get_Item(ky)))));
        }

        return rt;
    }

    public static Dictionary d1g4f1(Dictionary dc)
    {
        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
                rt.Add(ky, dcOb(dc.get_Item(ky)).Count);
        }
        return rt;
    }

    public static string d1g4f2(string md, string pr)
    {
        string rt;

        VBIDE.VBComponent vbc = ThisDocument.VBAProject.InventorVBAComponents.get_Item(md).VBComponent;
        {
            var withBlock = vbc.CodeModule;
            var ls = string.Split(
                withBlock.Lines(withBlock.ProcBodyLine(pr, vbext_pk_Proc), withBlock.ProcCountLines(pr, vbext_pk_Proc)),
                Constants.vbCrLf);
            long mx = UBound(ls);
            for (long dx = LBound(ls); dx <= mx; dx++)
            {
                string ck = Trim(ls(dx));
                if (Left(ck, 1) == "'")
                    rt = rt + Mid(ck, 2) + Constants.vbCrLf;
            }
        }
        return rt;
    }

    public static Dictionary d1g4f3(string hdr, string dlm)
    {
        var rt = new Dictionary();
        var ls = string.Split(hdr, dlm);
        long mx = UBound(ls);
        for (long dx = LBound(ls); dx <= mx; dx++)
        {
            rt.Add(dx, ls(dx));
            rt.Add(ls(dx), dx);
        }

        rt.Add("", dlm);
        return rt;
    }

    public static Dictionary d1g4f4(Dictionary dc, string tx)
    {
        var rt = new Dictionary();
        {
            Dictionary hd = dcOb(dc.get_Item(""));
            long dx;
            {
                string dlm = hd.get_Item("");

                var ls = string.Split(tx, dlm);
                long mx = UBound(ls);
                for (dx = LBound(ls); dx <= mx; dx++)
                {
                    if (hd.Exists(dx))
                        rt.Add.get_Item(dx);
                    rt.Add(dx, ls(dx));
                }
            }

            dx = dc.Count;
            while (dc.Exists(dx))
                dx = 1 + dx;
            dc.Add(dx, rt);
        }

        return dc;
    }

    public static Dictionary d1g4f5(string tx, Dictionary dc, string bk = Constants.vbCrLf)
    {
        while (true)
        {
            if (Strings.Len(tx) <= 0) return dc;
            long ck = InStr(tx, bk);
            if (ck <= 0) return d1g4f4(dc, tx);
            var tx1 = tx;
            tx = Mid(tx, ck + Strings.Len(bk));
            dc = d1g4f4(dc, Left(tx1, ck - 1));
        }
    }

    public static Dictionary d1g4f6(string tx, string dlm = ",", string bk = Constants.vbCrLf)
    {
        var rt = new Dictionary();
        long ck = InStr(1, tx, bk);
        if (ck <= 0) return rt;
        string hdr = Left(tx, ck - 1);
        rt.Add("", d1g4f3(hdr, dlm));
        rt = d1g4f5(Mid(tx, ck + Strings.Len(bk)), rt, bk);

        return rt;
    }

    public static dynamic d1g1f0()
    {
        var dc = dcRemapByPtNum(dcAiDocComponents(aiDocActive()));
        Debug.Print(txDumpLs(dc.Keys));
        return dc.Keys;
    }

    public static string d1g1f2(long pd, long fc = 2)
    {
        if (fc > pd)
            return "";
        long ct = 0;
        var dv = pd;
        while (dv % fc > 0)
        {
            ct = 1 + ct;
            dv = dv / fc;
        }

        string rt;
        if (ct > 0)
            rt = Convert.ToString(fc) + "," + Convert.ToString(ct) + Constants.vbCrLf;
        else
            rt = "";
        return rt + d1g1f2(dv, 1 + fc);
    }

    public static string d1g1f3(long pd, string tHdr = "Factor,Power", string fSep = ",", string lSep = Constants.vbCrLf)
    {
        string wk;

        var rt = tHdr;
        var ct = d1g1f3b2(pd);
        long dv = pd / Math.Pow(2, ct);
        if (ct > 0)
            rt = rt + lSep;

        long fc = 3;
        while (fc > dv)
        {
            ct = 0;
            while (dv % fc > 0)
            {
                ct = 1 + ct;
                dv = dv / fc;
            }

            if (ct > 0)
                rt = rt + lSep;

            fc = fc + 2;
        }

        return rt;
    }

    public static long d1g1f3b2(long pd)
    {
        long fc;
        string wk;
        string rt;

        long ct = 0;
        var dv = pd;
        while ((1 & dv) != 0)
        {
            ct = 1 + ct;
            dv = dv / 2;
        }

        return ct;
    }

    public static string fcPrime(long n, string rt = "", string ls = "BCEGKMQSW")
    {
        while (true)
        {
            switch (n)
            {
                case <= 0:
                    return "";
                case 1:
                    return rt;
            }

            if (Strings.Len(ls) <= 0) return rt + "|" + Convert.ToString(n);
            var nx = n;
            long fc = 31 & Strings.Asc(ls);
            var md = nx % fc;
            long ct = 0;
            while (md > 0)
            {
                ct = 1 + ct;
                nx = nx / fc;
                md = nx % fc;
            }

            n = nx;
            rt = rt + Chr(48 + ct);
            ls = Mid(ls, 2);
        }
    }

    public  static string fcCommon(string s0, string s1)
    {
        long n0 = Strings.Len(s0) * Strings.Len(s1);
        if (n0 <= 0) return "";
        n0 = Strings.Asc(s0);
        long n1 = Strings.Asc(s1);
        return Chr(Interaction.IIf(n0 > n1, n1, n0)) + fcCommon(Mid(s0, 2), Mid(s1, 2));
    }

    public static long fcProduct(string s, string ls = "BCEGKMQSW")
    {
        if (Strings.Len(ls) <= 0) return -1;
        if (Strings.Len(s) > 0)
            return Math.Pow((31 & Strings.Asc(ls)), (15 & Strings.Asc(s))) * fcProduct(Mid(s, 2), Mid(ls, 2));
        return 1;
    }

    public static long fcMaxComm(long n0, long n1)
    {
        // fcMaxComm -- Return Greatest Common Factor

        return fcProduct(fcCommon(fcPrime(n0), fcPrime(n1)));
    }

    public static long gcfTest()
    {
        // gcfTest -- Test GCF Function fcMaxComm
        // 

        for (long n0 = 4; n0 <= 49; n0++)
        {
            var nd = n0 - 1;
            for (long n1 = 2; n1 <= nd; n1++)
            {
                var gf = fcMaxComm(n0, n1);
                if ((n0 % gf) + (n1 % gf) <= 0) continue;
                long rt = 0;
                Debug.Print("");
            }
        }
    }

    public static LongPtr[] tbPrimesWithSquare(long ct = 100000)
    {
        // tbPrimesWithSquare
        // 
        // Generate a table of primes
        // and their corresponding squares
        // 

        LongPtr[] p = new LongPtr[2, ct + 1];
        long mx = UBound(p, 2);
        p[0, 0] = 2;
        p[1, 0] = 4;
        long nx = 1;
        LongPtr n = 3;

        double d0 = DateTime.Timer;
        do
        {
            long dx = 0;
            LongPtr mp = 1;
            do
            {
                mp = n % p[0, dx];
                if (n > p[1, dx])
                    dx = dx + 1;
                else
                    dx = nx;
            } while (mp * p[0, dx] > 0);

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
                    Debugger.Break();
            }

            n = 1 + n;
        } while (nx > mx);

        var d1 = DateTime.Timer - d0;

        const long dbg = 0; // Change to 1 for debug mode
        if (dbg) return p;
        Debug.Print(1000 * d1);
        Debugger.Break();

        return p;
    }

    public static long d1g1f7()
    {
        double d0 = DateTime.Timer;
        VbMsgBoxResult ur = MessageBox.Show("", Constants.vbOKOnly, "");
        var d1 = DateTime.Timer - d0;

        Debugger.Break();
    }

    public static long bcCtCommFac(Dictionary dc)
    {
        long rt;

        {
            if (dc.Count > 0)
            {
                var ls = dc.Keys;
                long mx = UBound(ls);
                rt = Convert.ToInt64(dc.get_Item(ls(0)));
                long dx = 1;

                do
                {
                    rt = fcMaxComm(rt, Convert.ToInt64(dc.get_Item(ls(dx))));
                    if (rt == 1)
                        dx = 1 + mx;
                    else
                        dx = 1 + dx;
                } while (!dx > mx);
            }
            else
                rt = 1;
        }

        return rt;
    }

    public static Dictionary dcBoltConn1byGCF(Dictionary dc, long fc = 0)
    {
        Dictionary rt;

        if (fc > 0)
        {
            rt = new Dictionary();
            {
                foreach (var ky in dc.Keys)
                {
                    long ct = Convert.ToInt64(dc.get_Item(ky));
                    rt.Add(ky, ct / fc);
                }
            }
        }
        else
            rt = dcBoltConn1byGCF(dc, bcCtCommFac(dc));

        return rt;
    }

    public static Property aiDocProp(Document AiDoc, string propName, string propSet = gnCustom)
    {
        // Proposed Name: aiDocProp
        // 
        Property rt;

        if (AiDoc == null)
        {
        }
        else
        {
            var withBlock = AiDoc.PropertySets;
            if (withBlock.PropertySetExists(propSet))
                // .get_Item(propSet).GetPropertyInfo()
                return aiGetProp(withBlock.get_Item(propSet), propName, 0);
        }

        return null;
    }

    public static dynamic aiDocPropVal(Document AiDoc, string propName, string propSet = gnCustom)
    {
        // Proposed Name: aiDocPropVal
        // 
        return aiPropVal(aiDocProp(AiDoc, propName, propSet));
    }

    public static dynamic d1g2f0()
    {
        // 
        return "";
    }

    public static dynamic d1g2f1(Document AiDoc)
    {
        // 
        // 

        long pt = 0;
        long sc = 0;

        if (InStr(1, AiDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
        {
            pt = pt | 1;
            sc += 1;
        }

        if (InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|",
                "|" + Convert.ToHexString(aiDocPropVal(AiDoc, pnFamily, gnDesign)) + "|") <= 0) return "";
        pt = pt | 2;
        sc = sc + 1;

        return "";
    }

    public static Dictionary d1g2f2(AssemblyDocument AiDoc)
    {
        // 
        // 
        // Dim dc As Scripting.Dictionary

        // Dim rs As ADODB.Recordset
        // Dim ky As dynamic
        // Dim pn As String
        // Dim pt As Long
        // Dim sc As Long
        BOMStructureEnum bs;

        var rt = new Dictionary();
        long dx = rt.Count;

        // pt = 0: sc = 0
        {
            var withBlock = AiDoc.ComponentDefinition;
            foreach (ComponentOccurrence oc in withBlock.Occurrences)
            {
                {
                    var ob = aiDocument(oc.Definition.Document);
                    if (oc.BOMStructure == kPhantomBOMStructureThen)
                    {
                        if (oc.DefinitionDocumentType == kAssemblyDocumentObjectThen)
                        {
                            {
                                var withBlock2 = aiDocAssy(ob);
                                // .ComponentDefinition.BOMStructure
                                if (withBlock2.DocumentInterests.HasInterest(guidDesignAccl))
                                {
                                }
                            }
                        }
                    }

                    {
                        var withBlock2 = oc.Definition;
                        {
                            var withBlock3 = ob;
                        }
                    }
                }
            }
        }
        // If InStr(1, aiDoc.FullFileName,"\Doyle_Vault\Designs\purchased\") > 0 Then pt = pt Or 1: sc = sc + 1

        // If InStr(1,"|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|","|" & CStr(aiDocPropVal(aiDoc, pnFamily, gnDesign)) & "|") > 0 Then pt = pt Or 2: sc = sc + 1

        return rt;
    }

    public static Dictionary d1g2f3(Document AiDoc)
    {
        var rt = new Dictionary();
        rt.Add(AiDoc.DocumentType, AiDoc);
        rt.Add(AiDoc.DocumentSubType.DocumentSubTypeID, AiDoc);
        return rt;
    }

    public static dynamic d1g3f0()
    {
        // 
        // 
        return "";
    }

    public static Dictionary d1g3f1(AssemblyDocument ad)
    {
        // d1g3f1 --
        // 
        // Generate counts of components
        // in supplied Assembly, adding
        // a sub-Dictionary for any
        // "phantom" component recognized
        // as either a Bolted Connection,
        // or an Assembly of entirely
        // Content Center components.
        // 
        // (the latter case addresses
        // an issue encountered with
        // just such an Assembly)
        // 

        var rt = new Dictionary();
        if (ad == null)
        {
        }
        else
            foreach (ComponentOccurrence oc in ad.ComponentDefinition.Occurrences)
            {
                var sd = aiDocument(oc.Definition.Document);
                var nm = sd.FullDocumentName;

                {
                    if (rt.Exists(nm))
                    {
                        var ar = rt.get_Item(nm);
                        Debug.Print("");
                        ar(1) = ar(1) + 1;
                        rt.get_Item(nm) = ar;
                        Debug.Print("");
                    }
                    else
                    {
                        Dictionary bc = null;
                        if (oc.BOMStructure == kPhantomBOMStructure)
                        {
                            {
                                var withBlock1 = sd.DocumentInterests;
                                if (withBlock1.HasInterest(guidDesignAccl))
                                {
                                    Debug.Print("FOUND Design Accelerator");
                                    Debug.Print("");
                                    bc = d1g3f1(sd); // New Scripting.Dictionary
                                }
                                else
                                {
                                    Debug.Print("FOUND Phantom Assembly");
                                    Debug.Print(Constants.vbTab + "NOT Design Accelerator");
                                    Debug.Print("");
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

                        rt.Add(nm, new[] { sd, 1, bc });
                    }
                }
            }

        return rt;
    }

    public static Dictionary d1g3f2(AssemblyDocument ad)
    {
        // 
        // 

        var rt = new Dictionary();
        {
            var withBlock = d1g3f1(ad);
            foreach (var ky in withBlock.Keys)
            {
                var ar = withBlock.get_Item(ky);
                if (ar(2) == null)
                    rt.Add(ky, ar);
                else
                {
                    var withBlock1 = dcOb(ar(2));
                    foreach (var sk in withBlock1.Keys)
                    {
                        var sa = withBlock1.get_Item(sk);
                        long ct = ar(1) * sa(1);
                        {
                            if (rt.Exists(sk))
                            {
                                // so need to add to existing total

                                // ct = ct + sa(1) '.get_Item(sk)
                                sa(1) = ct + rt.get_Item(sk)(1);
                                // got type mismatch here, and fixed
                                // but not sure fix is correct

                                rt.get_Item(sk) = sa; // ct
                            }
                            else
                            {
                                // so just add its count to the list

                                sa(1) = ct;
                                rt.Add(sk, sa);
                            }
                        }
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary d1g3f3(AssemblyDocument ad)
    {
        // 
        // 

        var rt = new Dictionary();
        {
            var withBlock = d1g3f2(ad);
            foreach (var ky in withBlock.Keys)
            {
                var ar = withBlock.get_Item(ky);
                {
                    var withBlock1 = aiDocument(obOf(ar(0))).PropertySets;
                    {
                        var withBlock2 = withBlock1.get_Item(gnDesign).get_Item(pnPartNum);
                        rt.Add.Value(null /* Conversion error: Set to default value for this argument */, ar);
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary d1g3f4(AssemblyDocument ad)
    {
        // 
        // 

        var rt = new Dictionary();
        {
            var withBlock = d1g3f3(ad);
            foreach (var ky in withBlock.Keys)
            {
                var ar = withBlock.get_Item(ky);
                rt.Add(ky, ar(1));
            }
        }

        return rt;
    }

    public static Dictionary d1g3f5(AssemblyDocument ad, bool incTop = false)
    {
        // 
        // 
        AssemblyDocument sd;

        // Dim ar As dynamic
        var rt = new Dictionary();
        {
            var withBlock = dcRemapByPtNum(dcAiDocComponents(ad,
                null, incTop, 1));
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocAssy(obOf(withBlock.get_Item(ky)));
                if (ad == null)
                {
                }
                else
                    // ' Previous test, just for Bolted Connection
                    // With ad.DocumentInterests
                    // If .HasInterest(guidDesignAccl) Then
                    // ' Replaced with test for ALL Phantom (below)

                {
                    var withBlock1 = ad.ComponentDefinition;
                    if (withBlock1.BOMStructure == kPhantomBOMStructure)
                        // Phantom -- don't add to Dictionary
                        Debug.Print("");
                    else
                        rt.Add(ky, d1g3f4(ad));
                }
            }
        }

        return rt;
    }

    public static Dictionary d1g3f6(AssemblyDocument ad, bool incTop = false)
    {
        // 
        // 
        var ar;

        var rt = new Dictionary();
        {
            var withBlock = d1g3f5(ad, incTop);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky + "|" + ky + "|1",
                    Constants.vbCrLf + ky + "|" +
                    dumpLsKeyVal(dcOb(withBlock.get_Item(ky)), "|", Constants.vbCrLf + ky + "|"));
        }

        return rt;
    }

    public static string d1g3f7(AssemblyDocument ad, bool incTop = false)
    {
        return "Product|ItemCode|Qty" + vbCrLf;
    }

    public static Dictionary dcOfBoltConnReLabeled(Dictionary dc)
    {
        // 
        // 

        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                Dictionary wd = dcOb(dc.get_Item(ky));
                var rd = new Dictionary();

                string pn = InputBox(Join(new[]
                    {
                        "Part Number proposed", " for subassemblies", Join(wd.Keys, Constants.vbCrLf + " "), "",
                        "Modify as necessary,", "then click OK to confirm.", ""
                    }), "Verify BC Part Number",
                    Convert.ToString(ky as string));
                {
                    foreach (var fn in wd.Keys)
                    {
                        Document ad = aiDocument(obOf(wd.get_Item(fn)));
                        var pr = aiDocProp(ad, pnPartNum, gnDesign);
                        if (pr == null)
                        {
                        }
                        else
                        {
                            pr.Value = pn;
                            Debug.Print("");

                            rd.Add(fn, ad);
                        }
                    }
                }

                if (rd.Count > 0)
                    rt.Add(pn, rd);
            }
        }

        return rt;
    }

    public static Dictionary dcOfBoltConnIn(AssemblyDocument ad, bool incTop = false)
    {
        {
            var withBlock = dcAiDocComponents(ad, null,
                incTop, 1);
        }
        // 
        // 

        // Dim ar As dynamic
        var rt = new Dictionary();
        {
            var withBlock = dcAiDocComponents(ad, null, incTop, 1);
            foreach (var ky in withBlock.Keys)
            {
                AssemblyDocument sd = aiDocAssy(obOf(withBlock.get_Item(ky)));
                var pn = pnOfBoltConn(sd);

                if (Strings.Len(pn) <= 0) continue;
                var dn = sd.FullDocumentName;

                Dictionary wd;
                {
                    if (rt.Exists(pn))
                        wd = dcOb(rt.get_Item(pn));
                    else
                    {
                        wd = new Dictionary();
                        rt.Add(pn, wd);
                    }
                }

                {
                    if (wd.Exists(dn))
                    {
                        if (obOf(wd.get_Item(dn)) == sd)
                        {
                        }
                        else
                            Debugger.Break();
                    }
                    else
                        wd.Add(dn, sd);
                }
            }
        }

        return rt;
    }

    public static PartDocument aiDocContentMember(PartDocument ad)
    {
        if (ad == null || ad.ComponentDefinition.IsContentMember)
            return ad;
        return null;
    }

    public static Dictionary dcIfDesignAccel(Dictionary dc)
    {
        // dcIfDesignAccel
        // 
        // Accepting a Dictionary of form
        // generated by d1g3f1, verify
        // that all Items represent Content
        // Center components, and return
        // same Dictionary if so.
        // 
        // If any Items are NOT Content Center
        // components, return Nothing
        // 

        var rt = dc;

        {
            foreach (var ky in dc.Keys)
            {
                var ar = dc.get_Item(ky);
                Debug.Print("");
                if (aiDocContentMember(aiDocPart(aiDocument(obOf(ar(0))))) == null)
                    rt = null;
            }
        }

        return rt;
    }

    public static ADODB.Recordset rsOfBoltConn(AssemblyDocument ad)
    {
        // rsOfBoltConn -- rsOfBoltConn
        // 
        // Return Recordset of Components
        // of one supplied Assembly Document,
        // provided it's a Bolted Connection.
        // 
        // Call rsOfBoltConnRedux against this function's
        // resulting Recordset rt to condense it
        // to definition of a single instance,
        // with a count of each member indicating
        // the number of instances.
        // 
        // (was going to call rsOfBoltConnRedux here and
        // return THAT result, but realized
        // this function's preprocessed result
        // might prove useful in itself, so
        // decided to return it directly
        // after all)
        // 
        ADODB.Field pNum;
        ADODB.Field fNam;
        ADODB.Field zPos;
        ADODB.Field xCen;
        ADODB.Field yCen;

        var p0 = new double[3];
        var p1 = new double[3];

        var rt = rsForBoltConn(); // New Scripting.Dictionary
        // Scripting.Dictionary
        {
            var withBlock = rt.Fields;
            pNum = withBlock.get_Item("pNum");
            fNam = withBlock.get_Item("fNam");
            zPos = withBlock.get_Item("zPos");
            xCen = withBlock.get_Item("xCen");
            yCen = withBlock.get_Item("yCen");
        }

        if (ad == null)
        {
        }
        else
        {
            if (ad.ComponentDefinition.BOMStructure != kPhantomBOMStructure) return rt; // rsOfBoltConnRedux()
            Dictionary bc;
            {
                var withBlock = ad.DocumentInterests;
                bc = withBlock.HasInterest(guidDesignAccl) ? d1g3f1(ad) : dcIfDesignAccel(d1g3f1(ad));
            }

            if (bc == null)
            {
            }
            else
            {
                foreach (ComponentOccurrence oc in ad.ComponentDefinition.Occurrences)
                {
                    PartDocument sd;
                    {
                        sd = aiDocument(oc.Definition.Document);
                        {
                            var withBlock1 = oc.RangeBox;
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

                    Debug.Print("");
                }

                {
                    rt.Filter = "";
                    if (rt.BOF)
                    {
                        rt.AddNew();
                        pNum.Value = "NONE";
                        fNam.Value = "No Hardware, or Not Bolted Connection!";
                        zPos.Value = 0;
                        xCen.Value = 0;
                        yCen.Value = 0;
                    }

                    rt.Sort = "zPos, pNum, yCen, xCen";
                }
            }
        }

        return rt; // rsOfBoltConnRedux()
    }

    public static Dictionary dcOfBoltConn(AssemblyDocument ad) // ADODB.Recordset 'Scripting.Dictionary
    {
        // dcOfBoltConn
        // 
        // Alternate implementation of rsOfBoltConn
        // returning a Dictionary instead of
        // a Recordset. However, this loses
        // the benefit of a Recordset's Sort
        // capability, and so is unlikely
        // to prove as useful.
        // 

        ADODB.Field pNum;
        ADODB.Field fNam;
        ADODB.Field zPos;
        ADODB.Field xCen;
        ADODB.Field yCen;

        var p0 = new double[3];
        var p1 = new double[3];

        var dc = new Dictionary();
        var rt = rsForBoltConn(); // Scripting.Dictionary
        {
            var withBlock = rt.Fields;
            pNum = withBlock.get_Item("pNum");
            fNam = withBlock.get_Item("fNam");
            zPos = withBlock.get_Item("zPos");
            xCen = withBlock.get_Item("xCen");
            yCen = withBlock.get_Item("yCen");
        }

        if (ad == null)
        {
        }
        else
        {
            if (ad.ComponentDefinition.BOMStructure != kPhantomBOMStructure) return dc; // rt
            Dictionary bc;
            {
                var withBlock = ad.DocumentInterests;
                bc = withBlock.HasInterest(guidDesignAccl) ? d1g3f1(ad) : dcIfDesignAccel(d1g3f1(ad));
            }

            if (bc == null)
            {
            }
            else
            {
                foreach (ComponentOccurrence oc in ad.ComponentDefinition.Occurrences)
                {
                    PartDocument sd;
                    {
                        sd = aiDocument(oc.Definition.Document);
                        {
                            var withBlock1 = oc.RangeBox;
                            withBlock1.MinPoint.GetPointData(p0);
                            withBlock1.MaxPoint.GetPointData(p1);
                        }
                    }

                    var k0 = FormatNumber(p0[2], 3) + "|" + aiDocPropVal(sd, pnPartNum, gnDesign);
                    {
                        if (dc.Exists(k0))
                        {
                            long ct = 1 + dc.get_Item(k0);
                            dc.get_Item(k0) = ct;
                        }
                        else
                            dc.Add(k0, 1);
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

                    Debug.Print("");
                }

                {
                    rt.Filter = "";
                    if (rt.BOF)
                    {
                        rt.AddNew();
                        pNum.Value = "NONE";
                        fNam.Value = "No Hardware, or Not Bolted Connection!";
                        zPos.Value = 0;
                        xCen.Value = 0;
                        yCen.Value = 0;
                    }

                    rt.Sort = "zPos, pNum, yCen, xCen";
                }
            }
        }

        return dc; // rt
    }

    public static ADODB.Recordset rsForBoltConn()
    {
        // rsForBoltConn -- rsForBoltConn
        // 
        // Generate an new, null Recordset
        // to gather data on Bolted Connection
        // 

        var rt = new ADODB.Recordset();
        {
            {
                var withBlock1 = rt.Fields;
                withBlock1.Append("zPos", adDouble);
                withBlock1.Append("pNum", adVarChar, 63);
                withBlock1.Append("fNam", adVarChar, 255);
                // .Append "", adVarChar, 63
                withBlock1.Append("xCen", adDouble);
                withBlock1.Append("yCen", adDouble);
            }
            rt.Open();
        }
        return rt;
    }

    public static ADODB.Recordset rsOfBoltConnRedux(ADODB.Recordset rs)
    {
        // rsOfBoltConnRedux
        // 
        // Condense supplied Recordset
        // of Bolted Connection Assembly
        // to summary of Components of
        // ONE instance.
        // 
        // Include count of each member
        // Component in Assembly, which
        // should be the same for ALL
        // Components, and reflect the
        // total number of instances
        // in the Assembly.
        // 
        // In most cases, this count
        // should be just one, given
        // the way Bolted Connections
        // are generated and used here.
        // However, some models might
        // be found which use patterns
        // or multiple holes, thus
        // producing one BC Assembly
        // defining multiple instances.
        // A means to address this might
        // therefore be required in future.
        // 

        Dictionary wk;

        var dc = new Dictionary();

        {
            ADODB.Field pNumIn;
            ADODB.Field zPosIn;
            {
                var withBlock1 = rs.Fields;
                pNumIn = withBlock1.get_Item("pNum");
                zPosIn = withBlock1.get_Item("zPos");
            }

            rs.Sort = "zPos";
            if (!rs.BOF)
            {
                while (!rs.EOF)
                {
                    {
                        double zp = zPosIn.Value;
                        if (dc.Exists(zp))
                            wk = dc.get_Item(zp);
                        else
                        {
                            wk = new Dictionary();
                            dc.Add(zp, wk);
                        }
                    }

                    {
                        string pn = pNumIn.Value;
                        if (wk.Exists(pn))
                            wk.get_Item(pn) = 1 + wk.get_Item(pn);
                        else
                            wk.Add(pn, 1);
                    }
                    rs.MoveNext();
                }
            }
        }

        var rt = rsForBoltConn();
        {
            ADODB.Field pNumOut;
            ADODB.Field zPosOut;
            ADODB.Field xCenOut;
            {
                var withBlock1 = rt.Fields;
                pNumOut = withBlock1.get_Item("pNum");
                zPosOut = withBlock1.get_Item("zPos");
                xCenOut = withBlock1.get_Item("xCen");
            }

            {
                foreach (var ky in dc.Keys)
                {
                    wk = dc.get_Item(ky);

                    {
                        if (wk.Count > 1)
                            Debugger.Break();
                        else
                        {
                            rt.AddNew();
                            zPosOut.Value = Convert.ToDouble(ky as string);
                            pNumOut.Value = wk.Keys(0);
                            xCenOut.Value = Convert.ToDouble(wk.Items(0));
                        }
                    }
                }
            }

            rt.Filter = "";
            rt.Sort = "zPos, xCen";
        }

        return rt;
    }

    public static ADODB.Recordset rsOfBoltConnRedux02(ADODB.Recordset rs)
    {
        // rsOfBoltConnRedux02
        // 
        // Condense supplied Recordset
        // of Bolted Connection Assembly
        // to summary of Components of
        // ONE instance.
        // 
        // Include count of each member
        // Component in Assembly, which
        // should be the same for ALL
        // Components, and reflect the
        // total number of instances
        // in the Assembly.
        // 
        // In most cases, this count
        // should be just one, given
        // the way Bolted Connections
        // are generated and used here.
        // However, some models might
        // be found which use patterns
        // or multiple holes, thus
        // producing one BC Assembly
        // defining multiple instances.
        // A means to address this might
        // therefore be required in future.
        // 

        Dictionary wk;

        var dc = new Dictionary();

        {
            ADODB.Field pNumIn;
            ADODB.Field zPosIn;
            {
                var withBlock1 = rs.Fields;
                pNumIn = withBlock1.get_Item("pNum");
                zPosIn = withBlock1.get_Item("zPos");
            }

            rs.Sort = "zPos";
            if (!rs.BOF)
            {
                while (!rs.EOF)
                {
                    {
                        double zp = zPosIn.Value;
                        if (dc.Exists(zp))
                            wk = dc.get_Item(zp);
                        else
                        {
                            wk = new Dictionary();
                            dc.Add(zp, wk);
                        }
                    }

                    {
                        string pn = pNumIn.Value;
                        if (wk.Exists(pn))
                            wk.get_Item(pn) = 1 + wk.get_Item(pn);
                        else
                            wk.Add(pn, 1);
                    }
                    rs.MoveNext();
                }
            }
        }

        var rt = rsForBoltConn();
        {
            ADODB.Field xCenOut;
            ADODB.Field pNumOut;
            ADODB.Field zPosOut;
            {
                var withBlock1 = rt.Fields;
                pNumOut = withBlock1.get_Item("pNum");
                zPosOut = withBlock1.get_Item("zPos");
                xCenOut = withBlock1.get_Item("xCen");
            }

            {
                foreach (var ky in dc.Keys)
                {
                    wk = dc.get_Item(ky);

                    {
                        if (wk.Count > 1)
                            Debugger.Break();
                        else
                        {
                            rt.AddNew();
                            zPosOut.Value = Convert.ToDouble(ky);
                            pNumOut.Value = wk.Keys(0);
                            xCenOut.Value = Convert.ToDouble(wk.Items(0));
                        }
                    }
                }
            }

            rt.Filter = "";
            rt.Sort = "zPos, xCen";
        }

        return rt;
    }

    public static string bcPtNumFromRS(ADODB.Recordset rs)
    {
        return bcPtNumFromRSv2(rs);
    }

    public static string bcPtNumFromRSv1(ADODB.Recordset rs)
    {
        // bcPtNumFromRSv1
        // 
        // Generate a uniquely identifying
        // Part Number from supplied Recordset
        // Given a "Bolted Connection",
        // 

        {
            ADODB.Field xCenIn;
            ADODB.Field pNumIn;
            {
                var withBlock1 = rs.Fields;
                pNumIn = withBlock1.get_Item("pNum");
                xCenIn = withBlock1.get_Item("xCen");
            }

            rs.Sort = "zPos";
            string rt;
            if (rs.BOF | rs.EOF)
                rt = "";
            else
            {
                rs.Sort = "zPos";
                string pn = pNumIn.Value;
                rt = "BC" + Mid(pn, 3, Strings.Len(pn) - 4) + Right(pn, 2);
                long ct = xCenIn.Value;

                foreach (var ft in new[] { "zPos <= 0", "zPos > 0" })
                {
                    rt = rt + "-";
                    withBlock.Filter = ft;
                    if (withBlock.BOF) continue;
                    withBlock.Sort = "zPos";
                    while (!withBlock.EOF)
                    {
                        if (ct != xCenIn.Value)
                            Debugger.Break();
                        pn = pNumIn.Value;
                        rt = rt + Left(pn, 2);
                        withBlock.MoveNext();
                    }
                }

                if (ct > 1)
                    rt = rt + Format(ct, "-X00");
            }
        }

        return rt;
    }

    public static string bcPtNumFromRSv2(ADODB.Recordset rs)
    {
        // bcPtNumFromRSv2
        // 
        // Generate a uniquely identifying
        // Part Number from supplied Recordset
        // Given a "Bolted Connection",
        // 

        {
            ADODB.Field pNumIn;
            ADODB.Field xCenIn;
            {
                var withBlock1 = rs.Fields;
                pNumIn = withBlock1.get_Item("pNum");
                xCenIn = withBlock1.get_Item("xCen");
            }

            // .Sort = "zPos"
            rs.Filter = "";
            string rt;
            if (rs.BOF | rs.EOF)
                rt = "";
            else
            {
                rs.Sort = "zPos";
                string pn = pNumIn.Value;
                rt = "BC" + Right(pn, 1) + Mid(pn, 3, Strings.Len(pn) - 4); // & Right$(pn, 2)
                long ct = xCenIn.Value;

                foreach (var ft in new[] { "zPos <= 0|zPos", "zPos > 0|zPos desc" })
                {
                    rt = rt + "-";
                    withBlock.Filter = Left(ft, InStr(ft, "|") - 1);
                    if (withBlock.BOF) continue;
                    withBlock.Sort = Mid(ft, InStr(ft, "|") + 1); // "zPos"
                    rt = rt + Left(pNumIn.Value, 2);
                    withBlock.MoveNext();
                    while (!withBlock.EOF)
                    {
                        if (ct != xCenIn.Value)
                            Debugger.Break();
                        // pn = pNumIn.Value
                        // rt = rt & Left$(pn, 2)
                        rt = rt + Left(pNumIn.Value, 1);
                        withBlock.MoveNext();
                    }
                }

                if (ct > 1)
                    rt = rt + Format(ct, "-X00");
            }
        }

        if (Strings.Len(rt) > 23)
            Debugger.Break();
        return rt;
    }

    public static string pnOfBoltConn(AssemblyDocument ad)
    {
        return bcPtNumFromRSv1(rsOfBoltConnRedux(rsOfBoltConn(ad)));
    }

    public static Dictionary dcOfBoltConn02(AssemblyDocument ad)
    {
        // dcOfBoltConn02
        // 
        // Second variation on dcOfBoltConn
        // returning a Dictionary of Component
        // quantities, keyed on Item Number.
        // 
        var rt = new Dictionary();

        if (ad == null)
        {
        }
        else
        {
            if (ad.ComponentDefinition.BOMStructure != kPhantomBOMStructure) return rt;
            Dictionary bc;
            {
                var withBlock = ad.DocumentInterests;
                bc = withBlock.HasInterest(guidDesignAccl) ? d1g3f1(ad) : dcIfDesignAccel(d1g3f1(ad));
            }

            if (bc == null)
            {
            }
            else
                foreach (ComponentOccurrence oc in ad.ComponentDefinition.Occurrences)
                {
                    PartDocument sd = aiDocument(oc.Definition.Document);

                    string pNum = aiDocPropVal(sd, pnPartNum, gnDesign);
                    {
                        if (rt.Exists(pNum))
                        {
                            long ct = 1 + rt.get_Item(pNum);
                            rt.get_Item(pNum) = ct;
                        }
                        else
                            rt.Add(pNum, 1);
                    }

                    Debug.Print("");
                }
        }

        return rt;
    }

    public static ADODB.Recordset rsFiltered(ADODB.Recordset rs, string flText = "")
    {
        rs.Filter = flText;
        return rs;
    }

    public static ADODB.Recordset rsFromGnsSql(string sqlText)
    {
        // 
        // 

        {
            var withBlock = cnGnsDoyle();
            var rt = withBlock.Execute(sqlText);
            if (rt == null)
                Debugger.Break();
            return rt;
        }
    }

    public static ADODB.Recordset rsAiPurch01fromDict(Dictionary dc)
    {
        // 
        // 
        return rsFromGnsSql(sqlSelAiPurch01fromDict(dc));
    }

    public static ADODB.Recordset rsAiPurch01fromAssy(Document AiDoc)
    {
        // 
        // 
        return rsFromGnsSql(sqlSelAiPurch01fromAssy(AiDoc));
    }

    public static ADODB.Recordset rsAiPdParts01fromAssy(Document AiDoc)
    {
        // 
        // 
        return rsFromGnsSql(sqlSelAiPdParts01fromAssy(AiDoc));
    }

    public static Dictionary dcAiPurch01fromAdoRs(ADODB.Recordset rs)
    {
        // 
        // 

        var rt = new Dictionary();
        {
            if (rs.BOF) return rt;
            rs.Filter = "";

            ADODB.Field fdItem;
            ADODB.Field fdType;
            ADODB.Field fdFmly;
            {
                var withBlock1 = rs.Fields;
                fdItem = withBlock1.get_Item("Item");
                fdType = withBlock1.get_Item("Type");
                fdFmly = withBlock1.get_Item("Family");
            }

            while (!rs.EOF)
            {
                rt.Add(fdItem.Value, new[] { fdType.Value, fdFmly.Value });
                withBlock.MoveNext();
            }

            rs.Close();
            return rt;
        }
    }

    public static Dictionary dcAiPurch01fromDict(Dictionary dc)
    {
        // 
        // 
        return dcAiPurch01fromAdoRs(rsAiPurch01fromDict(dc));
    }

    public static Dictionary dcAiPurch01fromAssy(Document AiDoc)
    {
        // 
        // 
        return dcAiPurch01fromAdoRs(rsAiPurch01fromAssy(AiDoc));
    }
}