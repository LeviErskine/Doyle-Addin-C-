class dvlVault
{
    /// 

    /// see module mod3 for other functions of possible use here

    /// named functions here originated there

    /// '

    public Variant ArrayFrom(Variant ls)
    {
        /// ArrayFrom -- return basic Variant Array
        /// from one of several various types
        /// of supplied Variant Values
        /// 
        Scripting.Dictionary dc;

        dc = dcOb(obOf(ls));
        if (dc == null)
        {
            if (IsObject(ls))
                ArrayFrom = Array();
            else if (IsArray(ls))
                ArrayFrom = ls;
            else
                ArrayFrom = Array(ls);
        }
        else
            ArrayFrom = dc.Keys;
    }

    public Scripting.Dictionary dcMapFSysVsVault(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        string bp;
        Variant fp;
        string vp;

        rt = new Scripting.Dictionary();

        bp = vaultBasePath();
        if (Strings.Len(bp) == 0)
            System.Diagnostics.Debugger.Break();// for debug/devel

        {
            var withBlock = dc // dcAiDocComponents(aiDocActive())
       ;
            foreach (var fp in withBlock.Keys)
            {
                vp = Replace(Replace(fp, bp, "$/"), @"\", "/");
                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(fp) | withBlock1.Exists(vp))
                        System.Diagnostics.Debugger.Break();
                    else
                    {
                        rt.Add(fp, vp);
                        rt.Add(vp, fp);
                    }
                }
            }
        }

        dcMapFSysVsVault = rt;
    }

    public Scripting.Dictionary dcRemapped2vaultPaths(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;
        string bp;

        rt = new Scripting.Dictionary();

        bp = vaultBasePath();
        if (Strings.Len(bp) == 0)
            System.Diagnostics.Debugger.Break();// for debug/devel

        {
            var withBlock = dc // dcAiDocComponents(aiDocActive())
       ;
            foreach (var ky in withBlock.Keys)
                rt.Add(Replace(Replace(ky, bp, "$/"), @"\", "/"), withBlock.Item(ky));
        }

        dcRemapped2vaultPaths = rt;
    }

    public string vaultBasePath()
    {
        {
            var withBlock = dcOb(nuILogicIfc().Apply("vltBasePath", new Scripting.Dictionary()));
            if (withBlock.Exists("OUT"))
                vaultBasePath = withBlock.Item("OUT");
            else
                vaultBasePath = "";
        }
    }

    public string vaultPropKeys()
    {
        {
            var withBlock = dcOb(nuILogicIfc().Apply("vltPropKeys", new Scripting.Dictionary()));
            if (withBlock.Exists("OUT"))
                vaultPropKeys = withBlock.Item("OUT");
            else
                vaultPropKeys = "";
        }
    }

    public Scripting.Dictionary dcOfDcByVltPathAndName(Scripting.Dictionary dc)
    {
        /// 
        /// 
        /// this one should probably call
        /// dcOfDcByNameAndPath against dc
        /// 
        /// actually, EACH should call
        /// some common function
        /// to perform similar task
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Variant ky;
        Variant ar;
        long bk;
        string bp;
        string fn;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcRemapped2vaultPaths(dc);
            foreach (var ky in withBlock.Keys)
            {
                ar = Array(withBlock.Item(ky));

                bk = InStrRev(ky, "/");
                if (bk == 0)
                    System.Diagnostics.Debugger.Break();
                else
                {
                    fn = Mid(ky, bk + 1);
                    bp = Left(ky, bk - 1);
                }

                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(bp))
                        withBlock1.Add(bp, new Scripting.Dictionary());

                    // gp =
                    {
                        var withBlock2 = dcOb(withBlock1.Item(bp));
                        withBlock2.Add(fn, ar(0));
                    }
                }
            }
        }

        dcOfDcByVltPathAndName = rt;
    }

    public Scripting.Dictionary dcOfDcByNameAndPath(Scripting.Dictionary dc)
    {
        /// 
        /// 
        /// closely related to
        /// dcOfDcByVltPathAndName
        /// (see above)
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Inventor.Document md;
        Variant ky;
        Variant ar;
        long bk;
        string fn;
        string bp;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc // dcRemapped2vaultPaths(dc)
       ;
            foreach (var ky in withBlock.Keys)
            {
                // ar = Array(.Item(ky))
                md = aiDocument(obOf(withBlock.Item(ky)));
                if (md == null)
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                else
                {
                    // Stop

                    bk = InStrRev(ky, @"\");
                    if (bk == 0)
                        System.Diagnostics.Debugger.Break();
                    else
                    {
                        bp = Left(ky, bk - 1);
                        fn = Mid(ky, bk + 1);
                    }

                    {
                        var withBlock1 = rt;
                        if (!withBlock1.Exists(fn))
                            withBlock1.Add(fn, new Scripting.Dictionary());

                        // gp =
                        {
                            var withBlock2 = dcOb(withBlock1.Item(fn));
                            withBlock2.Add(bp, md);
                        }
                    }
                }
            }
        }

        dcOfDcByNameAndPath = rt;
    }

    public Scripting.Dictionary d0g1f4d(Scripting.Dictionary dc)
    {
        /// d0g1f4d - categorize supplied Dictionary
        /// of Part/Assembly components
        /// by Vault Property Values
        /// 1 - takes same sort of
        /// Dictionary as d0g1f4c
        /// 2 - applies d0g1f4c to it
        /// 3 - rekeys the result
        /// 4 - transposes its sub Dictionaries
        /// 
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcOfDcRekeyedSecToPri(d0g1f4c(dc));
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcTransGrouped(dcOb(withBlock.Item(ky))));
        }

        d0g1f4d = rt;
    }

    public Scripting.Dictionary d0g1f4c(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        // Dim ag As Scripting.Dictionary

        Variant ls;
        Variant ky;

        Scripting.Dictionary sd;
        // Dim nm As Variant
        Variant pg;
        Variant rw;
        string p2;
        // Dim fl As Scripting.File

        rt = new Scripting.Dictionary();

        ls = dc.Keys;
        {
            var withBlock = nuILogicIfc();
            foreach (var ky in ls)
            {
                // send2clipBdWin10 ConvertToJson(nuILogicIfc()
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                                                             // pg =
                {
                    var withBlock1 = withBlock.Apply("dvl0", nuDcPopulator().Setting("PropName", "Name").Setting("Value", ky).Dictionary());
                    if (withBlock1.Exists("OUT"))
                        pg = withBlock1.Item("OUT");
                    else
                    {
                    }
                }
                // PropName", "FolderPath
                // FullPath
                // FullName
                // "$/Designs/doyle/(72) G3 Conveyor/I Parts/72-XXX-90403 G3 HD 8IN WRAP DRIVE 6IN END ROLLERS CONVEYOR BELT CRESCENT TOP ASSEMBLY"
                // , vbTab)
                /// REV[2023.03.03.1140]
                /// preceding pg assignment
                /// replaces the one following
                /// 
                // pg = .Apply("vlt04", nuDcPopulator().Setting("fullname", ky).Dictionary()).Item("OUT") '.DataFor(CStr(nm))
                // or "Full Path"
                foreach (var rw in pg)
                {
                    if (rw is Inventor.NameValueMap)
                        sd = dcFromAiNameValMap(obOf(rw));
                    else if (rw is Scripting.Dictionary)
                        sd = rw;
                    else
                    {
                        System.Diagnostics.Debugger.Break();
                        sd = null/* TODO Change to default(_) if this is not a reference type */;
                    }

                    if (sd == null)
                    {
                    }
                    else
                    {
                        p2 = sd.Item("fullname"); // .LocalForm()  'CStr(rw)
                        rt.Add(p2, sd);
                    }
                }
            }
        }

        d0g1f4c = rt;
    }

    public Scripting.Dictionary d0g1f4b(Variant ls)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary sd;
        Variant nm;
        Variant pg;
        Variant rw;
        string p2;
        Scripting.File fl;

        if (IsObject(ls))
        {
        }
        else if (IsArray(ls))
        {
        }
        else
            rt = d0g1f4b(Array(ls));

        rt = new Scripting.Dictionary();

        {
            var withBlock = nuILogicIfc() // nuIfcVault()
       ;
            foreach (var nm in ArrayFrom(ls))
            {
                pg = withBlock.Apply("vlt04", nuDcPopulator().Setting("PartNumber", nm).Dictionary()).Item("OUT"); // .DataFor(CStr(nm))
                foreach (var rw in pg) // Split(pg, vbNewLine)
                {
                    // Stop
                    if (rw is Inventor.NameValueMap)
                        sd = dcFromAiNameValMap(obOf(rw));
                    else if (rw is Scripting.Dictionary)
                        sd = rw;
                    else
                    {
                        System.Diagnostics.Debugger.Break();
                        sd = null/* TODO Change to default(_) if this is not a reference type */;
                    }

                    if (sd == null)
                    {
                    }
                    else
                    {
                        p2 = sd.Item("fullname"); // .LocalForm()  'CStr(rw)
                        rt.Add(p2, sd);
                        {
                            var withBlock1 = sd;
                            // .Add "ext", fnExt(p2)
                            withBlock1.Add("fileObj", fileIfPresent(p2));
                        }
                    }
                }
                {
                    var withBlock1 = rt;
                }
            }
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        if (false)
        {
        }

        d0g1f4b = rt;
    }

    public ADODB.Recordset d0g2f1d(Scripting.Dictionary dc) // Scripting.Dictionary
    {
        /// d0g2f1d --
        /// derived from d0g2f1b
        /// 
        Scripting.Dictionary rt;
        ADODB.Recordset rs;
        Scripting.Dictionary xt;
        Variant ls;
        Variant k0;
        Variant k1;
        Scripting.Dictionary i0;
        string fx;
        string ds;
        string pn;

        rt = new Scripting.Dictionary();

        ls = Array("Part Number", "Description");// ,"ext", "fullname"'

        rs = new ADODB.Recordset();
        {
            var withBlock = rs;
            {
                var withBlock1 = withBlock.Fields;
                foreach (var k1 in ls)
                    withBlock1.Append(k1, adVarChar, 127);
            }
            withBlock.Open();
        }

        {
            var withBlock = dc;
            foreach (var k0 in withBlock.Keys)
            {
                i0 = withBlock.Item(k0);

                rs.AddNew();

                foreach (var k1 in ls)
                {
                    {
                        var withBlock1 = i0;
                        ds = "";
                        if (withBlock1.Exists(k1))
                        {
                            if (IsEmpty(withBlock1.Item(k1)))
                            {
                            }
                            else
                                ds = withBlock1.Item(k1);
                        }
                    }

                    {
                        var withBlock1 = rs.Fields;
                        withBlock1.Item(k1) = ds;
                    }
                }
            }
        }
        rs.Filter = "";

        d0g2f1d = rs;
    }

    public Scripting.Dictionary dVg1f1(Variant argIn)
    {
        Scripting.Dictionary dc = new Scripting.Dictionary();
        Scripting.Dictionary rt;

        dc.Add("IN", argIn);
        {
            var withBlock = nuILogicIfc();
            rt = withBlock.Apply("vlt05", dc);
        }
        dVg1f1 = rt;
    }

    public Scripting.Dictionary dVg2f1(Scripting.Dictionary dc)
    {
        /// dVg2f1 - take Dictionary
        /// of Inventor Documents keyed
        /// to FullFileName as returned
        /// by dcAiDocComponents
        /// return Dictionary of Dictionaries
        /// of Inventor Documents keyed
        /// first to File Name only and
        /// then to ParentFolder Path
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary ls;
        Scripting.File fl;
        Variant ky;
        string nm;
        string fp;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                fl = fileIfPresent(System.Convert.ToHexString(ky));
                if (!fl == null)
                {
                    {
                        var withBlock1 = fl;
                        nm = withBlock1.Name;
                        fp = withBlock1.ParentFolder.Path;
                    }

                    {
                        var withBlock1 = rt;
                        if (!withBlock1.Exists(nm))
                            withBlock1.Add(nm, new Scripting.Dictionary());
                        ls = withBlock1.Item(nm);
                    }

                    {
                        var withBlock1 = ls;
                        if (!withBlock1.Exists(fp))
                            withBlock1.Add(fp, fl);
                    }
                }
            }
        }

        dVg2f1 = rt;
    }

    public Scripting.Dictionary dVg3f1(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Variant ob;
        string ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dVg1f1(dVg2f1(dc).Keys);
            if (withBlock.Exists("OUT"))
            {
                foreach (var ob in withBlock.Item("OUT"))
                {
                    wk = dcOb(ob);

                    if (wk == null)
                        System.Diagnostics.Debugger.Break();
                    else
                    {
                        var withBlock1 = wk;
                        if (withBlock1.Exists("fullname"))
                        {
                            ky = withBlock1.Item("fullname");
                            {
                                var withBlock2 = rt;
                                if (withBlock2.Exists(ky))
                                    System.Diagnostics.Debugger.Break();
                                else
                                    withBlock2.Add(ky, wk);
                            }
                        }
                        else
                            System.Diagnostics.Debugger.Break();
                    }
                }
            }
            else
                System.Diagnostics.Debugger.Break();
        }
        dVg3f1 = rt;
    }

    public Scripting.Dictionary dVg3f2(Scripting.Dictionary dc)
    {
        /// dVg3f2 -    '
        /// NOTE the following:
        /// dcMapFSysVsVault maps the full file names
        /// from the supplied Dictionary's Keys
        /// to their Vault paths/names,
        /// and vice-versa
        /// dVg3f1 returns a Dictionary
        /// keyed to Vault paths/names
        /// which must be translated
        /// to full file names
        /// dcKeysInCommon will return a Dictionary
        /// also keyed to Vault paths/names
        /// containing matching entries
        /// from the results of each
        /// of the prior two
        /// the Dictionary returned is keyed
        /// to the FIRST value in each
        /// entry from dcKeysInCommon,
        /// mapping it to the SECOND value
        /// in this way, each model's full file path
        /// is mapped to its Vault data
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant it;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcKeysInCommon(dcMapFSysVsVault(dc), dVg3f1(dc));
            foreach (var ky in withBlock.Keys)
            {
                it = withBlock.Item(ky);
                rt.Add(it(0), it(1));
            }
        }
        dVg3f2 = rt;
    }

    public Scripting.Dictionary dVg3f3(Scripting.Dictionary dc)
    {
        /// dVg3f3 -    given a Dictionary of Inventor Documents
        /// returns
        /// 
        Scripting.Dictionary d2;
        Scripting.Dictionary rt;
        Variant ky;
        Variant it;

        // rt =
        // New Scripting.Dictionary
        // d2 =
        // With
        rt = dcKeysCombined(dc, dVg3f2(dc));
        // For Each ky In .Keys
        // Stop
        // Next: End With

        dVg3f3 = rt;
    }

    /// END of module dvlVault

    /// 

    /// 
    private string dvlVault()
    {
    }
}