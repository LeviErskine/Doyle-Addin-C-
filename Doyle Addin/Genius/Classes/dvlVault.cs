using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class dvlVault
{
    // 

    // see module mod3 for other functions of possible use here

    // named functions here originated there

    // '

    public dynamic ArrayFrom(dynamic ls)
    {
        // ArrayFrom -- return basic dynamic Array
        // from one of several various types
        // of supplied dynamic Values
        // 

        Dictionary dc = dcOb(obOf(ls));
        if (dc != null) return dc.Keys;
        if (ls is not null)

            return Array.Empty<string>();
        if (false)
            return ls;
        return new
            string[] { (dynamic)null };
    }

    public Dictionary dcMapFSysVsVault(Dictionary dc)
    {
        var rt = new Dictionary();

        var bp = vaultBasePath();
        if (Strings.Len(bp) == 0)
            Debugger.Break(); // for debug/devel

        {
            foreach (var fp in dc.Keys)
            {
                string vp = Replace(Replace(fp, bp, "$/"), @"\", "/");
                {
                    if (rt.Exists(fp) | rt.Exists(vp))
                        Debugger.Break();
                    else
                    {
                        rt.Add(fp, vp);
                        rt.Add(vp, fp);
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary dcRemapped2vaultPaths(Dictionary dc)
    {
        var rt = new Dictionary();

        var bp = vaultBasePath();
        if (Strings.Len(bp) == 0)
            Debugger.Break(); // for debug/devel

        {
            foreach (var ky in dc.Keys)
                rt.Add(Replace(Replace(ky, bp, "$/"), @"\", "/"), dc.get_Item(ky));
        }

        return rt;
    }

    public string vaultBasePath()
    {
        {
            var withBlock = dcOb(nuILogicIfc().Apply("vltBasePath", new Dictionary()));
            if (withBlock.Exists("OUT"))
                return withBlock.get_Item("OUT");
            return "";
        }
    }

    public string vaultPropKeys()
    {
        {
            var withBlock = dcOb(nuILogicIfc().Apply("vltPropKeys", new Dictionary()));
            return withBlock.Exists("OUT") ? (string)withBlock.get_Item("OUT") : "";
        }
    }

    public Dictionary dcOfDcByVltPathAndName(Dictionary dc)
    {
        // 
        // 
        // this one should probably call
        // dcOfDcByNameAndPath against dc
        // 
        // actually, EACH should call
        // some common function
        // to perform similar task
        // 
        Dictionary gp;

        var rt = new Dictionary();

        {
            var withBlock = dcRemapped2vaultPaths(dc);
            foreach (var ky in withBlock.Keys)
            {
                dynamic ar = new[] { withBlock.get_Item(ky) };

                long bk = InStrRev(ky, "/");
                string bp;
                string fn;
                if (bk == 0)
                    Debugger.Break();
                else
                {
                    fn = Mid(ky, bk + 1);
                    bp = Left(ky, bk - 1);
                }

                {
                    if (!rt.Exists(bp))
                        rt.Add(bp, new Dictionary());

                    // gp =
                    {
                        var withBlock2 = dcOb(rt.get_Item(bp));
                        withBlock2.Add(fn, ar(0));
                    }
                }
            }

            ;
        }

        return rt;
    }

    public Dictionary dcOfDcByNameAndPath(Dictionary dc)
    {
        // 
        // 
        // closely related to
        // dcOfDcByVltPathAndName
        // (see above)
        // 
        Dictionary gp;
        dynamic ar;

        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                // ar = new string[] {.get_Item(ky))
                Document md = aiDocument(obOf(dc.get_Item(ky)));
                if (md == null)
                    Debug.Print(""); // Breakpoint Landing
                else
                {
                    // Stop

                    long bk = InStrRev(ky, @"\");
                    string fn;
                    string bp;
                    if (bk == 0)
                        Debugger.Break();
                    else
                    {
                        bp = Left(ky, bk - 1);
                        fn = Mid(ky, bk + 1);
                    }

                    {
                        if (!rt.Exists(fn))
                            rt.Add(fn, new Dictionary());

                        // gp =
                        {
                            var withBlock2 = dcOb(rt.get_Item(fn));
                            withBlock2.Add(bp, md);
                        }
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary d0g1f4d(Dictionary dc)
    {
        // d0g1f4d - categorize supplied Dictionary
        // of Part/Assembly components
        // by Vault Property Values
        // 1 - takes same sort of
        // Dictionary as d0g1f4c
        // 2 - applies d0g1f4c to it
        // 3 - rekeys the result
        // 4 - transposes its sub Dictionaries
        // 
        // 

        var rt = new Dictionary();

        {
            var withBlock = dcOfDcRekeyedSecToPri(d0g1f4c(dc));
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcTransGrouped(dcOb(withBlock.get_Item(ky))));
        }

        return rt;
    }

    public Dictionary d0g1f4c(Dictionary dc)
    {
        // Dim ag As Scripting.Dictionary

        // Dim nm As dynamic

        // Dim fl As Scripting.File
        var rt = new Dictionary();

        dynamic ls = dc.Keys;
        {
            var withBlock = nuILogicIfc();
            foreach (var ky in ls)
            {
                // send2clipBdWin10 ConvertToJson(nuILogicIfc()
                Debug.Print(""); // Breakpoint Landing
                // pg =
                dynamic pg;
                {
                    var withBlock1 = withBlock.Apply("dvl0",
                        nuDcPopulator().Setting("PropName", "Name").Setting("Value", ky).Dictionary());
                    if (withBlock1.Exists("OUT"))
                        pg = withBlock1.get_Item("OUT");
                }
                // PropName", "FolderPath
                // FullPath
                // FullName
                // "$/Designs/doyle/(72) G3 Conveyor/I Parts/72-XXX-90403 G3 HD 8IN WRAP DRIVE 6IN END ROLLERS CONVEYOR BELT CRESCENT TOP ASSEMBLY"
                // , vbTab)
                // REV[2023.03.03.1140]
                // preceding pg assignment
                // replaces the one following
                // 
                // pg = .Apply("vlt04", nuDcPopulator().Setting("fullname", ky).Dictionary()).get_Item("OUT") '.DataFor(CStr(nm))
                // or "Full Path"
                foreach (var rw in pg)
                {
                    Dictionary sd;
                    switch (rw)
                    {
                        case NameValueMap:
                            sd = dcFromAiNameValMap(obOf(rw));
                            break;
                        case Dictionary:
                            sd = rw;
                            break;
                        default:
                            Debugger.Break();
                            sd = null;
                            break;
                    }

                    if (sd == null)
                    {
                    }
                    else
                    {
                        string p2 = sd.get_Item("fullname");
                        rt.Add(p2, sd);
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary d0g1f4b(dynamic ls)
    {
        Dictionary rt;

        Scripting.File fl;

        if (ls is not null)
        {
        }
        else
            rt = d0g1f4b(new[] { ls });

        rt = new Dictionary();

        {
            var withBlock = nuILogicIfc() // nuIfcVault()
                ;
            foreach (var nm in ArrayFrom(ls))
            {
                dynamic pg = withBlock.Apply("vlt04", nuDcPopulator().Setting("PartNumber", nm).Dictionary())
                    .get_Item("OUT");
                foreach (var rw in pg) // Split(pg, vbCrLf)
                {
                    // Stop
                    Dictionary sd;
                    switch (rw)
                    {
                        case NameValueMap:
                            sd = dcFromAiNameValMap(obOf(rw));
                            break;
                        case Dictionary:
                            sd = rw;
                            break;
                        default:
                            Debugger.Break();
                            sd = null;
                            break;
                    }

                    if (sd == null)
                    {
                    }
                    else
                    {
                        string p2 = sd.get_Item("fullname");
                        rt.Add(p2, sd);
                        {
                            // .Add "ext", fnExt(p2)
                            sd.Add("fileObj", fileIfPresent(p2));
                        }
                    }
                }

                {
                    var withBlock1 = rt;
                }
            }
        }

        Debug.Print(""); // Breakpoint Landing
        if (false)
        {
        }

        return rt;
    }

    public ADODB.Recordset d0g2f1d(Dictionary dc) // Scripting.Dictionary
    {
        // d0g2f1d --
        // derived from d0g2f1b
        // 
        Dictionary xt;
        string fx;
        string pn;

        var rt = new Dictionary();

        dynamic ls = new[] { "Part Number", "Description" }; // ,"ext", "fullname"'

        var rs = new ADODB.Recordset();
        {
            {
                var withBlock1 = rs.Fields;
                foreach (var k1 in ls)
                    withBlock1.Append(k1, adVarChar, 127);
            }
            rs.Open();
        }

        {
            foreach (var k0 in dc.Keys)
            {
                Dictionary i0 = dc.get_Item(k0);

                rs.AddNew();

                foreach (var k1 in ls)
                {
                    string ds;
                    {
                        ds = "";
                        if (i0.Exists(k1))
                        {
                            if (IsEmpty(i0.get_Item(k1)))
                            {
                            }
                            else
                                ds = i0.get_Item(k1);
                        }
                    }

                    {
                        var withBlock1 = rs.Fields;
                        withBlock1.get_Item(k1) = ds;
                    }
                }
            }
        }
        rs.Filter = "";

        return rs;
    }

    public Dictionary dVg1f1(dynamic argIn)
    {
        var dc = new Dictionary();
        Dictionary rt;

        dc.Add("IN", argIn);
        {
            var withBlock = nuILogicIfc();
            rt = withBlock.Apply("vlt05", dc);
        }
        return rt;
    }

    public Dictionary dVg2f1(Dictionary dc)
    {
        // dVg2f1 - take Dictionary
        // of Inventor Documents keyed
        // to FullFileName as returned
        // by dcAiDocComponents
        // return Dictionary of Dictionaries
        // of Inventor Documents keyed
        // first to File Name only and
        // then to ParentFolder Path
        // 

        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                var fl = fileIfPresent(Convert.ToString(ky as string));
                if (fl != null) continue;
                string nm;
                string fp;
                {
                    nm = fl.Name;
                    fp = fl.ParentFolder.Path;
                }

                Dictionary ls;
                {
                    if (!rt.Exists(nm))
                        rt.Add(nm, new Dictionary());
                    ls = rt.get_Item(nm);
                }

                {
                    if (!ls.Exists(fp))
                        ls.Add(fp, fl);
                }
            }
        }

        return rt;
    }

    public Dictionary dVg3f1(Dictionary dc)
    {
        var rt = new Dictionary();
        {
            var withBlock = dVg1f1(dVg2f1(dc).Keys);
            if (withBlock.Exists("OUT"))
            {
                foreach (var ob in withBlock.get_Item("OUT"))
                {
                    Dictionary wk = dcOb(ob);

                    if (wk == null)
                        Debugger.Break();
                    else
                    {
                        if (wk.Exists("fullname"))
                        {
                            string ky = wk.get_Item("fullname");
                            {
                                if (rt.Exists(ky))
                                    Debugger.Break();
                                else
                                    rt.Add(ky, wk);
                            }
                        }
                        else
                            Debugger.Break();
                    }
                }
            }
            else
                Debugger.Break();
        }
        return rt;
    }

    public Dictionary dVg3f2(Dictionary dc)
    {
        // dVg3f2 - '
        // NOTE the following:
        // dcMapFSysVsVault maps the full file names
        // from the supplied Dictionary's Keys
        // to their Vault paths/names,
        // and vice-versa
        // dVg3f1 returns a Dictionary
        // keyed to Vault paths/names
        // which must be translated
        // to full file names
        // dcKeysInCommon will return a Dictionary
        // also keyed to Vault paths/names
        // containing matching entries
        // from the results of each
        // of the prior two
        // the Dictionary returned is keyed
        // to the FIRST value in each
        // entry from dcKeysInCommon,
        // mapping it to the SECOND value
        // in this way, each model's full file path
        // is mapped to its Vault data
        // 

        var rt = new Dictionary();

        {
            var withBlock = dcKeysInCommon(dcMapFSysVsVault(dc), dVg3f1(dc));
            foreach (var ky in withBlock.Keys)
            {
                var it = withBlock.get_Item(ky);
                rt.Add(it(0), it(1));
            }
        }
        return rt;
    }

    public Dictionary dVg3f3(Dictionary dc)
    {
        // dVg3f3 - given a Dictionary of Inventor Documents
        // returns
        // 
        Dictionary d2;
        dynamic ky;
        dynamic it;

        // rt =
        // New Scripting.Dictionary
        // d2 =
        // With
        var rt = dcKeysCombined(dc, dVg3f2(dc));

        // For Each ky In .Keys
        // Stop
        // Next: End With
        return rt;
    }

    // END of module dvlVault

    // 

    // 
    private string dvlVault()
    {
    }
}