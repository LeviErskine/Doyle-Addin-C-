using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class libVbCodeXport
{
    public long xprtModText()
    {
        VBIDE.VBProject vp;
        VBIDE.VBComponent vc;

        var dc = dcOfVbProjects(ThisApplication.VBAProjects); // ThisWorkbook.Application.VBE.VBProjects

        const string n1 = @"C:\Users\athompson\Documents\dvl\libExt.xlsm";
        Dictionary d1 = dc(n1);
        send2clipBdWin10(n1 + Constants.vbCrLf + dumpKeyedText(d1, d1));
        Debugger.Break();

        const string n2 = @"C:\Users\athompson\Documents\dvl\libExt-rcvr-2017-0608.xlsm";
        Dictionary d2 = dc(n2);
        send2clipBdWin10(n2 + Constants.vbCrLf + dumpKeyedText(d1, d2));
        Debugger.Break();
    }
    // C:\Users\athompson\Documents\dvl\libExt-rcvr-2017-0608.xlsm
    // C:\Users\athompson\Documents\dvl\libExt.xlsm

    public string txOfVbModule(VBIDE.CodeModule cm)
    {
        if (cm == null)
            return "";
        if (cm.CountOfLines > 0)
            return cm.Lines(1, cm.CountOfLines);
        return "";
    }

    public Dictionary lVCXg1f1(VBIDE.VBProject pj)
    {
        // lVCXg1f1 - generate Dictionary of
        // Dictionaries of collected text
        // of all procedures in each module
        // of given VBProject, keyed first
        // by module, and then by procedure
        // 

        var rt = new Dictionary();

        {
            var withBlock = dcOfVbModules(pj);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcOfVbProcs(obVbCodeMod(withBlock.get_Item(ky))));
        }

        return rt;
    }

    public Dictionary lVCXg1f2(VBIDE.VBProject pj)
    {
        // lVCXg1f2 - generate Dictionary of
        // procedures in given VBProject
        // keyed first by procedure name
        // and then by module name
        // 
        // this is accomplished using
        // function dcOfDcRekeyedSecToPri
        // to promote function names over
        // module names, with the expected
        // result being a Dictionary of
        // mostly single-entry Dictionaries.
        // 
        // each of these can then be replaced
        // with the text of its one entry
        // in a subsequent function
        // 
        // multi-entry Dictionaries might
        // have to be left as is
        // 
        // note that with all headers filed
        // under a blank key, at least one
        // multi-entry Dictionary is guaranteed
        // 
        return dcOfDcRekeyedSecToPri(lVCXg1f1(pj));
    }

    public Dictionary lVCXg1f3(Dictionary dc)
    {
        // lVCXg1f3 - return transformation
        // of supplied Dictionary of sort
        // returned by lVCXg1f2 as described
        // in that procedure's comments
        // 
        // each single-entry Dictionary is
        // replaced with the text of its entry
        // while multi-entry Dictionaries
        // are simply copied over
        // 
        // note that this function accepts
        // a Dictionary and NOT a VBProject
        // as lVCXg1f2 does. this permits other
        // functions to be applied to the result
        // of a single call to lVCXg1f2
        // and so reduce redundancy
        // 

        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                // If Len(ky) Then 'to filter out header
                Dictionary wk = dcOb(dc.get_Item(ky));
                if (wk == null)
                {
                    if (VarType(dc.get_Item(ky)) == Constants.vbString)
                        // just add to Dictionary as itself
                        rt.Add(ky, dc.get_Item(ky));
                    else
                        Debugger.Break();// problem!
                }
                else
                {
                    if (wk.Count > 1)
                        rt.Add(ky, wk);
                    else if (wk.Count < 1)
                        Debugger.Break(); // another problem!
                    else
                        rt.Add(ky, wk.Items(0));
                }
            }
        }

        return rt;
    }

    public string lVCXg2f1(string tx)
    {
        // lVCXg2f1 - "decontinuate" VB text
        // replace " _" and newline
        // at end of each continued
        // line with a vertical tab
        // 
        // thus reducing each continued
        // line sequence to a single line
        // while retaining a clear marker
        // for rebreaking, if necessary
        // 
        // lVCXg2f1 = Join(Split(tx," _" & vbCrLf)," _" & vbVerticalTab)
        return Replace(tx, " _" + Constants.vbCrLf, " _" + Constants.vbVerticalTab);
    }

    public string lVCXg2f2(string tx)
    {
        // lVCXg2f2 - "decomment" and "dequote" VB text
        // locate and remove any and all
        // remarks and string constants
        // from a line of VB text
        // '
        // a tricky operation, requiring detection
        // of the FIRST of either single or double
        // quotes (' or ") on a line.
        // '
        // subsequent procedure depends on WHICH
        // is discovered first. for a single quote,
        // the entire remainder of the line should
        // be dropped.
        // '
        // for a double quote, only the text prior
        // to the NEXT double quote (which MUST be
        // present) is dropped. the remainder must
        // then be searched for further reductions
        // '
        // REV[2023.04.20.1001]: recalling that TWO
        // double quotes INSIDE a string constant
        // form an "escape" sequence representing
        // ONE double quote, it is necessary to
        // replace all such instances with another
        // placeholder, in order to ensure the
        // correct closing quote is found. active
        // implementation has been modified
        // to achieve this
        // 
        return lVCXg2f2b(tx);
    }

    public string lVCXg2f2a(string tx)
    {
        // lVCXg2f2a - "decomment" and "dequote" VB text
        // initial implementation deactivated
        // to be held in reserve pending
        // verification of rewritten
        // version lVCXg2f2b
        // 
        string rt;
        long rmk; // location of first 'Rem'


        rmk = InStr(1, tx, "rem", vbTextCompare);


        if (rmk)
        {
            Debug.Print(tx);
            Debugger.Break(); // to review
        }

        long qt1 = InStr(1, tx, "'"); // location of first single quote (rem)
        long qt2 = InStr(1, tx, "\""); // location of first double quote (string)

        Debug.Print(tx);
        if (qt1 * qt2)
        {
            Debugger.Break();
            if (qt1 < qt2)
                rt = Left(tx, qt1);
            else
                Debugger.Break();
        }
        else if (qt1)
        {
            rt = Left(tx, qt1);
            Debug.Print(rt);
            Debugger.Break();
        }
        else if (qt2)
        {
            string[] ar = Split(tx, "\"", 3);
            rt = lVCXg2f2a(ar[0] + "$$" + ar[2]);
            Debug.Print(rt);
            Debugger.Break();
        }
        else
            rt = tx;

        return rt;
    }

    public string lVCXg2f2b(string tx)
    {
        // lVCXg2f2b - "decomment" and "dequote" VB text
        // currently active implementation @[2023.04.20.0957]
        // 
        string rt;
        var qt = new long[3]; // locations of first single and double quote
        long rmk; // location of first 'Rem'

        // Debug.Print "IN: "; tx 'while debugging only
        var rf = "'\"";
        long mx = Strings.Len(tx);
        for (long dx = 1; dx <= 2; dx++)
        {
            qt[dx] = InStr(1, tx, Strings.Mid(rf, dx, 1));
            if (qt[dx] == 0)
                qt[dx] = 1 + mx;
        }


        rmk = InStr(1, " " & tx, " rem ", vbTextCompare);


        if (rmk == 0)
            rmk = Strings.Len(tx) + 2;
        if (rmk < qt[1])
        {
            if (rmk < qt[2])
            {
                // If Mid$(tx, rmk + 3, 1) <= " " Then
                Debug.Print(tx);
                Debug.Print(Left(tx, rmk - 1));
                Debug.Print(Mid(tx, rmk));
                Debugger.Break(); // to review
            }
        }


        if (qt[1] < qt[2])
            rt = Left(tx, qt[1]);
        else if (qt[2] > mx)
            rt = tx;
        else
        {
            string[] ar = Split(tx, "\"", 2); // was 3

            // rt = lVCXg2f2b(ar(0) & "$$" & ar(2))
            // rt = ar(0) & """""" & lVCXg2f2b(ar(2))
            rt = ar[0] + "\"\"";

            // ar = Split(Join(Split(ar(1), """"""), vbFormFeed), """", 2)
            // REF: Replace("expr","find","rplc")
            ar = Split(Replace(ar[1], "\"\"", Constants.vbFormFeed), "\"", 2);

            if (UBound(ar) > 0)
                rt = rt + lVCXg2f2b(Replace(ar[1], Constants.vbFormFeed, "\"\""));
            else
                Debugger.Break();// problem!
        }

        return rt;
    }

    public Dictionary lVCXg2f3(string tx)
    {
        // lVCXg2f3 - decompose VB text to a "keyword" Dictionary
        // mapping each "keyword" to a count of instances
        // note that "keyword" includes not only words
        // reserved by VB, but any unbroken set of non-space
        // characters: variables, procedure names, etc.
        // '
        // the Keys returned in the resulting Dictionary
        // can then be matched against another Dictionary
        // listing all entities defined in a VB project,
        // as returned by other functions in this module,
        // thereby producting a first-level dependency map
        // '
        // note that the text supplied IS assumed to be
        // Visual Basic code, which should already be
        // "cleaned" of any "inactive" elements: comments
        // and the content of string literals. this should
        // limit the rate of "false positives," that is,
        // detection of entity names not actually required
        // by a procedure, but mentioned either in string
        // literals or comments the compiler does not parse.
        // '

        var rt = new Dictionary();

        var wk = tx;
        var ls = Join(new [] {Constants.vbCrLf, Constants.vbTab, "():&.!,[]"}, "");
        do
        {
            wk = Replace(wk, Left(ls, 1), " ");
            ls = Mid(ls, 2);
        }
        while (Strings.Len(ls));

        {
            var withBlock = rt;
            foreach (var ky in Split(wk, " "))
            {
                if (!Len(ky)) continue;
                if (withBlock.Exists(ky))
                    withBlock.get_Item(ky) = 1 + withBlock.get_Item(ky);
                else
                    withBlock.Add(ky, 1);
            }
        }

        return rt;
    }

    public string lVCXg3f1(string tx)
    {
        // lVCXg3f1 - "clean" supplied VB text
        // removing all comments and content
        // of string constants, and leaving
        // only comment markers and null
        // strings in their place
        // 
        // goal of this function is to remove
        // any "inactive" content from text
        // of VB procedure definitions, and
        // thereby limit the number of "false
        // positives" returned in a search
        // for procedural dependencies
        // 

        string[] wk = Split(Replace(tx, " _" + Constants.vbCrLf, " _" + Constants.vbVerticalTab), Constants.vbCrLf);

        long mx = UBound(wk);
        for (long dx = LBound(wk); dx <= mx; dx++)
            wk[dx] = lVCXg2f2(wk[dx]);

        return Replace(Strings.Join(wk, Constants.vbCrLf), " _" + Constants.vbVerticalTab, " _" + Constants.vbCrLf);
    }

    public Dictionary lVCXg3f2(Dictionary dc)
    {
        // lVCXg3f2 - "clean" all String Items
        // in supplied Dictionary, including
        // any sub Dictionaries, as VB text
        // 
        // any Items not recognized as String
        // or Dictionary Items are passed
        // through as is, at present
        // 
        // might want to reconsider this,
        // and probably make it optional
        // 

        // tx As String
        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                var ar = new[] { dc.get_Item(ky) };
                if (VarType(ar[0]) == Constants.vbString)
                    rt.Add(ky, lVCXg3f1(Convert.ToHexString(ar[0])));
                else if (ar[0] is Dictionary)
                    rt.Add(ky, lVCXg3f2(dcOb(ar[0])));
                else
                    rt.Add(ky, ar[0]);
            }
        }

        return rt;
    }

    public Dictionary lVCXg3f3(Dictionary dc)
    {
        // lVCXg3f3 - "clean" all String Items
        // in supplied Dictionary, including
        // any sub Dictionaries, as VB text
        // 
        // any Items not recognized as String
        // or Dictionary Items are passed
        // through as is, at present
        // 
        // might want to reconsider this,
        // and probably make it optional
        // 

        // tx As String
        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                dynamic[] ar = [dc.get_Item(ky)];
                if (VarType(ar[0]) == Constants.vbString)
                    rt.Add(ky, lVCXg3f1(Convert.ToHexString(ar[0])));
                else if (ar[0] is Dictionary)
                    rt.Add(ky, lVCXg3f3(dcOb(ar[0])));
                else
                    rt.Add(ky, ar[0]);
            }
        }

        return rt;
    }

    public Dictionary lVCXg3f4(Dictionary dc)
    {
        // lVCXg3f4 - generate "keyword" lists for all
        // String Items in supplied Dictionary,
        // including any sub Dictionaries
        // '
        // derived from lVCXg3f3, this function just
        // adds one level of processing, calling
        // lVCXg2f3 against the results of lVCXg3f1
        // '
        // it does NOT require input from lVCXg3f3
        // and would likely fail against such a source
        // '
        // any Items not recognized as String
        // or Dictionary Items are passed
        // through as is, at present
        // '
        // might want to reconsider this,
        // and probably make it optional
        // 

        // tx As String
        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                var ar = new[] { dc.get_Item(ky) };
                if (VarType(ar[0]) == Constants.vbString)
                    rt.Add(ky, lVCXg2f3(lVCXg3f1(Convert.ToHexString(ar[0]))));
                else if (ar[0] is Dictionary)
                    rt.Add(ky, lVCXg3f4(dcOb(ar[0])));
                else
                    rt.Add(ky, ar[0]);
            }
        }

        return rt;
    }

    public Dictionary lVCXg4f1(Dictionary dc, Dictionary rf = null, Dictionary rt = null)
    {
        // lVCXg4f1 - generate basic dependency list from
        // supplied Dictionary of "keyword" Dictionaries
        // keyed either to VB procedure names, or for
        // a subset of identically named procedures,
        // the module names of each implementation
        // '
        // note that this function does NOT disambiguate
        // dependencies on multiply defined names.
        // that is a task left to the client supplying
        // the initial Dictionary, on the assumption
        // that said client will have retained any
        // prior source used to generate it
        // '
        // note also that optional Dictionary parameters
        // rf and rt are NOT expected to be provided by
        // an outside client, but passed to a recursive
        // invocation when processing a multiply defined
        // procedure name. for this purpose, the original
        // source Dictionary dc is passed through rf to
        // ensure its availability to all recursive calls
        // '

        if (rf == null)
            rt = lVCXg4f1(dc, dc, rt);
        else if (rt == null)
            rt = lVCXg4f1(dc, rf, new Dictionary());
        else
        {
            foreach (var ky in dc.Keys)
            {
                Dictionary wk = dc.get_Item(ky);
                {
                    if (wk.Count > 0)
                    {
                        dynamic[] ar = [wk.get_Item(wk.Keys(0))];
                        // wk, NOT ar(0)
                        if (IsNumeric(ar[0]))
                            // intersect with reference Dictionary rf
                            // taking only the usage counts from wf
                            rt.Add(ky, dcKeysInCommon(wk, rf, 1));
                        else if (ar[0] is Dictionary)
                            // we have a subset of implementations
                            // keyed to module locations
                            // so need to go down a level
                            rt.Add(ky, lVCXg4f1(wk, rf, new Dictionary()));
                        else if (VarType(ar[0]) == Constants.vbString)
                        {
                            Debug.Print(""); // Breakpoint Landing
                            Debugger.Break(); // because this does NOT normally happen
                        }
                        else
                        {
                            Debugger.Break(); // because this shouldn't happen either
                            Debug.Print(""); // Breakpoint Landing
                        }
                    }
                }

                Debug.Print(""); // Breakpoint Landing
            }
        }

        return rt;
    }

    public Dictionary dxOfVbProcLocsInMod(VBIDE.CodeModule cm)
    {
        // dxOfVbProcLocsInMod -- Return Dictionary
        // of Procedure Locations in CodeModule.
        // derived from dcOfVbProcs as a "light"
        // alternative that only returns an "index"
        // of procedures, leaving the client to
        // extract their text as needed
        // 
        // REV[2023.05.05.1307]: copied from dcOfVbProcs
        // see that function for prior REVs
        // if and where applicable
        // 
        // Dim tx As String

        var rt = new Dictionary();
        dynamic ar = new [] {new dynamic[] {vbext_pk_Proc, ""}, new dynamic[] {vbext_pk_Get, ""}, new dynamic[] {vbext_pk_Let, "=#"}, new dynamic[] {vbext_pk_Set, "=@"}};

        {
            var withBlock = cm;
            long mx = withBlock.CountOfLines;
            long dx = withBlock.CountOfDeclarationLines;
            // REV[2023.05.05.1329]: modified
            // dx assignment for reuse below
            // and call .CountOfDeclarationLines
            // only once.
            rt.Add("", new [] {1, dx}); // REV[2023.05.05.1310]: replaced
            // .Lines with Array to capture start
            // line and line count of header
            // (AKA DeclarationLines)

            dx = 1 + dx;
            while (dx < mx)
            {
                var fw = dx;
                while (fw < mx)
                {
                    string ck = withBlock.ProcOfLine(fw, vbext_pk_Proc);
                    if (Strings.Len(ck) == 0)
                        fw = fw + 1;
                    else
                    {
                        long tp = 0;

                        do
                        {
                            Information.Err().Clear();
                            dx = withBlock.ProcStartLine(ck, ar(tp)(0));
                            if (Information.Err().Number)
                                dx = fw + 1;
                            if (dx != fw)
                                tp = tp + 1;
                            if (tp > 3)
                                Debugger.Break();
                        }
                        while (!Information.Err().Number == 0 & dx == fw) // should NOT happen...
                            ;
                        Information.Err().Clear();


                        fw = withBlock.ProcCountLines(ck, ar(tp)(0));
                        // tx = .Lines(dx, fw)
                        rt.Add(ck + ar(tp)(1), new [] {dx, fw}); // REV[2023.05.05.1337]: replaced tx
                        // .Lines with new [] {dx, fw) to capture
                        // start line and line count of procedure
                        dx = dx + fw;
                        fw = mx;
                    }
                }
            }
        }
        return rt;
    }

    public string vbProcTextFromPrj(string nm, VBIDE.VBProject pj = null)
    {
        // vbProcTextFromPrj
        // derived from vbTextOfProcInProject (sort of)
        // 
        // NOTE: To use this Function
        // from an external library,
        // the option has been removed
        // to call itself recursively
        // against ThisWorkbook. Since
        // ThisWorkbook would be the
        // library itself, a call against
        // it could result in a breach
        // of security.
        // 
        string ky;

        if (pj == null)
            ky = "";
        else
        {
            var withBlock = dxOfVbProcLocsInPrj(pj);
            if (withBlock.Exists(nm))
            {
                dynamic ar = new [] {null};

                Dictionary dc = withBlock.get_Item(nm);
                {
                    var withBlock1 = dc;
                    if (withBlock1.Count > 0)
                    {
                        ky = withBlock1.Keys(0);

                        if (withBlock1.Count > 1)
                            ky = userChoiceFromDc(dc, ky);

                        if (withBlock1.Exists(ky))
                            ar = withBlock1.get_Item(ky);
                    }
                }

                ky = "";
                if (UBound(ar) < 2) return ky;
                VBIDE.CodeModule cm = obVbCodeMod(obOf(ar(0)));
                if (!cm == null)
                    ky = cm.Lines(ar(1), ar(2));
            }
            else
                ky = "";
        }

        return ky;
    }

    public Dictionary dxOfVbProcLocsInPrj(VBIDE.VBProject pj)
    {
        // dxOfVbProcLocsInPrj - generate Dictionary of
        // Dictionaries of collected text
        // of all procedures in each module
        // of given VBProject, keyed first
        // by module, and then by procedure
        // 

        // Dim nm As String
        var rt = new Dictionary();

        {
            var withBlock = dcOfVbModules(pj);
            foreach (var kMd in withBlock.Keys)
            {
                VBIDE.CodeModule cm = obVbCodeMod(withBlock.get_Item(kMd));

                {
                    var withBlock1 = dxOfVbProcLocsInMod(cm);
                    foreach (var kPr in withBlock1.Keys)
                    {
                        Dictionary wk;
                        {
                            if (!rt.Exists(kPr))
                                rt.Add(kPr, new Dictionary());

                            wk = rt.get_Item(kPr);
                        }

                        var ar = withBlock1.get_Item(kPr);
                        {
                            var withBlock2 = wk;
                            if (withBlock2.Exists(kMd))
                                Debugger.Break(); // for problem
                            else
                                withBlock2.Add(kMd, new[] { cm, ar(0), ar(1) });
                        }
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary dcOfVbProcs(VBIDE.CodeModule cm)
    {
        // dcOfVbProcs -- Return Dictionary of Procedures
        // -- from supplied CodeModule
        // 
        // REV[2023.04.19.1146]: added code to capture
        // declaration text preceding all proc defs
        // REV[2023.02.15.0904]: modified to accommodate ,
        // , and Procedures. and Procedures
        // are stored under keys modified to indicate their
        // role: "=#" for indicates assignment to a value,
        // while "=@" for indicates an dynamic assignment.
        // NOTE: this new version, while now able to accommodate
        // Class Modules, is likely not the most efficient
        // in addressing the problem. Further development
        // might be warranted, should this prove an issue.
        // 

        var rt = new Dictionary();
        dynamic ar = new [] {new [] {vbext_pk_Proc, ""}, new [] {vbext_pk_Get, ""}, new [] {vbext_pk_Let, "=#"}, new [] {vbext_pk_Set, "=@"}};

        {
            var withBlock = cm;
            long mx = withBlock.CountOfLines;
            long dx = 1 + withBlock.CountOfDeclarationLines;
            // REV[2023.04.19.1146]: added following
            // to capture header, AKA declaration lines
            rt.Add("", withBlock.Lines(1, withBlock.CountOfDeclarationLines));
            while (dx < mx)
            {
                var fw = dx;
                while (fw < mx)
                {
                    string ck = withBlock.ProcOfLine(fw, vbext_pk_Proc);
                    if (Strings.Len(ck) > 0)
                    {
                        long tp = 0;

                        do
                        {
                            Information.Err().Clear();
                            dx = withBlock.ProcStartLine(ck, ar(tp)(0));
                            if (Information.Err().Number)
                                dx = fw + 1;
                            if (dx != fw)
                                tp = tp + 1;
                            if (tp > 3)
                                Debugger.Break();
                        }
                        while (!Information.Err().Number == 0 & dx == fw) // should NOT happen...
                            ;
                        Information.Err().Clear();


                        fw = withBlock.ProcCountLines(ck, ar(tp)(0));
                        string tx = withBlock.Lines(dx, fw);
                        rt.Add(ck + ar(tp)(1), tx);
                        dx = dx + fw;
                        fw = mx;
                    }
                    else
                        fw = fw + 1;
                }
            }
        }
        return rt;
    }

    public Dictionary dcOfVbProcs_obs2023_0419(VBIDE.CodeModule cm)
    {
        // dcOfVbProcs_obs2023_0419 -- Return Dictionary of Procedures
        // -- from supplied CodeModule
        // 
        // NOTE: This function ONLY looks for general Procedures.
        // It does NOT look for , , or Procedures.
        // It MIGHT NOT WORK properly against Class Modules!
        // 

        var rt = new Dictionary();
        {
            if (cm.Parent.Type != vbext_ct_StdModule) return rt;
            long mx = cm.CountOfLines;
            long dx = 1 + cm.CountOfDeclarationLines;
            // Debug.Print .Lines(1, .CountOfDeclarationLines) & "'''"

            while (dx < mx)
            {
                var fw = dx;
                while (fw < mx)
                {
                    string ck = cm.ProcOfLine(fw, vbext_pk_Proc);
                    if (Strings.Len(ck) > 0)
                    {
                        dx = cm.ProcStartLine(ck, vbext_pk_Proc);
                        fw = cm.ProcCountLines(ck, vbext_pk_Proc);
                        string tx = cm.Lines(dx, fw);
                        rt.Add(ck, tx);
                        dx = dx + fw;
                        fw = mx;
                    }
                    else
                        fw = fw + 1;
                }
            }
        }
        return rt;
    }

    public Dictionary dcOfVbModules(VBIDE.VBProject vb)
    {
        var rt = new Dictionary();
        {
            if (vb.Protection == vbext_pp_none)
            {
                foreach (VBIDE.VBComponent vc in vb.VBComponents)
                {
                    {
                        var withBlock1 = vc;
                        rt.Add.Name(null, withBlock1.CodeModule);
                    }
                }
            }
            else
                rt.Add("<PROTECTED>", new Dictionary());
        }
        return rt;
    }

    public Dictionary dcOfVbProcsFlat(VBIDE.VBProject pj)
    {
        // dcOfVbProcsFlat - generate Dictionary
        // of collected text of all procedures
        // in each module of given VBProject,
        // keyed by procedure name, or by
        // combination of module and procedure
        // name when more than one procedure
        // of same name is found
        // 
        // NOTE[2023.04.19.1256] the compromise
        // noted above is NOT ideal.
        // As the purpose of this function
        // is to produce a FLAT list
        // of procedure names for quick
        // searching purposes, the need
        // to modify duplicate names
        // for is likely to make it
        // difficult or impractical
        // to find all possible matches
        // 
        // 
        Dictionary dc;

        var rt = new Dictionary();

        {
            var withBlock = dcOfVbModules(pj);
            foreach (var kyMd in withBlock.Keys)
            {
                {
                    var withBlock1 = dcOfVbProcs(obOf(withBlock.get_Item(kyMd))) // dcOb
                        ;
                    foreach (var kyPr in withBlock1.Keys)
                    {
                        if (rt.Exists(kyPr))
                        {
                            Debug.Print(""); // breakpoint landing
                            // Stop
                            // going to need a better way
                            // to handle this situation
                            // but for now...
                            rt.Add(kyMd + "." + kyPr, withBlock1.get_Item(kyPr));
                        }
                        else
                            rt.Add(kyPr, withBlock1.get_Item(kyPr));
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary dcOfVbProjects(VBIDE.VBProjects pjs)
    {
        var rt = new Dictionary();
        {
            var withBlock = pjs;
            foreach (VBIDE.VBProject vb in pjs)
                rt.Add(vb.Filename, dcOfVbModules(vb));
        }
        return rt;
    }

    public Dictionary dcTxOfVbModule(Dictionary dc)
    {
        var rt = new Dictionary();

        if (dc == null)
        {
        }
        else
        {
            if (dc.Count <= 0) return rt;
            foreach (var ky in dc.Keys)
                rt.Add(ky, txOfVbModule(obVbCodeMod(dc.get_Item(ky))));
        }

        return rt;
    }

    public Dictionary dcTxOfVbProjMods(Dictionary dc)
    {
        var rt = new Dictionary();

        {
            if (dc.Count <= 0) return rt;
            foreach (var ky in dc.Keys)
                rt.Add(ky, dcTxOfVbModule(dcOb(dc.get_Item(ky))));
        }

        return rt;
    }

    public string vbTextOfProcInDict(string nm, Dictionary dc)
    {
        // vbTextOfProcInDict -- Retrieve text from Dictionary
        // 
        // This Function's name is probably
        // WAY unnecessarily specific.
        // 
        // The Function itself simply returns the String
        // found under the supplied key variable 'nm',
        // or an null String if none is found. This is
        // a fairly general type of Function, one which
        // could be named far more generically.
        // 
        // dcItemIfPresent looks a likely candidate,
        // although it might be a bit TOO general...
        // 
        return dc == null ?
            // Recursive call option removed.
            // See text of vbTextOfProcInProject
            // for details on security issue.
            // 
            "" : // vbTextOfProcInDict(nm, dcOfVbProcsFlat(ThisWorkbook.VBProject))
            Convert.ToHexString(dcItemIfPresent(dc, nm, Constants.vbString));
    }

    public string vbTextOfProcIn(string nm, VBIDE.CodeModule cm)
    {
        return vbTextOfProcInDict(nm, dcOfVbProcs(cm));
    }

    public string vbTextOfProcInProject(string nm, VBIDE.VBProject pj)
    {
        // vbTextOfProcInProject
        // 
        // NOTE: In order to use this Function
        // from an external library,
        // the option has been removed
        // to call itself recursively
        // against ThisWorkbook. Since
        // ThisWorkbook would be the
        // library itself, a call against
        // it could result in a breach
        // of security.
        // 
        return pj == null ? "" : vbTextOfProcInDict(nm, dcOfVbProcsFlat(pj));
    }

    public dynamic send2clipBd(dynamic src)
    {
        {
            
        }
        string ck = send2clipBdWin10(src);

        // With New MSForms.DataObject
        // .SetText src
        // .PutInClipboard
        // 
        // .GetFromClipboard
        // 
        // Do
        // Err.Clear
        // ck = .GetText
        // If Err.Number Then
        // If MessageBox.Show(' Join(new [] {' "Error Getting Text from DataObject.",' "A simple retry will usually succeed.",' "", "Go ahead and retry?"' ), vbCrLf),' vbYesNo, "Retry GetText?"' ) = vbNo Then
        // Err.Clear
        // End If
        // End If
        // Loop Until Err.Number = 0
        // 
        // End With
        if (ck == src)
        {
        }
        else
            Debugger.Break();

        return src;
    }

    public dynamic getFromClipBd(dynamic fmt = 1)
    {
        // ' 1 is the value of CF_TEXT, one of the clipboard format
        // ' enums which SHOULD be defined, but apparently aren't.
        // ' That is the effective default format used by GetText,
        // ' if none is given
        dynamic rt;
        {
            var withBlock = new DataObject();
            withBlock.GetFromClipboard();
            rt = withBlock.GetText(fmt);
        }
        return rt;
    }

    public string dumpKeyedText(Dictionary d1, Dictionary d2)
    {
        // ' Extract values from second dictionary
        // ' filed under keys from FIRST dictionary.
        // ' Theory is, the keys will always be
        // ' retrieved in the same order, as long as
        // ' no changes have been made between runs.
        // '
        // ' By supplying the same dictionary for
        // ' both d1 and d2, that dictionary's
        // ' content can be extracted, and then
        // ' a different d2's content can be
        // ' extracted in the same order.

        var rt = new Dictionary();
        foreach (var ky in d1.Keys)
            rt.Add("{" + ky + "}" + Constants.vbCrLf + d2(ky), 1);
        return Join(rt.Keys, Constants.vbCrLf);
    }
}