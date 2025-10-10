class SurroundingClass
{
    public long xprtModText()
    {
        VBIDE.VBProject vp;
        VBIDE.VBComponent vc;
        Scripting.Dictionary dc;
        Scripting.Dictionary d1;
        Scripting.Dictionary d2;
        string n1;
        string n2;

        dc = dcOfVbProjects(ThisApplication.VBAProjects); // ThisWorkbook.Application.VBE.VBProjects

        n1 = @"C:\Users\athompson\Documents\dvl\libExt.xlsm";
        d1 = dc(n1);
        send2clipBdWin10(n1 + Constants.vbNewLine + dumpKeyedText(d1, d1));
        System.Diagnostics.Debugger.Break();

        n2 = @"C:\Users\athompson\Documents\dvl\libExt-rcvr-2017-0608.xlsm";
        d2 = dc(n2);
        send2clipBdWin10(n2 + Constants.vbNewLine + dumpKeyedText(d1, d2));
        System.Diagnostics.Debugger.Break();
    }
    // C:\Users\athompson\Documents\dvl\libExt-rcvr-2017-0608.xlsm
    // C:\Users\athompson\Documents\dvl\libExt.xlsm

    public string txOfVbModule(VBIDE.CodeModule cm)
    {
        if (cm == null)
            txOfVbModule = "";
        else
        {
            var withBlock = cm;
            if (withBlock.CountOfLines > 0)
                txOfVbModule = withBlock.Lines(1, withBlock.CountOfLines);
            else
                txOfVbModule = "";
        }
    }

    public Scripting.Dictionary lVCXg1f1(VBIDE.VBProject pj)
    {
        /// lVCXg1f1 - generate Dictionary of
        /// Dictionaries of collected text
        /// of all procedures in each module
        /// of given VBProject, keyed first
        /// by module, and then by procedure
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcOfVbModules(pj);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcOfVbProcs(obVbCodeMod(withBlock.Item(ky))));
        }

        lVCXg1f1 = rt;
    }

    public Scripting.Dictionary lVCXg1f2(VBIDE.VBProject pj)
    {
        /// lVCXg1f2 - generate Dictionary of
        /// procedures in given VBProject
        /// keyed first by procedure name
        /// and then by module name
        /// 
        /// this is accomplished using
        /// function dcOfDcRekeyedSecToPri
        /// to promote function names over
        /// module names, with the expected
        /// result being a Dictionary of
        /// mostly single-entry Dictionaries.
        /// 
        /// each of these can then be replaced
        /// with the text of its one entry
        /// in a subsequent function
        /// 
        /// multi-entry Dictionaries might
        /// have to be left as is
        /// 
        /// note that with all headers filed
        /// under a blank key, at least one
        /// multi-entry Dictionary is guaranteed
        /// 
        lVCXg1f2 = dcOfDcRekeyedSecToPri(lVCXg1f1(pj));
    }

    public Scripting.Dictionary lVCXg1f3(Scripting.Dictionary dc)
    {
        /// lVCXg1f3 - return transformation
        /// of supplied Dictionary of sort
        /// returned by lVCXg1f2 as described
        /// in that procedure's comments
        /// 
        /// each single-entry Dictionary is
        /// replaced with the text of its entry
        /// while multi-entry Dictionaries
        /// are simply copied over
        /// 
        /// note that this function accepts
        /// a Dictionary and NOT a VBProject
        /// as lVCXg1f2 does. this permits other
        /// functions to be applied to the result
        /// of a single call to lVCXg1f2
        /// and so reduce redundancy
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc // lVCXg1f3(pj)
       ;
            foreach (var ky in withBlock.Keys)
            {
                // If Len(ky) Then 'to filter out header

                wk = dcOb(withBlock.Item(ky));
                if (wk == null)
                {
                    if (VarType(withBlock.Item(ky)) == Constants.vbString)
                        // just add to Dictionary as itself
                        rt.Add(ky, withBlock.Item(ky));
                    else
                        System.Diagnostics.Debugger.Break();// problem!
                }
                else
                {
                    var withBlock1 = wk;
                    if (withBlock1.Count > 1)
                        rt.Add(ky, wk);
                    else if (withBlock1.Count < 1)
                        System.Diagnostics.Debugger.Break(); // another problem!
                    else
                        rt.Add(ky, withBlock1.Items(0));
                }
            }
        }

        lVCXg1f3 = rt;
    }

    public string lVCXg2f1(string tx)
    {
        /// lVCXg2f1 - "decontinuate" VB text
        /// replace " _" and newline
        /// at end of each continued
        /// line with a vertical tab
        /// 
        /// thus reducing each continued
        /// line sequence to a single line
        /// while retaining a clear marker
        /// for rebreaking, if necessary
        /// 
        // lVCXg2f1 = Join(Split(tx," _" & vbNewLine)," _" & vbVerticalTab)
        lVCXg2f1 = Replace(tx, " _" + Constants.vbNewLine, " _" + Constants.vbVerticalTab);
    }

    public string lVCXg2f2(string tx)
    {
        /// lVCXg2f2 - "decomment" and "dequote" VB text
        /// locate and remove any and all
        /// remarks and string constants
        /// from a line of VB text
        /// '
        /// a tricky operation, requiring detection
        /// of the FIRST of either single or double
        /// quotes (' or ") on a line.
        /// '
        /// subsequent procedure depends on WHICH
        /// is discovered first. for a single quote,
        /// the entire remainder of the line should
        /// be dropped.
        /// '
        /// for a double quote, only the text prior
        /// to the NEXT double quote (which MUST be
        /// present) is dropped. the remainder must
        /// then be searched for further reductions
        /// '
        /// REV[2023.04.20.1001]: recalling that TWO
        /// double quotes INSIDE a string constant
        /// form an "escape" sequence representing
        /// ONE double quote, it is necessary to
        /// replace all such instances with another
        /// placeholder, in order to ensure the
        /// correct closing quote is found. active
        /// implementation has been modified
        /// to achieve this
        /// 
        lVCXg2f2 = lVCXg2f2b(tx);
    }

    public string lVCXg2f2a(string tx)
    {
        /// lVCXg2f2a - "decomment" and "dequote" VB text
        /// initial implementation deactivated
        /// to be held in reserve pending
        /// verification of rewritten
        /// version lVCXg2f2b
        /// 
        string rt;
        string[] ar;
        long qt1; // location of first single quote (rem)
        long qt2; // location of first double quote (string)
        long rmk; // location of first 'Rem'
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentOutOfRangeException: Index was out of range. Must be non-negative and less than the size of the collection. (Parameter 'index')
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitSimpleArgument(SimpleArgumentSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1069
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.<>c__DisplayClass83_0.<ConvertArguments>b__0(ArgumentSyntax a, Int32 i) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1045
   at System.Linq.Enumerable.SelectIterator[TSource,TResult](IEnumerable`1 source, Func`3 selector)+MoveNext()
   at System.Linq.Enumerable.WhereEnumerableIterator`1.MoveNext()
   at Microsoft.CodeAnalysis.CSharp.SyntaxFactory.SeparatedList[TNode](IEnumerable`1 nodes)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitArgumentList(ArgumentListSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1022
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitInvocationExpression(InvocationExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1431
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    
    rmk = InStr(1, tx, "rem", vbTextCompare)

 */
        if (rmk)
        {
            Debug.Print(tx);
            System.Diagnostics.Debugger.Break(); // to review
        }

        qt1 = InStr(1, tx, "'");
        qt2 = InStr(1, tx, "\"");

        Debug.Print(tx);
        if (qt1 * qt2)
        {
            System.Diagnostics.Debugger.Break();
            if (qt1 < qt2)
                rt = Left(tx, qt1);
            else
                System.Diagnostics.Debugger.Break();
        }
        else if (qt1)
        {
            rt = Left(tx, qt1);
            Debug.Print(rt);
            System.Diagnostics.Debugger.Break();
        }
        else if (qt2)
        {
            ar = Split(tx, "\"", 3);
            rt = lVCXg2f2a(ar[0] + "$$" + ar[2]);
            Debug.Print(rt);
            System.Diagnostics.Debugger.Break();
        }
        else
            rt = tx;

        lVCXg2f2a = rt;
    }

    public string lVCXg2f2b(string tx)
    {
        /// lVCXg2f2b - "decomment" and "dequote" VB text
        /// currently active implementation @[2023.04.20.0957]
        /// 
        string rt;
        string rf;
        string[] ar;
        long[] qt = new long[3]; // locations of first single and double quote
        long mx;
        long dx;
        long rmk; // location of first 'Rem'

        // Debug.Print "IN: "; tx 'while debugging only

        rf = "'\"";
        mx = Strings.Len(tx);
        for (dx = 1; dx <= 2; dx++)
        {
            qt[dx] = InStr(1, tx, Strings.Mid(rf, dx, 1));
            if (qt[dx] == 0)
                qt[dx] = 1 + mx;
        }
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentOutOfRangeException: Index was out of range. Must be non-negative and less than the size of the collection. (Parameter 'index')
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitSimpleArgument(SimpleArgumentSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1069
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.<>c__DisplayClass83_0.<ConvertArguments>b__0(ArgumentSyntax a, Int32 i) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1045
   at System.Linq.Enumerable.SelectIterator[TSource,TResult](IEnumerable`1 source, Func`3 selector)+MoveNext()
   at System.Linq.Enumerable.WhereEnumerableIterator`1.MoveNext()
   at Microsoft.CodeAnalysis.CSharp.SyntaxFactory.SeparatedList[TNode](IEnumerable`1 nodes)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitArgumentList(ArgumentListSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1022
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitInvocationExpression(InvocationExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1431
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    
    rmk = InStr(1, " " & tx, " rem ", vbTextCompare)

 */
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
                System.Diagnostics.Debugger.Break(); // to review
            }
        }


        if (qt[1] < qt[2])
            rt = Left(tx, qt[1]);
        else if (qt[2] > mx)
            rt = tx;
        else
        {
            ar = Split(tx, "\"", 2); // was 3

            // rt = lVCXg2f2b(ar(0) & "$$" & ar(2))
            // rt = ar(0) & """""" & lVCXg2f2b(ar(2))
            rt = ar[0] + "\"\"";

            // ar = Split(Join(Split(ar(1), """"""), vbFormFeed), """", 2)
            /// REF: Replace("expr","find","rplc")
            ar = Split(Replace(ar[1], "\"\"", Constants.vbFormFeed), "\"", 2);

            if (UBound(ar) > 0)
                rt = rt + lVCXg2f2b(Replace(ar[1], Constants.vbFormFeed, "\"\""));
            else
                System.Diagnostics.Debugger.Break();// problem!
        }

        lVCXg2f2b = rt;
    }

    public Scripting.Dictionary lVCXg2f3(string tx)
    {
        /// lVCXg2f3 - decompose VB text to a "keyword" Dictionary
        /// mapping each "keyword" to a count of instances
        /// note that "keyword" includes not only words
        /// reserved by VB, but any unbroken set of non-space
        /// characters: variables, procedure names, etc.
        /// '
        /// the Keys returned in the resulting Dictionary
        /// can then be matched against another Dictionary
        /// listing all entities defined in a VB project,
        /// as returned by other functions in this module,
        /// thereby producting a first-level dependency map
        /// '
        /// note that the text supplied IS assumed to be
        /// Visual Basic code, which should already be
        /// "cleaned" of any "inactive" elements: comments
        /// and the content of string literals. this should
        /// limit the rate of "false positives," that is,
        /// detection of entity names not actually required
        /// by a procedure, but mentioned either in string
        /// literals or comments the compiler does not parse.
        /// '
        Scripting.Dictionary rt;
        string wk;
        string ls;
        Variant ky;

        rt = new Scripting.Dictionary();

        wk = tx;
        ls = Join(Array(Constants.vbCrLf, Constants.vbTab, "():&.!,[]"), "");
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
                if (Len(ky))
                {
                    if (withBlock.Exists(ky))
                        withBlock.Item(ky) = 1 + withBlock.Item(ky);
                    else
                        withBlock.Add(ky, 1);
                }
            }
        }

        lVCXg2f3 = rt;
    }

    public string lVCXg3f1(string tx)
    {
        /// lVCXg3f1 - "clean" supplied VB text
        /// removing all comments and content
        /// of string constants, and leaving
        /// only comment markers and empty
        /// strings in their place
        /// 
        /// goal of this function is to remove
        /// any "inactive" content from text
        /// of VB procedure definitions, and
        /// thereby limit the number of "false
        /// positives" returned in a search
        /// for procedural dependencies
        /// 
        string[] wk;
        long mx;
        long dx;

        wk = Split(Replace(tx, " _" + Constants.vbNewLine, " _" + Constants.vbVerticalTab), Constants.vbNewLine);

        mx = UBound(wk);
        for (dx = LBound(wk); dx <= mx; dx++)
            wk[dx] = lVCXg2f2(wk[dx]);

        lVCXg3f1 = Replace(Strings.Join(wk, Constants.vbNewLine), " _" + Constants.vbVerticalTab, " _" + Constants.vbNewLine);
    }

    public Scripting.Dictionary lVCXg3f2(Scripting.Dictionary dc)
    {
        /// lVCXg3f2 - "clean" all String Items
        /// in supplied Dictionary, including
        /// any sub Dictionaries, as VB text
        /// 
        /// any Items not recognized as String
        /// or Dictionary Items are passed
        /// through as is, at present
        /// 
        /// might want to reconsider this,
        /// and probably make it optional
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant[] ar;
        // tx As String

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ar = Array(withBlock.Item(ky));
                if (VarType(ar[0]) == Constants.vbString)
                    rt.Add(ky, lVCXg3f1(System.Convert.ToHexString(ar[0])));
                else if (ar[0] is Scripting.Dictionary)
                    rt.Add(ky, lVCXg3f2(dcOb(ar[0])));
                else
                    rt.Add(ky, ar[0]);
            }
        }

        lVCXg3f2 = rt;
    }

    public Scripting.Dictionary lVCXg3f3(Scripting.Dictionary dc)
    {
        /// lVCXg3f3 - "clean" all String Items
        /// in supplied Dictionary, including
        /// any sub Dictionaries, as VB text
        /// 
        /// any Items not recognized as String
        /// or Dictionary Items are passed
        /// through as is, at present
        /// 
        /// might want to reconsider this,
        /// and probably make it optional
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant[] ar;
        // tx As String

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ar = Array(withBlock.Item(ky));
                if (VarType(ar[0]) == Constants.vbString)
                    rt.Add(ky, lVCXg3f1(System.Convert.ToHexString(ar[0])));
                else if (ar[0] is Scripting.Dictionary)
                    rt.Add(ky, lVCXg3f3(dcOb(ar[0])));
                else
                    rt.Add(ky, ar[0]);
            }
        }

        lVCXg3f3 = rt;
    }

    public Scripting.Dictionary lVCXg3f4(Scripting.Dictionary dc)
    {
        /// lVCXg3f4 - generate "keyword" lists for all
        /// String Items in supplied Dictionary,
        /// including any sub Dictionaries
        /// '
        /// derived from lVCXg3f3, this function just
        /// adds one level of processing, calling
        /// lVCXg2f3 against the results of lVCXg3f1
        /// '
        /// it does NOT require input from lVCXg3f3
        /// and would likely fail against such a source
        /// '
        /// any Items not recognized as String
        /// or Dictionary Items are passed
        /// through as is, at present
        /// '
        /// might want to reconsider this,
        /// and probably make it optional
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant[] ar;
        // tx As String

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ar = Array(withBlock.Item(ky));
                if (VarType(ar[0]) == Constants.vbString)
                    rt.Add(ky, lVCXg2f3(lVCXg3f1(System.Convert.ToHexString(ar[0]))));
                else if (ar[0] is Scripting.Dictionary)
                    rt.Add(ky, lVCXg3f4(dcOb(ar[0])));
                else
                    rt.Add(ky, ar[0]);
            }
        }

        lVCXg3f4 = rt;
    }

    public Scripting.Dictionary lVCXg4f1(Scripting.Dictionary dc, Scripting.Dictionary rf = null/* TODO Change to default(_) if this is not a reference type */, Scripting.Dictionary rt = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// lVCXg4f1 - generate basic dependency list from
        /// supplied Dictionary of "keyword" Dictionaries
        /// keyed either to VB procedure names, or for
        /// a subset of identically named procedures,
        /// the module names of each implementation
        /// '
        /// note that this function does NOT disambiguate
        /// dependencies on multiply defined names.
        /// that is a task left to the client supplying
        /// the initial Dictionary, on the assumption
        /// that said client will have retained any
        /// prior source used to generate it
        /// '
        /// note also that optional Dictionary parameters
        /// rf and rt are NOT expected to be provided by
        /// an outside client, but passed to a recursive
        /// invocation when processing a multiply defined
        /// procedure name. for this purpose, the original
        /// source Dictionary dc is passed through rf to
        /// ensure its availability to all recursive calls
        /// '
        Variant ky;
        Variant[] ar;
        Scripting.Dictionary wk;

        if (rf == null)
            rt = lVCXg4f1(dc, dc, rt);
        else if (rt == null)
            rt = lVCXg4f1(dc, rf, new Scripting.Dictionary());
        else
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                wk = withBlock.Item(ky);
                {
                    var withBlock1 = wk;
                    if (withBlock1.Count > 0)
                    {
                        ar = Array(withBlock1.Item(withBlock1.Keys(0)));

                        // wk, NOT ar(0)
                        if (IsNumeric(ar[0]))
                            // intersect with reference Dictionary rf
                            // taking only the usage counts from wf
                            rt.Add(ky, dcKeysInCommon(wk, rf, 1));
                        else if (ar[0] is Scripting.Dictionary)
                            // we have a subset of implementations
                            // keyed to module locations
                            // so need to go down a level
                            rt.Add(ky, lVCXg4f1(wk, rf, new Scripting.Dictionary()));
                        else if (VarType(ar[0]) == Constants.vbString)
                        {
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            System.Diagnostics.Debugger.Break(); // because this does NOT normally happen
                        }
                        else
                        {
                            System.Diagnostics.Debugger.Break(); // because this shouldn't happen either
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }
                    }
                }

                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }
        }

        lVCXg4f1 = rt;
    }

    public Scripting.Dictionary dxOfVbProcLocsInMod(VBIDE.CodeModule cm)
    {
        /// dxOfVbProcLocsInMod -- Return Dictionary
        /// of Procedure Locations in CodeModule.
        /// derived from dcOfVbProcs as a "light"
        /// alternative that only returns an "index"
        /// of procedures, leaving the client to
        /// extract their text as needed
        /// 
        /// REV[2023.05.05.1307]: copied from dcOfVbProcs
        /// see that function for prior REVs
        /// if and where applicable
        /// 
        Scripting.Dictionary rt;
        long mx;
        long dx;
        long fw;
        string ck;
        // Dim tx As String
        Variant ar;
        long tp;

        rt = new Scripting.Dictionary();
        ar = Array(Array(vbext_pk_Proc, ""), Array(vbext_pk_Get, ""), Array(vbext_pk_Let, "=#"), Array(vbext_pk_Set, "=@"));

        {
            var withBlock = cm;
            mx = withBlock.CountOfLines;
            dx = withBlock.CountOfDeclarationLines;
            /// REV[2023.05.05.1329]: modified
            /// dx assignment for reuse below
            /// and call .CountOfDeclarationLines
            /// only once.
            rt.Add("", Array(1, dx));          /// REV[2023.05.05.1310]: replaced
                                               /// .Lines with Array to capture start
                                               /// line and line count of header
                                               /// (AKA DeclarationLines)

            dx = 1 + dx;
            while (dx < mx)
            {
                fw = dx;
                while (fw < mx)
                {
                    ck = withBlock.ProcOfLine(fw, vbext_pk_Proc);
                    if (Strings.Len(ck) == 0)
                        fw = fw + 1;
                    else
                    {
                        tp = 0;

                        do
                        {
                            Information.Err.Clear();
                            dx = withBlock.ProcStartLine(ck, ar(tp)(0));
                            if (Information.Err.Number)
                                dx = fw + 1;
                            if (dx != fw)
                                tp = tp + 1;
                            if (tp > 3)
                                System.Diagnostics.Debugger.Break();
                        }
                        while (!Information.Err.Number == 0 & dx == fw) // should NOT happen...
    ;
                        Information.Err.Clear();


                        fw = withBlock.ProcCountLines(ck, ar(tp)(0));
                        // tx = .Lines(dx, fw)
                        rt.Add(ck + ar(tp)(1), Array(dx, fw));                     /// REV[2023.05.05.1337]: replaced tx
                                                                                   /// .Lines with Array(dx, fw) to capture
                                                                                   /// start line and line count of procedure
                        dx = dx + fw;
                        fw = mx;
                    }
                }
            }
        }
        dxOfVbProcLocsInMod = rt;
    }

    public string vbProcTextFromPrj(string nm, VBIDE.VBProject pj = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// vbProcTextFromPrj
        /// derived from vbTextOfProcInProject (sort of)
        /// 
        /// NOTE: In order to use this Function
        /// from an external library,
        /// the option has been removed
        /// to call itself recursively
        /// against ThisWorkbook. Since
        /// ThisWorkbook would be the
        /// library itself, a call against
        /// it could result in a breach
        /// of security.
        /// 
        Scripting.Dictionary dc;
        string ky;
        Variant ar;
        VBIDE.CodeModule cm;

        if (pj == null)
            ky = "";
        else
        {
            var withBlock = dxOfVbProcLocsInPrj(pj);
            if (withBlock.Exists(nm))
            {
                ar = Array(null/* TODO Change to default(_) if this is not a reference type */);

                dc = withBlock.Item(nm);
                {
                    var withBlock1 = dc;
                    if (withBlock1.Count > 0)
                    {
                        ky = withBlock1.Keys(0);

                        if (withBlock1.Count > 1)
                            ky = userChoiceFromDc(dc, ky);

                        if (withBlock1.Exists(ky))
                            ar = withBlock1.Item(ky);
                    }
                }

                ky = "";
                if (UBound(ar) >= 2)
                {
                    cm = obVbCodeMod(obOf(ar(0)));
                    if (!cm == null)
                        ky = cm.Lines(ar(1), ar(2));
                }
            }
            else
                ky = "";
        }

        vbProcTextFromPrj = ky;
    }

    public Scripting.Dictionary dxOfVbProcLocsInPrj(VBIDE.VBProject pj)
    {
        /// dxOfVbProcLocsInPrj - generate Dictionary of
        /// Dictionaries of collected text
        /// of all procedures in each module
        /// of given VBProject, keyed first
        /// by module, and then by procedure
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        VBIDE.CodeModule cm;
        Variant kMd;
        Variant kPr;
        Variant ar;
        // Dim nm As String

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcOfVbModules(pj);
            foreach (var kMd in withBlock.Keys)
            {
                cm = obVbCodeMod(withBlock.Item(kMd));

                {
                    var withBlock1 = dxOfVbProcLocsInMod(cm);
                    foreach (var kPr in withBlock1.Keys)
                    {
                        {
                            var withBlock2 = rt;
                            if (!withBlock2.Exists(kPr))
                                withBlock2.Add(kPr, new Scripting.Dictionary());

                            wk = withBlock2.Item(kPr);
                        }

                        ar = withBlock1.Item(kPr);
                        {
                            var withBlock2 = wk;
                            if (withBlock2.Exists(kMd))
                                System.Diagnostics.Debugger.Break(); // for problem
                            else
                                withBlock2.Add(kMd, Array(cm, ar(0), ar(1)));
                        }
                    }
                }
            }
        }

        dxOfVbProcLocsInPrj = rt;
    }

    public Scripting.Dictionary dcOfVbProcs(VBIDE.CodeModule cm)
    {
        /// dcOfVbProcs -- Return Dictionary of Procedures
        /// -- from supplied CodeModule
        /// 
        /// REV[2023.04.19.1146]: added code to capture
        /// declaration text preceding all proc defs
        /// REV[2023.02.15.0904]: modified to accommodate ,
        /// , and  Procedures.  and  Procedures
        /// are stored under keys modified to indicate their
        /// role: "=#" for  indicates assignment to a value,
        /// while "=@" for  indicates an Object assignment.
        /// NOTE: this new version, while now able to accommodate
        /// Class Modules, is likely not the most efficient
        /// in addressing the problem. Further development
        /// might be warranted, should this prove an issue.
        /// 
        Scripting.Dictionary rt;
        long mx;
        long dx;
        long fw;
        string ck;
        string tx;
        Variant ar;
        long tp;

        rt = new Scripting.Dictionary();
        ar = Array(Array(vbext_pk_Proc, ""), Array(vbext_pk_Get, ""), Array(vbext_pk_Let, "=#"), Array(vbext_pk_Set, "=@"));

        {
            var withBlock = cm;
            mx = withBlock.CountOfLines;
            dx = 1 + withBlock.CountOfDeclarationLines;
            /// REV[2023.04.19.1146]: added following
            /// to capture header, AKA declaration lines
            rt.Add("", withBlock.Lines(1, withBlock.CountOfDeclarationLines));
            while (dx < mx)
            {
                fw = dx;
                while (fw < mx)
                {
                    ck = withBlock.ProcOfLine(fw, vbext_pk_Proc);
                    if (Strings.Len(ck) > 0)
                    {
                        tp = 0;

                        do
                        {
                            Information.Err.Clear();
                            dx = withBlock.ProcStartLine(ck, ar(tp)(0));
                            if (Information.Err.Number)
                                dx = fw + 1;
                            if (dx != fw)
                                tp = tp + 1;
                            if (tp > 3)
                                System.Diagnostics.Debugger.Break();
                        }
                        while (!Information.Err.Number == 0 & dx == fw) // should NOT happen...
    ;
                        Information.Err.Clear();


                        fw = withBlock.ProcCountLines(ck, ar(tp)(0));
                        tx = withBlock.Lines(dx, fw);
                        rt.Add(ck + ar(tp)(1), tx);
                        dx = dx + fw;
                        fw = mx;
                    }
                    else
                        fw = fw + 1;
                }
            }
        }
        dcOfVbProcs = rt;
    }

    public Scripting.Dictionary dcOfVbProcs_obs2023_0419(VBIDE.CodeModule cm)
    {
        /// dcOfVbProcs_obs2023_0419     -- Return Dictionary of Procedures
        /// -- from supplied CodeModule
        /// 
        /// NOTE: This function ONLY looks for general Procedures.
        /// It does NOT look for , , or  Procedures.
        /// It MIGHT NOT WORK properly against Class Modules!
        /// 
        Scripting.Dictionary rt;
        long mx;
        long dx;
        long fw;
        string ck;
        string tx;

        rt = new Scripting.Dictionary();
        {
            var withBlock = cm;
            if (withBlock.Parent.Type == vbext_ct_StdModule)
            {
                mx = withBlock.CountOfLines;
                dx = 1 + withBlock.CountOfDeclarationLines;
                // Debug.Print .Lines(1, .CountOfDeclarationLines) & "'''"

                while (dx < mx)
                {
                    fw = dx;
                    while (fw < mx)
                    {
                        ck = withBlock.ProcOfLine(fw, vbext_pk_Proc);
                        if (Strings.Len(ck) > 0)
                        {
                            dx = withBlock.ProcStartLine(ck, vbext_pk_Proc);
                            fw = withBlock.ProcCountLines(ck, vbext_pk_Proc);
                            tx = withBlock.Lines(dx, fw);
                            rt.Add(ck, tx);
                            dx = dx + fw;
                            fw = mx;
                        }
                        else
                            fw = fw + 1;
                    }
                }
            }
            else
            {
            }
        }
        dcOfVbProcs_obs2023_0419 = rt;
    }

    public Scripting.Dictionary dcOfVbModules(VBIDE.VBProject vb)
    {
        Scripting.Dictionary rt;
        VBIDE.VBComponent vc;

        rt = new Scripting.Dictionary();
        {
            var withBlock = vb;
            if (withBlock.Protection == vbext_pp_none)
            {
                foreach (var vc in withBlock.VBComponents)
                {
                    {
                        var withBlock1 = vc;
                        rt.Add.Name(null/* Conversion error: Set to default value for this argument */, withBlock1.CodeModule);
                    }
                }
            }
            else
                rt.Add("<PROTECTED>", new Scripting.Dictionary());
        }
        dcOfVbModules = rt;
    }

    public Scripting.Dictionary dcOfVbProcsFlat(VBIDE.VBProject pj)
    {
        /// dcOfVbProcsFlat - generate Dictionary
        /// of collected text of all procedures
        /// in each module of given VBProject,
        /// keyed by procedure name, or by
        /// combination of module and procedure
        /// name when more than one procedure
        /// of same name is found
        /// 
        /// NOTE[2023.04.19.1256] the compromise
        /// noted above is NOT ideal.
        /// As the purpose of this function
        /// is to produce a FLAT list
        /// of procedure names for quick
        /// searching purposes, the need
        /// to modify duplicate names
        /// for is likely to make it
        /// difficult or impractical
        /// to find all possible matches
        /// 
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dc;
        Variant kyMd;
        Variant kyPr;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcOfVbModules(pj);
            foreach (var kyMd in withBlock.Keys)
            {
                {
                    var withBlock1 = dcOfVbProcs(obOf(withBlock.Item(kyMd))) // dcOb
          ;
                    foreach (var kyPr in withBlock1.Keys)
                    {
                        if (rt.Exists(kyPr))
                        {
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // breakpoint landing
                                                                                         // Stop
                                                                                         // going to need a better way
                                                                                         // to handle this situation
                                                                                         // but for now...
                            rt.Add(kyMd + "." + kyPr, withBlock1.Item(kyPr));
                        }
                        else
                            rt.Add(kyPr, withBlock1.Item(kyPr));
                    }
                }
            }
        }

        dcOfVbProcsFlat = rt;
    }

    public Scripting.Dictionary dcOfVbProjects(VBIDE.VBProjects pjs)
    {
        Scripting.Dictionary rt;
        VBIDE.VBProject vb;

        rt = new Scripting.Dictionary();
        {
            var withBlock = pjs;
            foreach (var vb in pjs)
                rt.Add(vb.Filename, dcOfVbModules(vb));
        }
        dcOfVbProjects = rt;
    }

    public Scripting.Dictionary dcTxOfVbModule(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        if (dc == null)
        {
        }
        else
        {
            var withBlock = dc;
            if (withBlock.Count > 0)
            {
                foreach (var ky in withBlock.Keys)
                    rt.Add(ky, txOfVbModule(obVbCodeMod(withBlock.Item(ky))));
            }
        }

        dcTxOfVbModule = rt;
    }

    public Scripting.Dictionary dcTxOfVbProjMods(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            if (withBlock.Count > 0)
            {
                foreach (var ky in withBlock.Keys)
                    rt.Add(ky, dcTxOfVbModule(dcOb(withBlock.Item(ky))));
            }
        }

        dcTxOfVbProjMods = rt;
    }

    public string vbTextOfProcInDict(string nm, Scripting.Dictionary dc)
    {
        /// vbTextOfProcInDict -- Retrieve text from Dictionary
        /// 
        /// This Function's name is probably
        /// WAY unnecessarily specific.
        /// 
        /// The Function itself simply returns the String
        /// found under the supplied key variable 'nm',
        /// or an empty String if none is found. This is
        /// a fairly general type of Function, one which
        /// could be named far more generically.
        /// 
        /// dcItemIfPresent looks a likely candidate,
        /// although it might be a bit TOO general...
        /// 
        if (dc == null)
            // Recursive call option removed.
            // See text of vbTextOfProcInProject
            // for details on security issue.
            // 
            vbTextOfProcInDict = ""; // vbTextOfProcInDict(nm, dcOfVbProcsFlat(ThisWorkbook.VBProject))
        else
            vbTextOfProcInDict = System.Convert.ToHexString(dcItemIfPresent(dc, nm, Constants.vbString));
    }

    public string vbTextOfProcIn(string nm, VBIDE.CodeModule cm)
    {
        vbTextOfProcIn = vbTextOfProcInDict(nm, dcOfVbProcs(cm));
    }

    public string vbTextOfProcInProject(string nm, VBIDE.VBProject pj)
    {
        /// vbTextOfProcInProject
        /// 
        /// NOTE: In order to use this Function
        /// from an external library,
        /// the option has been removed
        /// to call itself recursively
        /// against ThisWorkbook. Since
        /// ThisWorkbook would be the
        /// library itself, a call against
        /// it could result in a breach
        /// of security.
        /// 
        if (pj == null)
            vbTextOfProcInProject = "";
        else
            vbTextOfProcInProject = vbTextOfProcInDict(nm, dcOfVbProcsFlat(pj));
    }

    public Variant send2clipBd(Variant src)
    {
        string ck;

        ck = send2clipBdWin10(src);

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
        // If MsgBox('            Join(Array('                "Error Getting Text from DataObject.",'                "A simple retry will usually succeed.",'                "", "Go ahead and retry?"'            ), vbNewLine),'            vbYesNo, "Retry GetText?"'        ) = vbNo Then
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
            System.Diagnostics.Debugger.Break();

        send2clipBd = src;
    }

    public Variant getFromClipBd(Variant fmt = 1)
    {
        // '  1 is the value of CF_TEXT, one of the clipboard format
        // '  enums which SHOULD be defined, but apparently aren't.
        // '  That is the effective default format used by GetText,
        // '  if none is given
        Variant rt;
        {
            var withBlock = new MSForms.DataObject();
            withBlock.GetFromClipboard();
            rt = withBlock.GetText(fmt);
        }
        getFromClipBd = rt;
    }

    public string dumpKeyedText(Scripting.Dictionary d1, Scripting.Dictionary d2)
    {
        // '  Extract values from second dictionary
        // '  filed under keys from FIRST dictionary.
        // '  Theory is, the keys will always be
        // '  retrieved in the same order, as long as
        // '  no changes have been made between runs.
        // '
        // '  By supplying the same dictionary for
        // '  both d1 and d2, that dictionary's
        // '  content can be extracted, and then
        // '  a different d2's content can be
        // '  extracted in the same order.
        Variant ky;
        Scripting.Dictionary rt;

        rt = new Scripting.Dictionary();
        foreach (var ky in d1.Keys)
            rt.Add("{" + ky + "}" + Constants.vbNewLine + d2(ky), 1);
        dumpKeyedText = Join(rt.Keys, Constants.vbNewLine);
    }
}