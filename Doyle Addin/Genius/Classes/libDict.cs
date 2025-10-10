class libDict
{
    public Scripting.Dictionary dcNewIfNone(Scripting.Dictionary Dict)
    {
        if (Dict == null)
            dcNewIfNone = new Scripting.Dictionary();
        else
            dcNewIfNone = Dict;
    }

    public Scripting.Dictionary dcOfRsFields(ADODB.Recordset rs)
    {
        Scripting.Dictionary rt;
        ADODB.Field fd;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rs;
            if (withBlock.State == adStateOpen)
            {
                foreach (var fd in withBlock.Fields)
                    rt.Add(fd.Name, fd);
            }
        }
        dcOfRsFields = rt;
    }

    public Scripting.Dictionary dcDotted(Scripting.Dictionary Under = null/* TODO Change to default(_) if this is not a reference type */, Scripting.Dictionary Using = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcDotted -- return Dictionary with
        /// links to itself, under key ".",
        /// and under "..", either itself,
        /// or, if supplied, an optional
        /// "parent" Dictionary
        /// 
        /// this mimics the traditional linkage within
        /// POSIX-compliant and other file systems,
        /// where the "." and ".." names in each
        /// directory are assigned to itself
        /// and its parent, respecrively
        /// 
        /// !!WARNING!!
        /// this self- and back-linkage WILL cause
        /// endless loops in Dictionary traversal
        /// routines not prepared to deal with them!
        /// Be sure to review any procedure BEFORE
        /// calling against a Dictionary using
        /// this linkage!
        /// 
        Scripting.Dictionary rt;
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.<>c__DisplayClass83_0.<ConvertArguments>b__0(ArgumentSyntax a, Int32 i) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1045
   at System.Linq.Enumerable.SelectIterator[TSource,TResult](IEnumerable`1 source, Func`3 selector)+MoveNext()
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
    
 rt = dcNewIfNone(Using)

 */
        {
            var withBlock = rt;
            if (withBlock.Exists("."))
            {
                if (withBlock.Item(".") == rt)
                {
                }
                else
                    System.Diagnostics.Debugger.Break();
            }
            else
                withBlock.Add(".", rt);

            if (withBlock.Exists(".."))
            {
                if (withBlock.Item("..") is Scripting.Dictionary)
                    ;/* Cannot convert MultiLineIfBlockSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitMultiLineIfBlock(MultiLineIfBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 353
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                If Using Is Nothing Then
                Else
                    Stop
 .Item("..") = Using
                End If

 */
                else
                    System.Diagnostics.Debugger.Break();
            }
            else
                withBlock.Add("..", IIf(Under == null, rt, Under));
        }

        dcDotted = rt;
    }

    public Scripting.Dictionary dcUnDotted(Scripting.Dictionary dc)
    {
        /// dcUnDotted -- remove Keys "." and ".."
        /// from supplied Dictionary dc
        /// 
        /// no checks are made of the Items under these Keys.
        /// the Dictionary is assumed to have originated from
        /// or passed through a prior call to dcDotted, and
        /// thus include self- and back-linkage thereunder.
        /// 
        /// (a check system was considered and attempted,
        /// but deemed too unweildy, and so abandoned)
        /// 
        {
            var withBlock = dc;
            if (withBlock.Exists("."))
                withBlock.Remove(".");
            if (withBlock.Exists(".."))
                withBlock.Remove("..");
        }
        dcUnDotted = dc;
    }

    public Scripting.Dictionary dcFrom2Fields(ADODB.Recordset rs, string fnKey, string fnVal, string flt = "")
    {
        Scripting.Dictionary rt;
        ADODB.Field fdKey;
        ADODB.Field fdVal;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rs;
            {
                var withBlock1 = withBlock.Fields;
                fdKey = withBlock1.Item(fnKey);
                fdVal = withBlock1.Item(fnVal);
            }

            withBlock.Filter = flt;
            while (!withBlock.BOF | withBlock.EOF)
            {
                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(fdKey.Value))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(fdKey.Value, fdVal.Value);
                }
                withBlock.MoveNext();
            }
        }
        dcFrom2Fields = rt;
    }

    public Scripting.Dictionary dcFromAdoRS(ADODB.Recordset rs, string flt = "")
    {
        // , fnKey As String, fnVal As String
        // , Optional ovr As Long = -1
        /// dcFromAdoRS - return a Dictionary
        /// of tuples (rows) from an ADODB
        /// Recordset, keyed on order of
        /// encounter and processing.
        /// 
        /// NOTE that this Dictionary is NOT
        /// keyed on any particular Field.
        /// The wide range of situations which
        /// might be encountered suggests that
        /// indexing and keying on field values
        /// is best addressed in a separate,
        /// dedicated process.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary tp;
        // Dim fdKey As ADODB.Field
        ADODB.Field fdVal;
        Variant ky;
        Variant vl;
        string nm;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rs;
            // With .Fields
            // fdKey = .Item(fnKey)
            // End With

            withBlock.Filter = flt;
            while (!withBlock.BOF | withBlock.EOF)
            {
                ky = rt.Count; // fdKey.Value
                {
                    var withBlock1 = rt;
                    // If .Exists(ky) Then 'we have a collision!
                    // Stop 'and figure out what to do!
                    // Else
                    // .Add ky, dcFromAdoRSrow(rs)
                    withBlock1.Add(ky, new Scripting.Dictionary());
                    // End If
                    tp = dcOb(withBlock1.Item(ky));
                }

                foreach (var fdVal in withBlock.Fields)
                {
                    {
                        var withBlock1 = fdVal;
                        nm = withBlock1.Name;
                        vl = withBlock1.Value;
                    }

                    {
                        var withBlock1 = tp;
                        // If .Exists(nm) Then
                        // If ovr Then 'change if needed
                        // If .Item(nm) <> vl Then
                        // .Item(nm) = vl
                        // End If
                        // Else 'fuhgeddaboudit!
                        // End If
                        // Else
                        withBlock1.Add(nm, vl);
                    }
                }

                withBlock.MoveNext();
            }
        }
        dcFromAdoRS = rt;
    }

    public Scripting.Dictionary dcFromAdoRSrow(ADODB.Recordset rs, Variant nullVal = Null)
    {
        Scripting.Dictionary rt;
        ADODB.Field fd;
        string nm;
        bool ck;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rs;
            ck = withBlock.BOF | withBlock.EOF;
            foreach (var fd in withBlock.Fields)
            {
                {
                    var withBlock1 = fd;
                    nm = withBlock1.Name;
                    if (ck)
                        rt.Add(nm, nullVal);
                    else
                        rt.Add(nm, withBlock1.Value);
                }
            }
        }
        dcFromAdoRSrow = rt;
    }

    public Scripting.Dictionary dcDxFromRecSetDc(Scripting.Dictionary Dict)
    {
        /// dcDxFromRecSetDc -- Generate Dictionary
        /// of Indices from "RecordSet" Dictionary
        /// as returned by dcFromAdoRS
        /// 
        Scripting.Dictionary tp;
        Scripting.Dictionary dcDx;
        Scripting.Dictionary dcVl;
        Scripting.Dictionary dcTp;
        // '
        Variant k0;
        Variant k1;
        Variant k2;
        // '
        Variant vl;

        // rt = New Scripting.Dictionary

        dcDx = new Scripting.Dictionary();
        // '  the Dictionary of Indices

        {
            var withBlock = Dict;
            // '  Start scanning primary Keys
            // '  to begin overall process
            foreach (var k0 in withBlock.Keys)
            {
                // '  Retrieve "record" Dictionary
                // '  for next/current Key
                tp = dcOb(withBlock.Item(k0));
                if (tp == null)
                    // Stop
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                else
                {
                    var withBlock1 = tp;
                    // '  Scan "field" Keys of current "record"
                    // '  to identify and populate Index Dictionaries
                    foreach (var k1 in withBlock1.Keys)
                    {
                        // '  Retrieve "index" Dictionary for current
                        // '  "field". Generate new one, if not present.
                        // '
                        // '  (might want to support Key filtering
                        // '  to either exclude some "fields",
                        // '  or limit indexing to a list)
                        {
                            var withBlock2 = dcDx;
                            if (withBlock2.Exists(k1))
                            {
                            }
                            else
                                withBlock2.Add(k1, new Scripting.Dictionary());
                            dcVl = dcOb(withBlock2.Item(k1));
                        }

                        // '  Retrieve current "field" value, and return its
                        // '  Dictionary from the "field index" Dictionary.
                        // '
                        // '  Again, generate a new one, if needed.
                        vl = withBlock1.Item(k1);
                        {
                            var withBlock2 = dcVl;
                            if (withBlock2.Exists(vl))
                            {
                            }
                            else
                                withBlock2.Add(vl, new Scripting.Dictionary());
                            dcTp = dcOb(withBlock2.Item(vl));
                        }

                        // '  Add the current "record" to the recovered
                        // '  "field value" Dictionary. This SHOULD only
                        // '  add a link to the same "record" Dictionary,
                        // '  rather than duplicate the whole thing.
                        // '
                        // '  However, converting to JSON generates
                        // '  a new dump of the Dictionary structure
                        // '  wherever it appears, thus replicating it
                        // '  multiple times in the output.
                        // '
                        {
                            var withBlock2 = dcTp;
                            if (withBlock2.Exists(k0))
                                System.Diagnostics.Debugger.Break(); // for now. might still be okay
                            else
                                withBlock2.Add(k0, tp);

                            DoEvents();
                        }
                    }
                }
            }
        }

        {
            var withBlock = dcDx;
            if (withBlock.Exists(""))
                System.Diagnostics.Debugger.Break(); // because we have
            else
                withBlock.Add("", Dict);
        }

        dcDxFromRecSetDc = dcDx;
    }

    public Scripting.Dictionary dcRecSetDcDx4json(Scripting.Dictionary Dict)
    {
        /// dcRecSetDcDx4json -- Prep RecordSet
        /// Index Dictionary for JSON export.
        /// 
        /// Replaces each field/value index
        /// Dictionary with its Keys for export
        /// to JSON, to avoid replicating each
        /// original "record" Item in its entirety
        /// wherever it's referenced in the indices.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcFdIn;
        Scripting.Dictionary dcFdOut;
        Scripting.Dictionary dcVl;
        // '
        Variant k0;
        Variant vl;

        rt = new Scripting.Dictionary();
        // '  the Dictionary of Indices

        {
            var withBlock = Dict;
            // '  Start scanning field
            // '  names (top level Keys)
            // '  to begin transformation
            foreach (var k0 in withBlock.Keys)
            {
                // '  Retrieve next "field index" Dictionary
                dcFdIn = dcOb(withBlock.Item(k0));

                // '  Check for original RecordSet Dictionary
                if (k0 == "")
                    rt.Add("", dcFdIn);
                else
                {
                    // '  Generate corresponding "field index"
                    // '  output Dictionary
                    {
                        var withBlock1 = rt;
                        if (withBlock1.Exists(k0))
                            System.Diagnostics.Debugger.Break(); // because it should NOT
                        else
                            withBlock1.Add(k0, new Scripting.Dictionary());

                        dcFdOut = dcOb(withBlock1.Item(k0));
                    }

                    // '  Scan value Keys of current "field"
                    // '  to retrieve index Dictionaries
                    {
                        var withBlock1 = dcFdIn;
                        foreach (var vl in withBlock1.Keys)
                        {
                            // '  Retrieve Dictionary for current value
                            dcVl = dcOb(withBlock1.Item(vl));

                            {
                                var withBlock2 = dcFdOut;
                                if (withBlock2.Exists(vl))
                                    System.Diagnostics.Debugger.Break(); // because it should
                                else
                                    // '  Dump record Keys to output
                                    // '  field value index Dictionary
                                    withBlock2.Add(vl, dcVl.Keys);
                            }
                        }
                    }
                }
            }
        }

        dcRecSetDcDx4json = rt;
    }

    public Scripting.Dictionary dcOfSubDict(Scripting.Dictionary dc, Scripting.Dictionary rt = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcOfSubDict -- intended to return
        /// a "flat" Dictionary containing
        /// the supplied Dictionary and all
        /// Dictionary objects within it.
        /// 
        /// DO NOT ATTEMPT TO USE AT THIS TIME!!!
        /// Need to work out a way to tell
        /// if the supplied Dictionary is already in the returned
        /// 
        Variant ky;

        if (rt == null)
            dcOfSubDict = dcOfSubDict(dc, new Scripting.Dictionary());
        else if (dc == null)
            dcOfSubDict = rt;
        else
        {
        }
    }
}