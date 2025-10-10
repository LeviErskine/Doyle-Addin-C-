class SurroundingClass
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmSelectorV2";
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_GlobalNameSpace = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Creatable = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_PredeclaredId = true;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Exposed = false;



    private const string tagSelected = "%%%";
    private Scripting.Dictionary dc;

    private string msCancelHead;
    private string msCancelMain;
    private string msNoSelHead;
    private string msNoSelMain;
    private string msOkHead;
    private string msOkMain;
    // Private msCancelMain As String
    // Private msCancelMain As String

    public fmSelectorV2 SetMsgCancel(string Using)
    {
        msCancelMain = ; SetMsgCancel = this;
    }

    public fmSelectorV2 SetMsgNoSelection(string Using)
    {
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    msNoSelMain = Using

 */     SetMsgNoSelection = this;
    }

    public fmSelectorV2 SetMsgOK(string Using)
    {
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    msOkMain = Using

 */     SetMsgOK = this;
    }

    public fmSelectorV2 SetHdrCancel(string Using)
    {
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    msCancelHead = Using

 */     SetHdrCancel = this;
    }

    public fmSelectorV2 SetHdrNoSelection(string Using)
    {
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    msNoSelHead = Using

 */     SetHdrNoSelection = this;
    }

    public fmSelectorV2 SetHdrOK(string Using)
    {
        ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    msOkHead = Using

 */     SetHdrOK = this;
    }

    public fmSelectorV2 SelectIfIn(string Using)
    {
        long dx;

        {
            var withBlock = this.lsbSelection;
            dx = withBlock.ListIndex;
            ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 1392
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error Resume Next

 */
            Information.Err.Clear();
            ;/* Cannot convert AssignmentStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitAssignmentStatement(AssignmentStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 103
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        .Value = Using

 */
            if (Information.Err.Number)
            {
                withBlock.ListIndex = dx;
                // .Value = ""
                Information.Err.Clear();
            }
            ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 1579
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error GoTo 0

 */
        }

        SelectIfIn = this;
    }

    public fmSelectorV2 WithList(Variant Using)
    {
        Variant ky;
        string it;
        ;/* Cannot convert MultiLineIfBlockSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitSimpleArgument(SimpleArgumentSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1060
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.<>c__DisplayClass83_0.<ConvertArguments>b__0(ArgumentSyntax a, Int32 i) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1045
   at System.Linq.Enumerable.SelectIterator[TSource,TResult](IEnumerable`1 source, Func`3 selector)+MoveNext()
   at System.Linq.Enumerable.WhereEnumerableIterator`1.MoveNext()
   at Microsoft.CodeAnalysis.CSharp.SyntaxFactory.SeparatedList[TNode](IEnumerable`1 nodes)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitArgumentList(ArgumentListSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1022
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitInvocationExpression(InvocationExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1422
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitMultiLineIfBlock(MultiLineIfBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 353
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    
    If IsObject(Using) Then
        'Me.lsbSelection.List = Using
        If TypeOf Using Is Scripting.Dictionary Then
            dc = Using
        ElseIf TypeOf Using Is Inventor.NameValueMap Then
            dc = dcFromAiNameValMap(obOf(Using))
        Else
            'Stop
            Debug.Print ; 'Breakpoint Landing
             dc = Nothing
        End If
        
        If dc Is Nothing Then
            Me.lsbSelection.List = Array("<no items>")
        ElseIf dc.Count > 0 Then
            Me.lsbSelection.List = dc.Keys
        Else
            Me.lsbSelection.List = Array("<no items>")
        End If
    ElseIf IsArray(Using) Then
        dc = New Scripting.Dictionary

        For Each ky In Using
            dc.Add CStr(ky), ky
        Next
        
        Me.lsbSelection.List = dc.Keys
    Else
        'Stop
        Debug.Print ; 'Breakpoint Landing
    End If

 */
        WithList = this;
    }

    private void btnCancel_Click()
    {
        // '
        if (MsgBox(msCancelMain, Constants.vbYesNo, msCancelHead) == Constants.vbYes)
        {
            this.lsbSelection.ListIndex = -1;
            this.Hide();
        }
        else
        {
        }
    }

    private void btnOk_Click()
    {
        // '
        VbMsgBoxResult ck;
        long mx;
        long dx;
        long ct;

        string ls;

        {
            var withBlock = this.lsbSelection;
            if (withBlock.MultiSelect == fmMultiSelectSingle)
            {
                if (withBlock.ListIndex < 0)
                {
                    ck = MsgBox(msNoSelMain, Constants.vbYesNo, msNoSelHead);
                    if (ck == Constants.vbYes)
                        this.Hide();
                }
                else
                {
                    ck = MsgBox(Join(Split(msOkMain, tagSelected), withBlock.Value), Constants.vbYesNoCancel, msOkHead
              );
                    if (ck == Constants.vbYes)
                        this.Hide();
                    else if (ck == Constants.vbCancel)
                    {
                        withBlock.ListIndex = -1;
                        this.Hide();
                    }
                    else
                    {
                    }
                }
            }
            else
            {
                ls = lbxPickedStr(this.lsbSelection, Constants.vbNewLine);

                // ct = 0
                // mx = .ListCount - 1
                // For dx = 0 To mx
                // If .Selected(dx) Then ct = 1 + ct
                // Next

                if (Strings.Len(ls) > 0)
                {
                    ck = MsgBox(Join(Split(msOkMain, tagSelected), Constants.vbNewLine + ls + Constants.vbNewLine
), Constants.vbYesNoCancel, msOkHead
);
                    if (ck == Constants.vbYes)
                        this.Hide();
                    else if (ck == Constants.vbCancel)
                    {
                        withBlock.ListIndex = -1;
                        this.Hide();
                    }
                    else
                    {
                    }
                }
                else
                {
                    ck = MsgBox(msNoSelMain, Constants.vbYesNo, msNoSelHead);
                    if (ck == Constants.vbYes)
                        this.Hide();
                }
            }
        }
    }

    private void lsbSelection_Change()
    {
        /// 
        string ck;
        ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 4692
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    
    On Error Resume Next

 */
        Information.Err.Clear();
        ck = this.lsbSelection.Value;
        if (Information.Err.Number != 0)
            ck = "";
        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 4798
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    On Error GoTo 0

 */
        {
            var withBlock = dc;
            if (withBlock.Exists(ck))
            {
                ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 4871
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
            On Error Resume Next

 */
                Information.Err.Clear();
                tbxView.Value = System.Convert.ToHexString(withBlock.Item(ck));
                if (Information.Err.Number != 0)
                    tbxView.Value = "<not printable>";
                ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 5101
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
            
            On Error GoTo 0

 */
            }
            else
                tbxView.Value = "<no data>";
        }
    }

    private void lsbSelection_DblClick(MSForms.ReturnBoolean Cancel)
    {
        btnOk_Click();
    }

    private void UserForm_Initialize()
    {
        // 
        msCancelHead = "Cancel Operation?";
        msNoSelHead = "No Selection!";
        msOkHead = "Proceed?";
        // 
        msCancelMain = "Selection will be canceled.";
        msNoSelMain = Join(Array("Do you wish to cancel?", "(Click NO to return to list)"
        ), Constants.vbNewLine);
        msOkMain = Join(Array("Current selection is: ", tagSelected, "(Click CANCEL to quit with no selection)"
        ), Constants.vbNewLine);
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        /// 
        Cancel = 1;
        btnCancel_Click();
    }

    public string GetReply(Variant List = , object As = string == "%$#@"
    )
    {
        string rt;

        rt = "";
        ;/* Cannot convert WithBlockSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitSimpleArgument(SimpleArgumentSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1060
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.<>c__DisplayClass83_0.<ConvertArguments>b__0(ArgumentSyntax a, Int32 i) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1045
   at System.Linq.Enumerable.SelectIterator[TSource,TResult](IEnumerable`1 source, Func`3 selector)+MoveNext()
   at System.Linq.Enumerable.WhereEnumerableIterator`1.MoveNext()
   at Microsoft.CodeAnalysis.CSharp.SyntaxFactory.SeparatedList[TNode](IEnumerable`1 nodes)
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitArgumentList(ArgumentListSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1022
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitInvocationExpression(InvocationExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1422
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitWithBlock(WithBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 510
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    With Me.WithList(List).SelectIfIn(Default)
        '.lsbSelection.List = lsWorkbooks()
        .Show 1
        If .lsbSelection.MultiSelect = fmMultiSelectSingle Then
            rt = .lsbSelection.Text
        Else
            rt = lbxPickedStr(.lsbSelection, vbNewLine)
        End If
    End With

 */
        GetReply = rt;
    }

    private void UserForm_Resize()
    {
    }
}