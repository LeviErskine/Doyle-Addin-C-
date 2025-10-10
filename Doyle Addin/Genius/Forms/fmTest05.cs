class SurroundingClass
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmTest05";
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_GlobalNameSpace = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Creatable = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_PredeclaredId = true;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Exposed = false;


    public event SentEventHandler Sent;

    public delegate void SentEventHandler(VbMsgBoxResult Signal);

    public event GroupIsEventHandler GroupIs;

    public delegate void GroupIsEventHandler(string Now);

    public event ItemIsEventHandler ItemIs;

    public delegate void ItemIsEventHandler(string Now);

    private Scripting.Dictionary dcHolding;

    private const string txVersion = "";
    /// 

    /// 

    public fmTest05 Holding(object Obj
    )
    {
        /// Holding -- Hold onto supplied
        /// Object until terminated,
        /// or directed to drop it.
        /// 
        /// not sure about this one.
        /// purpose is to keep a
        /// client interface "alive"
        /// while the form itself
        /// remains active.
        /// 
        {
            var withBlock = dcHolding;
            if (withBlock.Exists(Obj))
            {
            }
            else
                withBlock.Add(Obj, withBlock.Count);
        }

        Holding = this;
    }

    public fmTest05 Dropping(object Obj
    )
    {
        {
            var withBlock = dcHolding;
            if (withBlock.Exists(Obj))
                withBlock.Remove(Obj);
            else
            {
            }
        }

        Dropping = this;
    }

    public string GroupNow()
    {
        MSForms.Tab tb;
        long dx;

        {
            var withBlock = tbsItemGrps;
            dx = withBlock.Value;
            tb = withBlock.Tabs.Item(dx);
            GroupNow = tb.Name;
        }
    }

    public fmTest05 InGroup(string GrpId
    ) // fmIfcTest05A
    {
        MSForms.Tab tb;

        {
            var withBlock = tbsItemGrps;
            ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 1452
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error Resume Next

 */
            Information.Err.Clear();

            tb = withBlock.Tabs.Item(GrpId);
            if (Information.Err.Number == 0)
                withBlock.Value = tb.Index;
            else
            {
            }

            Information.Err.Clear();
            ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 1824
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error GoTo 0

 */
        }

        InGroup = this;
    }

    public string ItemNow()
    {
        // With lbxItems
        ItemNow = lbxItems.Value;
    }

    public fmTest05 OnItem(string ItemId
    ) // fmIfcTest05A
    {
        // Dim tb As MSForms.Tab

        {
            var withBlock = lbxItems;
            ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 2124
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error Resume Next

 */
            Information.Err.Clear();

            // tb = .Tabs.Item(ItemId)
            withBlock.Value = ItemId;
            if (Information.Err.Number == 0)
                // .Value = tb.Index
                // Stop
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            else
                System.Diagnostics.Debugger.Break();

            Information.Err.Clear();
            ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 2426
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error GoTo 0

 */
        }

        OnItem = this;
    }

    private void cmdEndCancel_Click()
    {
        Sent?.Invoke(Constants.vbCancel);
    }

    private void cmdEndSave_Click()
    {
        Sent?.Invoke(Constants.vbOK);
    }

    private void cmdOpenItem_Click()
    {
        Sent?.Invoke(Constants.vbRetry);
    }

    private void lbxItems_Change()
    {
        ItemIs?.Invoke(lbxItems.Value);
    }

    private void tbsItemGrps_Change()
    {
        GroupIs?.Invoke(GroupNow);
    }

    private void tbsItemGrps_BeforeDropOrPaste(long Index, MSForms.ReturnBoolean Cancel, MSForms.fmAction Action, MSForms.DataObject Data, float X, float Y, MSForms.ReturnEffect Effect, int Shift
    )
    {
        // will keep this one as is, for now
        // not sure what you can actually drop
        // onto a tab group
        System.Diagnostics.Debugger.Break();
    }

    private void lbxItems_MouseMove(int Button, int Shift, float X, float Y
    )
    {
        /// keeping this one here, since it basically governs
        /// drag-and-drop behavior from a local control.
        /// might try to see if this is actually needed.
        /// one would think this kind of behavior
        /// would occur automatically.
        MSForms.DataObject dt;
        int ef;

        if (Button == 1)
        {
            dt = new MSForms.DataObject();
            dt.SetText(lbxItems.Value);
            ef = dt.StartDrag();
        }
    }

    private void UserForm_Initialize()
    {
        dcHolding = new Scripting.Dictionary();
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        Sent?.Invoke(Constants.vbAbort);
    }

    private void UserForm_Terminate()
    {
        dcHolding.RemoveAll();
        dcHolding = null;
    }
}