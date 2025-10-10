class SurroundingClass
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmTest2";
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_GlobalNameSpace = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Creatable = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_PredeclaredId = true;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Exposed = false;


    private Inventor.Document ad;

    private Inventor.PropertySet psDsn;
    private Inventor.PropertySet psUsr;

    private Scripting.Dictionary dcDsn;
    private Scripting.Dictionary dcUsr;

    private Inventor.Property prFam;
    private Inventor.Property prStk;

    // Private dmFmHt As Long
    // Private dmFmWd As Long
    // 'Private dmLbMsHt As Long
    private long dmLbMsWd;
    // 'Private dmDfFmMsHt As Long
    // 'Private dmDfFmMsWd As Long
    private long dmFmHt2cmdTop;

    private VbMsgBoxResult rtAnswer;

    public VbMsgBoxResult AskAbout(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */, string txPre = "", string txPost = ""
    )
    {
        /// AskAbout -- prompt User for action
        /// to take on supplied Document
        /// UPDATE[2021.12.13]
        /// Document parameter now Optional.
        /// will attempt to use previously
        /// registered Document when none
        /// supplied. Warning/error message
        /// will be presented if no Document
        /// is registered OR supplied.
        /// 
        stdole.IPictureDisp pc;
        string pn;
        string sn;
        string pd;
        float dj; // use to adjust
                  // form height and positions
                  // of command buttons

        rtAnswer = Constants.vbCancel;
        if (!AiDoc == null)
            ad = AiDoc;

        if (ad == null)
        {
            MsgBox("Review or Update requested"
+ Constants.vbNewLine + "but no Document provided!"
+ Constants.vbNewLine + ""
+ Constants.vbNewLine + ""
, Constants.vbOKOnly, "No Document!");
            rtAnswer = Constants.vbNo;
        }
        else if (aiDocPartFromCCtr(ad) == null)
        {
            // ad = AiDoc
            {
                var withBlock = ad;
                pc = withBlock.Thumbnail;
                psDsn = withBlock.PropertySets(gnDesign);
                psUsr = withBlock.PropertySets(gnCustom);

                dcDsn = dcAiPropsInSet(psDsn);
                dcUsr = dcAiPropsInSet(psUsr);

                prFam = psDsn.Item(pnFamily);
                {
                    var withBlock1 = dcUsr;
                    if (withBlock1.Exists(pnRawMaterial))
                        prStk = psUsr.Item(pnRawMaterial);
                    else
                    {
                        ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 2266
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                    On Error Resume Next

 */
                        Information.Err.Clear();
                        prStk = psUsr.Add("", pnRawMaterial);
                        if (Information.Err.Number)
                        {
                            Debug.Print(Information.Err.Number, Information.Err.Description);
                            System.Diagnostics.Debugger.Break();
                        }
                        else
                            withBlock1.Add(pnRawMaterial, prStk);
                        ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 2630
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                    On Error GoTo 0

 */
                    }
                }

                if (!prStk == null)
                    sn = prStk.Value;
                pn = psDsn.Item(pnPartNum).Value;
                pd = psDsn.Item(pnDesc).Value;
            }

            {
                var withBlock = this;
                withBlock.Caption = "Please Review Item: " + pn;

                if (pc == null)
                {
                }
                else
                    withBlock.imThmNail.Picture = pc;

                dj = fmHtAdjust(lblHtAdjust(withBlock.lbMsg, Interaction.IIf(Strings.Len(txPre) > 0, txPre + Constants.vbNewLine + Constants.vbNewLine, ""
    ) + Join(Array(pn + ": " + pd, pnCatWebLink + ": " + psDsn.Item(pnCatWebLink).Value, pnMaterial + ": " + psDsn.Item(pnMaterial).Value
    ), Constants.vbNewLine + Constants.vbNewLine)
                    + Interaction.IIf(Strings.Len(txPost) > 0, Constants.vbNewLine + Constants.vbNewLine + txPost, ""
    )
    ));
                // .dbFamily.Value = prFam.Value

                withBlock.Show(1);
            }
        }
        else
        {
            MsgBox(ad.DisplayName
          + Constants.vbNewLine + "is a Content Center part"
          + Constants.vbNewLine + "and cannot be updated."
          + Constants.vbNewLine + ""
          + Constants.vbNewLine + ""
      , Constants.vbOKOnly, "Can't Update!"); rtAnswer = Constants.vbYes;
        }

        AskAbout = rtAnswer; // vbYes ' = 1
    }

    public fmTest2 Using(Inventor.Document AiDoc
    )
    {
        /// NEWMETHOD[2021.12.13]
        /// Using -- assign supplied Document
        /// for use in all subsequent calls
        /// to AskAbout without one.
        /// 
        rtAnswer = Constants.vbCancel;

        if (!AiDoc == null)
            ad = AiDoc;

        using ( == this)
        {
        }
    }

    public Inventor.Document Document(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */
    )
    {
        /// NEWMETHOD[2021.12.13]
        /// Document -- return currently active Document
        /// 
        if (AiDoc == null)
            Document = ad;
        else
            Document = this.Using(AiDoc).Document;
    }

    private float fmHtAdjust(long by)
    {
        long cmdTop;

        {
            var withBlock = this;
            withBlock.Height = withBlock.Height + by;

            withBlock.cmdLt.Top = withBlock.Height - dmFmHt2cmdTop;
            withBlock.cmdCt.Top = withBlock.cmdLt.Top;
            withBlock.cmdRt.Top = withBlock.cmdLt.Top;

            fmHtAdjust = withBlock.Height;
        }
    }

    private float lblHtAdjust(MSForms.Label lb, string tx
    )
    {
        MSForms.Control ct;
        bool au;
        float wd;
        float ht;

        ct = lb;
        {
            var withBlock = ct;
            wd = withBlock.Width;
            ht = withBlock.Height;

            {
                var withBlock1 = lb;
                au = withBlock1.AutoSize;
                withBlock1.Caption = tx;
                withBlock1.AutoSize = true;
                ct.Width = dmLbMsWd;
                withBlock1.AutoSize = au;
            }

            lblHtAdjust = Int(withBlock.Height - ht);
        }
    }

    private void cmdCt_Click()
    {
        rtAnswer = Constants.vbNo;
        this.Hide();
    }

    private void cmdLt_Click()
    {
        rtAnswer = Constants.vbYes;
        this.Hide();
    }

    private void cmdRt_Click()
    {
        rtAnswer = Constants.vbCancel;
        this.Hide();
    }

    private void UserForm_Initialize()
    {
        /// 
        {
            var withBlock = this;
            // dmFmHt = .Height
            // dmFmWd = .Width
            {
                var withBlock1 = withBlock.lbMsg;
                // dmLbMsHt = .Height
                dmLbMsWd = withBlock1.Width;
            }
            dmFmHt2cmdTop = withBlock.Height - withBlock.cmdLt.Top;
        }
        // dmDfFmMsWd = dmFmWd - dmLbMsWd
        // dmDfFmMsHt = dmFmHt - dmLbMsHt
        rtAnswer = Constants.vbCancel;
    }

    private void UserForm_Click()
    {
    }

    private void UserForm_Layout()
    {
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        this.Hide();
    }

    private void UserForm_Terminate()
    {
    }
}