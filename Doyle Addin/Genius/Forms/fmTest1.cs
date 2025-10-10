class SurroundingClass
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmTest1";
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_GlobalNameSpace = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Creatable = false;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_PredeclaredId = true;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Exposed = false;


    private ADODB.Connection cn;
    private ADODB.Recordset rsFam;
    private ADODB.Recordset rsPrt;
    private ADODB.Recordset rsItm;

    private Scripting.Dictionary dc;
    private Inventor.Document ad;

    private Inventor.PropertySet psDsn;
    private Inventor.PropertySet psUsr;

    private Scripting.Dictionary dcDsn;
    private Scripting.Dictionary dcUsr;

    private Inventor.Property prFam;
    private Inventor.Property prStk;
    private Inventor.Property prThk;

    public VbMsgBoxResult AskAbout(Inventor.Document AiDoc, string txMsg = ""
    )
    {
        stdole.IPictureDisp pc;
        VbMsgBoxResult ck;
        string pn;    // part number
        string sn;    // material (stock) number
        string sf;    // material (stock) family
        string pd;    // part description
        float df;

        ad = AiDoc;
        {
            var withBlock = ad;
            ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 1060
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error Resume Next

 */
            Information.Err.Clear();
            pc = withBlock.Thumbnail;
            if (Information.Err.Number == 0)
            {
            }
            else
            {
            }
            ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 1214
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorGoToStatement(OnErrorGoToStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorGoToStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        On Error GoTo 0

 */
            psDsn = withBlock.PropertySets(gnDesign);
            psUsr = withBlock.PropertySets(gnCustom);

            dcDsn = dcAiPropsInSet(psDsn);
            dcUsr = dcAiPropsInSet(psUsr);

            prFam = psDsn.Item(pnFamily);

            // '  Get Sheet Metal Thickness Property
            prThk = aiPropShtMetalThickness(ad);
            // '  NOTE: Function returns Nothing
            // '      if Part is NOT Sheet Metal!

            {
                var withBlock1 = dcUsr;
                if (withBlock1.Exists(pnRawMaterial))
                    prStk = psUsr.Item(pnRawMaterial);
                else
                {
                    ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 1751
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
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(pnRawMaterial, prStk);
                    ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 2019
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

            /// REV[2022.04.28.1615]
            /// added initializtion of Dictionary dc
            /// with initial raw material setting.
            /// sn now assigned from the Dictionary.
            /// NOTE: probably want to  initial
            /// values in a separate "recovery"
            /// Dictionary to be restored if
            /// the User chooses to cancel.
            /// Also, see function/method dcUpd.
            /// looks like it gets called when
            /// something changes. Easy to miss!
            dc.Item(pnRawMaterial) = prStk.Value;
            sn = dc.Item(pnRawMaterial);
            pn = psDsn.Item(pnPartNum).Value;
            pd = psDsn.Item(pnDesc).Value;
        }

        {
            var withBlock = this;
            withBlock.Caption = "Please Review Part Number: " + pn;

            if (pc == null)
            {
            }
            else
                withBlock.imThmNail.Picture = pc;

            {
                var withBlock1 = withBlock.lbMsg;
                withBlock1.Caption = pn + ": " + pd
                   + Constants.vbNewLine + txMsg + Interaction.IIf(Strings.Len(txMsg) > 0, Constants.vbNewLine, "")
                   + ft1g0f0(pnCatWebLink, psDsn.Item(pnCatWebLink)) + Constants.vbNewLine
                   + ft1g0f0(pnMaterial, psDsn.Item(pnMaterial)) + Constants.vbNewLine
                   + ft1g0f0(pnThickness, prThk) + Constants.vbNewLine
                   + "";
                ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 3260
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                '& vbNewLine _
                & vbNewLine & pnThickness & ": " & psUsr.Item(pnThickness).Value _

 */
            }
            df = mdl1g1f2(withBlock.lbMsg);
            if (df > 0)
            {
                mdl1g1f3.lbMtFamily(null/* Conversion error: Set to default value for this argument */, 0, df);
                mdl1g1f3.lbxFamily(null/* Conversion error: Set to default value for this argument */, 0, df);
            }

            withBlock.dbFamily.Value = prFam.Value;

            if (Strings.Len(sn) > 0)
            {
                {
                    var withBlock1 = cn.Execute("select Family from vgMfiItems where Item = '"
+ Replace(sn, "'", "''") + "'"
);
                    /// REV[2022.08.19.1359]
                    /// temporarily replacing direct use of sn
                    /// with call to Replace single quotes
                    /// in string with doubled single quotes
                    /// (NOT double quotes!) to "escape" the
                    /// character in a string value.
                    /// '
                    /// will ultimately want to produce some
                    /// sort of 'handler' to preprocess values
                    /// for use in SQL commands to avoid errors
                    /// that arise from this sort of thing.
                    if (withBlock1.BOF | withBlock1.EOF)
                        sf = "";
                    else
                        sf = withBlock1.Fields(0).Value;
                }

                if (Strings.Len(sf) == 0)
                    // EITHER doesn't have a Family,
                    // OR is not (yet) in Genius.
                    // SO, let's just ...
                    sf = "DSHEET"; // as a default!
                else
                {
                }
                ;/* Cannot convert OnErrorResumeNextStatementSyntax, CONVERSION ERROR: Conversion for OnErrorResumeNextStatement not implemented, please report this issue in 'On Error Resume Next' at character 5379
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitOnErrorResumeNextStatement(OnErrorResumeNextStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.OnErrorResumeNextStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 

            On Error Resume Next

 */
                Information.Err.Clear();
                withBlock.lbxFamily.Value = sf;
                if (Information.Err.Number)
                {
                    Debug.Print("FAILED TO  MATERIAL FAMILY " + sf);
                    ck = MsgBox(Join(Array("Part Number " + pn, "uses Material " + sn
, "which is a" + IIf(InStr(1, "AEIOU", UCase(Left(sf, 1))), "n ", " "
) + sf + " Item."
, ""
, "This interface does not presently"
, "support Materials from this Family."
, ""
, "You might not be able to find the correct"
, "Material for this Part, and might wish"
, "to avoid changing it here."
, ""
, "Do you wish to proceed anyway?"
), Constants.vbNewLine), Constants.vbYesNoCancel + Constants.vbExclamation + Constants.vbDefaultButton2, "Material Family not Supported"
);
                    if (ck == Constants.vbCancel)
                        System.Diagnostics.Debugger.Break();
                }
                else
                {
                    Information.Err.Clear();
                    withBlock.lbxItem.Value = sn;

                    /// REV[2022.05.06.1329]
                    /// added intermediate error handler
                    /// to capture failure in Material
                    /// Family selector to adopt new Value.
                    /// it re-implements process of Event
                    /// handler Sub lbxFamily_Change
                    /// against variable 'sf' directly
                    /// in an effort to force population
                    /// of Material list.
                    if (Information.Err.Number)
                    {
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        Information.Err.Clear();
                        rsItm.Filter = "Family = '" + sf + "'";
                        withBlock.lbxItem.List = m0g3f1(rsItm);
                        withBlock.lbxItem.Value = sn;
                    }
                    /// something MIGHT have happened
                    /// to prevent normal Value update
                    /// when lbxFamily is  above.
                    /// further investigation may be
                    /// warranted.

                    if (Information.Err.Number)
                    {
                        Debug.Print("FAILED TO  MATERIAL " + sn);
                        ck = MsgBox(Join(Array("!!WARNING!!", ""
, "Active Material " + sn
, "for Part Number " + pn
, "could NOT be selected,"
, "and might be unavailable."
, ""
, "You might wish to avoid"
, "making Material changes"
, "to this Part here."
, ""
, "Do you wish to proceed anyway?"
), Constants.vbNewLine), Constants.vbYesNoCancel + Constants.vbExclamation, "Active Material Not Found!"
);
                        if (ck == Constants.vbCancel)
                            System.Diagnostics.Debugger.Break();
                    }
                    else
                    {
                        ck = Constants.vbYes;
                        lbxItem_Change();
                        // lbxFamily_Change
                        rsItm.Filter = "Family = '" + sf + "'";
                        withBlock.lbxItem.List = m0g3f1(rsItm);
                    }
                }
                ;/* Cannot convert OnErrorGoToStatementSyntax, CONVERSION ERROR: Conversion for OnErrorGoToZeroStatement not implemented, please report this issue in 'On Error GoTo 0' at character 8792
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
                ck = Constants.vbYes;

            if (ck == Constants.vbYes)
                withBlock.Show(1);
        }
        AskAbout = ck; // vbYes ' = 1
    }

    private string ft1g0f0(string pn, Inventor.Property pr
    )
    {
        if (pr == null)
            ft1g0f0 = "";
        else
            ft1g0f0 = Constants.vbNewLine + pn + ": " + pr.Value;
    }

    private void dbFamily_Change()
    {
        Debug.Print(dcUpd(pnFamily, dbFamily.Value));
    }
    // Me.lbxItem.ColumnWidths = "84 pt;6 pt;180 pt"
    // Me.lbxItem.ColumnWidths = "84 pt;48 pt;216 pt"

    private void lbMsg_DblClick(MSForms.ReturnBoolean Cancel)
    {
        System.Diagnostics.Debugger.Break();
    }

    private void lbxFamily_Change()
    {
        {
            var withBlock = this;
            rsItm.Filter = "Family = '" + withBlock.lbxFamily.Value + "'";
            withBlock.lbxItem.List = m0g3f1(rsItm);
        }
    }

    public Scripting.Dictionary ItemData()
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, withBlock.Item(ky));
        }
        ItemData = rt;
    }

    public Scripting.Dictionary Synch()
    {
        {
            var withBlock = dc;
            if (withBlock.Exists(pnFamily))
                prFam.Value = dc.Item(pnFamily);
            if (withBlock.Exists(pnRawMaterial))
                prStk.Value = dc.Item(pnRawMaterial);
        }

        Synch = this.ItemData();
    }

    private string dcUpd(string ky, Variant vl)
    {
        string rt;

        if (IsNull(vl))
            dcUpd = dcUpd(ky, "");
        else
        {
            var withBlock = dc;
            if (withBlock.Exists(ky))
            {
                rt = System.Convert.ToHexString(withBlock.Item(ky));
                withBlock.Item(ky) = vl;
                dcUpd = "CHANGE[" + ky + "] FROM '" + rt
                + "' TO '" + System.Convert.ToHexString(withBlock.Item(ky)) + "'";
            }
            else
            {
                withBlock.Add(ky, vl);
                dcUpd = "[" + ky + "] TO '"
+ System.Convert.ToHexString(withBlock.Item(ky)) + "'";
            }
        }
    }

    private void lbxItem_Change()
    {
        Debug.Print(dcUpd(pnRawMaterial, lbxItem.Value));
    }

    private void UserForm_Initialize()
    {
        dc = new Scripting.Dictionary();
        cn = cnGnsDoyle();

        {
            var withBlock = cn;
            ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 10904
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
        ' rsFam = .Execute(Join(Array( _
            "select Family, Description1", _
            "from vgMfiFamilies", _
            "order by Family" _
        ), " ")) ', _

 */
            ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 11080
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
            "where FamilyGroup = 'RAW'"

 */         rsPrt = withBlock.Execute(Join(Array("select Family, FamilyGroup, Description1", "from vgMfiFamilies", "order by Family"
), Constants.vbNewLine)); // , _
            ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 11306
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 41
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.VisitEmptyStatement(EmptyStatementSyntax node)
   at Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax.Accept[TResult](VisualBasicSyntaxVisitor`1 visitor)
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
            "where FamilyGroup = 'PARTS'"

 */         rsItm = withBlock.Execute(Join(Array("Select I.Item, I.Family, I.Description1, I.Specification1", "From vgMfiItems as I", "Inner Join vgMfiFamilies as F", "On I.Family = F.Family", "Where F.FamilyGroup = 'RAW'", "order by Family, Item"
), " "));
        }

        {
            var withBlock = this;
            rsPrt.Filter = "FamilyGroup = 'RAW'";
            withBlock.lbxFamily.List = m0g3f1(rsPrt); // rsFam

            rsPrt.Filter = "FamilyGroup = 'PARTS'";
            withBlock.dbFamily.List = m0g3f1(rsPrt);
        }
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        this.Hide();
    }

    private void UserForm_Terminate()
    {
        cn.Close();
        cn = null;
    }
}