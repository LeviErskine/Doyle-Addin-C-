class SurroundingClass
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmTestStockSel0";
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
    private ADODB.Recordset rsItm;

    private void lbxFamily_Change()
    {
        {
            var withBlock = this;
            rsItm.Filter = "Family = '" + withBlock.lbxFamily.Value + "'";
            withBlock.lbxItem.List = m0g3f1(rsItm);
        }
    }

    private void UserForm_Initialize()
    {
        cn = cnGnsDoyle();
        {
            var withBlock = cn;
            rsFam = withBlock.Execute(Join(Array("select Family, Description1", "from vgMfiFamilies", "where FamilyGroup = 'RAW'"
   ), " "));
            rsItm = withBlock.Execute(Join(Array("Select I.Family, I.Item, I.Description1", "From vgMfiItems as I", "Inner Join vgMfiFamilies as F", "On I.Family = F.Family", "Where F.FamilyGroup = 'RAW'"
    ), " "));
        }

        this.lbxFamily.List = m0g3f1(rsFam);
    }
}