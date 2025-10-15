namespace Doyle_Addin.Genius.Forms;

class fmTestStockSel0 : Form
{
    private var VB_Name = "fmTestStockSel0";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

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

    private void InitializeComponent()
    {
        cn = cnGnsDoyle();
        {
            var withBlock = cn;
            rsFam = withBlock.Execute(Join(new[]
            {
                "select Family, Description1", "from vgMfiFamilies",
                "where FamilyGroup = 'RAW'"
            }, " "));
            rsItm = withBlock.Execute(Join(new[]
            {
                "Select I.Family, I.Item, I.Description1",
                "From vgMfiItems as I",
                "Inner Join vgMfiFamilies as F", "On I.Family = F.Family", "Where F.FamilyGroup = 'RAW'"
            }, " "));
        }

        this.lbxFamily.List = m0g3f1(rsFam);
    }
}