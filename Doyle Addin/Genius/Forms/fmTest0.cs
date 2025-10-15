class fmTest0 : Form
{
    private var VB_Name = "fmTest0";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

    private var VB_Exposed = false;

    public long ft0g0f0(stdole.StdPicture im)
    {
        {
            var withBlock = this;
            withBlock.imTNail.Picture = im;
            withBlock.Show(1);
        }
        return 0;
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        this.Hide();
    }
}