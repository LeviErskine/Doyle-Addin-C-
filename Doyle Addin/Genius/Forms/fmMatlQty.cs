class fmMatlQty
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    private var VB_Name = "fmMatlQty";
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

    private void cmdCancel_Click()
    {
        Sent?.Invoke(Constants.vbCancel);
    }

    private void cmdOk_Click()
    {
        Sent?.Invoke(Constants.vbOK);
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        cmdCancel_Click();
    }
}