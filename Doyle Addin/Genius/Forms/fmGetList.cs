Class fmGetList
{
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    Private var VB_Name = "fmGetList";
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    Private var VB_GlobalNameSpace = False;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    Private var VB_Creatable = False;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    Private var VB_PredeclaredId = True;
    /* TODO ERROR: Skipped SkippedTokensTrivia */
    Private var VB_Exposed = False;



    // Event CheckOut(Cancel As Long)

    Private String bg;
    Private String rt;

    Public String AskUser(String Using = ""
    )
    {
        bg = ; txIn.Value = bg;     // initialize text box
        Show(vbModal); AskUser = rt;        // return final result
    }

    Private void CheckOut(Long NoChg)
    {
        VbMsgBoxResult ck;

        If (NoChg == 0)
        {
            ck = MsgBox("Use this List?", Constants.vbYesNo + Constants.vbQuestion, "Confirm"
);
            If (ck == Constants.vbYes)
                rt = txIn.Value;
        }
        Else
        {
            ck = MsgBox("Cancel this Entry?", Constants.vbYesNo + Constants.vbQuestion, "Cancel"
      );
            If (ck == Constants.vbYes)
                rt = bg;
        }

        If (ck == Constants.vbYes)
            this.Hide();
    }

    Private void cmdCancel_Click()
    {
        CheckOut(1);
    }

    Private void cmdOk_Click()
    {
        CheckOut(0);
    }

    Private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        CheckOut(1);
    }

    Private void UserForm_Initialize()
    {
    }

    Private void UserForm_Terminate()
    {
    }
}