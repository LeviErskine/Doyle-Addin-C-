using Microsoft.VisualBasic;

public class fmSelectorList : Form
{
    private var VB_Name = "fmSelectorList";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

    private var VB_Exposed = false;

    private const string tagSelected = "%%%";

    private string msCancelHead;
    private string msCancelMain;
    private string msNoSelHead;
    private string msNoSelMain;
    private string msOkHead;

    private string msOkMain;
    // Private msCancelMain As String
    // Private msCancelMain As String

    public fmSelectorList SetMsgCancel(string Using)
    {
        msCancelMain = ;
        return this;
    }

    public fmSelectorList SetMsgNoSelection(string Using)
    {
        msNoSelMain = Using;

        return this;
    }

    public fmSelectorList SetMsgOK(string Using)
    {
        msOkMain = Using;

        return this;
    }

    public fmSelectorList SetHdrCancel(string Using)
    {
        msCancelHead = Using;

        return this;
    }

    public fmSelectorList SetHdrNoSelection(string Using)
    {
        msNoSelHead = Using;

        return this;
    }

    public fmSelectorList SetHdrOK(string Using)
    {
        msOkHead = Using;

        return this;
    }

    public fmSelectorList SelectIfIn(string Using)
    {
        long dx;

        {
            var withBlock = this.lsbSelection;
            dx = withBlock.ListIndex;

            Information.Err().Clear();
                .Value = Using

            if (!Information.Err().Number) return this;
            withBlock.ListIndex = dx;
            // .Value = ""
            Information.Err().Clear();
        }

        return this;
    }

    public fmSelectorList WithList(dynamic Using)
    {
        if IsArray(Using)
        Me.lsbSelection.List = Using;
        else
        //Stop
        Debug.Print(""); //Breakpoint Landing
        end if

        return this;
    }

    private void btnCancel_Click()
    {
        // '
        if (MessageBox.Show(msCancelMain, Constants.vbYesNo, msCancelHead) == Constants.vbYes)
        {
            this.lsbSelection.ListIndex = -1;
            Hide();
        }
    }

    private void btnOk_Click()
    {
        // '
        long mx;
        long dx;
        long ct;

        {
            var withBlock = this.lsbSelection;
            VbMsgBoxResult ck;
            if (withBlock.MultiSelect == fmMultiSelectSingle)
            {
                if (withBlock.ListIndex < 0)
                {
                    ck = MessageBox.Show(msNoSelMain, Constants.vbYesNo, msNoSelHead);
                    if (ck == Constants.vbYes)
                        Hide();
                }
                else
                {
                    ck = MessageBox.Show(Join(Split(msOkMain, tagSelected), withBlock.Value), Constants.vbYesNoCancel,
                        msOkHead
                    );
                    if (ck == Constants.vbYes)
                        Hide();
                    else if (ck == Constants.vbCancel)
                    {
                        withBlock.ListIndex = -1;
                        Hide();
                    }
                }
            }
            else
            {
                string ls = lbxPickedStr(this.lsbSelection, Constants.vbCrLf);

                // ct = 0
                // mx = .ListCount - 1
                // For dx = 0 To mx
                // If .Selected(dx) Then ct = 1 + ct
                // Next

                if (Strings.Len(ls) > 0)
                {
                    ck = MessageBox.Show(Join(Split(msOkMain, tagSelected), Constants.vbCrLf + ls + Constants.vbCrLf
                        ), Constants.vbYesNoCancel, msOkHead
                    );
                    if (ck == Constants.vbYes)
                        Hide();
                    else if (ck == Constants.vbCancel)
                    {
                        withBlock.ListIndex = -1;
                        Hide();
                    }
                }
                else
                {
                    ck = MessageBox.Show(msNoSelMain, Constants.vbYesNo, msNoSelHead);
                    if (ck == Constants.vbYes)
                        Hide();
                }
            }
        }
    }

    private void lsbSelection_DblClick(bool Cancel)
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
        msNoSelMain = Join(new[]
        {
            "Do you wish to cancel?", "(Click NO to return to list)"
        }, Constants.vbCrLf);
        msOkMain = Join(new[]
        {
            "Current selection is: ", tagSelected, "(Click CANCEL to quit with no selection)"
        }, Constants.vbCrLf);
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        // 
        Cancel = 1;
        btnCancel_Click();
    }

    public string GetReply(dynamic List = , dynamic As = string == "%$#@"
    )
    {
        const string rt = "";

        with Me.WithList(List).SelectIfIn(Default)
            //.lsbSelection.List = lsWorkbooks()
            .Show(1);
        if
            .lsbSelection.MultiSelect = fmMultiSelectSingle
        rt =  .lsbSelection.Text;
        else
        rt = lbxPickedStr(.lsbSelection, vbCrLf)
        end if
        end with

        return rt;
    }

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        SuspendLayout();
        // 
        // fmSelectorList
        // 
        ClientSize = new System.Drawing.Size(324, 281);
        ResumeLayout(false);
    }
}