using Microsoft.VisualBasic;

public class fmSelectorV2 : Form
{
    private var VB_Name = "fmSelectorV2";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

    private var VB_Exposed = false;

    private const string tagSelected = "%%%";
    private Dictionary dc;

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
        msCancelMain = ;
        return this;
    }

    public fmSelectorV2 SetMsgNoSelection(string Using)
    {
        msNoSelMain = Using;

        return this;
    }

    public fmSelectorV2 SetMsgOK(string Using)
    {
        msOkMain = Using;

        return this;
    }

    public fmSelectorV2 SetHdrCancel(string Using)
    {
        msCancelHead = Using;

        return this;
    }

    public fmSelectorV2 SetHdrNoSelection(string Using)
    {
        msNoSelHead = Using;

        return this;
    }

    public fmSelectorV2 SetHdrOK(string Using)
    {
        msOkHead = Using;

        return this;
    }

    public fmSelectorV2 SelectIfIn(string Using)
    {
        long dx;

        {
            var withBlock = this.lsbSelection;
            dx = withBlock.ListIndex;

            Information.Err().Clear();
                .Value = Using;

            if (!Information.Err().Number) return this;
            withBlock.ListIndex = dx;
            // .Value = ""
            Information.Err().Clear();
        }

        return this;
    }

    public fmSelectorV2 WithList(dynamic Using)
    {
        string it;

        if (Using is not null) ;
        {
            this.lsbSelection.List = new[] { "<no items>" };
            return this;
        }
        //Me.lsbSelection.List = Using
        if (Using is Dictionary)
        {
            {
                dc = Using;

                if (Using is NameValueMap)
                {
                    dc = Using;

                    if (Using is Array)
                    {
                        dc = dcFromAiNameValMap(obOf(Using))

                        else
                        //Stop
                        {
                            Debug.Print("")
                            //Breakpoint Landing
                            dc = null
                        }
                    }
                }
            }
        }

        if (dc is null) ;
        {
            Me.lsbSelection.List = new[]
            {
                "<no items>"
            };
            else if (dc.Count > 0)
            {
                Me.lsbSelection.List = dc.Keys;
            }
            else
            {
                Me.lsbSelection.List = new[]
                {
                    "<no items>"
                };
            }
        }

        else if (Using is Array)
        {
            dc = new Scripting.Dictionary;
            for each(ky in Using);
            dc.Add(CStr(ky), ky);
            next
            Me.lsbSelection.List = dc.Keys;
        }
        else
        {
            //Stop
            Debug.Print;
            //Breakpoint Landing
        }

        return this;
    }
}

private void btnCancel_Click()
{
    // '
    if (MessageBox.Show(msCancelMain, Constants.vbYesNo, msCancelHead) != Constants.vbYes) return;
    this.lsbSelection.ListIndex = -1;
    this.Hide();
}

private void btnOk_Click()
{
    // '
    long mx;
    long dx;
    long ct;

    {
        var withBlock = this.lsbSelection;
        MsgBoxResult ck;
        if (withBlock.MultiSelect == fmMultiSelectSingle)
        {
            if (withBlock.ListIndex < 0)
            {
                ck = MessageBox.Show(msNoSelMain, Constants.vbYesNo, msNoSelHead);
                if (ck == Constants.vbYes)
                    this.Hide();
            }
            else
            {
                ck = MessageBox.Show(Join(Split(msOkMain, tagSelected), withBlock.Value), Constants.vbYesNoCancel,
                    msOkHead
                );
                switch (ck)
                {
                    case Constants.vbYes:
                        this.Hide();
                        break;
                    case Constants.vbCancel:
                        withBlock.ListIndex = -1;
                        this.Hide();
                        break;
                    case MsgBoxResult.Ok:
                    case MsgBoxResult.Abort:
                    case MsgBoxResult.Retry:
                    case MsgBoxResult.Ignore:
                    case MsgBoxResult.No:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }
        else
        {
            var ls = lbxPickedStr(this.lsbSelection, Constants.vbCrLf);

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
                switch (ck)
                {
                    case Constants.vbYes:
                        this.Hide();
                        break;
                    case Constants.vbCancel:
                        withBlock.ListIndex = -1;
                        this.Hide();
                        break;
                    case MsgBoxResult.Ok:
                    case MsgBoxResult.Abort:
                    case MsgBoxResult.Retry:
                    case MsgBoxResult.Ignore:
                    case MsgBoxResult.No:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
            else
            {
                ck = MessageBox.Show(msNoSelMain, Constants.vbYesNo, msNoSelHead);
                if (ck == Constants.vbYes)
                    this.Hide();
            }
        }
    }
}

private void lsbSelection_Change()
{
    // 
    string ck;

    Information.Err().Clear();
    ck = this.lsbSelection.Value;
    if (Information.Err().Number != 0)
        ck = "";

    {
        var withBlock = dc;
        if (withBlock.Exists(ck))
        {
            Information.Err().Clear();
            tbxView.Value = Convert.ToHexString(withBlock.get_Item(ck));
            if (Information.Err().Number != 0)
                tbxView.Value = "<not printable>";
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
    string rt;
    rt = "";
    ;
    With Me.WithList(List).SelectIfIn(Default)
    '.lsbSelection.List = lsWorkbooks()
        .Show 1
    If.lsbSelection.MultiSelect = fmMultiSelectSingle Then
        rt = .lsbSelection.Text
    Else
        rt = lbxPickedStr(.lsbSelection, vbCrLf)
    End If
    End With
    return rt;
}

private void UserForm_Resize()
{
}
}