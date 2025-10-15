using Microsoft.VisualBasic;

class fmTest2 : Form
{
    private var VB_Name = "fmTest2";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

    private var VB_Exposed = false;

    private Document ad;

    private PropertySet psDsn;
    private PropertySet psUsr;

    private Dictionary dcDsn;
    private Dictionary dcUsr;

    private Property prFam;
    private Property prStk;

    // Private dmFmHt As Long
    // Private dmFmWd As Long
    // 'Private dmLbMsHt As Long
    private long dmLbMsWd;

    // 'Private dmDfFmMsHt As Long
    // 'Private dmDfFmMsWd As Long
    private long dmFmHt2cmdTop;

    private VbMsgBoxResult rtAnswer;

    public VbMsgBoxResult AskAbout(Document AiDoc = null, string txPre = "", string txPost = ""
    )
    {
        // AskAbout -- prompt User for action
        // to take on supplied Document
        // UPDATE[2021.12.13]
        // Document parameter now Optional.
        // will attempt to use previously
        // registered Document when none
        // supplied. Warning/error message
        // will be presented if no Document
        // is registered OR supplied.
        // 
        string sn;
        float dj; // use to adjust
        // form height and positions
        // of command buttons

        rtAnswer = Constants.vbCancel;
        if (!AiDoc == null)
            ad = AiDoc;

        if (ad == null)
        {
            MessageBox.Show("Review or Update requested"
                            + Constants.vbCrLf + "but no Document provided!"
                            + Constants.vbCrLf + ""
                            + Constants.vbCrLf + ""
                , Constants.vbOKOnly, "No Document!");
            rtAnswer = Constants.vbNo;
        }
        else if (aiDocPartFromCCtr(ad) == null)
        {
            // ad = AiDoc
            stdole.IPictureDisp pc;
            string pn;
            string pd;
            {
                var withBlock = ad;
                pc = withBlock.Thumbnail;
                psDsn = withBlock.PropertySets(gnDesign);
                psUsr = withBlock.PropertySets(gnCustom);

                dcDsn = dcAiPropsInSet(psDsn);
                dcUsr = dcAiPropsInSet(psUsr);

                prFam = psDsn.get_Item(pnFamily);
                {
                    var withBlock1 = dcUsr;
                    if (withBlock1.Exists(pnRawMaterial))
                        prStk = psUsr.get_Item(pnRawMaterial);
                    else
                    {
                        Information.Err().Clear();
                        prStk = psUsr.Add("", pnRawMaterial);
                        if (Information.Err().Number)
                        {
                            Debug.Print(Information.Err().Number, Information.Err().Description);
                            Debugger.Break();
                        }
                        else
                            withBlock1.Add(pnRawMaterial, prStk);
                    }
                }

                if (!prStk == null)
                    sn = prStk.Value;
                pn = psDsn.get_Item(pnPartNum).Value;
                pd = psDsn.get_Item(pnDesc).Value;
            }

            {
                var withBlock = this;
                withBlock.Caption = "Please Review Item: " + pn;

                if (pc == null)
                {
                }
                else
                    withBlock.imThmNail.Picture = pc;

                dj = fmHtAdjust(lblHtAdjust(withBlock.lbMsg, Interaction.IIf(Strings.Len(txPre) > 0,
                        txPre + Constants.vbCrLf + Constants.vbCrLf, ""
                    ) + Join(new[]
                      {
                          pn + ": " + pd, pnCatWebLink + ": " + psDsn.get_Item(pnCatWebLink).Value,
                          pnMaterial + ": " + psDsn.get_Item(pnMaterial).Value
                      }, Constants.vbCrLf + Constants.vbCrLf)
                      + Interaction.IIf(Strings.Len(txPost) > 0, Constants.vbCrLf + Constants.vbCrLf + txPost, ""
                      )
                ));
                // .dbFamily.Value = prFam.Value

                withBlock.Show(1);
            }
        }
        else
        {
            MessageBox.Show(ad.DisplayName
                            + Constants.vbCrLf + "is a Content Center part"
                            + Constants.vbCrLf + "and cannot be updated."
                            + Constants.vbCrLf + ""
                            + Constants.vbCrLf + ""
                , Constants.vbOKOnly, "Can't Update!");
            rtAnswer = Constants.vbYes;
        }

        return rtAnswer; // vbYes ' = 1
    }

    public fmTest2 Using(Document AiDoc
    )
    {
        // NEWMETHOD[2021.12.13]
        // Using -- assign supplied Document
        // for use in all subsequent calls
        // to AskAbout without one.
        // 
        rtAnswer = Constants.vbCancel;

        if (!AiDoc == null)
            ad = AiDoc;

        using ( == this);
        {
        }
    }

    public Document Document(Document AiDoc = null
    )
    {
        // NEWMETHOD[2021.12.13]
        // Document -- return currently active Document
        // 
        if (AiDoc == null)
            return ad;
        return Using(AiDoc).Document;
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

            return withBlock.Height;
        }
    }

    private float lblHtAdjust(Label lb, string tx
    )
    {
        Control ct = lb;
        {
            float ht = ct.Height;

            {
                var au = lb.AutoSize;
                lb.Caption = tx;
                lb.AutoSize = true;
                ct.Width = dmLbMsWd;
                lb.AutoSize = au;
            }

            return Int(ct.Height - ht);
        }
    }

    private void cmdCt_Click()
    {
        rtAnswer = Constants.vbNo;
        Hide();
    }

    private void cmdLt_Click()
    {
        rtAnswer = Constants.vbYes;
        Hide();
    }

    private void cmdRt_Click()
    {
        rtAnswer = Constants.vbCancel;
        Hide();
    }

    private void UserForm_Initialize()
    {
        // 
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
        Hide();
    }

    private void UserForm_Terminate()
    {
    }
}