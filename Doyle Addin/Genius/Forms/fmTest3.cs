class fmTest3 : Form
{
    private var VB_Name = "fmTest3";

    private var VB_GlobalNameSpace = false;

    private var VB_Creatable = false;

    private var VB_PredeclaredId = true;

    private var VB_Exposed = false;

    private void UserForm_Click()
    {
    }

    private void UserForm_DblClick(MSForms.ReturnBoolean Cancel)
    {
        Debug.Print(ft3g1f0(frmShp, "QUICK BROWN FOX JUMPED LAZY DOG"));
    }

    private void UserForm_Initialize()
    {
        long dx;

        long tp = 18;
        const long gp = 0;

        for (dx = 1; dx <= 3; dx++)
        {
            var cp = "CB" + Convert.ToString(dx);

            Control ct = frmShp.Controls.Add("Forms.CheckBox.1", cp, true);
            {
                var withBlock = ct;
                withBlock.Height = 18;
                withBlock.Width = 96;
                withBlock.Left = 18;
                withBlock.Top = tp;

                tp = tp + withBlock.Height + gp;
            }

            CheckBox cb = ct;
            {
                var withBlock = cb;
                withBlock.Caption = cp;
            }
        }
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode)
    {
        Cancel = 1;
        this.Hide();

        Debug.Print(ft3g0f0(frmShp, "Check"));
        CheckBox cb;
        foreach (var cb in frmShp.Controls)
        {
            // Stop
            if (cb.Value)
                Debug.Print(cb.Caption);
        }
    }

    private string ft3g0f0(Frame src, string fdName
    )
    {
        string rt;

        foreach (Control ct in src.Controls)
        {
            if (ct is not CheckBox) continue;
            CheckBox cb = ct;
            if (!cb.Value) continue;
            if (Strings.Len(rt) > 0)
                // rt = rt & " OR "
                rt = rt + ", ";
            // rt = rt & fdName & " = '" & cb.Caption & "'"
            rt = rt + "'" + cb.Caption + "'";
        }

        // ft3g0f0 = "(" & rt & ")"
        return fdName + " IN (" + rt + ")";
    }

    private long ft3g1f0(Frame frm, string ls, string bk = " "
    )
    {
        string rt;
        long ct;

        long tp = 18;
        const long gp = 0;

        {
            var withBlock = frm.Controls // .Remove
                ;
            while (withBlock.Count > 0)
                withBlock.Remove(0);

            foreach (var cp in Split(ls, bk))
            {
                if (Len(cp) <= 0) continue;
                Control ctrl = withBlock.Add("Forms.CheckBox.1", cp, true);
                {
                    var withBlock1 = ctrl;
                    withBlock1.Height = 18;
                    withBlock1.Width = 96;
                    withBlock1.Left = 18;
                    withBlock1.Top = tp;

                    tp = tp + withBlock1.Height + gp;
                }

                CheckBox cb = ctrl;
                {
                    var withBlock1 = cb;
                    withBlock1.Caption = cp;
                }

                ct = ct + 1;
            }
        }

        // ft3g1f0 = "(" & rt & ")"
        // ft3g1f0 = ls & " IN (" & rt & ")"
        return ct;
    }
}