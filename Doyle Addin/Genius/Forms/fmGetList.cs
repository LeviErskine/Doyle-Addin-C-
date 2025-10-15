using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Forms;

class fmGetList : Form
{
    // Event CheckOut(Cancel As Long)

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        lbTxIn = new System.Windows.Forms.Label();
        txIn = new System.Windows.Forms.TextBox();
        cmdCancel = new System.Windows.Forms.Button();
        cmdOk = new System.Windows.Forms.Button();
        SuspendLayout();
        // 
        // lbTxIn
        // 
        lbTxIn.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)0));
        lbTxIn.Location = new System.Drawing.Point(12, 12);
        lbTxIn.Name = "lbTxIn";
        lbTxIn.Size = new System.Drawing.Size(121, 15);
        lbTxIn.TabIndex = 0;
        lbTxIn.Text = "Paste or Type List Here";
        lbTxIn.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
        // 
        // txIn
        // 
        txIn.Location = new System.Drawing.Point(12, 30);
        txIn.Multiline = true;
        txIn.Name = "txIn";
        txIn.Size = new System.Drawing.Size(150, 300);
        txIn.TabIndex = 1;
        // 
        // cmdCancel
        // 
        cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
        cmdCancel.Location = new System.Drawing.Point(12, 336);
        cmdCancel.Name = "cmdCancel";
        cmdCancel.Size = new System.Drawing.Size(75, 23);
        cmdCancel.TabIndex = 2;
        cmdCancel.Text = "Cancel";
        cmdCancel.UseVisualStyleBackColor = true;
        cmdCancel.Click += cmdCancel_Click;
        // 
        // cmdOk
        // 
        cmdOk.DialogResult = System.Windows.Forms.DialogResult.OK;
        cmdOk.Location = new System.Drawing.Point(85, 336);
        cmdOk.Name = "cmdOk";
        cmdOk.Size = new System.Drawing.Size(77, 23);
        cmdOk.TabIndex = 3;
        cmdOk.Text = "OK";
        cmdOk.UseVisualStyleBackColor = true;
        cmdOk.Click += cmdOk_Click;
        // 
        // fmGetList
        // 
        AcceptButton = cmdOk;
        CancelButton = cmdCancel;
        ClientSize = new System.Drawing.Size(172, 366);
        Controls.Add(cmdOk);
        Controls.Add(cmdCancel);
        Controls.Add(txIn);
        Controls.Add(lbTxIn);
        MaximizeBox = false;
        MinimizeBox = false;
        ShowIcon = false;
        Text = "List Entry";
        FormClosed += fmGetList_FormClosed;
        ResumeLayout(false);
        PerformLayout();
    }

    private System.Windows.Forms.Button cmdCancel;
    private System.Windows.Forms.Button cmdOk;
    private System.Windows.Forms.TextBox txIn;
    private System.Windows.Forms.Label lbTxIn;
    private string bg;
    private string rt;

    public string AskUser(string Using = "")
    {
        bg = txIn.Text = bg; // initialize text box
        Show(Modal);
        return rt; // return final result
    }

    private void CheckOut(bool NoChg)
    {
        if (NoChg) return;
        var ck = MessageBox.Show(@"Use this List?", @"Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (ck == (DialogResult)Constants.vbYes)
        {
            rt = txIn.Text;
        }
        else
            ck = MessageBox.Show(@"Cancel this Entry?", @"Cancel", MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

        if (ck != (DialogResult)Constants.vbYes) return;
        rt = bg;
        Hide();
    }

    private void cmdCancel_Click()
    {
        CheckOut(true);
    }

    private void cmdOk_Click()
    {
        CheckOut(false);
    }

    private void fmGetList_FormClosed(object sender, FormClosedEventArgs e, int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        CheckOut(true);
    }
}