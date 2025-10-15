using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Forms;

public class fmMatlQty : Form
{

    public event SentEventHandler Sent;

    public delegate void SentEventHandler(MsgBoxResult Signal);

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        lblPartNumber = new System.Windows.Forms.Label();
        imThmNail = new System.Windows.Forms.PictureBox();
        lblPartInfo = new System.Windows.Forms.Label();
        lblMatlNumber = new System.Windows.Forms.Label();
        lblMatlInfo = new System.Windows.Forms.Label();
        lblMatlQty = new System.Windows.Forms.Label();
        txbMatlQty = new System.Windows.Forms.TextBox();
        cbxUnitQty = new System.Windows.Forms.ComboBox();
        lbxMatlQty = new System.Windows.Forms.ListBox();
        cmdOK = new System.Windows.Forms.Button();
        cmdCancel = new System.Windows.Forms.Button();
        ((System.ComponentModel.ISupportInitialize)imThmNail).BeginInit();
        SuspendLayout();
        // 
        // lblPartNumber
        // 
        lblPartNumber.Location = new System.Drawing.Point(12, 9);
        lblPartNumber.Name = "lblPartNumber";
        lblPartNumber.Size = new System.Drawing.Size(100, 23);
        lblPartNumber.TabIndex = 0;
        lblPartNumber.Text = "<partNumber>";
        // 
        // imThmNail
        // 
        imThmNail.Location = new System.Drawing.Point(12, 35);
        imThmNail.Name = "imThmNail";
        imThmNail.Size = new System.Drawing.Size(314, 171);
        imThmNail.TabIndex = 1;
        imThmNail.TabStop = false;
        // 
        // lblPartInfo
        // 
        lblPartInfo.Location = new System.Drawing.Point(12, 209);
        lblPartInfo.Name = "lblPartInfo";
        lblPartInfo.Size = new System.Drawing.Size(295, 30);
        lblPartInfo.TabIndex = 2;
        lblPartInfo.Text = "label1";
        // 
        // lblMatlNumber
        // 
        lblMatlNumber.Location = new System.Drawing.Point(12, 239);
        lblMatlNumber.Name = "lblMatlNumber";
        lblMatlNumber.Size = new System.Drawing.Size(100, 23);
        lblMatlNumber.TabIndex = 3;
        lblMatlNumber.Text = "label1";
        // 
        // lblMatlInfo
        // 
        lblMatlInfo.Location = new System.Drawing.Point(12, 262);
        lblMatlInfo.Name = "lblMatlInfo";
        lblMatlInfo.Size = new System.Drawing.Size(100, 23);
        lblMatlInfo.TabIndex = 4;
        lblMatlInfo.Text = "label1";
        // 
        // lblMatlQty
        // 
        lblMatlQty.Location = new System.Drawing.Point(12, 285);
        lblMatlQty.Name = "lblMatlQty";
        lblMatlQty.Size = new System.Drawing.Size(100, 23);
        lblMatlQty.TabIndex = 5;
        lblMatlQty.Text = "label1";
        // 
        // txbMatlQty
        // 
        txbMatlQty.Location = new System.Drawing.Point(12, 311);
        txbMatlQty.Name = "txbMatlQty";
        txbMatlQty.Size = new System.Drawing.Size(181, 23);
        txbMatlQty.TabIndex = 6;
        // 
        // cbxUnitQty
        // 
        cbxUnitQty.FormattingEnabled = true;
        cbxUnitQty.Location = new System.Drawing.Point(12, 340);
        cbxUnitQty.Name = "cbxUnitQty";
        cbxUnitQty.Size = new System.Drawing.Size(121, 23);
        cbxUnitQty.TabIndex = 7;
        // 
        // lbxMatlQty
        // 
        lbxMatlQty.FormattingEnabled = true;
        lbxMatlQty.ItemHeight = 15;
        lbxMatlQty.Location = new System.Drawing.Point(206, 269);
        lbxMatlQty.Name = "lbxMatlQty";
        lbxMatlQty.Size = new System.Drawing.Size(120, 94);
        lbxMatlQty.TabIndex = 8;
        // 
        // cmdOK
        // 
        cmdOK.Location = new System.Drawing.Point(12, 369);
        cmdOK.Name = "cmdOK";
        cmdOK.Size = new System.Drawing.Size(75, 23);
        cmdOK.TabIndex = 9;
        cmdOK.Text = "button1";
        cmdOK.UseVisualStyleBackColor = true;
        cmdOK.Click += cmdOK_Click;
        // 
        // cmdCancel
        // 
        cmdCancel.Location = new System.Drawing.Point(122, 369);
        cmdCancel.Name = "cmdCancel";
        cmdCancel.Size = new System.Drawing.Size(62, 32);
        cmdCancel.TabIndex = 10;
        cmdCancel.Text = "button2";
        cmdCancel.UseVisualStyleBackColor = true;
        cmdCancel.Click += cmdCancel_Click;
        // 
        // fmMatlQty
        // 
        ClientSize = new System.Drawing.Size(350, 421);
        Controls.Add(cmdCancel);
        Controls.Add(cmdOK);
        Controls.Add(lbxMatlQty);
        Controls.Add(cbxUnitQty);
        Controls.Add(txbMatlQty);
        Controls.Add(lblMatlQty);
        Controls.Add(lblMatlInfo);
        Controls.Add(lblMatlNumber);
        Controls.Add(lblPartInfo);
        Controls.Add(imThmNail);
        Controls.Add(lblPartNumber);
        MaximizeBox = false;
        MinimizeBox = false;
        ShowIcon = false;
        Text = "Set/Verify Material Quantity";
        ((System.ComponentModel.ISupportInitialize)imThmNail).EndInit();
        ResumeLayout(false);
        PerformLayout();
    }

    private Button cmdOK;
    private Button cmdCancel;

    private ListBox lbxMatlQty;

    private ComboBox cbxUnitQty;

    private TextBox txbMatlQty;

    private Label lblMatlQty;

    private Label lblMatlInfo;

    private Label lblMatlNumber;

    private System.Windows.Forms.Label lblPartInfo;

    private PictureBox imThmNail;

    private Label lblPartNumber;

    private void cmdOK_Click(object sender, EventArgs e)
    {
        Sent?.Invoke(Constants.vbOK);
    }

    private void cmdCancel_Click(object sender, EventArgs e)
    {
        Sent?.Invoke(Constants.vbCancel);
    }
    
}