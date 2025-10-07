using System.Diagnostics;

namespace Doyle_Addin.Options
{
    [Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
    public partial class UserOptionsForm : System.Windows.Forms.Form
    {

        // Form overrides dispose to clean up the component list.
        /// <inheritdoc />
        [DebuggerNonUserCode()]
        protected override void Dispose(bool disposing)
        {
            try
            {
                if (disposing && this.components is not null)
                {
                    this.components.Dispose();
                }
            }
            finally
            {
                base.Dispose(disposing);
            }
        }

        // Required by the Windows Form Designer
        private System.ComponentModel.IContainer components;

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        [DebuggerStepThrough()]
        private void InitializeComponent()
        {
            System.Windows.Forms.Label PrintText;
            System.Windows.Forms.Label DXFText;
            this.FutureMsg = new System.Windows.Forms.TextBox();
            this.BtnSave = new System.Windows.Forms.Button();
            BtnSave.Click += new EventHandler(BtnSave_Click);
            this.PrintExportFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.DXFexportFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.Destinationbackground = new System.Windows.Forms.Panel();
            this.DXFexLoc = new System.Windows.Forms.TextBox();
            this.PEXLoc = new System.Windows.Forms.TextBox();
            this.PrintExportLocationButton = new System.Windows.Forms.Button();
            PrintExportLocationButton.Click += new EventHandler(PrintExportLocationButton_Click);
            this.DXFExportLocationButton = new System.Windows.Forms.Button();
            DXFExportLocationButton.Click += new EventHandler(DXFExportLocationButton_Click);
            this.BtnCncl = new System.Windows.Forms.Button();
            BtnCncl.Click += new EventHandler(BtnCncl_Click);
            this.SCBackground = new System.Windows.Forms.Panel();
            this.FeaturesPanel = new System.Windows.Forms.Panel();
            this.ChkObsoletePrint = new System.Windows.Forms.CheckBox();
            PrintText = new System.Windows.Forms.Label();
            DXFText = new System.Windows.Forms.Label();
            this.Destinationbackground.SuspendLayout();
            this.SCBackground.SuspendLayout();
            this.FeaturesPanel.SuspendLayout();
            SuspendLayout();
            // 
            // PrintText
            // 
            PrintText.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            PrintText.Location = new System.Drawing.Point(8, 10);
            PrintText.Name = "PrintText";
            PrintText.Size = new System.Drawing.Size(129, 15);
            PrintText.TabIndex = 9;
            PrintText.Text = "Print Export Location";
            // 
            // DXFText
            // 
            DXFText.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            DXFText.Location = new System.Drawing.Point(8, 36);
            DXFText.Name = "DXFText";
            DXFText.Size = new System.Drawing.Size(129, 17);
            DXFText.TabIndex = 10;
            DXFText.Text = "DXF Export Location";
            // 
            // FutureMsg
            // 
            this.FutureMsg.BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            this.FutureMsg.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.FutureMsg.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.FutureMsg.Enabled = false;
            this.FutureMsg.Font = new System.Drawing.Font("Tahoma", 12.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            this.FutureMsg.ForeColor = System.Drawing.Color.White;
            this.FutureMsg.HideSelection = false;
            this.FutureMsg.Location = new System.Drawing.Point(3, 102);
            this.FutureMsg.Name = "FutureMsg";
            this.FutureMsg.ReadOnly = true;
            this.FutureMsg.Size = new System.Drawing.Size(344, 20);
            this.FutureMsg.TabIndex = 11;
            this.FutureMsg.Text = "More options planned for the future";
            this.FutureMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnSave
            // 
            this.BtnSave.Dock = System.Windows.Forms.DockStyle.Left;
            this.BtnSave.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            this.BtnSave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            this.BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BtnSave.Font = new System.Drawing.Font("Tahoma", 10.0f);
            this.BtnSave.Location = new System.Drawing.Point(3, 3);
            this.BtnSave.Name = "BtnSave";
            this.BtnSave.Padding = new System.Windows.Forms.Padding(3);
            this.BtnSave.Size = new System.Drawing.Size(162, 31);
            this.BtnSave.TabIndex = 5;
            this.BtnSave.Text = "Save";
            this.BtnSave.UseVisualStyleBackColor = true;
            // 
            // PrintExportFolder
            // 
            this.PrintExportFolder.InitialDirectory = @"P:\";
            this.PrintExportFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            this.PrintExportFolder.SelectedPath = @"P:\";
            // 
            // DXFexportFolder
            // 
            this.DXFexportFolder.InitialDirectory = @"X:\";
            this.DXFexportFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            this.DXFexportFolder.SelectedPath = @"X:\";
            // 
            // Destinationbackground
            // 
            this.Destinationbackground.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.Destinationbackground.Controls.Add(this.DXFexLoc);
            this.Destinationbackground.Controls.Add(DXFText);
            this.Destinationbackground.Controls.Add(this.PEXLoc);
            this.Destinationbackground.Controls.Add(PrintText);
            this.Destinationbackground.Controls.Add(this.PrintExportLocationButton);
            this.Destinationbackground.Controls.Add(this.DXFExportLocationButton);
            this.Destinationbackground.Dock = System.Windows.Forms.DockStyle.Top;
            this.Destinationbackground.Font = new System.Drawing.Font("Tahoma", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            this.Destinationbackground.Location = new System.Drawing.Point(3, 3);
            this.Destinationbackground.Name = "Destinationbackground";
            this.Destinationbackground.Padding = new System.Windows.Forms.Padding(3);
            this.Destinationbackground.Size = new System.Drawing.Size(344, 63);
            this.Destinationbackground.TabIndex = 9;
            // 
            // DXFexLoc
            // 
            this.DXFexLoc.AllowDrop = true;
            this.DXFexLoc.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.DXFexLoc.BackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            this.DXFexLoc.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.DXFexLoc.ForeColor = System.Drawing.Color.White;
            this.DXFexLoc.Location = new System.Drawing.Point(132, 37);
            this.DXFexLoc.Name = "DXFexLoc";
            this.DXFexLoc.PlaceholderText = "X:/";
            this.DXFexLoc.Size = new System.Drawing.Size(127, 15);
            this.DXFexLoc.TabIndex = 9;
            // 
            // PEXLoc
            // 
            this.PEXLoc.AllowDrop = true;
            this.PEXLoc.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.PEXLoc.BackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            this.PEXLoc.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.PEXLoc.ForeColor = System.Drawing.Color.White;
            this.PEXLoc.Location = new System.Drawing.Point(132, 10);
            this.PEXLoc.Name = "PEXLoc";
            this.PEXLoc.PlaceholderText = "P:/";
            this.PEXLoc.Size = new System.Drawing.Size(127, 15);
            this.PEXLoc.TabIndex = 7;
            this.PEXLoc.WordWrap = false;
            // 
            // PrintExportLocationButton
            // 
            this.PrintExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.PrintExportLocationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            this.PrintExportLocationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            this.PrintExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.PrintExportLocationButton.Location = new System.Drawing.Point(262, 7);
            this.PrintExportLocationButton.Margin = new System.Windows.Forms.Padding(0);
            this.PrintExportLocationButton.Name = "PrintExportLocationButton";
            this.PrintExportLocationButton.Size = new System.Drawing.Size(69, 20);
            this.PrintExportLocationButton.TabIndex = 6;
            this.PrintExportLocationButton.TabStop = false;
            this.PrintExportLocationButton.Text = "Browse";
            this.PrintExportLocationButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.PrintExportLocationButton.UseCompatibleTextRendering = true;
            this.PrintExportLocationButton.UseVisualStyleBackColor = true;
            // 
            // DXFExportLocationButton
            // 
            this.DXFExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            this.DXFExportLocationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            this.DXFExportLocationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            this.DXFExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.DXFExportLocationButton.Location = new System.Drawing.Point(262, 34);
            this.DXFExportLocationButton.Margin = new System.Windows.Forms.Padding(0);
            this.DXFExportLocationButton.Name = "DXFExportLocationButton";
            this.DXFExportLocationButton.Size = new System.Drawing.Size(69, 20);
            this.DXFExportLocationButton.TabIndex = 8;
            this.DXFExportLocationButton.TabStop = false;
            this.DXFExportLocationButton.Text = "Browse";
            this.DXFExportLocationButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            this.DXFExportLocationButton.UseCompatibleTextRendering = true;
            this.DXFExportLocationButton.UseVisualStyleBackColor = true;
            // 
            // BtnCncl
            // 
            this.BtnCncl.Dock = System.Windows.Forms.DockStyle.Right;
            this.BtnCncl.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            this.BtnCncl.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            this.BtnCncl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.BtnCncl.Font = new System.Drawing.Font("Tahoma", 10.0f);
            this.BtnCncl.Location = new System.Drawing.Point(179, 3);
            this.BtnCncl.Name = "BtnCncl";
            this.BtnCncl.Padding = new System.Windows.Forms.Padding(3);
            this.BtnCncl.Size = new System.Drawing.Size(162, 31);
            this.BtnCncl.TabIndex = 12;
            this.BtnCncl.Text = "Cancel";
            this.BtnCncl.UseVisualStyleBackColor = true;
            // 
            // SCBackground
            // 
            this.SCBackground.BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            this.SCBackground.Controls.Add(this.BtnSave);
            this.SCBackground.Controls.Add(this.BtnCncl);
            this.SCBackground.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.SCBackground.Font = new System.Drawing.Font("Tahoma", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            this.SCBackground.Location = new System.Drawing.Point(3, 122);
            this.SCBackground.Name = "SCBackground";
            this.SCBackground.Padding = new System.Windows.Forms.Padding(3);
            this.SCBackground.Size = new System.Drawing.Size(344, 37);
            this.SCBackground.TabIndex = 13;
            // 
            // FeaturesPanel
            // 
            this.FeaturesPanel.Controls.Add(this.ChkObsoletePrint);
            this.FeaturesPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.FeaturesPanel.Font = new System.Drawing.Font("Tahoma", 9f);
            this.FeaturesPanel.Location = new System.Drawing.Point(3, 66);
            this.FeaturesPanel.Name = "FeaturesPanel";
            this.FeaturesPanel.Padding = new System.Windows.Forms.Padding(8, 5, 3, 3);
            this.FeaturesPanel.Size = new System.Drawing.Size(344, 36);
            this.FeaturesPanel.TabIndex = 14;
            // 
            // ChkObsoletePrint
            // 
            this.ChkObsoletePrint.AutoSize = true;
            this.ChkObsoletePrint.Dock = System.Windows.Forms.DockStyle.Top;
            this.ChkObsoletePrint.Location = new System.Drawing.Point(8, 5);
            this.ChkObsoletePrint.Name = "ChkObsoletePrint";
            this.ChkObsoletePrint.Size = new System.Drawing.Size(333, 18);
            this.ChkObsoletePrint.TabIndex = 0;
            this.ChkObsoletePrint.Text = "Enable Obsolete Print";
            this.ChkObsoletePrint.UseVisualStyleBackColor = true;
            // 
            // UserOptionsForm
            // 
            AcceptButton = this.BtnSave;
            AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            CancelButton = this.BtnCncl;
            ClientSize = new System.Drawing.Size(350, 162);
            Controls.Add(this.FutureMsg);
            Controls.Add(this.FeaturesPanel);
            Controls.Add(this.Destinationbackground);
            Controls.Add(this.SCBackground);
            Font = new System.Drawing.Font("Tahoma", 8.25f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ForeColor = System.Drawing.Color.White;
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            Name = "UserOptionsForm";
            Padding = new System.Windows.Forms.Padding(3);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Options";
            this.Destinationbackground.ResumeLayout(false);
            this.Destinationbackground.PerformLayout();
            this.SCBackground.ResumeLayout(false);
            this.FeaturesPanel.ResumeLayout(false);
            this.FeaturesPanel.PerformLayout();
            Load += new EventHandler(UserOptionsForm_Load);
            ResumeLayout(false);
            PerformLayout();
        }

        internal System.Windows.Forms.Button BtnSave;
        internal System.Windows.Forms.FolderBrowserDialog PrintExportFolder;
        internal System.Windows.Forms.FolderBrowserDialog DXFexportFolder;
        internal System.Windows.Forms.Panel Destinationbackground;
        internal System.Windows.Forms.Label DXFText;
        internal System.Windows.Forms.TextBox PEXLoc;
        internal System.Windows.Forms.Label PrintText;
        internal System.Windows.Forms.Button PrintExportLocationButton;
        internal System.Windows.Forms.TextBox DXFexLoc;
        internal System.Windows.Forms.Button DXFExportLocationButton;
        internal System.Windows.Forms.TextBox FutureMsg;
        internal System.Windows.Forms.Button BtnCncl;
        internal System.Windows.Forms.Panel SCBackground;
        internal System.Windows.Forms.Panel FeaturesPanel;
        internal System.Windows.Forms.CheckBox ChkObsoletePrint;
    }
}