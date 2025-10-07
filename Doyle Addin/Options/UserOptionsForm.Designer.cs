using System;
using System.Diagnostics;

namespace Doyle_Addin.Options
{
    /// <inheritdoc />
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
                if (disposing && components is not null)
                {
                    components.Dispose();
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
            FutureMsg = new System.Windows.Forms.TextBox();
            BtnSave = new System.Windows.Forms.Button();
            BtnSave.Click += new EventHandler(BtnSave_Click);
            PrintExportFolder = new System.Windows.Forms.FolderBrowserDialog();
            DXFexportFolder = new System.Windows.Forms.FolderBrowserDialog();
            Destinationbackground = new System.Windows.Forms.Panel();
            DXFexLoc = new System.Windows.Forms.TextBox();
            PEXLoc = new System.Windows.Forms.TextBox();
            PrintExportLocationButton = new System.Windows.Forms.Button();
            PrintExportLocationButton.Click += new EventHandler(PrintExportLocationButton_Click);
            DXFExportLocationButton = new System.Windows.Forms.Button();
            DXFExportLocationButton.Click += new EventHandler(DXFExportLocationButton_Click);
            BtnCncl = new System.Windows.Forms.Button();
            BtnCncl.Click += new EventHandler(BtnCncl_Click);
            SCBackground = new System.Windows.Forms.Panel();
            FeaturesPanel = new System.Windows.Forms.Panel();
            ChkObsoletePrint = new System.Windows.Forms.CheckBox();
            PrintText = new System.Windows.Forms.Label();
            DXFText = new System.Windows.Forms.Label();
            Destinationbackground.SuspendLayout();
            SCBackground.SuspendLayout();
            FeaturesPanel.SuspendLayout();
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
            FutureMsg.BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            FutureMsg.BorderStyle = System.Windows.Forms.BorderStyle.None;
            FutureMsg.Dock = System.Windows.Forms.DockStyle.Bottom;
            FutureMsg.Enabled = false;
            FutureMsg.Font = new System.Drawing.Font("Tahoma", 12.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            FutureMsg.ForeColor = System.Drawing.Color.White;
            FutureMsg.HideSelection = false;
            FutureMsg.Location = new System.Drawing.Point(3, 102);
            FutureMsg.Name = "FutureMsg";
            FutureMsg.ReadOnly = true;
            FutureMsg.Size = new System.Drawing.Size(344, 20);
            FutureMsg.TabIndex = 11;
            FutureMsg.Text = "More options planned for the future";
            FutureMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnSave
            // 
            BtnSave.Dock = System.Windows.Forms.DockStyle.Left;
            BtnSave.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            BtnSave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            BtnSave.Font = new System.Drawing.Font("Tahoma", 10.0f);
            BtnSave.Location = new System.Drawing.Point(3, 3);
            BtnSave.Name = "BtnSave";
            BtnSave.Padding = new System.Windows.Forms.Padding(3);
            BtnSave.Size = new System.Drawing.Size(162, 31);
            BtnSave.TabIndex = 5;
            BtnSave.Text = "Save";
            BtnSave.UseVisualStyleBackColor = true;
            // 
            // PrintExportFolder
            // 
            PrintExportFolder.InitialDirectory = @"P:\";
            PrintExportFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            PrintExportFolder.SelectedPath = @"P:\";
            // 
            // DXFexportFolder
            // 
            DXFexportFolder.InitialDirectory = @"X:\";
            DXFexportFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            DXFexportFolder.SelectedPath = @"X:\";
            // 
            // Destinationbackground
            // 
            Destinationbackground.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            Destinationbackground.Controls.Add(DXFexLoc);
            Destinationbackground.Controls.Add(DXFText);
            Destinationbackground.Controls.Add(PEXLoc);
            Destinationbackground.Controls.Add(PrintText);
            Destinationbackground.Controls.Add(PrintExportLocationButton);
            Destinationbackground.Controls.Add(DXFExportLocationButton);
            Destinationbackground.Dock = System.Windows.Forms.DockStyle.Top;
            Destinationbackground.Font = new System.Drawing.Font("Tahoma", 9.0f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            Destinationbackground.Location = new System.Drawing.Point(3, 3);
            Destinationbackground.Name = "Destinationbackground";
            Destinationbackground.Padding = new System.Windows.Forms.Padding(3);
            Destinationbackground.Size = new System.Drawing.Size(344, 63);
            Destinationbackground.TabIndex = 9;
            // 
            // DXFexLoc
            // 
            DXFexLoc.AllowDrop = true;
            DXFexLoc.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            DXFexLoc.BackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            DXFexLoc.BorderStyle = System.Windows.Forms.BorderStyle.None;
            DXFexLoc.ForeColor = System.Drawing.Color.White;
            DXFexLoc.Location = new System.Drawing.Point(132, 37);
            DXFexLoc.Name = "DXFexLoc";
            DXFexLoc.PlaceholderText = "X:/";
            DXFexLoc.Size = new System.Drawing.Size(127, 15);
            DXFexLoc.TabIndex = 9;
            // 
            // PEXLoc
            // 
            PEXLoc.AllowDrop = true;
            PEXLoc.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            PEXLoc.BackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            PEXLoc.BorderStyle = System.Windows.Forms.BorderStyle.None;
            PEXLoc.ForeColor = System.Drawing.Color.White;
            PEXLoc.Location = new System.Drawing.Point(132, 10);
            PEXLoc.Name = "PEXLoc";
            PEXLoc.PlaceholderText = "P:/";
            PEXLoc.Size = new System.Drawing.Size(127, 15);
            PEXLoc.TabIndex = 7;
            PEXLoc.WordWrap = false;
            // 
            // PrintExportLocationButton
            // 
            PrintExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            PrintExportLocationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            PrintExportLocationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            PrintExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            PrintExportLocationButton.Location = new System.Drawing.Point(262, 7);
            PrintExportLocationButton.Margin = new System.Windows.Forms.Padding(0);
            PrintExportLocationButton.Name = "PrintExportLocationButton";
            PrintExportLocationButton.Size = new System.Drawing.Size(69, 20);
            PrintExportLocationButton.TabIndex = 6;
            PrintExportLocationButton.TabStop = false;
            PrintExportLocationButton.Text = "Browse";
            PrintExportLocationButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            PrintExportLocationButton.UseCompatibleTextRendering = true;
            PrintExportLocationButton.UseVisualStyleBackColor = true;
            // 
            // DXFExportLocationButton
            // 
            DXFExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            DXFExportLocationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            DXFExportLocationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            DXFExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            DXFExportLocationButton.Location = new System.Drawing.Point(262, 34);
            DXFExportLocationButton.Margin = new System.Windows.Forms.Padding(0);
            DXFExportLocationButton.Name = "DXFExportLocationButton";
            DXFExportLocationButton.Size = new System.Drawing.Size(69, 20);
            DXFExportLocationButton.TabIndex = 8;
            DXFExportLocationButton.TabStop = false;
            DXFExportLocationButton.Text = "Browse";
            DXFExportLocationButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage;
            DXFExportLocationButton.UseCompatibleTextRendering = true;
            DXFExportLocationButton.UseVisualStyleBackColor = true;
            // 
            // BtnCncl
            // 
            BtnCncl.Dock = System.Windows.Forms.DockStyle.Right;
            BtnCncl.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            BtnCncl.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            BtnCncl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            BtnCncl.Font = new System.Drawing.Font("Tahoma", 10.0f);
            BtnCncl.Location = new System.Drawing.Point(179, 3);
            BtnCncl.Name = "BtnCncl";
            BtnCncl.Padding = new System.Windows.Forms.Padding(3);
            BtnCncl.Size = new System.Drawing.Size(162, 31);
            BtnCncl.TabIndex = 12;
            BtnCncl.Text = "Cancel";
            BtnCncl.UseVisualStyleBackColor = true;
            // 
            // SCBackground
            // 
            SCBackground.BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            SCBackground.Controls.Add(BtnSave);
            SCBackground.Controls.Add(BtnCncl);
            SCBackground.Dock = System.Windows.Forms.DockStyle.Bottom;
            SCBackground.Font = new System.Drawing.Font("Tahoma", 9f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            SCBackground.Location = new System.Drawing.Point(3, 122);
            SCBackground.Name = "SCBackground";
            SCBackground.Padding = new System.Windows.Forms.Padding(3);
            SCBackground.Size = new System.Drawing.Size(344, 37);
            SCBackground.TabIndex = 13;
            // 
            // FeaturesPanel
            // 
            FeaturesPanel.Controls.Add(ChkObsoletePrint);
            FeaturesPanel.Dock = System.Windows.Forms.DockStyle.Top;
            FeaturesPanel.Font = new System.Drawing.Font("Tahoma", 9f);
            FeaturesPanel.Location = new System.Drawing.Point(3, 66);
            FeaturesPanel.Name = "FeaturesPanel";
            FeaturesPanel.Padding = new System.Windows.Forms.Padding(8, 5, 3, 3);
            FeaturesPanel.Size = new System.Drawing.Size(344, 36);
            FeaturesPanel.TabIndex = 14;
            // 
            // ChkObsoletePrint
            // 
            ChkObsoletePrint.AutoSize = true;
            ChkObsoletePrint.Dock = System.Windows.Forms.DockStyle.Top;
            ChkObsoletePrint.Location = new System.Drawing.Point(8, 5);
            ChkObsoletePrint.Name = "ChkObsoletePrint";
            ChkObsoletePrint.Size = new System.Drawing.Size(333, 18);
            ChkObsoletePrint.TabIndex = 0;
            ChkObsoletePrint.Text = "Enable Obsolete Print";
            ChkObsoletePrint.UseVisualStyleBackColor = true;
            // 
            // UserOptionsForm
            // 
            AcceptButton = BtnSave;
            AutoScaleDimensions = new System.Drawing.SizeF(6f, 13f);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            CancelButton = BtnCncl;
            ClientSize = new System.Drawing.Size(350, 162);
            Controls.Add(FutureMsg);
            Controls.Add(FeaturesPanel);
            Controls.Add(Destinationbackground);
            Controls.Add(SCBackground);
            Font = new System.Drawing.Font("Tahoma", 8.25f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ForeColor = System.Drawing.Color.White;
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            Name = "UserOptionsForm";
            Padding = new System.Windows.Forms.Padding(3);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Options";
            Destinationbackground.ResumeLayout(false);
            Destinationbackground.PerformLayout();
            SCBackground.ResumeLayout(false);
            FeaturesPanel.ResumeLayout(false);
            FeaturesPanel.PerformLayout();
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