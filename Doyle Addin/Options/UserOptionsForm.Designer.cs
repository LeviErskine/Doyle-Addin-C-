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
#pragma warning disable CS0649 // Field is never assigned to, and will always have its default value
        private System.ComponentModel.IContainer components;
#pragma warning restore CS0649 // Field is never assigned to, and will always have its default value

        // NOTE: The following procedure is required by the Windows Form Designer
        // It can be modified using the Windows Form Designer.  
        // Do not modify it using the code editor.
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            FutureMsg = new System.Windows.Forms.TextBox();
            BtnSave = new System.Windows.Forms.Button();
            PrintExportFolder = new System.Windows.Forms.FolderBrowserDialog();
            DXFexportFolder = new System.Windows.Forms.FolderBrowserDialog();
            DXFexLoc = new System.Windows.Forms.TextBox();
            PEXLoc = new System.Windows.Forms.TextBox();
            PrintExportLocationButton = new System.Windows.Forms.Button();
            DXFExportLocationButton = new System.Windows.Forms.Button();
            BtnCncl = new System.Windows.Forms.Button();
            SCBackground = new System.Windows.Forms.Panel();
            ChkObsoletePrint = new System.Windows.Forms.CheckBox();
            Destinationbackground = new System.Windows.Forms.TableLayoutPanel();
            DXFText = new System.Windows.Forms.Label();
            PrintText = new System.Windows.Forms.Label();
            FeaturesPanel = new System.Windows.Forms.TableLayoutPanel();
            SCBackground.SuspendLayout();
            Destinationbackground.SuspendLayout();
            FeaturesPanel.SuspendLayout();
            SuspendLayout();
            // 
            // FutureMsg
            // 
            FutureMsg.BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            FutureMsg.BorderStyle = System.Windows.Forms.BorderStyle.None;
            FutureMsg.Dock = System.Windows.Forms.DockStyle.Bottom;
            FutureMsg.Enabled = false;
            FutureMsg.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            FutureMsg.ForeColor = System.Drawing.Color.White;
            FutureMsg.HideSelection = false;
            FutureMsg.Location = new System.Drawing.Point(12, 82);
            FutureMsg.Name = "FutureMsg";
            FutureMsg.ReadOnly = true;
            FutureMsg.Size = new System.Drawing.Size(336, 20);
            FutureMsg.TabIndex = 11;
            FutureMsg.Text = "More options planned for the future";
            FutureMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // BtnSave
            // 
            BtnSave.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            BtnSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            BtnSave.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            BtnSave.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            BtnSave.Font = new System.Drawing.Font("Tahoma", 10F);
            BtnSave.Location = new System.Drawing.Point(3, 3);
            BtnSave.MaximumSize = new System.Drawing.Size(140, 30);
            BtnSave.MinimumSize = new System.Drawing.Size(140, 30);
            BtnSave.Name = "BtnSave";
            BtnSave.Padding = new System.Windows.Forms.Padding(3);
            BtnSave.Size = new System.Drawing.Size(140, 30);
            BtnSave.TabIndex = 5;
            BtnSave.Text = "Save";
            BtnSave.UseVisualStyleBackColor = true;
            BtnSave.Click += BtnSave_Click;
            // 
            // PrintExportFolder
            // 
            PrintExportFolder.InitialDirectory = "P:\\";
            PrintExportFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            PrintExportFolder.SelectedPath = "P:\\";
            // 
            // DXFexportFolder
            // 
            DXFexportFolder.InitialDirectory = "X:\\";
            DXFexportFolder.RootFolder = Environment.SpecialFolder.MyComputer;
            DXFexportFolder.SelectedPath = "X:\\";
            // 
            // DXFexLoc
            // 
            DXFexLoc.AllowDrop = true;
            DXFexLoc.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            DXFexLoc.BackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            DXFexLoc.BorderStyle = System.Windows.Forms.BorderStyle.None;
            DXFexLoc.ForeColor = System.Drawing.Color.White;
            DXFexLoc.Location = new System.Drawing.Point(123, 27);
            DXFexLoc.Name = "DXFexLoc";
            DXFexLoc.PlaceholderText = "X:/";
            DXFexLoc.Size = new System.Drawing.Size(122, 14);
            DXFexLoc.TabIndex = 9;
            // 
            // PEXLoc
            // 
            PEXLoc.AllowDrop = true;
            PEXLoc.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            PEXLoc.BackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            PEXLoc.BorderStyle = System.Windows.Forms.BorderStyle.None;
            PEXLoc.ForeColor = System.Drawing.Color.White;
            PEXLoc.Location = new System.Drawing.Point(123, 4);
            PEXLoc.Name = "PEXLoc";
            PEXLoc.PlaceholderText = "P:/";
            PEXLoc.Size = new System.Drawing.Size(122, 14);
            PEXLoc.TabIndex = 7;
            PEXLoc.WordWrap = false;
            // 
            // PrintExportLocationButton
            // 
            PrintExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            PrintExportLocationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            PrintExportLocationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            PrintExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            PrintExportLocationButton.Location = new System.Drawing.Point(248, 1);
            PrintExportLocationButton.Margin = new System.Windows.Forms.Padding(0);
            PrintExportLocationButton.MinimumSize = new System.Drawing.Size(0, 20);
            PrintExportLocationButton.Name = "PrintExportLocationButton";
            PrintExportLocationButton.Size = new System.Drawing.Size(88, 20);
            PrintExportLocationButton.TabIndex = 6;
            PrintExportLocationButton.TabStop = false;
            PrintExportLocationButton.Text = "Browse";
            PrintExportLocationButton.UseCompatibleTextRendering = true;
            PrintExportLocationButton.UseVisualStyleBackColor = true;
            PrintExportLocationButton.Click += PrintExportLocationButton_Click;
            // 
            // DXFExportLocationButton
            // 
            DXFExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            DXFExportLocationButton.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            DXFExportLocationButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            DXFExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            DXFExportLocationButton.Location = new System.Drawing.Point(248, 24);
            DXFExportLocationButton.Margin = new System.Windows.Forms.Padding(0);
            DXFExportLocationButton.MinimumSize = new System.Drawing.Size(0, 20);
            DXFExportLocationButton.Name = "DXFExportLocationButton";
            DXFExportLocationButton.Size = new System.Drawing.Size(88, 20);
            DXFExportLocationButton.TabIndex = 8;
            DXFExportLocationButton.TabStop = false;
            DXFExportLocationButton.Text = "Browse";
            DXFExportLocationButton.UseCompatibleTextRendering = true;
            DXFExportLocationButton.UseVisualStyleBackColor = true;
            DXFExportLocationButton.Click += DXFExportLocationButton_Click;
            // 
            // BtnCncl
            // 
            BtnCncl.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            BtnCncl.Dock = System.Windows.Forms.DockStyle.Right;
            BtnCncl.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(134, 145, 161);
            BtnCncl.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(44, 51, 64);
            BtnCncl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            BtnCncl.Font = new System.Drawing.Font("Tahoma", 10F);
            BtnCncl.Location = new System.Drawing.Point(193, 3);
            BtnCncl.MaximumSize = new System.Drawing.Size(140, 30);
            BtnCncl.MinimumSize = new System.Drawing.Size(140, 30);
            BtnCncl.Name = "BtnCncl";
            BtnCncl.Padding = new System.Windows.Forms.Padding(3);
            BtnCncl.Size = new System.Drawing.Size(140, 30);
            BtnCncl.TabIndex = 12;
            BtnCncl.Text = "Cancel";
            BtnCncl.UseVisualStyleBackColor = true;
            BtnCncl.Click += BtnCncl_Click;
            // 
            // SCBackground
            // 
            SCBackground.BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            SCBackground.Controls.Add(BtnSave);
            SCBackground.Controls.Add(BtnCncl);
            SCBackground.Dock = System.Windows.Forms.DockStyle.Bottom;
            SCBackground.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            SCBackground.Location = new System.Drawing.Point(12, 102);
            SCBackground.Name = "SCBackground";
            SCBackground.Padding = new System.Windows.Forms.Padding(3);
            SCBackground.Size = new System.Drawing.Size(336, 36);
            SCBackground.TabIndex = 13;
            // 
            // ChkObsoletePrint
            // 
            ChkObsoletePrint.AutoSize = true;
            ChkObsoletePrint.Location = new System.Drawing.Point(3, 3);
            ChkObsoletePrint.Name = "ChkObsoletePrint";
            ChkObsoletePrint.Size = new System.Drawing.Size(129, 17);
            ChkObsoletePrint.TabIndex = 0;
            ChkObsoletePrint.Text = "Enable Obsolete Print";
            ChkObsoletePrint.UseVisualStyleBackColor = true;
            // 
            // Destinationbackground
            // 
            Destinationbackground.ColumnCount = 3;
            Destinationbackground.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 48.3333321F));
            Destinationbackground.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 51.6666679F));
            Destinationbackground.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 87F));
            Destinationbackground.Controls.Add(PEXLoc, 1, 0);
            Destinationbackground.Controls.Add(PrintExportLocationButton, 2, 0);
            Destinationbackground.Controls.Add(DXFexLoc, 1, 1);
            Destinationbackground.Controls.Add(DXFText, 0, 1);
            Destinationbackground.Controls.Add(DXFExportLocationButton, 2, 1);
            Destinationbackground.Controls.Add(PrintText, 0, 0);
            Destinationbackground.Dock = System.Windows.Forms.DockStyle.Top;
            Destinationbackground.Location = new System.Drawing.Point(12, 12);
            Destinationbackground.Name = "Destinationbackground";
            Destinationbackground.RowCount = 2;
            Destinationbackground.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            Destinationbackground.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            Destinationbackground.Size = new System.Drawing.Size(336, 46);
            Destinationbackground.TabIndex = 15;
            // 
            // DXFText
            // 
            DXFText.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            DXFText.Location = new System.Drawing.Point(3, 24);
            DXFText.MinimumSize = new System.Drawing.Size(115, 20);
            DXFText.Name = "DXFText";
            DXFText.Size = new System.Drawing.Size(115, 20);
            DXFText.TabIndex = 10;
            DXFText.Text = "DXF Export Location";
            DXFText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // PrintText
            // 
            PrintText.Anchor = System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right;
            PrintText.Location = new System.Drawing.Point(3, 1);
            PrintText.MinimumSize = new System.Drawing.Size(115, 20);
            PrintText.Name = "PrintText";
            PrintText.Size = new System.Drawing.Size(115, 20);
            PrintText.TabIndex = 9;
            PrintText.Text = "Print Export Location";
            PrintText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // FeaturesPanel
            // 
            FeaturesPanel.AutoSize = true;
            FeaturesPanel.ColumnCount = 1;
            FeaturesPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            FeaturesPanel.Controls.Add(ChkObsoletePrint, 0, 0);
            FeaturesPanel.Dock = System.Windows.Forms.DockStyle.Top;
            FeaturesPanel.Location = new System.Drawing.Point(12, 58);
            FeaturesPanel.Name = "FeaturesPanel";
            FeaturesPanel.RowCount = 1;
            FeaturesPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            FeaturesPanel.Size = new System.Drawing.Size(336, 23);
            FeaturesPanel.TabIndex = 16;
            // 
            // UserOptionsForm
            // 
            AcceptButton = BtnSave;
            AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            BackColor = System.Drawing.Color.FromArgb(59, 68, 83);
            CancelButton = BtnCncl;
            ClientSize = new System.Drawing.Size(360, 150);
            Controls.Add(FeaturesPanel);
            Controls.Add(FutureMsg);
            Controls.Add(SCBackground);
            Controls.Add(Destinationbackground);
            Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0);
            ForeColor = System.Drawing.Color.White;
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            MaximumSize = new System.Drawing.Size(360, 150);
            MinimumSize = new System.Drawing.Size(360, 150);
            Name = "UserOptionsForm";
            Padding = new System.Windows.Forms.Padding(12);
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Options";
            Load += UserOptionsForm_Load;
            SCBackground.ResumeLayout(false);
            Destinationbackground.ResumeLayout(false);
            Destinationbackground.PerformLayout();
            FeaturesPanel.ResumeLayout(false);
            FeaturesPanel.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        internal System.Windows.Forms.Button BtnSave;
        internal System.Windows.Forms.FolderBrowserDialog PrintExportFolder;
        internal System.Windows.Forms.FolderBrowserDialog DXFexportFolder;
        private System.Windows.Forms.Label DXFText;
        private System.Windows.Forms.TextBox PEXLoc;
        private System.Windows.Forms.Label PrintText;
        private System.Windows.Forms.Button PrintExportLocationButton;
        private System.Windows.Forms.TextBox DXFexLoc;
        private System.Windows.Forms.Button DXFExportLocationButton;
        private System.Windows.Forms.TextBox FutureMsg;
        private System.Windows.Forms.Button BtnCncl;
        private System.Windows.Forms.Panel SCBackground;
        internal System.Windows.Forms.CheckBox ChkObsoletePrint;
        private System.Windows.Forms.TableLayoutPanel Destinationbackground;
        private System.Windows.Forms.TableLayoutPanel FeaturesPanel;
    }
}