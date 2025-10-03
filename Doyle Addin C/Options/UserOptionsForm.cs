using System;
using Color = System.Drawing.Color;
using System.Windows.Forms;
using Inventor;

namespace Doyle_Addin.Options
{
    public partial class UserOptionsForm
    {
        private UserOptions options = new UserOptions();

        public UserOptionsForm()
        {
            this.InitializeComponent();
            options = UserOptions.Load();
        }

        private void UserOptionsForm_Load(object sender, EventArgs e)
        {
            options = UserOptions.Load();
            this.PEXLoc.Text = options.PrintExportLocation;
            this.DXFexLoc.Text = options.DxfExportLocation;
            this.ChkObsoletePrint.Checked = options.EnableObsoletePrint;

            // Apply theme colors
            ApplyThemeColors();
        }

        private void ApplyThemeColors()
        {
            ThemeManager oThemeManager;
            oThemeManager = Doyle_Addin.GlobalsHelpers.ThisApplication.ThemeManager;
            Theme oTheme;
            oTheme = oThemeManager.ActiveTheme;
            if (oTheme.Name == "LightTheme")
            {
                // Dark theme colors
                this.FutureMsg.BackColor = Color.FromArgb(245, 245, 245);
                this.FutureMsg.ForeColor = Color.Black;
                this.BtnSave.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                this.BtnSave.FlatAppearance.MouseOverBackColor = Color.White;
                this.PEXLoc.BackColor = Color.White;
                this.PEXLoc.ForeColor = Color.Black;
                this.PrintExportLocationButton.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                this.PrintExportLocationButton.FlatAppearance.MouseOverBackColor = Color.White;
                this.PrintExportLocationButton.ForeColor = Color.Black;
                this.DXFexLoc.BackColor = Color.White;
                this.DXFexLoc.ForeColor = Color.Black;
                this.DXFExportLocationButton.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                this.DXFExportLocationButton.FlatAppearance.MouseOverBackColor = Color.White;
                BackColor = Color.FromArgb(245, 245, 245);
                ForeColor = Color.Black;
                this.SCBackground.BackColor = Color.FromArgb(245, 245, 245);
                this.BtnCncl.FlatAppearance.MouseOverBackColor = Color.White;
                this.BtnCncl.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                this.FeaturesPanel.BackColor = Color.FromArgb(245, 245, 245);
                this.ChkObsoletePrint.BackColor = Color.FromArgb(245, 245, 245);
                this.ChkObsoletePrint.ForeColor = Color.Black;
            }
            else
            {
                // Dark theme colors (keep existing dark theme as default)
                this.FeaturesPanel.BackColor = Color.FromArgb(59, 68, 83);
                this.ChkObsoletePrint.BackColor = Color.FromArgb(59, 68, 83);
                this.ChkObsoletePrint.ForeColor = Color.White;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            options.PrintExportLocation = this.PEXLoc.Text;
            options.DxfExportLocation = this.DXFexLoc.Text;
            options.EnableObsoletePrint = this.ChkObsoletePrint.Checked;
            options.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void PrintExportLocationButton_Click(object sender, EventArgs e)
        {
            var folderBrowser = new FolderBrowserDialog() { Description = "Select Print Export Location" };
            folderBrowser.ShowDialog();
            this.PEXLoc.Text = folderBrowser.SelectedPath;
        }

        private void DXFExportLocationButton_Click(object sender, EventArgs e)
        {
            var folderBrowser = new FolderBrowserDialog() { Description = "Select DXF Export Location" };
            folderBrowser.ShowDialog();
            // Assuming you have a text box to display the selected path
            this.DXFexLoc.Text = folderBrowser.SelectedPath;
        }

        private void BtnCncl_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}