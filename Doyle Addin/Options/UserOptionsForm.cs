using System.Diagnostics;
using Inventor;
using Color = System.Drawing.Color;

namespace Doyle_Addin.Options
{
    /// <inheritdoc />
    public partial class UserOptionsForm
    {
        private UserOptions options;

        /// <inheritdoc />
        public UserOptionsForm()
        {
            InitializeComponent();
            options = UserOptions.Load();
        }

        private void UserOptionsForm_Load(object sender, EventArgs e)
        {
            options = UserOptions.Load();
            PEXLoc.Text = options.PrintExportLocation;
            DXFexLoc.Text = options.DxfExportLocation;
            ChkObsoletePrint.Checked = options.EnableObsoletePrint;

            // Apply theme colors
            ApplyThemeColors();
        }

        private void ApplyThemeColors()
        {
            var oThemeManager = typeof(ThemeManager).InvokeMember("Application",
                System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Static |
                System.Reflection.BindingFlags.NonPublic,
                null, null, null) as ThemeManager;
            Debug.Assert(oThemeManager != null, nameof(oThemeManager) + " != null");
            var oTheme = oThemeManager.ActiveTheme;
            if (oTheme.Name == "LightTheme")
            {
                // Dark theme colors
                FutureMsg.BackColor = Color.FromArgb(245, 245, 245);
                FutureMsg.ForeColor = Color.Black;
                BtnSave.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                BtnSave.FlatAppearance.MouseOverBackColor = Color.White;
                PEXLoc.BackColor = Color.White;
                PEXLoc.ForeColor = Color.Black;
                PrintExportLocationButton.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                PrintExportLocationButton.FlatAppearance.MouseOverBackColor = Color.White;
                PrintExportLocationButton.ForeColor = Color.Black;
                DXFexLoc.BackColor = Color.White;
                DXFexLoc.ForeColor = Color.Black;
                DXFExportLocationButton.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                DXFExportLocationButton.FlatAppearance.MouseOverBackColor = Color.White;
                BackColor = Color.FromArgb(245, 245, 245);
                ForeColor = Color.Black;
                SCBackground.BackColor = Color.FromArgb(245, 245, 245);
                BtnCncl.FlatAppearance.MouseOverBackColor = Color.White;
                BtnCncl.FlatAppearance.BorderColor = Color.FromArgb(186, 186, 186);
                FeaturesPanel.BackColor = Color.FromArgb(245, 245, 245);
                ChkObsoletePrint.BackColor = Color.FromArgb(245, 245, 245);
                ChkObsoletePrint.ForeColor = Color.Black;
            }
            else
            {
                // Dark theme colors (keep the existing dark theme as default)
                FeaturesPanel.BackColor = Color.FromArgb(59, 68, 83);
                ChkObsoletePrint.BackColor = Color.FromArgb(59, 68, 83);
                ChkObsoletePrint.ForeColor = Color.White;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            options.PrintExportLocation = PEXLoc.Text;
            options.DxfExportLocation = DXFexLoc.Text;
            options.EnableObsoletePrint = ChkObsoletePrint.Checked;
            options.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void PrintExportLocationButton_Click(object sender, EventArgs e)
        {
            var folderBrowser = new FolderBrowserDialog() { Description = @"Select Print Export Location" };
            folderBrowser.ShowDialog();
            PEXLoc.Text = folderBrowser.SelectedPath;
        }

        private void DXFExportLocationButton_Click(object sender, EventArgs e)
        {
            var folderBrowser = new FolderBrowserDialog() { Description = @"Select DXF Export Location" };
            folderBrowser.ShowDialog();
            // Assuming you have a text box to display the selected path
            DXFexLoc.Text = folderBrowser.SelectedPath;
        }

        // ReSharper disable once IdentifierTypo
        private void BtnCncl_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}