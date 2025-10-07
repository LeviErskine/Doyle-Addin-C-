using System;
using Color = System.Drawing.Color;
using System.Windows.Forms;
using Doyle_Addin.My_Project;

namespace Doyle_Addin.Options;

public partial class UserOptionsForm
{
    private UserOptions options;

    /// <inheritdoc />
    public UserOptionsForm()
    {
        InitializeComponent();
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
        var oThemeManager = GlobalsHelpers.ThisApplication.ThemeManager;
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


// 3. Create a helper method to avoid repeating button styling

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
        SelectFolderPath(PEXLoc, "Select Print Export Location");
    }

    private void DXFExportLocationButton_Click(object sender, EventArgs e)
    {
        SelectFolderPath(DXFexLoc, "Select DXF Export Location");
    }

    // New helper method to handle the shared logic
    private static void SelectFolderPath(TextBox targetTextBox, string description)
    {
        using var folderBrowser = new FolderBrowserDialog();
        folderBrowser.Description = description;
        // Only update the textbox if the user clicks "OK"
        if (folderBrowser.ShowDialog() == DialogResult.OK)
        {
            targetTextBox.Text = folderBrowser.SelectedPath;
        }
    }

    private void BtnCncl_Click(object sender, EventArgs e)

    {
        DialogResult = DialogResult.Cancel;

        Close();
    }
}