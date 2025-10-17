using System;
using Ookii.Dialogs.Wpf; // From the WpfFolderDialog NuGet package
using System.Windows;
using System.Windows.Controls;
using Doyle_Addin.My_Project;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Markup;

namespace Doyle_Addin.Options;

/// <summary>
/// 
/// </summary>
public partial class UserOptionsWindow
{
    private UserOptions options;

    /// <inheritdoc />
    public UserOptionsWindow()
    {
        // Preload theme resources into the window resources so styles resolve during InitializeComponent
        try
        {
            PreloadThemeResources();
        }
        catch
        {
            // ignore preload failures; ApplyThemeColors will try again on Loaded
        }

        InitializeComponent();
    }

    // Load the theme dictionary into this window's resources before InitializeComponent runs
    private void PreloadThemeResources()
    {
        try
        {
            var oThemeManager = GlobalsHelpers.ThisApplication?.ThemeManager;
            var oTheme = oThemeManager?.ActiveTheme;
            var themeName = (oTheme?.Name == "LightTheme") ? "LightTheme" : "DarkTheme";

            ResourceDictionary rd = null;
            try
            {
                var asmName = Assembly.GetExecutingAssembly().GetName().Name;
                var packUri = new Uri($"pack://application:,,,/{asmName};component/Options/Themes/{themeName}.xaml",
                    UriKind.Absolute);
                rd = new ResourceDictionary { Source = packUri };
            }
            catch
            {
                var asmLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
                var themePath = Path.Combine(asmLocation, "Options", "Themes", themeName + ".xaml");
                if (File.Exists(themePath))
                {
                    using var fs = File.OpenRead(themePath);
                    rd = (ResourceDictionary)XamlReader.Load(fs);
                }
            }

            if (rd != null)
            {
                // Add to window merged dictionaries so InitializeComponent sees resources
                this.Resources.MergedDictionaries.Add(rd);
            }
        }
        catch
        {
            // swallow: fallback will occur in ApplyThemeColors
        }
    }

    private void UserOptionsWindow_Loaded(object sender, RoutedEventArgs e)
    {
        options = UserOptions.Load();

        // Set the DataContext for potential data binding
        this.DataContext = options;

        // Manual property setting for non-bound controls
        // PexLoc.Text = options.PrintExportLocation;
        // DxFexLoc.Text = options.DxfExportLocation;
        // ChkObsoletePrint.IsChecked is now handled by data binding

        ApplyThemeColors();
    }

    private void ApplyThemeColors()
    {
        var oThemeManager = GlobalsHelpers.ThisApplication.ThemeManager;
        var oTheme = oThemeManager.ActiveTheme;

        // Decide theme file name
        var themeName = (oTheme?.Name == "LightTheme") ? "LightTheme" : "DarkTheme";

        // Load the ResourceDictionary first (pack URI preferred, then disk fallback)
        ResourceDictionary rd = null;
        try
        {
            var asmName = Assembly.GetExecutingAssembly().GetName().Name;
            var packUri = new Uri($"pack://application:,,,/{asmName};component/Options/Themes/{themeName}.xaml",
                UriKind.Absolute);
            rd = new ResourceDictionary { Source = packUri };
        }
        catch
        {
            // Fallback: try to load the XAML from disk next to the executing assembly
            try
            {
                var asmLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
                var themePath = Path.Combine(asmLocation, "Options", "Themes", themeName + ".xaml");
                if (File.Exists(themePath))
                {
                    using var fs = File.OpenRead(themePath);
                    rd = (ResourceDictionary)XamlReader.Load(fs);
                }
            }
            catch
            {
                // ignore - leave rd as null
            }
        }

        if (rd != null)
        {
            // If we have an Application-level resources, copy theme resources there so DynamicResource lookups succeed
            if (Application.Current?.Resources != null)
            {
                var appRes = Application.Current.Resources;
                foreach (var key in rd.Keys)
                {
                    appRes[key] = rd[key];
                }
            }
            else
            {
                // Replace any existing theme dictionaries in window resources with the new one
                // Avoid clearing until we have rd to prevent transient missing-resource warnings
                var toRemove = this.Resources.MergedDictionaries.Where(md =>
                    md.Source != null && md.Source.OriginalString.Contains("/Options/Themes/")).ToList();

                foreach (var r in toRemove)
                    this.Resources.MergedDictionaries.Remove(r);

                this.Resources.MergedDictionaries.Add(rd);
            }
        }

        // Ensure the window's Foreground uses the ControlForegroundBrush from the theme
        this.SetResourceReference(Window.ForegroundProperty, "ControlForegroundBrush");

        // Many visual properties are now driven by DynamicResource in XAML; no per-control assignment required here.
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        // Bound properties are already in the `options` instance; just persist them.
        options.Save();
        DialogResult = true;
        Close();
    }

    private void PrintExportLocationButton_Click(object sender, RoutedEventArgs e)
    {
        SelectFolderPath(PexLoc, "Select Print Export Location");
    }

    private void DXFExportLocationButton_Click(object sender, RoutedEventArgs e)
    {
        SelectFolderPath(DxFexLoc, "Select DXF Export Location");
    }

    // Helper method using the modern WPF folder browser dialog
    private void SelectFolderPath(TextBox targetTextBox, string description)
    {
        var folderBrowser = new VistaFolderBrowserDialog
        {
            Description = description,
            UseDescriptionForTitle = true
        };

        // When ShowDialog() returns true, the user has selected a folder
        if (folderBrowser.ShowDialog(this).GetValueOrDefault())
        {
            targetTextBox.Text = folderBrowser.SelectedPath;
        }
    }

    private void BtnCncl_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;

        Close();
    }
}