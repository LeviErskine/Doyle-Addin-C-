using TextBox = Wpf.Ui.Controls.TextBox;

// From the WpfFolderDialog NuGet package

namespace DoyleAddin.Options;

using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Navigation;
using Ookii.Dialogs.Wpf;

/// <summary>
/// </summary>
public partial class UserOptionsWindow
{
	private const string LightTheme = "LightTheme";
	private const string DarkTheme = "DarkTheme";
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
			var oThemeManager = ThisApplication?.ThemeManager;
			var oTheme        = oThemeManager?.ActiveTheme;
			var themeName     = oTheme?.Name == LightTheme ? LightTheme : DarkTheme;

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
				var themePath   = Path.Combine(asmLocation, "Options", "Themes", themeName + ".xaml");
				if (File.Exists(themePath))
				{
					using var fs = File.OpenRead(themePath);
					rd = (ResourceDictionary)XamlReader.Load(fs);
				}
			}

			if (rd != null)
				// Add to window merged dictionaries so InitializeComponent sees resources
				Resources.MergedDictionaries.Add(rd);
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
		DataContext = options;

		ApplyThemeColors();
	}

	private void ApplyThemeColors()
	{
		var themeName          = GetThemeName();
		var resourceDictionary = LoadThemeResource(themeName);

		if (resourceDictionary != null) ApplyThemeResources(resourceDictionary);

		SetResourceReference(ForegroundProperty, "ControlForegroundBrush");
	}

	private static string GetThemeName()
	{
		var oTheme = ThisApplication.ThemeManager.ActiveTheme;
		return oTheme?.Name == LightTheme ? LightTheme : DarkTheme;
	}

	private static ResourceDictionary LoadThemeResource(string themeName)
	{
		return LoadFromPackUri(themeName) ?? LoadFromDisk(themeName);
	}

	private static ResourceDictionary LoadFromPackUri(string themeName)
	{
		try
		{
			var asmName = Assembly.GetExecutingAssembly().GetName().Name;
			var packUri = new Uri($"pack://application:,,,/{asmName};component/Options/Themes/{themeName}.xaml",
				UriKind.Absolute);
			return new ResourceDictionary { Source = packUri };
		}
		catch
		{
			return [];
		}
	}

	private static ResourceDictionary LoadFromDisk(string themeName)
	{
		try
		{
			var asmLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
			var themePath   = Path.Combine(asmLocation, "Options", "Themes", themeName + ".xaml");

			if (!File.Exists(themePath))
				return [];

			using var fs = File.OpenRead(themePath);
			return (ResourceDictionary)XamlReader.Load(fs);
		}
		catch
		{
			return [];
		}
	}

	private void ApplyThemeResources(ResourceDictionary resourceDictionary)
	{
		if (Application.Current?.Resources != null)
			CopyToApplicationResources(resourceDictionary);
		else
			AddToWindowResources(resourceDictionary);
	}

	private static void CopyToApplicationResources(ResourceDictionary resourceDictionary)
	{
		var appRes                                               = Application.Current.Resources;
		foreach (var key in resourceDictionary.Keys) appRes[key] = resourceDictionary[key];
	}

	private void AddToWindowResources(ResourceDictionary resourceDictionary)
	{
		RemoveExistingThemeDictionaries();
		Resources.MergedDictionaries.Add(resourceDictionary);
	}

	private void RemoveExistingThemeDictionaries()
	{
		var toRemove = Resources.MergedDictionaries.Where(md =>
			md.Source != null && md.Source.OriginalString.Contains("/Options/Themes/")).ToList();

		foreach (var r in toRemove)
			Resources.MergedDictionaries.Remove(r);
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
			Description            = description,
			UseDescriptionForTitle = true
		};

		// When ShowDialog() returns true, the user has selected a folder
		if (folderBrowser.ShowDialog(this).GetValueOrDefault()) targetTextBox.Text = folderBrowser.SelectedPath;
	}

	private void BtnCncl_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;

		Close();
	}

	private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
	{
		Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
		e.Handled = true;
		Close();
	}
}