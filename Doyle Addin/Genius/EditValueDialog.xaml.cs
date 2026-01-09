#nullable enable

#region

using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

#endregion

namespace Doyle_Addin.Genius;

public partial class EditValueDialog
{
	private const string LightTheme = "LightTheme";
	private const string DarkTheme = "DarkTheme";

	public EditValueDialog()
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
		ValueTextBox.Focus();
	}

	public EditValueDialog(string? propertyName, string? currentValue) : this()
	{
		if (propertyName != null) PropertyNameText.Text = $"Property: {propertyName}";
		if (currentValue == null) return;
		ValueTextBox.Text = currentValue;
		ValueTextBox.SelectAll();
	}

	public string PropertyValue { get; private set; } = string.Empty;

	// Load the theme dictionary into this window's resources before InitializeComponent runs
	private void PreloadThemeResources()
	{
		try
		{
			var oThemeManager = ThisApplication?.ThemeManager;
			var themeName     = oThemeManager is { ActiveTheme.Name: LightTheme } ? LightTheme : DarkTheme;

			ResourceDictionary? rd = null;
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

	private void ValueTextBox_TextChanged(object sender, TextChangedEventArgs e)
	{
		PropertyValue = ValueTextBox.Text;
	}

	private void OKButton_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = true;
		Close();
	}

	private void CancelButton_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}
}