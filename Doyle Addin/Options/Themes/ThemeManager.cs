namespace DoyleAddin.Options.Themes;

using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Markup;

/// <summary>
///     Manages dynamic theme loading and application for WPF windows
/// </summary>
/// <example>
///     To use this in any WPF window, call ApplyTheme in the Loaded event:
///     <code>
/// private void Window_Loaded(object sender, RoutedEventArgs e)
/// {
///     ThemeManager.ApplyTheme(this);
/// }
/// </code>
/// </example>
public static class ThemeManager
{
	private const string LightTheme = "LightTheme";
	private const string DarkTheme = "DarkTheme";
	private const string ThemesFolder = "Options/Themes";
	private const string ThemeFileExtension = ".xaml";


	/// <summary>
	///     Applies the appropriate theme to the specified framework element based on Inventor's active theme
	/// </summary>
	/// <param name="element">The framework element to apply the theme to</param>
	public static void ApplyTheme(FrameworkElement element)
	{
		if (element == null) return;

		var themeName          = GetThemeName();
		var themeDictionary    = LoadResourceDictionary(themeName);
		var elementsDictionary = LoadFromDisk("Elements");

		if (themeDictionary != null && elementsDictionary != null)
		{
			// Merge theme resources into Elements dictionary so DynamicResource works

			foreach (var key in themeDictionary.Keys) elementsDictionary.Add(key, themeDictionary[key]);
			ApplyThemeResources(element, elementsDictionary);
		}
		else
		{
			if (themeDictionary != null)
				ApplyThemeResources(element, themeDictionary);

			if (elementsDictionary != null)
				ApplyThemeResources(element, elementsDictionary);
		}
	}


	/// <summary>
	///     Gets the current theme name from Inventor's ThemeManager
	/// </summary>
	/// <returns>LightTheme or DarkTheme</returns>
	private static string GetThemeName()
	{
		if (ThisApplication?.ThemeManager?.ActiveTheme == null) return DarkTheme; // Default fallback


		var themeName = ThisApplication.ThemeManager.ActiveTheme.Name;
		return themeName == LightTheme ? LightTheme : DarkTheme;
	}


	/// <summary>
	///     Loads a resource dictionary, trying pack URI first then falling back to disk
	/// </summary>
	/// <param name="resourceName">The name of the resource (without extension)</param>
	/// <returns>The loaded ResourceDictionary or null if failed</returns>
	private static ResourceDictionary LoadResourceDictionary(string resourceName)
	{
		return LoadFromPackUri(resourceName) ?? LoadFromDisk(resourceName);
	}


	/// <summary>
	///     Loads a resource dictionary from pack URI
	/// </summary>
	/// <param name="resourceName">The name of the resource (without extension)</param>
	/// <returns>The loaded ResourceDictionary or null if failed</returns>
	private static ResourceDictionary LoadFromPackUri(string resourceName)
	{
		try
		{
			var asmName = Assembly.GetExecutingAssembly().GetName().Name;
			var packUri = new Uri(
				$"pack://application:,,,/{asmName};component/Options/Themes/{resourceName}.xaml",
				UriKind.Absolute);
			return new ResourceDictionary { Source = packUri };
		}
		catch
		{
			return null;
		}
	}


	/// <summary>
	///     Loads a resource dictionary from disk
	/// </summary>
	/// <param name="resourceName">The name of the resource (without extension)</param>
	/// <returns>The loaded ResourceDictionary or null if failed</returns>
	private static ResourceDictionary LoadFromDisk(string resourceName)
	{
		try
		{
			var asmLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ?? ".";
			var resourcePath = Path.Combine(asmLocation, ThemesFolder.Replace('/', Path.DirectorySeparatorChar),
				resourceName + ThemeFileExtension);

			if (!File.Exists(resourcePath)) return null;

			using var fs = File.OpenRead(resourcePath);
			return (ResourceDictionary)XamlReader.Load(fs);
		}
		catch
		{
			return null;
		}
	}


	/// <summary>
	///     Applies a resource dictionary to the specified framework element
	/// </summary>
	/// <param name="element">The framework element to apply resources to</param>
	/// <param name="resourceDictionary">The resource dictionary to apply</param>
	private static void ApplyThemeResources(FrameworkElement element, ResourceDictionary resourceDictionary)
	{
		if (resourceDictionary == null) return;

		AddToWindowResources(element, resourceDictionary);
	}


	/// <summary>
	///     Adds a resource dictionary to framework element's merged dictionaries
	/// </summary>
	/// <param name="element">The framework element to add resources to</param>
	/// <param name="resourceDictionary">The resource dictionary to add</param>
	private static void AddToWindowResources(FrameworkElement element, ResourceDictionary resourceDictionary)
	{
		// Only remove existing dictionaries with the same name, not all theme dictionaries

		var resourceName = Path.GetFileNameWithoutExtension(resourceDictionary.Source?.OriginalString);
		var toRemove = element.Resources.MergedDictionaries.Where(md =>
		{
			var sourceName = Path.GetFileNameWithoutExtension(md.Source?.OriginalString);
			return sourceName == resourceName;
		}).ToList();

		foreach (var r in toRemove)
			element.Resources.MergedDictionaries.Remove(r);

		element.Resources.MergedDictionaries.Add(resourceDictionary);
	}
}