namespace DoyleAddin.Options;

using System.Windows;
using System.Windows.Controls;
using Ookii.Dialogs.Wpf;
using Themes;

/// <summary>
/// </summary>
public partial class UserOptionsWindow
{
	private UserOptions options;

	/// <inheritdoc />
	public UserOptionsWindow()
	{
		InitializeComponent();
	}

	private void UserOptionsWindow_Loaded(object sender, RoutedEventArgs e)
	{
		options = UserOptions.Load();

		// Set the DataContext for potential data binding
		DataContext = options;

		// Apply the theme
		ThemeManager.ApplyTheme(this);
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

		// When ShowDialog() returns true, user has selected a folder
		if (folderBrowser.ShowDialog(this).GetValueOrDefault()) targetTextBox.Text = folderBrowser.SelectedPath;
	}

	private void BtnCncl_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}
}