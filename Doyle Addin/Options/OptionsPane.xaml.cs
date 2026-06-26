namespace DoyleAddin.Options;

using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;
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

		UpdateBetaButtons();
	}

	private void UpdateBetaButtons()
	{
		OptInBeta.Visibility  = options.IsBetaEnabled ? Visibility.Collapsed : Visibility.Visible;
		OptOutBeta.Visibility = options.IsBetaEnabled ? Visibility.Visible : Visibility.Collapsed;
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

	private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
	{
		Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
		e.Handled = true;
		Close();
	}

	private void BtnCncl_Click(object sender, RoutedEventArgs e)
	{
		DialogResult = false;
		Close();
	}

	private void OptInBeta_Click(object sender, RoutedEventArgs e)
	{
		var result = MessageBox.Show(
			"This will close Inventor and install the latest Beta version of the Doyle Add-in. Continue?",
			"Opt In to Beta",
			MessageBoxButton.YesNo,
			MessageBoxImage.Warning);

		if (result != MessageBoxResult.Yes) return;

		options.IsBetaEnabled = true;
		options.Save();
		UpdateBetaButtons();

		Process.Start(new ProcessStartInfo
		{
			FileName        = "Doyle Addin Installer Beta.bat",
			UseShellExecute = true
		});

		Application.Current.Shutdown();
	}

	private void OptOutBeta_Click(object sender, RoutedEventArgs e)
	{
		var result = MessageBox.Show(
			"This will close Inventor and install the latest stable Release of the Doyle Add-in. Continue?",
			"Opt Out of Beta",
			MessageBoxButton.YesNo,
			MessageBoxImage.Warning);

		if (result != MessageBoxResult.Yes) return;

		options.IsBetaEnabled = false;
		options.Save();
		UpdateBetaButtons();

		Process.Start(new ProcessStartInfo
		{
			FileName        = "Doyle Addin Installer.bat",
			UseShellExecute = true
		});

		Application.Current.Shutdown();
	}
}