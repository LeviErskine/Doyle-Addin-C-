Imports System.Windows.Forms
Imports Inventor
Imports Color = System.Drawing.Color

Namespace Options
	Public Class UserOptionsForm
		Private options As New UserOptions

		Public Sub New()
			InitializeComponent()
			options = UserOptions.Load()
		End Sub

		Private Sub UserOptionsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
			options = UserOptions.Load()
			PEXLoc.Text = options.PrintExportLocation
			DXFexLoc.Text = options.DxfExportLocation
			ChkObsoletePrint.Checked = options.EnableObsoletePrint

			' Apply theme colors
			ApplyThemeColors()
		End Sub

		Private Sub ApplyThemeColors()
			Dim oThemeManager As ThemeManager
			oThemeManager = ThisApplication.ThemeManager
			Dim oTheme As Theme
			oTheme = oThemeManager.ActiveTheme
			If oTheme.Name = "LightTheme" Then
				' Dark theme colors
				FutureMsg.BackColor = Color.FromArgb(CByte(245), CByte(245), CByte(245))
				FutureMsg.ForeColor = Color.Black
				BtnSave.FlatAppearance.BorderColor = Color.FromArgb(CByte(186), CByte(186), CByte(186))
				BtnSave.FlatAppearance.MouseOverBackColor = Color.White
				PEXLoc.BackColor = Color.White
				PEXLoc.ForeColor = Color.Black
				PrintExportLocationButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(186), CByte(186), CByte(186))
				PrintExportLocationButton.FlatAppearance.MouseOverBackColor = Color.White
				PrintExportLocationButton.ForeColor = Color.Black
				DXFexLoc.BackColor = Color.White
				DXFexLoc.ForeColor = Color.Black
				DXFExportLocationButton.FlatAppearance.BorderColor = Color.FromArgb(CByte(186), CByte(186), CByte(186))
				DXFExportLocationButton.FlatAppearance.MouseOverBackColor = Color.White
				BackColor = Color.FromArgb(CByte(245), CByte(245), CByte(245))
				ForeColor = Color.Black
				SCBackground.BackColor = Color.FromArgb(CByte(245), CByte(245), CByte(245))
				BtnCncl.FlatAppearance.MouseOverBackColor = Color.White
				BtnCncl.FlatAppearance.BorderColor = Color.FromArgb(CByte(186), CByte(186), CByte(186))
				FeaturesPanel.BackColor = Color.FromArgb(CByte(245), CByte(245), CByte(245))
				ChkObsoletePrint.BackColor = Color.FromArgb(CByte(245), CByte(245), CByte(245))
				ChkObsoletePrint.ForeColor = Color.Black
			Else
				' Dark theme colors (keep existing dark theme as default)
				FeaturesPanel.BackColor = Color.FromArgb(CByte(59), CByte(68), CByte(83))
				ChkObsoletePrint.BackColor = Color.FromArgb(CByte(59), CByte(68), CByte(83))
				ChkObsoletePrint.ForeColor = Color.White
			End If
		End Sub

		Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
			options.PrintExportLocation = PEXLoc.Text
			options.DxfExportLocation = DXFexLoc.Text
			options.EnableObsoletePrint = ChkObsoletePrint.Checked
			options.Save()
			DialogResult = DialogResult.OK
			Close()
		End Sub

		Private Sub PrintExportLocationButton_Click(sender As Object, e As EventArgs) Handles PrintExportLocationButton.Click
			Dim folderBrowser As New FolderBrowserDialog With {
				    .Description = "Select Print Export Location"
				    }
			folderBrowser.ShowDialog()
			PEXLoc.Text = folderBrowser.SelectedPath
		End Sub

		Private Sub DXFExportLocationButton_Click(sender As Object, e As EventArgs) Handles DXFExportLocationButton.Click
			Dim folderBrowser As New FolderBrowserDialog With {
				    .Description = "Select DXF Export Location"
				    }
			folderBrowser.ShowDialog()
			' Assuming you have a text box to display the selected path
			DXFexLoc.Text = folderBrowser.SelectedPath
		End Sub

		Private Sub BtnCncl_Click(sender As Object, e As EventArgs) Handles BtnCncl.Click
			DialogResult = DialogResult.Cancel
			Close()
		End Sub
	End Class
End Namespace