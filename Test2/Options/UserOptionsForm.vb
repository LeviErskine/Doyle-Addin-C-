Imports System.Windows.Forms
Imports Inventor

Namespace Options
    Public Class UserOptionsForm
        Private _options As New UserOptions
        Public Sub New()
            InitializeComponent()
            _options = UserOptions.Load()
        End Sub

        Private Sub UserOptionsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
            _options = UserOptions.Load()
            PEXLoc.Text = _options.PrintExportLocation
            DXFexLoc.Text = _options.DxfExportLocation

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
                FutureMsg.BackColor = Drawing.Color.FromArgb(CByte(245), CByte(245), CByte(245))
                FutureMsg.ForeColor = Drawing.Color.Black
                BtnSave.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(186), CByte(186), CByte(186))
                BtnSave.FlatAppearance.MouseOverBackColor = Drawing.Color.White
                PEXLoc.BackColor = Drawing.Color.White
                PEXLoc.ForeColor = Drawing.Color.Black
                PrintExportLocationButton.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(186), CByte(186), CByte(186))
                PrintExportLocationButton.FlatAppearance.MouseOverBackColor = Drawing.Color.White
                PrintExportLocationButton.ForeColor = Drawing.Color.Black
                DXFexLoc.BackColor = Drawing.Color.White
                DXFexLoc.ForeColor = Drawing.Color.Black
                DXFExportLocationButton.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(186), CByte(186), CByte(186))
                DXFExportLocationButton.FlatAppearance.MouseOverBackColor = Drawing.Color.White
                BackColor = Drawing.Color.FromArgb(CByte(245), CByte(245), CByte(245))
                ForeColor = Drawing.Color.Black
                SCBackground.BackColor = Drawing.Color.FromArgb(CByte(245), CByte(245), CByte(245))
                BtnCncl.FlatAppearance.MouseOverBackColor = Drawing.Color.White
                BtnCncl.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(186), CByte(186), CByte(186))
            Else
            End If
        End Sub

        Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
            _options.PrintExportLocation = PEXLoc.Text
            _options.DxfExportLocation = DXFexLoc.Text
            _options.Save()
            Close()
        End Sub

        Private Sub PrintExportLocationButton_Click(sender As Object, e As EventArgs) _
            Handles PrintExportLocationButton.Click
            Dim folderBrowser As New FolderBrowserDialog With {
                    .Description = "Select Print Export Location"
                    }
            folderBrowser.ShowDialog()
            PEXLoc.Text = folderBrowser.SelectedPath
        End Sub

        Private Sub DXFExportLocationButton_Click(sender As Object, e As EventArgs) _
            Handles DXFExportLocationButton.Click
            Dim folderBrowser As New FolderBrowserDialog With {
                    .Description = "Select DXF Export Location"
                    }
            folderBrowser.ShowDialog()
            ' Assuming you have a text box to display the selected path
            DXFexLoc.Text = folderBrowser.SelectedPath
        End Sub
        Private Sub BtnCncl_Click(sender As Object, e As EventArgs) Handles BtnCncl.Click
            Close()
        End Sub
    End Class
End Namespace