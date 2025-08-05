Imports System.Windows.Forms

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
    End Class
End NameSpace