Imports System.Windows.Forms

Public Class UserOptionsForm
    Private Options As New UserOptions
    Public Sub New()
        InitializeComponent()
        Options = UserOptions.Load()
    End Sub

    Private Sub UserOptionsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Options = UserOptions.Load()
        PEXLoc.Text = Options.PrintExportLocation
        DXFexLoc.Text = Options.DXFExportLocation
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As EventArgs) Handles BtnSave.Click
        Options.PrintExportLocation = PEXLoc.Text
        Options.DXFExportLocation = DXFexLoc.Text
        Options.Save()
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
End Class