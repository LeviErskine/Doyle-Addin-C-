<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class UserOptionsForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        Dim Label1 As System.Windows.Forms.Label
        Dim Label2 As System.Windows.Forms.Label
        Dim TextBox1 As System.Windows.Forms.TextBox
        BtnSave = New System.Windows.Forms.Button()
        PrintExportFolder = New System.Windows.Forms.FolderBrowserDialog()
        DXFexportFolder = New System.Windows.Forms.FolderBrowserDialog()
        Panel1 = New System.Windows.Forms.Panel()
        PEXLoc = New System.Windows.Forms.TextBox()
        PrintExportLocationButton = New System.Windows.Forms.Button()
        DXFexLoc = New System.Windows.Forms.TextBox()
        DXFExportLocationButton = New System.Windows.Forms.Button()
        LayerColorChooser = New System.Windows.Forms.ColorDialog()
        PDFToImageBindingSource = New System.Windows.Forms.BindingSource(components)
        Label1 = New System.Windows.Forms.Label()
        Label2 = New System.Windows.Forms.Label()
        TextBox1 = New System.Windows.Forms.TextBox()
        Panel1.SuspendLayout()
        CType(PDFToImageBindingSource, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.Location = New System.Drawing.Point(8, 7)
        Label1.Name = "Label1"
        Label1.Size = New System.Drawing.Size(117, 15)
        Label1.TabIndex = 9
        Label1.Text = "Print Export Location"
        ' 
        ' Label2
        ' 
        Label2.AutoSize = True
        Label2.Location = New System.Drawing.Point(12, 35)
        Label2.Name = "Label2"
        Label2.Size = New System.Drawing.Size(113, 15)
        Label2.TabIndex = 10
        Label2.Text = "DXF Export Location"
        ' 
        ' BtnSave
        ' 
        BtnSave.Dock = System.Windows.Forms.DockStyle.Bottom
        BtnSave.Location = New System.Drawing.Point(3, 82)
        BtnSave.Name = "BtnSave"
        BtnSave.Size = New System.Drawing.Size(386, 35)
        BtnSave.TabIndex = 5
        BtnSave.Text = "Save"
        BtnSave.UseVisualStyleBackColor = True
        ' 
        ' PrintExportFolder
        ' 
        PrintExportFolder.InitialDirectory = "P:\"
        PrintExportFolder.RootFolder = Environment.SpecialFolder.MyComputer
        PrintExportFolder.SelectedPath = "P:\"
        ' 
        ' DXFexportFolder
        ' 
        DXFexportFolder.InitialDirectory = "X:\"
        DXFexportFolder.RootFolder = Environment.SpecialFolder.MyComputer
        DXFexportFolder.SelectedPath = "X:\"
        ' 
        ' Panel1
        ' 
        Panel1.AutoSize = True
        Panel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Panel1.Controls.Add(Label2)
        Panel1.Controls.Add(PEXLoc)
        Panel1.Controls.Add(Label1)
        Panel1.Controls.Add(PrintExportLocationButton)
        Panel1.Controls.Add(DXFexLoc)
        Panel1.Controls.Add(DXFExportLocationButton)
        Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Panel1.Location = New System.Drawing.Point(3, 3)
        Panel1.Name = "Panel1"
        Panel1.Size = New System.Drawing.Size(386, 58)
        Panel1.TabIndex = 9
        ' 
        ' PEXLoc
        ' 
        PEXLoc.AllowDrop = True
        PEXLoc.Location = New System.Drawing.Point(131, 3)
        PEXLoc.Name = "PEXLoc"
        PEXLoc.PlaceholderText = "P:/"
        PEXLoc.Size = New System.Drawing.Size(147, 23)
        PEXLoc.TabIndex = 7
        PEXLoc.WordWrap = False
        ' 
        ' PrintExportLocationButton
        ' 
        PrintExportLocationButton.Location = New System.Drawing.Point(284, 3)
        PrintExportLocationButton.Name = "PrintExportLocationButton"
        PrintExportLocationButton.Size = New System.Drawing.Size(86, 23)
        PrintExportLocationButton.TabIndex = 6
        PrintExportLocationButton.Text = "Browse"
        PrintExportLocationButton.UseVisualStyleBackColor = True
        ' 
        ' DXFexLoc
        ' 
        DXFexLoc.AllowDrop = True
        DXFexLoc.Location = New System.Drawing.Point(131, 32)
        DXFexLoc.Name = "DXFexLoc"
        DXFexLoc.PlaceholderText = "X:/"
        DXFexLoc.Size = New System.Drawing.Size(147, 23)
        DXFexLoc.TabIndex = 9
        ' 
        ' DXFExportLocationButton
        ' 
        DXFExportLocationButton.Location = New System.Drawing.Point(284, 32)
        DXFExportLocationButton.Name = "DXFExportLocationButton"
        DXFExportLocationButton.Size = New System.Drawing.Size(86, 23)
        DXFExportLocationButton.TabIndex = 8
        DXFExportLocationButton.Text = "Browse"
        DXFExportLocationButton.UseVisualStyleBackColor = True
        ' 
        ' LayerColorChooser
        ' 
        LayerColorChooser.AnyColor = True
        LayerColorChooser.ShowHelp = True
        LayerColorChooser.SolidColorOnly = True
        ' 
        ' PDFToImageBindingSource
        ' 
        PDFToImageBindingSource.DataSource = GetType(PDFToImage)
        ' 
        ' TextBox1
        ' 
        TextBox1.BackColor = Drawing.SystemColors.Control
        TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        TextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        TextBox1.Font = New System.Drawing.Font("Segoe UI", 12F)
        TextBox1.Location = New System.Drawing.Point(3, 61)
        TextBox1.Name = "TextBox1"
        TextBox1.Size = New System.Drawing.Size(386, 22)
        TextBox1.TabIndex = 11
        TextBox1.Text = "More options planned for the future"
        TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        ' 
        ' UserOptionsForm
        ' 
        AutoScaleDimensions = New System.Drawing.SizeF(7F, 15F)
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        ClientSize = New System.Drawing.Size(392, 120)
        Controls.Add(TextBox1)
        Controls.Add(Panel1)
        Controls.Add(BtnSave)
        Name = "UserOptionsForm"
        Padding = New System.Windows.Forms.Padding(3)
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Text = "Options"
        Panel1.ResumeLayout(False)
        Panel1.PerformLayout()
        CType(PDFToImageBindingSource, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub
    Friend WithEvents FlowLayoutPanel1 As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents BtnSave As System.Windows.Forms.Button
    Friend WithEvents PrintExportFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents DXFexportFolder As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents PEXLoc As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PrintExportLocationButton As System.Windows.Forms.Button
    Friend WithEvents DXFexLoc As System.Windows.Forms.TextBox
    Friend WithEvents DXFExportLocationButton As System.Windows.Forms.Button
    Friend WithEvents LayerColorChooser As System.Windows.Forms.ColorDialog
    Friend WithEvents PDFToImageBindingSource As System.Windows.Forms.BindingSource
End Class
