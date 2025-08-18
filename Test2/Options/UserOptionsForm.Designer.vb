Namespace Options
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
        <System.Diagnostics.DebuggerStepThrough()>
        Private Sub InitializeComponent()
            Dim PrintText As System.Windows.Forms.Label
            Dim DXFText As System.Windows.Forms.Label
            FutureMsg = New System.Windows.Forms.TextBox()
            BtnSave = New System.Windows.Forms.Button()
            PrintExportFolder = New System.Windows.Forms.FolderBrowserDialog()
            DXFexportFolder = New System.Windows.Forms.FolderBrowserDialog()
            Destinationbackground = New System.Windows.Forms.Panel()
            DXFexLoc = New System.Windows.Forms.TextBox()
            PEXLoc = New System.Windows.Forms.TextBox()
            PrintExportLocationButton = New System.Windows.Forms.Button()
            DXFExportLocationButton = New System.Windows.Forms.Button()
            BtnCncl = New System.Windows.Forms.Button()
            SCBackground = New System.Windows.Forms.Panel()
            PrintText = New System.Windows.Forms.Label()
            DXFText = New System.Windows.Forms.Label()
            Destinationbackground.SuspendLayout()
            SCBackground.SuspendLayout()
            SuspendLayout()
            ' 
            ' PrintText
            ' 
            PrintText.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            PrintText.Location = New System.Drawing.Point(8, 10)
            PrintText.Name = "PrintText"
            PrintText.Size = New System.Drawing.Size(129, 15)
            PrintText.TabIndex = 9
            PrintText.Text = "Print Export Location"
            ' 
            ' DXFText
            ' 
            DXFText.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            DXFText.Location = New System.Drawing.Point(8, 36)
            DXFText.Name = "DXFText"
            DXFText.Size = New System.Drawing.Size(129, 17)
            DXFText.TabIndex = 10
            DXFText.Text = "DXF Export Location"
            ' 
            ' FutureMsg
            ' 
            FutureMsg.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(68), CByte(83))
            FutureMsg.BorderStyle = System.Windows.Forms.BorderStyle.None
            FutureMsg.Dock = System.Windows.Forms.DockStyle.Bottom
            FutureMsg.Enabled = False
            FutureMsg.Font = New System.Drawing.Font("Tahoma", 12.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(0))
            FutureMsg.ForeColor = Drawing.Color.White
            FutureMsg.HideSelection = False
            FutureMsg.Location = New System.Drawing.Point(3, 65)
            FutureMsg.Name = "FutureMsg"
            FutureMsg.ReadOnly = True
            FutureMsg.Size = New System.Drawing.Size(344, 20)
            FutureMsg.TabIndex = 11
            FutureMsg.Text = "More options planned for the future"
            FutureMsg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
            ' 
            ' BtnSave
            ' 
            BtnSave.Dock = System.Windows.Forms.DockStyle.Left
            BtnSave.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(134), CByte(145), CByte(161))
            BtnSave.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(CByte(44), CByte(51), CByte(64))
            BtnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            BtnSave.Font = New System.Drawing.Font("Tahoma", 10.0F)
            BtnSave.Location = New System.Drawing.Point(3, 3)
            BtnSave.Name = "BtnSave"
            BtnSave.Padding = New System.Windows.Forms.Padding(3)
            BtnSave.Size = New System.Drawing.Size(162, 31)
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
            ' Destinationbackground
            ' 
            Destinationbackground.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            Destinationbackground.Controls.Add(DXFexLoc)
            Destinationbackground.Controls.Add(DXFText)
            Destinationbackground.Controls.Add(PEXLoc)
            Destinationbackground.Controls.Add(PrintText)
            Destinationbackground.Controls.Add(PrintExportLocationButton)
            Destinationbackground.Controls.Add(DXFExportLocationButton)
            Destinationbackground.Dock = System.Windows.Forms.DockStyle.Top
            Destinationbackground.Font = New System.Drawing.Font("Tahoma", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(0))
            Destinationbackground.Location = New System.Drawing.Point(3, 3)
            Destinationbackground.Name = "Destinationbackground"
            Destinationbackground.Padding = New System.Windows.Forms.Padding(3)
            Destinationbackground.Size = New System.Drawing.Size(344, 63)
            Destinationbackground.TabIndex = 9
            ' 
            ' DXFexLoc
            ' 
            DXFexLoc.AllowDrop = True
            DXFexLoc.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            DXFexLoc.BackColor = Drawing.Color.FromArgb(CByte(44), CByte(51), CByte(64))
            DXFexLoc.BorderStyle = System.Windows.Forms.BorderStyle.None
            DXFexLoc.ForeColor = Drawing.Color.White
            DXFexLoc.Location = New System.Drawing.Point(132, 37)
            DXFexLoc.Name = "DXFexLoc"
            DXFexLoc.PlaceholderText = "X:/"
            DXFexLoc.Size = New System.Drawing.Size(127, 15)
            DXFexLoc.TabIndex = 9
            ' 
            ' PEXLoc
            ' 
            PEXLoc.AllowDrop = True
            PEXLoc.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            PEXLoc.BackColor = Drawing.Color.FromArgb(CByte(44), CByte(51), CByte(64))
            PEXLoc.BorderStyle = System.Windows.Forms.BorderStyle.None
            PEXLoc.ForeColor = Drawing.Color.White
            PEXLoc.Location = New System.Drawing.Point(132, 10)
            PEXLoc.Name = "PEXLoc"
            PEXLoc.PlaceholderText = "P:/"
            PEXLoc.Size = New System.Drawing.Size(127, 15)
            PEXLoc.TabIndex = 7
            PEXLoc.WordWrap = False
            ' 
            ' PrintExportLocationButton
            ' 
            PrintExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            PrintExportLocationButton.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(134), CByte(145), CByte(161))
            PrintExportLocationButton.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(CByte(44), CByte(51), CByte(64))
            PrintExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            PrintExportLocationButton.Location = New System.Drawing.Point(262, 7)
            PrintExportLocationButton.Margin = New System.Windows.Forms.Padding(0)
            PrintExportLocationButton.Name = "PrintExportLocationButton"
            PrintExportLocationButton.Size = New System.Drawing.Size(69, 20)
            PrintExportLocationButton.TabIndex = 6
            PrintExportLocationButton.TabStop = False
            PrintExportLocationButton.Text = "Browse"
            PrintExportLocationButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
            PrintExportLocationButton.UseCompatibleTextRendering = True
            PrintExportLocationButton.UseVisualStyleBackColor = True
            ' 
            ' DXFExportLocationButton
            ' 
            DXFExportLocationButton.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right
            DXFExportLocationButton.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(134), CByte(145), CByte(161))
            DXFExportLocationButton.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(CByte(44), CByte(51), CByte(64))
            DXFExportLocationButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            DXFExportLocationButton.Location = New System.Drawing.Point(262, 34)
            DXFExportLocationButton.Margin = New System.Windows.Forms.Padding(0)
            DXFExportLocationButton.Name = "DXFExportLocationButton"
            DXFExportLocationButton.Size = New System.Drawing.Size(69, 20)
            DXFExportLocationButton.TabIndex = 8
            DXFExportLocationButton.TabStop = False
            DXFExportLocationButton.Text = "Browse"
            DXFExportLocationButton.TextImageRelation = System.Windows.Forms.TextImageRelation.TextAboveImage
            DXFExportLocationButton.UseCompatibleTextRendering = True
            DXFExportLocationButton.UseVisualStyleBackColor = True
            ' 
            ' BtnCncl
            ' 
            BtnCncl.Dock = System.Windows.Forms.DockStyle.Right
            BtnCncl.FlatAppearance.BorderColor = Drawing.Color.FromArgb(CByte(134), CByte(145), CByte(161))
            BtnCncl.FlatAppearance.MouseOverBackColor = Drawing.Color.FromArgb(CByte(44), CByte(51), CByte(64))
            BtnCncl.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            BtnCncl.Font = New System.Drawing.Font("Tahoma", 10.0F)
            BtnCncl.Location = New System.Drawing.Point(179, 3)
            BtnCncl.Name = "BtnCncl"
            BtnCncl.Padding = New System.Windows.Forms.Padding(3)
            BtnCncl.Size = New System.Drawing.Size(162, 31)
            BtnCncl.TabIndex = 12
            BtnCncl.Text = "Cancel"
            BtnCncl.UseVisualStyleBackColor = True
            ' 
            ' SCBackground
            ' 
            SCBackground.BackColor = Drawing.Color.FromArgb(CByte(59), CByte(68), CByte(83))
            SCBackground.Controls.Add(BtnSave)
            SCBackground.Controls.Add(BtnCncl)
            SCBackground.Dock = System.Windows.Forms.DockStyle.Bottom
            SCBackground.Font = New System.Drawing.Font("Tahoma", 9.0F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(0))
            SCBackground.Location = New System.Drawing.Point(3, 85)
            SCBackground.Name = "SCBackground"
            SCBackground.Padding = New System.Windows.Forms.Padding(3)
            SCBackground.Size = New System.Drawing.Size(344, 37)
            SCBackground.TabIndex = 13
            ' 
            ' UserOptionsForm
            ' 
            AutoScaleDimensions = New System.Drawing.SizeF(6.0F, 13.0F)
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            BackColor = Drawing.Color.FromArgb(CByte(59), CByte(68), CByte(83))
            ClientSize = New System.Drawing.Size(350, 125)
            Controls.Add(FutureMsg)
            Controls.Add(Destinationbackground)
            Controls.Add(SCBackground)
            Font = New System.Drawing.Font("Tahoma", 8.25F, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point, CByte(0))
            ForeColor = Drawing.Color.White
            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
            Name = "UserOptionsForm"
            Padding = New System.Windows.Forms.Padding(3)
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Text = "Options"
            Destinationbackground.ResumeLayout(False)
            Destinationbackground.PerformLayout()
            SCBackground.ResumeLayout(False)
            ResumeLayout(False)
            PerformLayout()
        End Sub
        Friend WithEvents BtnSave As System.Windows.Forms.Button
        Friend WithEvents PrintExportFolder As System.Windows.Forms.FolderBrowserDialog
        Friend WithEvents DXFexportFolder As System.Windows.Forms.FolderBrowserDialog
        Friend WithEvents Destinationbackground As System.Windows.Forms.Panel
        Friend WithEvents DXFText As System.Windows.Forms.Label
        Friend WithEvents PEXLoc As System.Windows.Forms.TextBox
        Friend WithEvents PrintText As System.Windows.Forms.Label
        Friend WithEvents PrintExportLocationButton As System.Windows.Forms.Button
        Friend WithEvents DXFexLoc As System.Windows.Forms.TextBox
        Friend WithEvents DXFExportLocationButton As System.Windows.Forms.Button
        Friend WithEvents FutureMsg As System.Windows.Forms.TextBox
        Friend WithEvents BtnCncl As System.Windows.Forms.Button
        Friend WithEvents SCBackground As System.Windows.Forms.Panel
    End Class
End NameSpace