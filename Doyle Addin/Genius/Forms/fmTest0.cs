
Attribute VB_Name = "fmTest0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Public Function ft0g0f0(im As stdole.StdPicture) As Long
    With Me
        .imTNail.Picture = im
        .Show 1
    End With
    ft0g0f0 = 0
End Function

Private Sub UserForm_QueryClose( _
    Cancel As Integer, CloseMode As Integer _
)
    Cancel = 1
    Me.Hide
End Sub
