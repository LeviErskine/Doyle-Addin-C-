
Attribute VB_Name = "fmTest3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub UserForm_Click()
    '
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Debug.Print ft3g1f0(frmShp, "QUICK BROWN FOX JUMPED LAZY DOG")
End Sub

Private Sub UserForm_Initialize()
    Dim ct As MSForms.Control
    Dim cb As MSForms.CheckBox
    Dim dx As Long
    Dim tp As Long
    Dim gp As Long
    Dim cp As String
    
    tp = 18
    gp = 0
    
    For dx = 1 To 3
        cp = "CB" & CStr(dx)

        ct = frmShp.Controls.Add("Forms.CheckBox.1", cp, True) 'Me
        With ct
            .Height = 18
            .Width = 96
            .Left = 18
            .Top = tp

            tp = tp + .Height + gp
        End With

        cb = ct
        With cb
            .Caption = cp
        End With
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Cancel = 1
    Me.Hide()

    Debug.Print ft3g0f0(frmShp, "Check")

    Dim cb As MSForms.CheckBox
    For Each cb In frmShp.Controls
        'Stop
        If cb.Value Then
            Debug.Print cb.Caption
        End If
    Next
End Sub

Private Function ft3g0f0(
    src As MSForms.Frame, fdName As String
) As String
    Dim ct As MSForms.Control
    Dim cb As MSForms.CheckBox
    Dim rt As String

    For Each ct In src.Controls
        If TypeOf ct Is MSForms.CheckBox Then
            cb = ct
            If cb.Value Then
                If Len(rt) > 0 Then
                    'rt = rt & " OR "
                    rt = rt & ", "
                End If
                'rt = rt & fdName & " = '" & cb.Caption & "'"
                rt = rt & "'" & cb.Caption & "'"
            End If
        End If
    Next

    'ft3g0f0 = "(" & rt & ")"
    ft3g0f0 = fdName & " IN (" & rt & ")"
End Function

Private Function ft3g1f0(
    frm As MSForms.Frame, ls As String,
    Optional bk As String = " "
) As Long
    Dim ctrl As MSForms.Control
    Dim cb As MSForms.CheckBox
    Dim rt As String
    Dim cp As Variant
    Dim ct As Long

    Dim tp As Long
    Dim gp As Long

    tp = 18
    gp = 0

    With frm.Controls '.Remove
        Do While .Count > 0
            .Remove 0
        Loop

        For Each cp In Split(ls, bk)
            If Len(cp) > 0 Then
                ctrl = .Add("Forms.CheckBox.1", cp, True)
                With ctrl
                    .Height = 18
                    .Width = 96
                    .Left = 18
                    .Top = tp

                    tp = tp + .Height + gp
                End With

                cb = ctrl
                With cb
                    .Caption = cp
                End With
                
                ct = ct + 1
            End If
        Next
    End With
    
    'ft3g1f0 = "(" & rt & ")"
    'ft3g1f0 = ls & " IN (" & rt & ")"
    ft3g1f0 = ct
End Function
