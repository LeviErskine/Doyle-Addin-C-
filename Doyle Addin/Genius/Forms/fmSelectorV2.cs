VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmSelectorV2 
   Caption         =   "fmSelectorV2"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "fmSelectorV2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmSelectorV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const tagSelected As String = "%%%"
Private dc As Scripting.Dictionary

Private msCancelHead As String
Private msCancelMain As String
Private msNoSelHead As String
Private msNoSelMain As String
Private msOkHead As String
Private msOkMain As String
'Private msCancelMain As String
'Private msCancelMain As String

Public Function SetMsgCancel(Using As String) As fmSelectorV2
    msCancelMain = Using
     SetMsgCancel = Me
End Function

Public Function SetMsgNoSelection(Using As String) As fmSelectorV2
    msNoSelMain = Using
     SetMsgNoSelection = Me
End Function

Public Function SetMsgOK(Using As String) As fmSelectorV2
    msOkMain = Using
     SetMsgOK = Me
End Function

Public Function SetHdrCancel(Using As String) As fmSelectorV2
    msCancelHead = Using
     SetHdrCancel = Me
End Function

Public Function SetHdrNoSelection(Using As String) As fmSelectorV2
    msNoSelHead = Using
     SetHdrNoSelection = Me
End Function

Public Function SetHdrOK(Using As String) As fmSelectorV2
    msOkHead = Using
     SetHdrOK = Me
End Function

Public Function SelectIfIn(Using As String) As fmSelectorV2
    Dim dx As Long
    
    With Me.lsbSelection
        dx = .ListIndex
        On Error Resume Next
        Err.Clear
        .Value = Using
        If Err.Number Then
            .ListIndex = dx
            '.Value = ""
            Err.Clear
        End If
        On Error GoTo 0
    End With

    SelectIfIn = Me
End Function

Public Function WithList(Using As Variant) As fmSelectorV2
    Dim ky As Variant
    Dim it As String
    
    If IsObject(Using) Then
        'Me.lsbSelection.List = Using
        If TypeOf Using Is Scripting.Dictionary Then
            dc = Using
        ElseIf TypeOf Using Is Inventor.NameValueMap Then
            dc = dcFromAiNameValMap(obOf(Using))
        Else
            'Stop
            Debug.Print ; 'Breakpoint Landing
             dc = Nothing
        End If
        
        If dc Is Nothing Then
            Me.lsbSelection.List = Array("<no items>")
        ElseIf dc.Count > 0 Then
            Me.lsbSelection.List = dc.Keys
        Else
            Me.lsbSelection.List = Array("<no items>")
        End If
    ElseIf IsArray(Using) Then
        dc = New Scripting.Dictionary

        For Each ky In Using
            dc.Add CStr(ky), ky
        Next
        
        Me.lsbSelection.List = dc.Keys
    Else
        'Stop
        Debug.Print ; 'Breakpoint Landing
    End If

    WithList = Me
End Function

Private Sub btnCancel_Click()
    ''
    If MsgBox( _
        msCancelMain, vbYesNo, msCancelHead _
    ) = vbYes Then
        Me.lsbSelection.ListIndex = -1
        Me.Hide
    Else
        'Do nothing
    End If
End Sub

Private Sub btnOk_Click()
    ''
    Dim ck As VbMsgBoxResult
    Dim mx As Long
    Dim dx As Long
    Dim ct As Long
    
    Dim ls As String
    
    With Me.lsbSelection
        If .MultiSelect = fmMultiSelectSingle Then
            If .ListIndex < 0 Then
                ck = MsgBox(msNoSelMain, vbYesNo, msNoSelHead)
                If ck = vbYes Then Me.Hide
            Else
                ck = MsgBox( _
                    Join(Split(msOkMain, tagSelected), .Value), _
                    vbYesNoCancel, msOkHead _
                )
                If ck = vbYes Then
                    Me.Hide
                ElseIf ck = vbCancel Then
                    .ListIndex = -1
                    Me.Hide
                Else
                    'Do nothing
                End If
            End If
        Else
            ls = lbxPickedStr(Me.lsbSelection, vbNewLine)
            
            'ct = 0
            'mx = .ListCount - 1
            'For dx = 0 To mx
                'If .Selected(dx) Then ct = 1 + ct
            'Next
            
            If Len(ls) > 0 Then
                ck = MsgBox(Join( _
                    Split(msOkMain, tagSelected), _
                    vbNewLine & ls & vbNewLine _
                    ), vbYesNoCancel, msOkHead _
                )
                If ck = vbYes Then
                    Me.Hide
                ElseIf ck = vbCancel Then
                    .ListIndex = -1
                    Me.Hide
                Else
                    'Do nothing
                End If
            Else
                ck = MsgBox(msNoSelMain, vbYesNo, msNoSelHead)
                If ck = vbYes Then Me.Hide
            End If
        End If
    End With
End Sub

Private Sub lsbSelection_Change()
    '''
    '''
    Dim ck As String
    
    On Error Resume Next
    Err.Clear
    ck = Me.lsbSelection.Value
    If Err.Number <> 0 Then ck = ""
    On Error GoTo 0
    
    With dc
        If .Exists(ck) Then
            On Error Resume Next
            
            Err.Clear
            tbxView.Value = CStr(.Item(ck))
            If Err.Number <> 0 Then
                tbxView.Value = "<not printable>"
            End If
            
            On Error GoTo 0
        Else
            tbxView.Value = "<no data>"
        End If
    End With
End Sub

Private Sub lsbSelection_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    btnOk_Click
End Sub

Private Sub UserForm_Initialize()
    '
    msCancelHead = "Cancel Operation?"
    msNoSelHead = "No Selection!"
    msOkHead = "Proceed?"
    '
    msCancelMain = "Selection will be canceled."
    msNoSelMain = Join(Array( _
        "Do you wish to cancel?", _
        "(Click NO to return to list)" _
    ), vbNewLine)
    msOkMain = Join(Array( _
        "Current selection is: ", tagSelected, _
        "(Click CANCEL to quit with no selection)" _
    ), vbNewLine)
    '
End Sub

Private Sub UserForm_QueryClose( _
    Cancel As Integer, CloseMode As Integer _
)
    '''
    Cancel = 1
    btnCancel_Click
End Sub

Public Function GetReply( _
    Optional List As Variant, _
    Optional Default As String = "%$#@" _
) As String
    Dim rt As String
    
    rt = ""
    With Me.WithList(List).SelectIfIn(Default)
        '.lsbSelection.List = lsWorkbooks()
        .Show 1
        If .lsbSelection.MultiSelect = fmMultiSelectSingle Then
            rt = .lsbSelection.Text
        Else
            rt = lbxPickedStr(.lsbSelection, vbNewLine)
        End If
    End With
    GetReply = rt
End Function

Private Sub UserForm_Resize()
    '''
End Sub
