VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmEmpty 
   Caption         =   "fmEmpty"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "fmEmpty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmEmpty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Watchers As Scripting.Dictionary

Event CloseRequested(CloseMode As Integer)

Public Function Itself() As fmEmpty
    Set Itself = Me
End Function

Public Function Notify(ob As Object, _
    Optional ky As Variant = Empty _
) As Variant
    Dim dx As Long
    
    With Watchers
    If IsEmpty(ky) Then
        dx = .Count
        Do While .Exists(dx)
            dx = 1 + dx
        Loop
        Notify = Notify(ob, dx)
    Else
        If .Exists(ky) Then
            Notify = Empty
        Else
            .Add ky, ob
            Notify = ky
        End If
    End If: End With
End Function

Public Function NoMsgs(nm As Variant) As Variant
    With Watchers
    If .Exists(nm) Then
        .Remove nm
        NoMsgs = nm
    Else
        NoMsgs = Empty
    End If: End With
End Function

Private Sub UserForm_Initialize()
    Set Watchers = New Scripting.Dictionary
End Sub

Private Sub UserForm_QueryClose( _
    Cancel As Integer, CloseMode As Integer _
)
    Dim ck As VbMsgBoxResult
    
    Cancel = 1
    If Watchers.Count > 0 Then
        RaiseEvent CloseRequested(CloseMode)
    Else
        ck = MsgBox(Join(Array( _
            "Review any selections", _
            "and select Yes if ready.", _
            "Otherwise, select No." _
        ), vbNewLine), vbYesNo, _
            "Close Form?" _
        )
        If ck = vbYes Then
            Me.Hide
        Else
        End If
    End If
End Sub

Private Sub UserForm_Terminate()
    Watchers.RemoveAll
    Set Watchers = Nothing
End Sub
