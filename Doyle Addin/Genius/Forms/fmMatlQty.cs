VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmMatlQty 
   Caption         =   "Set/Verify Material Quantity"
   ClientHeight    =   6960
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4230
   OleObjectBlob   =   "fmMatlQty.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmMatlQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Event Sent(Signal As VbMsgBoxResult)

Private Sub cmdCancel_Click()
    RaiseEvent Sent(vbCancel)
End Sub

Private Sub cmdOk_Click()
    RaiseEvent Sent(vbOK)
End Sub

Private Sub UserForm_QueryClose( _
    Cancel As Integer, _
    CloseMode As Integer _
)
    Cancel = 1
    cmdCancel_Click
End Sub
