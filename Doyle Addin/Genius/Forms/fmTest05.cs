VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} fmTest05 
   Caption         =   "UserForm1"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10860
   OleObjectBlob   =   "fmTest05.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "fmTest05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event Sent(Signal As VbMsgBoxResult)
Public Event GroupIs(Now As String)
Public Event ItemIs(Now As String)

Private dcHolding As Scripting.Dictionary

Private Const txVersion As String = ""
'''
'''
'''

Public Function Holding( _
    Obj As Object _
) As fmTest05
    '''
    ''' Holding -- Hold onto supplied
    '''     Object until terminated,
    '''     or directed to drop it.
    '''
    '''     not sure about this one.
    '''     purpose is to keep a
    '''     client interface "alive"
    '''     while the form itself
    '''     remains active.
    '''
    With dcHolding
        If .Exists(Obj) Then
        Else
            .Add Obj, .Count
        End If
    End With

    Holding = Me
End Function

Public Function Dropping(
    Obj As Object
) As fmTest05
    With dcHolding
        If .Exists(Obj) Then
            .Remove Obj
        Else
        End If
    End With

    Dropping = Me
End Function

Public Function GroupNow() As String
    Dim tb As MSForms.Tab
    Dim dx As Long

    With tbsItemGrps
        dx = .Value
        tb = .Tabs.Item(dx)
        GroupNow = tb.Name
    End With
End Function

Public Function InGroup(
    GrpId As String
) As fmTest05 'fmIfcTest05A
    Dim tb As MSForms.Tab

    With tbsItemGrps
        On Error Resume Next
        Err.Clear()

        tb = .Tabs.Item(GrpId)
        If Err.Number = 0 Then
            .Value = tb.Index
        Else
            ''' might be a good idea to raise
            ''' a "Fault" event here, to permit
            ''' a client entity to respond
            ''' to an error/fault condition
        End If

        Err.Clear()
        On Error GoTo 0
    End With

    InGroup = Me
End Function

Public Function ItemNow() As String
    'With lbxItems
    ItemNow = lbxItems.Value
    'End With
End Function

Public Function OnItem(
    ItemId As String
) As fmTest05 'fmIfcTest05A
    'Dim tb As MSForms.Tab

    With lbxItems
        On Error Resume Next
        Err.Clear()

        ' tb = .Tabs.Item(ItemId)
        .Value = ItemId
        If Err.Number = 0 Then
            '.Value = tb.Index
            'Stop
            Debug.Print() ; 'Breakpoint Landing
        Else
            Stop
        End If

        Err.Clear()
        On Error GoTo 0
    End With

    OnItem = Me
End Function

Private Sub cmdEndCancel_Click()
    RaiseEvent Sent(vbCancel)
End Sub

Private Sub cmdEndSave_Click()
    RaiseEvent Sent(vbOK)
End Sub

Private Sub cmdOpenItem_Click()
    RaiseEvent Sent(vbRetry)
    'might make a different event
    'to trigger file open, perphaps
    'with active group/item
End Sub

Private Sub lbxItems_Change()
    RaiseEvent ItemIs(lbxItems.Value)
End Sub

Private Sub tbsItemGrps_Change()
    RaiseEvent GroupIs(GroupNow)
End Sub

Private Sub tbsItemGrps_BeforeDropOrPaste(
    ByVal Index As Long,
    ByVal Cancel As MSForms.ReturnBoolean,
    ByVal Action As MSForms.fmAction,
    ByVal Data As MSForms.DataObject,
    ByVal X As Single, ByVal Y As Single,
    ByVal Effect As MSForms.ReturnEffect,
    ByVal Shift As Integer
)
    'will keep this one as is, for now
    'not sure what you can actually drop
    'onto a tab group
    Stop
End Sub

Private Sub lbxItems_MouseMove(
    ByVal Button As Integer,
    ByVal Shift As Integer,
    ByVal X As Single,
    ByVal Y As Single
)
    ''' keeping this one here, since it basically governs
    ''' drag-and-drop behavior from a local control.
    ''' might try to see if this is actually needed.
    ''' one would think this kind of behavior
    ''' would occur automatically.
    Dim dt As MSForms.DataObject
    Dim ef As Integer

    If Button = 1 Then
        dt = New MSForms.DataObject
        dt.SetText lbxItems.Value
        ef = dt.StartDrag()
    End If
End Sub

Private Sub UserForm_Initialize()
    dcHolding = New Scripting.Dictionary
End Sub

Private Sub UserForm_QueryClose(
    Cancel As Integer,
    CloseMode As Integer
)
    Cancel = 1
    RaiseEvent Sent(vbAbort)
    'Me.Hide
End Sub

Private Sub UserForm_Terminate()
    dcHolding.RemoveAll
    dcHolding = Nothing
End Sub
