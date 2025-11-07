

Private dcGrpIn As Scripting.Dictionary
Private dcGrpOut As Scripting.Dictionary

Private Sub Class_Initialize()
    Set dcGrpIn = New Scripting.Dictionary
    Set dcGrpOut = New Scripting.Dictionary
End Sub

Public Function Itself() As kyPick
    Set Itself = Me
End Function

Public Function WithInDc( _
    Dict As Scripting.Dictionary _
) As kyPick
    Set dcGrpIn = dcNewIfNone(Dict)
    Set WithInDc = Me
End Function

Public Function WithOutDc( _
    Dict As Scripting.Dictionary _
) As kyPick
    Set dcGrpOut = dcNewIfNone(Dict)
    Set WithOutDc = Me
End Function

Public Function AfterScanning( _
    dSrc As Scripting.Dictionary _
) As kyPick
    Dim ky As Variant
    
    With dSrc: For Each ky In .Keys
        With dcFor(.Item(ky))
        If .Exists(ky) Then
            Stop
        Else
            .Add ky, dSrc.Item(ky)
        End If
        End With
    Next: End With
    Set AfterScanning = Me
End Function

Public Function dcIn() As Scripting.Dictionary
    Set dcIn = dcGrpIn
End Function

Public Function dcOut() As Scripting.Dictionary
    Set dcOut = dcGrpOut
End Function

Public Function dcFor( _
    Item As Variant _
) As Scripting.Dictionary
    If IsObject(Item) Then
        Set dcFor = dcGrpIn
    Else
        Set dcFor = dcGrpOut
    End If
End Function

