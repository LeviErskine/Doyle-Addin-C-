

'Private dcGrpIn As Scripting.Dictionary
'Private dcGrpOut As Scripting.Dictionary
Private dcPicker As kyPick

''' Sample Usage:
'Debug.Print txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).OutGroup().Keys)
'Debug.Print txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).WithSel(New kyPickAiDocWithRM, 1).SubScanning().OutGroup().Keys)

Private Sub Class_Initialize()
    'Set dcGrpIn = New Scripting.Dictionary
    'Set dcGrpOut = New Scripting.Dictionary
    Set dcPicker = New kyPick
End Sub

Public Function WithInDc( _
    Dict As Scripting.Dictionary _
) As dcSplitter
    Set dcPicker = dcPicker.WithInDc(Dict)
    Set WithInDc = Me
End Function

Public Function WithOutDc( _
    Dict As Scripting.Dictionary _
) As dcSplitter
    Set dcPicker = dcPicker.WithOutDc(Dict)
    Set WithOutDc = Me
End Function

Public Function WithSel(Selector As kyPick, _
    Optional KeepData As Long = 0 _
) As dcSplitter
    If KeepData = 0 Then
        Set dcPicker = Selector
    Else
        Set dcPicker = Selector.WithInDc( _
            dcPicker.dcIn).WithOutDc( _
            dcPicker.dcOut _
        )
    End If
    Set WithSel = Me
End Function

Public Function InGroup() As Scripting.Dictionary
    Set InGroup = dcPicker.dcIn
End Function

Public Function OutGroup() As Scripting.Dictionary
    Set OutGroup = dcPicker.dcOut
End Function

Public Function Scanning(SrcDict As Scripting.Dictionary) As dcSplitter
    Dim ky As Variant
    
    With SrcDict
        For Each ky In .Keys
            With dcPicker.dcFor(.Item(ky))
                If .Exists(ky) Then
                    Stop
                Else
                    .Add ky, SrcDict.Item(ky)
                End If
            End With
        Next
    End With
    Set Scanning = Me
End Function

Public Function SubScanning(Optional WantOut As Long = 0) As dcSplitter
    Dim dcSub As Scripting.Dictionary
    
    If WantOut = 0 Then
        Set dcSub = dcPicker.dcIn
    Else
        Set dcSub = dcPicker.dcOut
    End If
    Set dcPicker = dcPicker.WithInDc(New Scripting.Dictionary).WithOutDc(New Scripting.Dictionary)
    
    Set SubScanning = Scanning(dcSub)
End Function

