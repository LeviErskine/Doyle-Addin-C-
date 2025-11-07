

Public Function Repeat( _
    Count As Long, Text As String _
) As String
    Repeat = Replace(Space$(Count), " ", Text)
End Function

Public Function txBlk( _
    Lines As Long, Chars As Long, _
    Optional Use As String = "+" _
) As String
    txBlk = Mid$(Repeat( _
        Lines, vbNewLine _
        & String$(Chars, "+") _
    ), 1 + Len(vbNewLine))
End Function

Public Sub MakeActivePurchased()
    Dim md As Inventor.Document
    Dim ck As VbMsgBoxResult
    
    Set md = ThisApplication.ActiveDocument
    If md Is ThisDocument Then
        ck = vbNo
    Else
        ck = mkAiDocPurchased(md)
    End If
        
    If ck = vbOK Then
        ck = MsgBox(Join(Array( _
            "Model BOM Structure", _
            "now Purchased." _
        ), vbNewLine), _
            vbOKOnly + vbInformation, _
            "Success!" _
        )
    ElseIf ck = vbNo Then
        ck = MsgBox(Join(Array( _
            "Document is not", _
            "a valid Model.", _
            "", _
            "Please select a", _
            "Part or Assembly." _
        ), vbNewLine), _
            vbOKOnly + vbExclamation, _
            "No Model" _
        )
    ElseIf ck = vbAbort Then
        ck = MsgBox(Join(Array( _
            "Failed to update", _
            "model's BOM Structure!", _
            "", _
            "Check for locks", _
            "or other issues." _
        ), vbNewLine), _
            vbOKOnly + vbCritical, _
            "Change Failed!" _
        )
    Else
        ck = MsgBox(Join(Array( _
            "Change Operation returned", _
            "unexpected result code.", _
            "", _
            "Please review model status." _
        ), vbNewLine), _
            vbOKOnly + vbQuestion, _
            "Result Unknown" _
        )
    End If
End Sub

Public Function mkAiDocPurchased( _
    AiDoc As Inventor.Document _
) As VbMsgBoxResult
    Dim ck As VbMsgBoxResult
    
    If TypeOf AiDoc Is Inventor.PartDocument Then
        ck = mkAiPartPurchased(AiDoc)
    ElseIf TypeOf AiDoc Is Inventor.AssemblyDocument Then
        ck = mkAiAssyPurchased(AiDoc)
    Else
        ck = vbNo
    End If
    
    mkAiDocPurchased = ck
End Function

Public Function mkAiPartPurchased( _
    AiDoc As Inventor.PartDocument _
) As VbMsgBoxResult
    If AiDoc Is Nothing Then
        mkAiPartPurchased = vbNo
    Else
    With AiDoc.ComponentDefinition
        On Error Resume Next
        Err.Clear
        .BOMStructure = kPurchasedBOMStructure
        If Err.Number = 0 Then
            mkAiPartPurchased = vbOK
        Else
            mkAiPartPurchased = vbAbort
        End If
        On Error GoTo 0
    End With: End If
End Function

Public Function mkAiAssyPurchased( _
    AiDoc As Inventor.AssemblyDocument _
) As VbMsgBoxResult
    If AiDoc Is Nothing Then
        mkAiAssyPurchased = vbNo
    Else
    With AiDoc.ComponentDefinition
        On Error Resume Next
        Err.Clear
        .BOMStructure = kPurchasedBOMStructure
        If Err.Number = 0 Then
            mkAiAssyPurchased = vbOK
        Else
            mkAiAssyPurchased = vbAbort
        End If
        On Error GoTo 0
    End With: End If
End Function

Public Function dcTemplate0A( _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If dc Is Nothing Then
        Set rt = dcTemplate0A( _
            New Scripting.Dictionary _
        )
    Else
        Set rt = dc
    End If
    
    Set dcTemplate0A = rt
End Function

Public Function send2clipBd_OBSOLETE(src As Variant) As Variant
    With New MSForms.DataObject
        .SetText src
        .PutInClipboard
    End With
    send2clipBd_OBSOLETE = src
End Function

