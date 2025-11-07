

Public Function dImG1f1iPart(md As Inventor.PartDocument) As Inventor.PartDocument
    If md Is Nothing Then
        Set dImG1f1iPart = Nothing
    Else
    With md.ComponentDefinition
    If .IsiPartFactory Then
        If .iPartFactory Is Nothing Then
            Stop
            Set dImG1f1iPart = Nothing
        Else
            Set dImG1f1iPart = aiDocPart( _
            .iPartFactory.Parent)
        End If
    ElseIf .IsiPartMember Then
        If .iPartMember Is Nothing Then
            Stop
            Set dImG1f1iPart = Nothing
        Else
            Set dImG1f1iPart = dImG1f1iPart( _
            .iPartMember.ParentFactory.Parent)
        End If
    'ElseIf .IsContentMember Then
    'ElseIf .IsModelStateFactory Then
    'ElseIf .IsModelStateMember Then
    Else
        Set dImG1f1iPart = Nothing
    End If: End With: End If
    'dImG0f1 = "REV[2023.01.19.1046]"
End Function

Public Function dImG1f2iPart(md As Inventor.Document) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim pt As Inventor.PartDocument
    Dim ck As Inventor.PartDocument
    
    Set rt = New Scripting.Dictionary
    
    With dcAiPartDocs(dcAiDocComponents(md))
        For Each ky In .Keys
            Set pt = aiDocPart(.Item(ky))
            If pt Is Nothing Then
            Else
                Set ck = dImG1f1iPart(pt)
                If ck Is Nothing Then
                Else
                    'ck.File.FullFileName
                    'Stop
                    Debug.Print ; 'Breakpoint Landing
                End If
            End If
        Next
    End With
    
    Set dImG1f2iPart = rt
'Debug.Print ConvertToJson(dImG1f2iPart(aiDocActive()), vbTab)
End Function

Public Function dImG0f0() As String
    dImG0f0 = "REV[2023.01.19.1046]"
End Function
