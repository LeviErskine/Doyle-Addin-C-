

Public Function ckCtCtr() As Long
    '
End Function

Public Function ccc0g0f0(ck As Inventor.Document) As Inventor.Document
    Dim pt As Inventor.PartDocument
    
    Set pt = aiDocPart(ck)
    If pt Is Nothing Then
        Set ccc0g0f0 = pt
    ElseIf pt.ComponentDefinition.IsContentMember Then
        Set ccc0g0f0 = pt
    ElseIf pt.PropertySets.Count > 4 Then
        Set ccc0g0f0 = pt
    ElseIf 0 Then
    Else
        Set ccc0g0f0 = Nothing
    End If
End Function
