

Private dc As Scripting.Dictionary

Private Sub Class_Initialize()
    Set dc = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    dc.RemoveAll
    Set dc = Nothing
End Sub

Public Function PrepFrom(AiDoc As Inventor.Document) As aiPropLink
    Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    Dim psName As String
    Dim prName As String
    
    With AiDoc
        For Each ps In .PropertySets
            psName = ps.Name
            For Each pr In ps
                prName = pr.Name
                If dc.Exists(prName) Then
                    Stop
                Else
                    dc.Add prName, psName
                End If
            Next
        Next
    End With
    
    Set PrepFrom = Me
End Function
