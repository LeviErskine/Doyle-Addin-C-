

Public Function dcAddIns4inventor( _
    Optional app As Inventor.Application = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim adIn As Inventor.ApplicationAddIn
    
    If app Is Nothing Then
        Set dcAddIns4inventor = dcAddIns4inventor(ThisApplication)
    Else
        Set rt = New Scripting.Dictionary
        For Each adIn In app.ApplicationAddIns
            rt.Add adIn.ClassIdString, adIn
            DoEvents
        Next
        Set dcAddIns4inventor = rt
    End If
End Function

Public Function addIn4inventor( _
    clsId As String _
) As Inventor.ApplicationAddIn
    With dcAddIns4inventor()
        If .Exists(clsId) Then
            Set addIn4inventor = .Item(clsId)
        Else
            Set addIn4inventor = Nothing
        End If
    End With
End Function

Public Function addInILogic() As Inventor.ApplicationAddIn
    Set addInILogic = addIn4inventor( _
        "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}" _
    )
End Function

