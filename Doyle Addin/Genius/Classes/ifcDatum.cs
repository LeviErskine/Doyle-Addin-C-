

Private obNow As Object
Private valNow As Variant
Private valWas As Variant
'''
'''
'''

Public Function Connect(This As Object) As ifcDatum
    If This Is Nothing Then
        Set Connect = Me
    ElseIf TypeOf This Is ifcDatum Then
        Set Connect = This
    ElseIf False Then
    Else
        Set Connect = obIfcDatum(This)
    End If
End Function

Public Function MakeValue(This As Variant) As ifcDatum
    Set MakeValue = Me
End Function

Public Function Commit() As ifcDatum
    Set Commit = Me
End Function

Public Function Itself() As ifcDatum
    Set Itself = Me
End Function

Public Function Connected( _
    Optional ToThis As Object = Nothing _
) As Boolean
    If ToThis Is Nothing Then
        Connected = Not obNow Is Nothing
    Else
        Connected = obNow Is ToThis
    End If
End Function

Public Function Value() As Variant
    Value = valNow
End Function

Private Sub Class_Initialize()
    valNow = ""
End Sub
