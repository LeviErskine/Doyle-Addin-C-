

Public Function sqlTextInVBcode(nm As String) As String
    Dim ar As Variant
    
    ar = Split(nm, "'''SQL'''")
    If UBound(ar) < 1 Then
        sqlTextInVBcode = ""
    Else
        'sqlTextInVBcode = ar(1)
        sqlTextInVBcode = Join(Split( _
            ar(1), vbNewLine & "'" _
        ), vbNewLine)
    End If
End Function

Public Function sqlTextInDict( _
    nm As String, dc As Scripting.Dictionary _
) As String
    sqlTextInDict = sqlTextInVBcode( _
        vbTextOfProcInDict(nm, dc) _
    )
End Function

Public Function sqlTextInProject( _
    nm As String, pj As VBIDE.VBProject _
) As String
    sqlTextInProject = sqlTextInDict( _
        nm, dcOfVbProcsFlat(pj) _
    )
End Function

