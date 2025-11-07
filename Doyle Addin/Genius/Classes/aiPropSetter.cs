

Private ls() As String

Public Function PropList() As String()
    PropList = ls
End Function

Public Function dcPropsIn(ad As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set ps = ad.PropertySets.Item(gnCustom)
    
    If dc Is Nothing Then
        Set dcPropsIn = dcPropsIn(ad, _
            New Scripting.Dictionary _
        )
    Else
        If IsArray(ls) Then
            For Each ky In ls
                Set pr = aiGetProp(ps, CStr(ky), 1)
                If pr Is Nothing Then
                    'nothing we can do (as yet?)
                Else
                    dc.Add ky, pr
                End If
            Next
            Set dcPropsIn = dc
        ElseIf VarType(ls) = vbString Then
            Stop 'shouldn't wind up here
            Set dcPropsIn = dcPropsIn(ad, dc)
        Else
            Stop 'or here, either
        End If
    End If
End Function

Private Sub Class_Initialize()
    ls = Split("andrew patrick thompson", " ")
End Sub
