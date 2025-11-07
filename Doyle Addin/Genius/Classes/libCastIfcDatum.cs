

Private Const txVersion As String = "module libCastIfcDatum REV[2022.03.18.1136]"
'''
'''
'''

Public Function obIfcDatum(ob As Object) As ifcDatum
    If TypeOf ob Is Inventor.Property Then
        With New ifcAiProperty
        Set obIfcDatum = .Connect(obAiProp(ob))
        End With
    Else
        With New ifcDatum
        Set obIfcDatum = .Connect(ob)
        End With
    End If
End Function

'''
''' END of Module libCastIfcDatum
'''
Public Function libCastIfcDatum() As String
    libCastIfcDatum = txVersion
End Function
