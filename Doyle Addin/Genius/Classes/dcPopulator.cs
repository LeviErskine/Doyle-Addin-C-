Private dc As Scripting.Dictionary

Private Function checkDC( _
    Optional dcIn As Scripting.Dictionary = Nothing, _
    Optional Opts As Long = 0 _
) As Scripting.Dictionary
    ''
    ''
    If dcIn Is Nothing Then
        If dc Is Nothing Then
            Set dc = New Scripting.Dictionary
        End If
    Else
        If dc Is Nothing Then
            Set dc = dcIn
        Else
            Stop
            If Opts And 1 Then
                'Planning Key Value Replacement here
            End If
        End If
    End If
    
    Set checkDC = dc
    ''
    ''
End Function

Public Function Using( _
    Dict As Scripting.Dictionary, _
    Optional Opts As Long = 0 _
) As dcPopulator
    ''
    ''
    With checkDC(Dict, Opts)
    End With
    
    Set Using = Me
End Function

Public Function Setting( _
    Key As Variant, Item As Variant _
) As dcPopulator
    With checkDC()
        If .Exists(Key) Then .Remove Key
        .Add Key, Item
    End With
    
    Set Setting = Me
End Function

Public Function Count() As Long
    Count = checkDC().Count
End Function

Public Function Dictionary( _
    Optional Dict As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Set Dictionary = checkDC(Dict)
End Function

Public Function Matching( _
    KeySet As Variant _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If IsArray(KeySet) Then
        Set rt = New Scripting.Dictionary
        
        With checkDC()
        For Each ky In KeySet
        If .Exists(ky) Then
            If rt.Exists(ky) Then
                'don't need to match it twice!
            Else
                rt.Add ky, .Item(ky)
            End If
        End If
        Next: End With
    Else
        Set rt = Matching(Array(KeySet))
    End If
    
    Set Matching = rt
End Function

Public Function Exists(Key As Variant) As Boolean
    With Dictionary()
        Exists = .Exists(Key)
    End With
End Function

Public Function Item(Key As Variant) As Variant
    Dim rt As Variant
    
    With Dictionary()
    If .Exists(Key) Then
        rt = Array(.Item(Key))
        If IsObject(rt(0)) Then
            Set Item = rt(0)
        Else
            Item = rt(0)
        End If
    Else
        Item = Empty
    End If: End With
End Function

'''
''' OPTIONAL section for Inventor.NameValueMap
''' the following functions will ONLY compile
''' within the Autodesk Inventor VBA environment,
''' or other environment which supports the same
''' NameValueMap classes and structures.
'''
''' It should be disabled or deleted for use
''' outside of any such environment.
'''
Public Function UsingNameValMap( _
    NVMap As Inventor.NameValueMap, _
    Optional Opts As Long = 0 _
) As dcPopulator
    ''
    ''
    Set UsingNameValMap = Using( _
    dcFromAiNameValMap(NVMap), Opts)
End Function

Public Function NameValMap( _
    Optional NVMap As Inventor.NameValueMap = Nothing _
) As Inventor.NameValueMap
    Set NameValMap = dc2aiNameValMap(Dictionary(), NVMap)
End Function
'''
''' END of OPTIONAL Inventor.NameValueMap section
'''
'''

