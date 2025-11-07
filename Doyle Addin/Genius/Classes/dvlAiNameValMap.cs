
'''
''' dvlAiNameValMap -- functions to streamline
'''     translation of data from Dictionary
'''     Objects to Inventor NameValueMap
'''     Objects, and vice-versa.
'''
''' NOTE: these functions MIGHT be supplanted by
'''     their addition to or implementation in
'''     Class Module dcPopulator
'''

Public Function dc2aiNameValMap(dc As Scripting.Dictionary, _
    Optional mp As Inventor.NameValueMap = Nothing _
) As Inventor.NameValueMap
    Dim rt As Inventor.NameValueMap
    Dim ky As Variant
    Dim it As Variant
    Dim nm As String
    Dim ck As VbMsgBoxResult
    
    If mp Is Nothing Then
        With ThisApplication.TransientObjects
        Set rt = dc2aiNameValMap( _
            dc, .CreateNameValueMap _
        )   'NameValueMap cannot
            'be created with New
        End With
    Else
        Set rt = mp
        With dc: For Each ky In .Keys
            nm = CStr(ky)
            it = Array(.Item(ky))
            
            If IsObject(it(0)) Then
                ''' Object handling not
                ''' implemented as yet.
                ''' A general solution is
                ''' not likely possible.
                
                ''' UPDATE[2022.07.05.1319]
                ''' it appears that a NameValueMap
                ''' CAN include another NameValueMap
                ''' as a Value, thus enabling multi-
                ''' level NameValueMaps. Whether
                ''' other Objects can be so contained
                ''' is likely a question best left
                ''' unexplored for now, but it seems
                ''' at least sub-Dictionaries and
                ''' NameValueMaps can be processed.
                If TypeOf it(0) Is Scripting.Dictionary Then
                    rt.Add nm, dc2aiNameValMap(obOf(it(0)))
                ElseIf TypeOf it(0) Is Inventor.NameValueMap Then
                    rt.Add nm, it(0)
                Else
                End If
            Else
                On Error Resume Next
                
                Err.Clear
                rt.Add nm, it(0)
                
                If Err.Number Then
                    ck = MsgBox(Join(Array( _
                        "Key """ & nm, _
                        """ Value (" & CStr(it(0)) & ")", _
                        "could not be set.", _
                        "The Key will not", _
                        "be assigned.", _
                        "", _
                        "Click OK to continue", _
                        "(Cancel to debug)" _
                    ), vbNewLine), _
                        vbOKCancel + vbExclamation, _
                        "Assignment Error!" _
                    )
                    If ck = vbCancel Then
                        Stop 'to Debug
                    End If
                End If
                
                On Error GoTo 0
            End If
        Next: End With
    End If
    
    Set dc2aiNameValMap = rt
End Function

Public Function dcFromAiNameValMap( _
    mp As Inventor.NameValueMap, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim nm As String
    Dim mx As Long
    Dim dx As Long
    
    If dc Is Nothing Then
        Set rt = dcFromAiNameValMap( _
            mp, New Scripting.Dictionary _
        )
    Else
        Set rt = dc 'mp
        With mp 'dc
            mx = .Count
            For dx = 1 To mx
                nm = .Name(dx)
                
                rt.Add nm, itemForDcOr(.Item(nm))
            Next
        End With
    End If
    
    Set dcFromAiNameValMap = rt
End Function

Public Function itemForDcOr(it As Variant, _
    Optional tp As Long = 0 _
) As Variant
    '''
    ''' itemForDcOr -- given item it, return
    '''     transformation according to type,
    '''     and type of result desired for
    '''     Dictionary and NameValueMap Objects,
    '''     according to value of tp:
    '''         NameValueMap for tp = 1
    '''         Dictionary for any other
    '''         value, including default 0
    '''     all other types of item are returned
    '''     as is, including Objects other than
    '''     Dictionary and NameValueMap
    '''
    Dim rt As Variant
    Dim ck As Variant
    Dim mx As Long
    Dim dx As Long
    Dim dc As Scripting.Dictionary
    
    If IsArray(it) Then
        mx = UBound(it)
        If mx < LBound(it) Then
            rt = Array()
        Else
            ReDim rt(mx)
            
            For dx = 0 To mx
                ck = Array(itemForDcOr(it(dx), tp))
                If IsObject(ck(0)) Then
                    Set rt(dx) = ck(0)
                Else
                    rt(dx) = ck(0)
                End If
            Next
        End If
    ElseIf IsObject(it) Then
        If TypeOf it Is Inventor.NameValueMap Then
            Set rt = dcFromAiNameValMap(obOf(it))
        ElseIf TypeOf it Is Scripting.Dictionary Then
            Set dc = it
            With New dcPopulator
                For Each ck In dc.Keys
                .Setting ck, itemForDcOr(dc.Item(ck), tp)
                Next
                
                If tp = 1 Then 'NameValMap wanted
                    Set rt = .NameValMap()
                Else 'assume Dictionary
                    Set rt = .Dictionary()
                End If
            End With
        Else
            Set rt = it
        End If
    Else
        rt = it
    End If
    
    If IsObject(rt) Then
        Set itemForDcOr = rt
    Else
        itemForDcOr = rt
    End If
End Function

Public Function nuAiNameValMap( _
    Optional Using As Inventor.Document = Nothing _
) As Inventor.NameValueMap
    If Using Is Nothing Then
        With ThisApplication.TransientObjects
        Set nuAiNameValMap = .CreateNameValueMap
        End With
    Else
        Set nuAiNameValMap = dc2aiNameValMap(Using)
    End If
End Function
'''
'''
'''
