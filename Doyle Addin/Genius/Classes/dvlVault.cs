
''' '
''' see module mod3 for other functions of possible use here
''' named functions here originated there
''' '

Public Function ArrayFrom(ls As Variant) As Variant
    '''
    ''' ArrayFrom -- return basic Variant Array
    '''     from one of several various types
    '''     of supplied Variant Values
    '''
    Dim dc As Scripting.Dictionary
    
    Set dc = dcOb(obOf(ls))
    If dc Is Nothing Then
        If IsObject(ls) Then
            ArrayFrom = Array()
        ElseIf IsArray(ls) Then
            ArrayFrom = ls
        Else
            ArrayFrom = Array(ls)
        End If
    Else
        ArrayFrom = dc.Keys
    End If
End Function

Public Function dcMapFSysVsVault( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim bp As String
    Dim fp As Variant
    Dim vp As String
    
    Set rt = New Scripting.Dictionary
    
    bp = vaultBasePath()
    If Len(bp) = 0 Then
        Stop 'for debug/devel
        'not sure what to do
        'with no path found
        
        'default here is
        '"C:/Doyle_Vault/"
        'but don't want
        'to assume
    End If
    
    With dc 'dcAiDocComponents(aiDocActive())
    For Each fp In .Keys
        vp = Replace(Replace( _
            fp, bp, "$/" _
            ), "\", "/" _
        )
        With rt
        If .Exists(fp) Or .Exists(vp) Then
            Stop
        Else
            rt.Add fp, vp
            rt.Add vp, fp
        End If: End With
    Next: End With
    
    Set dcMapFSysVsVault = rt
End Function

Public Function dcRemapped2vaultPaths( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim bp As String
    
    Set rt = New Scripting.Dictionary
    
    bp = vaultBasePath()
    If Len(bp) = 0 Then
        Stop 'for debug/devel
        'not sure what to do
        'with no path found
        
        'default here is
        '"C:/Doyle_Vault/"
        'but don't want
        'to assume
    End If
    
    With dc 'dcAiDocComponents(aiDocActive())
    For Each ky In .Keys
    rt.Add Replace(Replace( _
        ky, bp, "$/" _
        ), "\", "/" _
    ), .Item(ky)
    Next: End With
    
    Set dcRemapped2vaultPaths = rt
End Function

Public Function vaultBasePath() As String
    With dcOb(nuILogicIfc().Apply( _
    "vltBasePath", New Scripting.Dictionary _
    ))
    If .Exists("OUT") Then
        vaultBasePath = .Item("OUT")
    Else
        vaultBasePath = ""
    End If: End With
End Function

Public Function vaultPropKeys() As String
    With dcOb(nuILogicIfc().Apply( _
    "vltPropKeys", New Scripting.Dictionary _
    ))
    If .Exists("OUT") Then
        vaultPropKeys = .Item("OUT")
    Else
        vaultPropKeys = ""
    End If: End With
End Function

Public Function dcOfDcByVltPathAndName( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    '''
    '''
    ''' this one should probably call
    ''' dcOfDcByNameAndPath against dc
    '''
    ''' actually, EACH should call
    ''' some common function
    ''' to perform similar task
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim ky As Variant
    Dim ar As Variant
    Dim bk As Long
    Dim bp As String
    Dim fn As String
    
    Set rt = New Scripting.Dictionary
    
    With dcRemapped2vaultPaths(dc)
    For Each ky In .Keys
        ar = Array(.Item(ky))
        
        bk = InStrRev(ky, "/")
        If bk = 0 Then
            Stop
        Else
            fn = Mid$(ky, bk + 1)
            bp = Left$(ky, bk - 1)
            'Stop
        End If
        
        With rt
            If Not .Exists(bp) Then
            .Add bp, New Scripting.Dictionary
            End If
            
            'Set gp =
            With dcOb(.Item(bp))
                .Add fn, ar(0)
            End With
        End With
        
        
    Next: End With
    
    Set dcOfDcByVltPathAndName = rt
End Function

Public Function dcOfDcByNameAndPath( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    '''
    '''
    ''' closely related to
    ''' dcOfDcByVltPathAndName
    ''' (see above)
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim md As Inventor.Document
    Dim ky As Variant
    Dim ar As Variant
    Dim bk As Long
    Dim fn As String
    Dim bp As String
    
    Set rt = New Scripting.Dictionary
    
    With dc 'dcRemapped2vaultPaths(dc)
    For Each ky In .Keys
        'ar = Array(.Item(ky))
        Set md = aiDocument(obOf(.Item(ky)))
        If md Is Nothing Then
            Debug.Print ; 'Breakpoint Landing
        Else
            'Stop
            
            bk = InStrRev(ky, "\")
            If bk = 0 Then
                Stop
            Else
                bp = Left$(ky, bk - 1)
                fn = Mid$(ky, bk + 1)
                'Stop
            End If
            
            With rt
                If Not .Exists(fn) Then
                .Add fn, New Scripting.Dictionary
                End If
                
                'Set gp =
                With dcOb(.Item(fn))
                    .Add bp, md 'ar(0)
                End With
            End With
        End If
    Next: End With
    
    Set dcOfDcByNameAndPath = rt
'send2clipBdWin10 ConvertToJson(dcOfDcByNameAndPath(dcAiDocComponents(aiDocActive())), vbTab)
End Function

Public Function d0g1f4d( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' d0g1f4d - categorize supplied Dictionary
    '''     of Part/Assembly components
    '''     by Vault Property Values
    '''     1 - takes same sort of
    '''         Dictionary as d0g1f4c
    '''     2 - applies d0g1f4c to it
    '''     3 - rekeys the result
    '''     4 - transposes its sub Dictionaries
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dcOfDcRekeyedSecToPri(d0g1f4c(dc))
    For Each ky In .Keys
        rt.Add ky, dcTransGrouped(dcOb(.Item(ky)))
    Next: End With
    
    Set d0g1f4d = rt
End Function

Public Function d0g1f4c( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    'Dim ag As Scripting.Dictionary
    
    Dim ls As Variant
    Dim ky As Variant
    
    Dim sd As Scripting.Dictionary
    'Dim nm As Variant
    Dim pg As Variant
    Dim rw As Variant
    Dim p2 As String
    'Dim fl As Scripting.File
    
    Set rt = New Scripting.Dictionary
    
    ls = dc.Keys
    With nuILogicIfc()
    For Each ky In ls
        'send2clipBdWin10 ConvertToJson(nuILogicIfc()
        Debug.Print ; 'Breakpoint Landing
        'pg =
        With .Apply("dvl0", nuDcPopulator( _
            ).Setting("PropName", "Name" _
            ).Setting("Value", ky _
        ).Dictionary())
            If .Exists("OUT") Then
                pg = .Item("OUT")
            Else
            End If
        End With
            'PropName", "FolderPath
            '   FullPath
            '   FullName
            '"$/Designs/doyle/(72) G3 Conveyor/I Parts/72-XXX-90403 G3 HD 8IN WRAP DRIVE 6IN END ROLLERS CONVEYOR BELT CRESCENT TOP ASSEMBLY"
            ', vbTab)
        '''
        ''' REV[2023.03.03.1140]
        ''' preceding pg assignment
        ''' replaces the one following
        '''
        'pg = .Apply("vlt04", nuDcPopulator( _
            ).Setting("fullname", ky _
            ).Dictionary() _
        ).Item("OUT") '.DataFor(CStr(nm))
        'or "Full Path"
        For Each rw In pg
            If TypeOf rw Is Inventor.NameValueMap Then
                Set sd = dcFromAiNameValMap(obOf(rw))
            ElseIf TypeOf rw Is Scripting.Dictionary Then
                Set sd = rw
            Else
                Stop
                Set sd = Nothing
            End If
            
            If sd Is Nothing Then
            Else
                p2 = sd.Item("fullname") '.LocalForm()  'CStr(rw)
                rt.Add p2, sd 'Array(nm, fnExt(p2), fl)
                'With sd
                '    '.Add "ext", fnExt(p2)
                '    '.Add "fileObj", fileIfPresent(p2)
                'End With
            End If
        Next
    Next: End With
    
    Set d0g1f4c = rt
End Function

Public Function d0g1f4b(ls As Variant) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    Dim nm As Variant
    Dim pg As Variant
    Dim rw As Variant
    Dim p2 As String
    Dim fl As Scripting.File
    
    If IsObject(ls) Then
    ElseIf IsArray(ls) Then
    Else
        Set rt = d0g1f4b(Array(ls))
        ''' NOTE this isn't going to work.
        ''' next Set rt below is going
        ''' to wipe this one right out!
    End If
    
    Set rt = New Scripting.Dictionary
    
    With nuILogicIfc() 'nuIfcVault()
    For Each nm In ArrayFrom(ls)
        pg = .Apply("vlt04", _
            nuDcPopulator().Setting( _
                "PartNumber", nm _
            ).Dictionary() _
        ).Item("OUT") '.DataFor(CStr(nm))
        For Each rw In pg 'Split(pg, vbNewLine)
            'Stop
            If TypeOf rw Is Inventor.NameValueMap Then
                Set sd = dcFromAiNameValMap(obOf(rw))
            ElseIf TypeOf rw Is Scripting.Dictionary Then
                Set sd = rw
            Else
                Stop
                Set sd = Nothing
            End If
            
            If sd Is Nothing Then
            Else
                p2 = sd.Item("fullname") '.LocalForm()  'CStr(rw)
                rt.Add p2, sd 'Array(nm, fnExt(p2), fl)
            With sd
                '.Add "ext", fnExt(p2)
                .Add "fileObj", fileIfPresent(p2)
            End With
            End If
        Next
        With rt
        End With
    Next: End With
    
    Debug.Print ; 'Breakpoint Landing
    If False Then
    'for each k0 in rt.Keys:set sd = rt.Item(k0):debug.Print sd.Item("Folder Path"):next
    End If
    
    Set d0g1f4b = rt
'send2clipBdWin10 ConvertToJson(d0g2f1b(d0g1f4b(Split(nu_FmGetList().AskUser(), vbNewLine))), vbTab)
'send2clipBdWin10 ConvertToJson(d0g2f1b(d0g1f4b(Split(nu_FmGetList().AskUser(Join(dcOb(dcAiDocCompSetsByPtNum(aiDocActive()).Item(1)).Keys, vbNewLine)), vbNewLine))), vbTab)
End Function

Public Function d0g2f1d( _
    dc As Scripting.Dictionary _
) As ADODB.Recordset 'Scripting.Dictionary
    '''
    ''' d0g2f1d --
    '''     derived from d0g2f1b
    '''
    Dim rt As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim xt As Scripting.Dictionary
    Dim ls As Variant
    Dim k0 As Variant
    Dim k1 As Variant
    Dim i0 As Scripting.Dictionary
    Dim fx As String
    Dim ds As String
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    
    ls = Array( _
        "Part Number", _
        "Description" _
    ) _
    ' _
        ,"ext", "fullname" _
    '
    
    Set rs = New ADODB.Recordset
    With rs
        With .Fields: For Each k1 In ls
            .Append k1, adVarChar, 127
        Next: End With
        .Open
    End With
    
    With dc: For Each k0 In .Keys
        Set i0 = .Item(k0)
        
        rs.AddNew
        
        For Each k1 In ls
            With i0
                ds = ""
                If .Exists(k1) Then
                    If IsEmpty(.Item(k1)) Then
                    Else
                        ds = .Item(k1)
                    End If
                End If
            End With
            
            With rs.Fields
                .Item(k1) = ds
            End With
        Next
    Next: End With
    rs.Filter = ""
    
    Set d0g2f1d = rs
'send2clipBdWin10 d0g2f1d(d0g1f4b(Split(nu_FmGetList().AskUser(), vbNewLine))).GetString(adClipString, , vbTab, vbNewLine)
'send2clipBdWin10 rsFiltered(d0g2f1d(d0g1f4b(Split(nu_FmGetList().AskUser(), vbNewLine))), "Description <> ''").GetString(adClipString, , vbTab, vbNewLine)
'send2clipBdWin10 "select i.ItemID Id, ls.pn, ls.ds Description1, i.Description1 oldDescription1 from (values" & vbNewLine & vbTab & "('" & rsFiltered(d0g2f1d(d0g1f4b(Split(nu_FmGetList().AskUser(), vbNewLine))), "Description <> ''").GetString(adClipString, , "', '", "')," & vbNewLine & vbTab & "('") & "', '')" & vbNewLine & ") as ls(pn, ds) inner join vgMfiItems i on ls.pn = i.Item"
End Function

Public Function dVg1f1(argIn As Variant) As Scripting.Dictionary
    Dim dc As New Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    dc.Add "IN", argIn
    With nuILogicIfc()
        Set rt = .Apply("vlt05", dc)
    End With
    Set dVg1f1 = rt
End Function

Public Function dVg2f1(dc As Scripting.Dictionary) As Scripting.Dictionary
    '''
    ''' dVg2f1 - take Dictionary
    '''     of Inventor Documents keyed
    '''     to FullFileName as returned
    '''     by dcAiDocComponents
    ''' return Dictionary of Dictionaries
    '''     of Inventor Documents keyed
    '''     first to File Name only and
    '''     then to ParentFolder Path
    '''
    Dim rt As Scripting.Dictionary
    Dim ls As Scripting.Dictionary
    Dim fl As Scripting.File
    Dim ky As Variant
    Dim nm As String
    Dim fp As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each ky In .Keys
        Set fl = fileIfPresent(CStr(ky))
        If Not fl Is Nothing Then
            With fl
                nm = .Name
                fp = .ParentFolder.Path
            End With
            
            With rt
                If Not .Exists(nm) Then
                    .Add nm, New Scripting.Dictionary
                End If
                Set ls = .Item(nm)
            End With
                
            With ls
            If Not .Exists(fp) Then
                .Add fp, fl
            End If: End With
        End If
    Next: End With
    
    Set dVg2f1 = rt
End Function

Public Function dVg3f1(dc As Scripting.Dictionary) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ob As Variant
    Dim ky As String
    
    Set rt = New Scripting.Dictionary
    With dVg1f1(dVg2f1(dc).Keys)
    If .Exists("OUT") Then
        For Each ob In .Item("OUT")
            Set wk = dcOb(ob)
            
            If wk Is Nothing Then
                Stop
            Else
                With wk
                If .Exists("fullname") Then
                    ky = .Item("fullname")
                    With rt
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add ky, wk
                    End If: End With
                Else
                    Stop
                End If: End With
            End If
        Next
    Else
        Stop
    End If: End With
    Set dVg3f1 = rt
'send2clipBdWin10 ConvertToJson(dVg3f1(dcAiDocComponents(aiDocActive())), vbTab)
End Function

Public Function dVg3f2(dc As Scripting.Dictionary) As Scripting.Dictionary
    '''
    ''' dVg3f2 -    '
    '''     NOTE the following:
    '''     dcMapFSysVsVault maps the full file names
    '''         from the supplied Dictionary's Keys
    '''         to their Vault paths/names,
    '''         and vice-versa
    '''     dVg3f1 returns a Dictionary
    '''         keyed to Vault paths/names
    '''         which must be translated
    '''         to full file names
    '''     dcKeysInCommon will return a Dictionary
    '''         also keyed to Vault paths/names
    '''         containing matching entries
    '''         from the results of each
    '''         of the prior two
    '''     the Dictionary returned is keyed
    '''         to the FIRST value in each
    '''         entry from dcKeysInCommon,
    '''         mapping it to the SECOND value
    '''     in this way, each model's full file path
    '''         is mapped to its Vault data
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dcKeysInCommon( _
        dcMapFSysVsVault(dc), _
        dVg3f1(dc) _
    )
    For Each ky In .Keys
        it = .Item(ky)
        rt.Add it(0), it(1)
    Next: End With
    Set dVg3f2 = rt
End Function

Public Function dVg3f3(dc As Scripting.Dictionary) As Scripting.Dictionary
    '''
    ''' dVg3f3 -    given a Dictionary of Inventor Documents
    '''             returns
    '''
    Dim d2 As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    
    'Set rt =
    'New Scripting.Dictionary
    'Set d2 =
    'With
    Set rt = dcKeysCombined(dc, dVg3f2(dc))
    'For Each ky In .Keys
    '    Stop
    'Next: End With
    
    Set dVg3f3 = rt
End Function

''' END of module dvlVault
'''
'''
Private Function dvlVault() As String

End Function
