

Public Function bomViewStruct( _
    pd As Inventor.AssemblyDocument _
) As Inventor.BOMView
    '''
    ''' bomViewStruct -- Get Structured BOM View
    '''     for supplied Assembly, if available
    '''
    Dim bv As Inventor.BOMView
    'Dim br As Inventor.BOMRow
    
    If pd Is Nothing Then
        Set bv = Nothing
    Else
        With pd 'aiDocAssy(aiDocActive())
            On Error Resume Next
            Err.Clear
            Set bv = .ComponentDefinition.BOM.BOMViews.Item("Structured")
            
            If Err.Number = 0 Then 'we're okay
            Else 'we got nothin'
                Set bv = Nothing
            End If
            
            On Error GoTo 0
            'Stop
        End With
    End If
    Set bomViewStruct = bv
End Function

Public Function dBVg1f1(itmPath As String) As String()
    Dim rt(1) As String
    Dim bk As Long
    
    bk = InStrRev(itmPath, ".")
    
    If bk > 0 Then
        rt(0) = Left$(itmPath, bk - 1)
        rt(1) = Mid$(itmPath, bk + 1)
    Else
        rt(0) = ""
        rt(1) = itmPath
    End If
    
    dBVg1f1 = rt
End Function

Public Function bomLnumBkDn( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim ls() As String
    
    With dc
        If .Exists("path") Then
            ls = dBVg1f1(dc.Item("path"))
            .Item("base") = ls(0)
            .Item("seq") = ls(1)
        End If
    End With
    Set bomLnumBkDn = dc
End Function

Public Function bomLineInfo( _
    brw As Inventor.BOMRow _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    With brw
        rt.Add "bomStruct", .BOMStructure
        rt.Add "path", .ItemNumber
        'rt.Add "seq", .ItemNumber
        
        rt.Add "qty", .ItemQuantity
        rt.Add "qtTotal", .TotalQuantity
        rt.Add "qtUnit", "EA"
        
        rt.Add "mrg", .Merged
        rt.Add "pro", .Promoted
        rt.Add "rol", .RolledUp
        
        '.ChildRows
    End With
    Set bomLineInfo = bomLnumBkDn(rt)
End Function

Public Function dBVg1f2( _
    AiDoc As Inventor.Document, _
    wk As Scripting.Dictionary, _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim bs As String
    Dim pn As String
    Dim rm As String
    Dim k0 As String
    Dim k1 As Variant
    
    With wk
        bs = .Item("path")
        pn = .Item("ptNum")
    End With
    
    With dcOfPropsInAiDoc(AiDoc)
    If .Exists("RM") Then
        rm = aiProperty(.Item("RM")).Value
        k0 = pn & "|" & rm
        
        Set rt = New Scripting.Dictionary
        rt.Add "bomStruct", kPurchasedBOMStructure
        
        rt.Add "path", bs & ".1"
        
        If .Exists("RMQTY") Then
            rt.Add "qty", aiProperty(.Item("RMQTY")).Value
        Else
            rt.Add "qty", -1
        End If
        rt.Add "qtTotal", rt.Item("qty")
        
        If .Exists("RMUNIT") Then
            rt.Add "qtUnit", aiProperty(.Item("RMUNIT")).Value
        Else
            rt.Add "qtUnit", "EA"
        End If
        
        rt.Add "mrg", False
        rt.Add "pro", False
        rt.Add "rol", False
        
        rt.Add "base", bs
        rt.Add "seq", "1"
        
        rt.Add "ptNum", rm
        'rt.Add "aiDoc", ""
        
        If dc.Exists(k0) Then
            Set ck = dcOb(dc.Item(k0))
            'send2clipBd ConvertToJson( _
                dcWBQbyCmpResult( _
                dcCmpTextOf2dc( _
                ck, rt _
            )), vbTab)
            With dcWBQbyCmpResult( _
                dcCmpTextOf2dc( _
                ck, rt _
            ))
                If .Exists("!=") Then
                    With dcOb(.Item("!="))
                        For Each k1 In .Keys
                            ck.Item(k1) = ck.Item(k1) _
                                & vbTab & rt.Item(k1)
                            'Debug.Print Join(Array( _
                                k0, k1, ck.Item(k1) _
                            ), "|")
                            'Stop
                        Next
                    End With
                End If
            End With
            Debug.Print ; 'Breakpoint Landing
            'Stop
        Else
            dc.Add pn & "|" & rm, rt
        End If
    Else
    End If
    End With
    
    'Stop
    Set dBVg1f2 = dc
End Function

Public Function bomItemInfo( _
    rw As Inventor.BOMRow, _
    Optional dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    'Dim wk As Scripting.Dictionary
    Dim df As Inventor.ComponentDefinition
    Dim pt As Inventor.Document
    Dim pn As String
    Dim ck As String
    Dim fn As String
    Dim ct As Long
    
    If dc Is Nothing Then
        Set rt = bomLineInfo(rw)
    Else
        Set rt = dc
    End If
    
    With rw
        With .ComponentDefinitions
            ct = .Count
            If ct > 0 Then
                With .Item(1)
                    Set pt = aiDocument(.Document)
                    pn = aiDocPartNum(pt)
                    fn = pt.FullDocumentName
                End With
                
                'With dc
                '    If .Exists(pn) Then
                '        Stop
                '    Else
                '        .Add pn, New Scripting.Dictionary
                '    End If
                '
                '    'Set wk = dcOb(.Item(pn))
                'End With
            Else
                Stop
                pn = ""
            End If
            
            rt.Add "ptNum", pn
        End With
        
        If ct > 1 Then
            Stop
            fn = ""
            For Each df In .ComponentDefinitions
                Set pt = aiDocument(.Document)
                ck = aiDocPartNum(pt)
                If ck = pn Then
                    fn = fn & vbNewLine & pt.FullDocumentName
                    'With wk
                    '    If .Exists(fn) Then
                    '        Stop
                    '    Else
                    '        .Add fn, pt
                    '    End If
                    'End With
                Else
                    Stop
                End If
            Next
        End If
        
        rt.Add "aiDoc", fn
    End With
    
    Set bomItemInfo = rt
End Function

Public Function dBVg7f4( _
    aiAssy As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' not sure what doing with this one
    ''' further developemt on hold
    '''
    'Dim pd As Inventor.AssemblyDocument
    Dim rt As Scripting.Dictionary
    
    Dim bv As Inventor.BOMView
    Dim br As Inventor.BOMRow
    
    If dc Is Nothing Then
        Set rt = dBVg7f4(aiAssy, New Scripting.Dictionary)
    ElseIf aiAssy Is Nothing Then
        Set rt = dc
    Else
        With aiAssy
            Set bv = .ComponentDefinition.BOM.BOMViews.Item("Structured")
            With bv
                Stop
            End With
        End With
    End If
    Set dBVg7f4 = rt
End Function

Public Function bomInfoBkDn( _
    rwSet As Inventor.BOMRowsEnumerator, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional fn As String = "" _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim rw As Inventor.BOMRow
    Dim pn As String
    Dim ck As String
    
    If dc Is Nothing Then
        Set rt = bomInfoBkDn(rwSet, New Scripting.Dictionary, fn)
    Else
        Set rt = dc
        If rwSet Is Nothing Then 'likely a Part
            'might need/want to look for raw material here
        Else 'it's an Assembly
            For Each rw In rwSet
                'Stop
                DoEvents
                Set wk = bomItemInfo(rw)
                wk.Add "ptOf", fn
                pn = CStr(wk.Item("ptNum"))
                ck = fn & "|" & pn
                With rt
                    If .Exists(ck) Then
                        'Stop
                        'debug.Print ConvertToJson(dcCmpTextOf2dc(wk,dcOb(.Item(ck))),vbTab)
                        'debug.Print ConvertToJson(
                        With dcWBQbyCmpResult(dcCmpTextOf2dc(wk, dcOb(.Item(ck)))) ',vbTab)
                            With dcOb(.Item("!="))
                                .Remove "path"
                                .Remove "base"
                                If .Count > 0 Then
                                    Debug.Print "MISMATCH: "; txDumpLs(.Keys, ", ")
                                    Stop
                                End If
                            End With
                        End With
                    Else
                        .Add ck, wk
                    End If
                End With
                
                With rw
                    If .ChildRows Is Nothing Then
                        Set dc = dBVg1f2( _
                            ThisApplication.Documents.ItemByName( _
                                wk.Item("aiDoc") _
                            ), wk, dc)
                    Else
                        DoEvents
                        Set dc = bomInfoBkDn( _
                            .ChildRows, dc, pn _
                        )
                    End If
                End With
                DoEvents
            Next
        End If
    End If
    Set bomInfoBkDn = rt
'Debug.Print txDumpLs(bomInfoBkDn(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows).Keys)
'send2clipBd ConvertToJson(bomInfoBkDn(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows), vbTab)
End Function

Public Function dcOfBomsFromAiStructured( _
    rwSet As Inventor.BOMRowsEnumerator, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional fn As String = "" _
) As Scripting.Dictionary
    '''
    ''' dcOfBomsFromAiStructured --
    '''     generate Dictionary of BOMs:
    '''     one for each distinct Assembly in
    '''     supplied Inventor BOM (structured)
    '''
    '''     returned as Dictionary of Assembly
    '''     sub Dictionaries, each keyed to its
    '''     Part Number and containing a set of
    '''     Item sub Dictionaries, again keyed
    '''     to Item P/N. Each Item sub Dictionary
    '''     represents a BOM line item
    '''
    Dim rt As Scripting.Dictionary
    
    Set dcOfBomsFromAiStructured = _
        dBV0g0f4(dBV0g0f3(dBV0g0f1( _
        rwSet, dc, fn _
    )))
    Debug.Print ; 'Breakpoint Landing
'Debug.Print ConvertToJson(dBV0g0f2(dcOb(dcOb(rt.Item("19-240-79925")).Item("19-240-90004"))), vbTab)
'Debug.Print dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows).Count
End Function

Public Function dBV0g0f1( _
    rwSet As Inventor.BOMRowsEnumerator, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional fn As String = "" _
) As Scripting.Dictionary
    '''
    ''' dBV0g0f1 -- retrieve BOM data
    '''     from a BOMRowsEnumerator
    '''     and its child row enumerators
    '''
    Dim rt As Scripting.Dictionary
    Dim pd As Scripting.Dictionary
    Dim it As Scripting.Dictionary
    Dim dt As Scripting.Dictionary
    'Dim fd As Scripting.Dictionary
    '''
    'Dim kyIt As Variant
    '''
    Dim rw As Inventor.BOMRow
    Dim pn As String
    Dim th As String
    
    If dc Is Nothing Then
        Set rt = dBV0g0f1(rwSet, _
        New Scripting.Dictionary _
        , fn)
    Else
        Set rt = dc
        If rwSet Is Nothing Then
            'Stop
            Debug.Print ; 'Breakpoint Landing
        Else
        
        With rt
            If Not .Exists(fn) Then
                .Add fn, New Scripting.Dictionary
            End If
            Set pd = .Item(fn)
        End With
        
        For Each rw In rwSet
            DoEvents
            
            Set dt = bomItemInfo(rw)
            With dt
                pn = CStr(.Item("ptNum"))
                th = CStr(.Item("path"))
            End With
            
            With pd
                If Not .Exists(pn) Then
                    .Add pn, New Scripting.Dictionary
                End If
                Set it = .Item(pn)
            End With
            
            With it
                If .Exists(th) Then
                    Stop
                Else
                    .Add th, dt
                End If
            End With
            
            Set rt = dBV0g0f1(rw.ChildRows, rt, pn)
            DoEvents
        Next
        End If
    End If
    
    'If Len(fn) = 0 Then Stop
    Set dBV0g0f1 = rt
End Function

Public Function dBV0g0f2( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim dcField As Scripting.Dictionary
    Dim k0path As Variant
    Dim k1Field As Variant
    Dim itValue As Variant
    
    Set rt = New Scripting.Dictionary
    With dc: For Each k0path In .Keys
        With dcOb(.Item(k0path))
        For Each k1Field In .Keys
            itValue = .Item(k1Field)
            
            With rt
                If Not .Exists(k1Field) Then
                    .Add k1Field, New Scripting.Dictionary
                End If
                Set dcField = .Item(k1Field)
            End With
            
            With dcField
                If Not .Exists(itValue) Then
                    .Add itValue, New Scripting.Dictionary
                End If
                
                With dcOb(.Item(itValue))
                    If .Exists(k0path) Then
                    Else
                        .Add k0path, 1
                    End If
                End With
            End With
        Next: End With
    Next: End With
    
    Set dBV0g0f2 = rt
End Function

Public Function dBV0g0f3( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dBV0g0f3 -- summarize BOM line item fields
    '''
    Dim rt As Scripting.Dictionary
    Dim dcProd As Scripting.Dictionary
    Dim k0Prod As Variant
    Dim k1Item As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each k0Prod In .Keys
        With rt
            .Add k0Prod, New Scripting.Dictionary
            Set dcProd = .Item(k0Prod)
        End With
        
        With dcOb(.Item(k0Prod))
        For Each k1Item In .Keys
            dcProd.Add k1Item, _
            dBV0g0f2(.Item(k1Item))
        Next: End With
    Next: End With
    
    Set dBV0g0f3 = rt
End Function

Public Function dBV0g0f4( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dBV0g0f4 -- reduce results of dBV0g0f3
    '''     to single values per field
    '''     for each Item under each Product
    '''
    Dim rt As Scripting.Dictionary
    Dim dcProd As Scripting.Dictionary
    Dim dcItem As Scripting.Dictionary
    Dim k0Prod As Variant
    Dim k1Item As Variant
    Dim k2Feld As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each k0Prod In .Keys 'Products
        With rt
            .Add k0Prod, New Scripting.Dictionary
            Set dcProd = .Item(k0Prod)
        End With
        
        With dcOb(.Item(k0Prod)) 'Items
        For Each k1Item In .Keys
            With dcProd
                .Add k1Item, New Scripting.Dictionary
                Set dcItem = .Item(k1Item)
            End With
            
            With dcOb(.Item(k1Item)) 'Fields
                On Error Resume Next
                .Remove "path"
                .Remove "base"
                On Error GoTo 0
                
                For Each k2Feld In .Keys
                    With dcOb(.Item(k2Feld)) 'Value(s)
                    If .Count > 1 Then
                        Stop
                    Else
                        dcItem.Add k2Feld, _
                        .Keys(0) '.Item()
                    End If: End With
                Next
            End With
        Next: End With
    Next: End With
    
    Set dBV0g0f4 = rt
End Function

Public Function dBV0g0f5( _
    dc As Scripting.Dictionary, _
    Optional dlm As String = "|" _
) As Scripting.Dictionary
    '''
    ''' dBV0g0f5 -- reduce results of dBV0g0f3
    '''     to single values per field
    '''     for each Item under each Product
    '''
    Dim rt As Scripting.Dictionary
    Dim dcItem As Scripting.Dictionary
    Dim k0Prod As Variant
    Dim k1Item As Variant
    Dim k2Feld As Variant
    Dim rw As String
    Dim co As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each k0Prod In .Keys 'Products
        If Len(k0Prod) > 0 Then
            With dcOb(.Item(k0Prod)) 'Items
            For Each k1Item In .Keys
                Set dcItem = .Item(k1Item)
                
                With dcItem
                    rw = k0Prod & dlm & k1Item
                    For Each k2Feld In Array( _
                        "seq", "qty", "qtUnit" _
                    )
                        If .Exists(k2Feld) Then
                            co = CStr(.Item(k2Feld))
                        Else
                            co = ""
                        End If
                        
                        rw = rw & dlm & co
                    Next
                    Debug.Print ; 'Breakpoint Landing
                End With
                
                rt.Add rw, dcItem
            Next: End With
        Else 'top-level assembly
            'skip (for now)
        End If
    Next: End With
    
    Set dBV0g0f5 = rt
'send2clipBdWin10 txDumpLs(dBV0g0f5(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows)).Keys)
End Function

Public Function csvOfBomsFromDc( _
    dc As Scripting.Dictionary, _
    Optional dlm As String = "|" _
) As String
    'Product|Item|ItemOrder|QuantityInConversionUnit|ConversionUnit
    'NOTE[2021.08.20]: want to change 'Item'
    '   to 'ItemCode' for compatibility with
    '   current Genius BOM import format.
    '   Will hold off for now.
    csvOfBomsFromDc = Join(Array( _
        "Product", "Item", "ItemOrder", _
        "QuantityInConversionUnit", _
        "ConversionUnit" _
    ), dlm) & vbNewLine & txDumpLs( _
        dBV0g0f5(dc, dlm).Keys _
    )
'send2clipBdWin10 csvOfBomsFromDc(dcOfBomsFromAiStructured(bomViewStruct(aiDocAssy(aiDocActive())).BOMRows))
End Function

Public Function csvOfBomsFromAiStructured( _
    AiDoc As Inventor.Document, _
    Optional dlm As String = "|" _
) As String
    csvOfBomsFromAiStructured = _
        csvOfBomsFromDc( _
        dcOfBomsFromAiStructured( _
        bomViewStruct(aiDocAssy( _
        aiDocActive())).BOMRows _
    ), dlm)
'send2clipBdWin10 csvOfBomsFromAiStructured(aiDocActive())
End Function
            
'''
'''
'''
Private Function dvlBomView() As String
    dvlBomView = "dvlBomView"
End Function

            '        'debug.Print ConvertToJson(dcCmpTextOf2dc(fd,dcOb(.Item(dt))),vbTab)
            '        'debug.Print ConvertToJson(
            '        With dcWBQbyCmpResult(dcCmpTextOf2dc(fd, dcOb(.Item(dt)))) ',vbTab)
            '            With dcOb(.Item("!="))
            '                .Remove "path"
            '                .Remove "base"
            '                If .Count > 0 Then
            '                    Debug.Print "MISMATCH: "; txDumpLs(.Keys, ", ")
            '                    Stop
            '                End If
            '            End With
            '        End With
            '    Else
            '        .Add dt, fd
            
            'With rw
            '    If .ChildRows Is Nothing Then
            '        Set dc = dBVg1f2( _
            '            ThisApplication.Documents.ItemByName( _
            '                fd.Item("aiDoc") _
            '            ), fd, dc)
            '    Else
            '        DoEvents
            '        Set dc = bomInfoBkDn( _
            '            .ChildRows, dc, pn _
            '        )
            '    End If
            'End With


