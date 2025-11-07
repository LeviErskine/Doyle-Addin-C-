

Public Function d0g2f1b( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim xt As Scripting.Dictionary
    Dim k0 As Variant
    Dim i0 As Scripting.Dictionary
    Dim fx As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each k0 In .Keys
        Set i0 = .Item(k0)
        fx = i0.Item("ext")
        With rt
            If Not .Exists(fx) Then 'i0(1)
            .Add fx, New Scripting.Dictionary 'i0(1)
            End If
            
            Set xt = .Item(fx) 'i0(1)
        End With
        
        xt.Add k0, i0 'Array(i0(0), i0(2))
    Next: End With
    
    Set d0g2f1b = rt
'send2clipBdWin10 ConvertToJson(d0g2f1b(d0g1f3()), vbTab)
End Function

Public Function d0g2f1c( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' d0g2f1c --
    '''     derived from d0g2f1b
    '''
    Dim rt As Scripting.Dictionary
    Dim xt As Scripting.Dictionary
    Dim k0 As Variant
    Dim k1 As Variant
    Dim i0 As Scripting.Dictionary
    Dim fx As String
    Dim ds As String
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each k0 In .Keys
        Set i0 = .Item(k0)
        Set xt = New Scripting.Dictionary
        
        For Each k1 In Array( _
            "Part Number", _
            "Description", _
            "ext", "fullname" _
        )
            ds = ""
            If i0.Exists(k1) Then '"Description"
                If IsEmpty(i0.Item(k1)) Then '"Description"
                Else
                    ds = i0.Item(k1) '"Description"
                End If
            End If
            xt.Add k1, ds
        Next
        fx = xt.Item("Part Number") 'i0.Item("ext")
        
        With rt
            If Not .Exists(fx) Then 'i0(1)
            .Add fx, New Scripting.Dictionary 'i0(1)
            End If
            
            'Set xt =
            dcOb(.Item(fx)).Add xt.Item("fullname"), xt 'i0(1)
        End With
        
        'xt.Add k0, i0 'Array(i0(0), i0(2))
    Next: End With
    
    Set d0g2f1c = rt
'send2clipBdWin10 ConvertToJson(d0g2f1c(d0g1f3()), vbTab)
End Function

Public Function m3g0f0()
    Dim ky As Variant
    Dim dt As Inventor.DocumentTypeEnum
    Dim ad As Inventor.Document
    
    With dcAssyCompAndSub(aiDocAssy( _
        ThisApplication.ActiveDocument _
    ).ComponentDefinition.Occurrences)
        For Each ky In .Keys
            Set ad = aiDocument(.Item(ky))
            dt = ad.DocumentType
            If ad.NeedsMigrating Then Debug.Print ky
            'If dt = kAssemblyDocumentObject Then
                'With aiDocAssy(ad)
                    ''is .ComponentDefinition.IsiAssemblyMember
                    
                'End With
            'ElseIf dt = kPartDocumentObject Then
            'Else
            'End If
        Next
    End With
End Function

Public Function m3g0f1migrate(dc As Scripting.Dictionary) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            If aiDocument(.Item(ky)).NeedsMigrating Then
                rt.Add ky, .Item(ky)
            End If
        Next
    End With
    Set m3g0f1migrate = rt
End Function
'Debug.Print Join(m3g0f1migrate(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, vbNewLine)

Public Function m3g0f1factories(dc As Scripting.Dictionary) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set rt = m3g0f3(m3g0f2(.Item(ky)), rt)
        Next
    End With
    Set m3g0f1factories = rt
End Function

Public Function m3g0f3(ad As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If dc Is Nothing Then
        Set m3g0f3 = m3g0f3(ad, New Scripting.Dictionary)
    Else
        Set rt = dc
        If ad Is Nothing Then
        Else
            If rt.Exists(ad.FullFileName) Then
            Else
                rt.Add ad.FullDocumentName, ad
            End If
        End If
    End If
    
    Set m3g0f3 = rt
End Function

Public Function m3g0f2(ad As Inventor.Document) As Inventor.Document
    Dim dt As Inventor.DocumentTypeEnum
    
    dt = ad.DocumentType
    If dt = kAssemblyDocumentObject Then
        Set m3g0f2 = m3g0f2a(ad)
    ElseIf dt = kPartDocumentObject Then
        Set m3g0f2 = m3g0f2p(ad)
    Else
        Set m3g0f2 = Nothing
    End If 'm3g0f2
End Function

Public Function m3g0f2a(ad As Inventor.AssemblyDocument) As Inventor.Document
    If ad.ComponentDefinition.IsiAssemblyFactory Then
        Set m3g0f2a = ad
    ElseIf ad.ComponentDefinition.IsiAssemblyMember Then
        Set m3g0f2a = m3g0f2a(ad.ComponentDefinition.iAssemblyMember.ParentFactory.Parent.Document)
    Else
        Set m3g0f2a = Nothing
    End If
End Function

Public Function m3g0f2p(ad As Inventor.PartDocument) As Inventor.Document
    If ad.ComponentDefinition.IsiPartFactory Then
        Set m3g0f2p = ad
    ElseIf ad.ComponentDefinition.IsiPartMember Then
        Set m3g0f2p = m3g0f2p(ad.ComponentDefinition.iPartMember.ParentFactory.Parent)
    Else
        Set m3g0f2p = Nothing
    End If
End Function

Public Function m3g1f1() As Scripting.Dictionary
    ''
    ''  Test time taken for several operations
    ''  involving collection of Item data from Genius
    ''  and correlation with Inventor Model/Assembly
    ''
    Dim ad As Inventor.Document
    'Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim dcGns As Scripting.Dictionary
    Dim dcInv As Scripting.Dictionary
    Dim tm As Single
    Dim ms As Single
    
    Set ad = ThisApplication.ActiveDocument
    tm = Timer
    Set rs = cnGnsDoyle().Execute("select Item, Family from vgMfiItems")
    ms = Timer - tm
    Debug.Print "Query Genius for Items: " & CStr(ms) & "sec": Stop
    
    tm = Timer
    Set dcGns = dcFrom2Fields(rs, "Item", "Family")
    ms = Timer - tm
    Debug.Print "Generate Dictionary from Result: " & CStr(ms) & "sec": Stop
    
    tm = Timer
    Set dcInv = m3g1f2(ad)
    ms = Timer - tm
    Debug.Print "Generate Dictionary from Assembly: " & CStr(ms) & "sec": Stop
    
    tm = Timer
    With dcKeysInCommon(dcGns, dcInv)
        ms = Timer - tm
        Debug.Print "Join Dictionaries: " & CStr(ms) & "sec": Stop
        
        Stop
    End With
    Debug.Print
End Function

Public Function m3g1f2( _
    ad As Inventor.AssemblyDocument, _
    Optional ct As Long = 0 _
) As Scripting.Dictionary
    '''
    Set m3g1f2 = dcRemapByPtNum( _
        dcAiDocComponents(ad, , ct) _
    )
End Function

Public Function m3g1f3( _
    rs As ADODB.Recordset _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim lsFd As Scripting.Dictionary
    Dim lsNm As Variant
    Dim dt As Variant
    Dim tx As String
    Dim mxCo As Long
    Dim dxCo As Long
    Dim mxRw As Long
    Dim dxRw As Long
    
    Set rt = New Scripting.Dictionary
    With rs
        If .State = adStateClosed Then
        Else
            With .Fields
                Set lsFd = New Scripting.Dictionary
                tx = ""
                mxCo = .Count - 1
                For dxCo = 0 To mxCo
                    tx = tx & vbTab & .Item(dxCo).Name
                    lsFd.Add .Item(dxCo).Name, dxCo
                Next
                lsNm = Split(Mid$(tx, 2), vbTab)
            End With
            
            If .BOF And .EOF Then
                '''
            Else
                'dt = .GetRows
                'dt = split(left$(.GetRows
                'mxRw = UBound(dt, 2)
                'For dxRw = 0 To mxRw
                    'Stop
                'Next
                With m3g1f4(.GetString( _
                    adClipString, , _
                    vbTab, vbVerticalTab _
                ))
                End With
            End If
        End If
    End With
    
'm3g1f3 cnGnsDoyle().Execute((sqlOf_ERC_PTOSIZE))
'm3g1f3 cnGnsDoyle().Execute((sqlOf_03R4LC09_NOCOND))
End Function

Public Function m3g1f4(txData As String) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim dcCo As Scripting.Dictionary
    Dim dcKy As Scripting.Dictionary
    Dim lsDc() As Scripting.Dictionary
    Dim lsDt() As String
    Dim lsRw() As String
    Dim ck As String
    Dim mxRw As Long
    Dim dxRw As Long
    Dim mxCo As Long
    Dim dxCo As Long
    
    Set rt = New Scripting.Dictionary
    
    lsDt = Split(Left$( _
        txData, InStrRev( _
            txData, vbVerticalTab _
        ) - 1), _
    vbVerticalTab)
    
    mxCo = 0
    mxRw = UBound(lsDt)
    For dxRw = 0 To mxRw
        lsRw = Split(lsDt(dxRw), vbTab)
        If mxCo = 0 Then
            mxCo = UBound(lsRw)
            ReDim lsDc(mxCo)
            rt.Add "COLIDX", lsDc
        End If
        
        If mxCo = UBound(lsRw) Then
            For dxCo = 0 To mxCo
                Set dcCo = lsDc(dxCo)
                If dcCo Is Nothing Then
                    Set dcCo = New Scripting.Dictionary
                    Set lsDc(dxCo) = dcCo
                End If
                
                With dcCo
                    ck = lsRw(dxCo)
                    If .Exists(ck) Then
                        Set dcKy = .Item(ck)
                        '.Item(ck) = 1 + .Item(ck)
                    Else
                        Set dcKy = New Scripting.Dictionary
                        .Add ck, dcKy '1
                    End If
                    
                    dcKy.Add dxRw, dxRw
                End With
            Next
        Else 'we got a faulty row
            Stop
        End If
    Next
    
    Set m3g1f4 = rt
End Function

