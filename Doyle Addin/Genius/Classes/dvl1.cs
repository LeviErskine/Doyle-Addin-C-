
'
'''
''' Development Module dvl1 -- (generic until renamed)
''' begun 2019.08.21
''' by Andrew Thompson ()
'''
''' Initial Purpose: Begin design of new Genius Properties Generator/Populator
'''

Public Function d1g0f0() As Variant
    d1g0f0 = 0
End Function

Public Function d1g4f0( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dcRemapByPtNum( _
        dcAiDocComponents(AiDoc) _
    )
        For Each ky In .Keys
            'rt.Add ky, dcProps4genius( _
                aiDocument(.Item(ky)), , 0 _
            )
            rt.Add ky, dcAiPropValsFromDc( _
                dcOfPropsInAiDoc( _
                    aiDocument(.Item(ky)) _
                ) _
            )
        Next
    End With
    
    Set d1g4f0 = rt
'Debug.Print dumpLsKeyVal(d1g4f1(d1g4f0(ThisApplication.ActiveDocument)), " - ")
'send2clipBd ConvertToJson(d1g4f0(ThisApplication.ActiveDocument), vbTab)
End Function

Public Function d1g4f1( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            rt.Add ky, dcOb(.Item(ky)).Count
        Next
    End With
    Set d1g4f1 = rt
End Function

Public Function d1g4f2(md As String, pr As String) As String
    Dim vbc As VBIDE.VBComponent
    Dim rt As Variant
    Dim ls As Variant
    Dim mx As Long
    Dim dx As Long
    Dim ck As String
    
    Set vbc = ThisDocument.VBAProject.InventorVBAComponents.Item(md).VBComponent
    With vbc.CodeModule
        ls = Split(.Lines( _
            .ProcBodyLine(pr, vbext_pk_Proc), _
            .ProcCountLines(pr, vbext_pk_Proc) _
        ), vbNewLine)
        mx = UBound(ls)
        For dx = LBound(ls) To mx
            ck = Trim$(ls(dx))
            If Left$(ck, 1) = "'" Then
                rt = rt & Mid$(ck, 2) & vbNewLine
            End If
        Next
    End With
    d1g4f2 = rt

'Debug.Print d1g4f2("dvl1", "d1g4f2")
'Debug.Print d1g4f2("zzCsv000", "zc0g0f1")
'Debug.Print d1g4f2("zzCsv000", "zc0g0f2")
End Function

Public Function d1g4f3( _
    hdr As String, dlm As String _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ls As Variant
    Dim mx As Long
    Dim dx As Long
    
    Set rt = New Scripting.Dictionary
    ls = Split(hdr, dlm)
    mx = UBound(ls)
    For dx = LBound(ls) To mx
        rt.Add dx, ls(dx)
        rt.Add ls(dx), dx
    Next
    rt.Add "", dlm
    Set d1g4f3 = rt
End Function

Public Function d1g4f4( _
    dc As Scripting.Dictionary, _
    tx As String _
) As Scripting.Dictionary
    Dim hd As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ls As Variant
    Dim mx As Long
    Dim dx As Long
    Dim dlm As String
    
    Set rt = New Scripting.Dictionary
    With dc
        Set hd = dcOb(.Item(""))
        With hd
            dlm = .Item("")
            
            ls = Split(tx, dlm)
            mx = UBound(ls)
            For dx = LBound(ls) To mx
                If .Exists(dx) Then
                    rt.Add .Item(dx), ls(dx)
                End If
                rt.Add dx, ls(dx)
            Next
        End With
        
        dx = .Count
        Do While .Exists(dx)
            dx = 1 + dx
        Loop
        .Add dx, rt
    End With
    
    Set d1g4f4 = dc
End Function

Public Function d1g4f5(tx As String, _
    dc As Scripting.Dictionary, _
    Optional bk As String = vbNewLine _
) As Scripting.Dictionary
    Dim ck As Long
    
    If Len(tx) > 0 Then
        ck = InStr(tx, bk)
        If ck > 0 Then
            Set d1g4f5 = d1g4f5( _
                Mid$(tx, ck + Len(bk)), _
                d1g4f4(dc, Left$(tx, ck - 1)), _
            bk)
        Else
            Set d1g4f5 = d1g4f4(dc, tx)
        End If
    Else
        Set d1g4f5 = dc
    End If
End Function

Public Function d1g4f6(tx As String, _
    Optional dlm As String = ",", _
    Optional bk As String = vbNewLine _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim hdr As String
    Dim ck As Long
    
    Set rt = New Scripting.Dictionary
    ck = InStr(1, tx, bk)
    If ck > 0 Then
        hdr = Left$(tx, ck - 1)
        rt.Add "", d1g4f3(hdr, dlm)
        Set rt = d1g4f5( _
            Mid$(tx, ck + Len(bk)), _
            rt, bk _
        )
    Else
        'handle straight text here?
    End If
    
    Set d1g4f6 = rt
'Debug.Print ConvertToJson(d1g4f6(d1g4f2("zzCsv000", "zc0g0f2"), "|"), "  ")
End Function

Public Function d1g1f0() As Variant
    Dim dc As Scripting.Dictionary
    
    Set dc = dcRemapByPtNum( _
        dcAiDocComponents(aiDocActive()) _
    )
    Debug.Print txDumpLs(dc.Keys)

    d1g1f0 = dc.Keys
'Debug.Print txDumpLs(dcRemapByPtNum(dcAiDocComponents(ThisApplication.ActiveDocument)).Keys)
End Function

Public Function d1g1f2(pd As Long, _
    Optional fc As Long = 2 _
) As String
    Dim ct As Long
    Dim dv As Long
    Dim rt As String
    
    If fc > pd Then
        d1g1f2 = ""
    Else
        ct = 0
        dv = pd
        Do Until dv Mod fc > 0
            ct = 1 + ct
            dv = dv \ fc
        Loop
        If ct > 0 Then
            rt = CStr(fc) & "," _
               & CStr(ct) & vbNewLine
        Else
            rt = ""
        End If
        d1g1f2 = rt & d1g1f2(dv, 1 + fc)
    End If
'
'
End Function

Public Function d1g1f3(pd As Long, _
    Optional tHdr As String = "Factor,Power", _
    Optional fSep As String = ",", _
    Optional lSep As String = vbNewLine _
) As String
    Dim fc As Long
    Dim ct As Long
    Dim dv As Long
    Dim wk As String
    Dim rt As String
    
    rt = tHdr
    ct = d1g1f3b2(pd)
    dv = pd \ 2 ^ ct
    If ct > 0 Then rt _
        = rt & lSep _
        & "2" & fSep _
        & CStr(ct)
    
    fc = 3
    Do Until fc > dv
        ct = 0
        Do Until dv Mod fc > 0
            ct = 1 + ct
            dv = dv \ fc
        Loop
        
        If ct > 0 Then rt _
            = rt & lSep _
            & CStr(fc) & fSep _
            & CStr(ct)
        
        fc = fc + 2
    Loop
    d1g1f3 = rt
'
'Debug.Print d1g1f3(489168, "1&", " ^ ", " * ")
'   Some odd behavior on this one.
'   The result, as printed:
'       1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
'   can be processed in immediate mode
'   to return the original value.
'   However, there MUST be a space BEFORE
'   the caret (^) in order for it to be
'   interpreted correctly. Otherwise, it
'   seems to be recognized as some sort
'   of separator, as seen in these examples:
'
'Debug.Print "Debug.Print " & d1g1f3(489168, "1&", "^", "*")
'Debug.Print 1& * 2^; 4 * 3^; 2 * 43^; 1 * 79^; 1
' 2  12  86  79  1
'Debug.Print "Debug.Print " & d1g1f3(489168, "1&", "^", " * ")
'Debug.Print 1& * 2^; 4 * 3^; 2 * 43^; 1 * 79^; 1
' 2  12  86  79  1
'Debug.Print "Debug.Print " & d1g1f3(489168, "1&", " ^", "*")
'Debug.Print 1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
'489168
'Debug.Print "Debug.Print " & d1g1f3(489168, "1&", "^ ", "*")
'Debug.Print 1& * 2^; 4 * 3^; 2 * 43^; 1 * 79^; 1
' 2  12  86  79  1
'
'   It LOOKS like the caret is treated as some sort
'   of type indicator, like !, #, % and & are used
'   to indicate single, double, integer and long
'   values. This would also seem to be supported
'   by the presence of semicolons in the samples
'   above. Those were NOT PRINTED in immediate mode!
'   Instead, they were almost certainly added when
'   copied into the editor. They will be left as
'   they are, here, to show the interpreter does this.
'
'   Follow-up confirms speculation: the caret is
'   indeed a type indicator, but ONLY in 64-bit VBA.
'   It indicates a LongLong value. See references:
'   https://stackoverflow.com/questions/51264287/
'       vba-power-operator-not-working-as-expected-in-64-bit-vba
'   https://docs.microsoft.com/en-us/office/vba/language/
'       reference/user-interface-help/longlong-data-type
'
'   Further confirmation from testing in Excel VBA,
'   which here at Doyle is still a 32-bit installation:
'
'Debug.Print "Debug.Print " & d1g1f3(489168, "1&", "^ ", "*")
'Debug.Print 1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
'489168
'Debug.Print "Debug.Print " & d1g1f3(489168, "1&", "^", "*")
'Debug.Print 1& * 2 ^ 4 * 3 ^ 2 * 43 ^ 1 * 79 ^ 1
'489168
'
'   These are the same examples which produced lists
'   of numbers in Inventor VBA, and were "fixed"
'   on pasting into this code before commenting.
'   Actually, these strings WERE fixed, but only
'   by inserting spaces where none were previously.
'   In any case, this looks like mystery solved.
'   On top of that, however, it would appear that
'   Inventor VBA is a 64-bit implementation, which
'   means it SHOULD support the LongLong data type.
'   THAT might prove interesting to explore...
'
'
End Function

Public Function d1g1f3b2(pd As Long) As Long
    Dim fc As Long
    Dim ct As Long
    Dim dv As Long
    Dim wk As String
    Dim rt As String
    
    ct = 0
    dv = pd
    Do Until 1 And dv
        ct = 1 + ct
        dv = dv \ 2
    Loop
    
    d1g1f3b2 = ct
'
'
End Function

Public Function fcPrime(n As Long _
    , Optional rt As String = "" _
    , Optional ls As String = "BCEGKMQSW" _
) As String
    Dim nx As Long
    Dim md As Long
    Dim fc As Long
    Dim ct As Long
    
    If n > 0 Then
        If n = 1 Then
            fcPrime = rt
        ElseIf Len(ls) > 0 Then
            nx = n
            fc = 31 And Asc(ls)
            md = nx Mod fc
            ct = 0
            Do Until md > 0
                ct = 1 + ct
                nx = nx \ fc
                md = nx Mod fc
            Loop
            fcPrime = fcPrime(nx, _
                rt & Chr$(48 + ct), _
                Mid$(ls, 2) _
            )
        Else
            fcPrime = rt & "|" & CStr(n)
        End If
    Else
        fcPrime = ""
    End If
End Function

Public Function fcCommon(s0 As String, s1 As String) As String
    Dim n0 As Long
    Dim n1 As Long
    
    n0 = Len(s0) * Len(s1)
    If n0 > 0 Then
        n0 = Asc(s0)
        n1 = Asc(s1)
        fcCommon = Chr$(IIf(n0 > n1, n1, n0)) _
            & fcCommon(Mid$(s0, 2), Mid$(s1, 2))
    Else
        fcCommon = ""
    End If
End Function

Public Function fcProduct(s As String _
    , Optional ls As String = "BCEGKMQSW" _
) As Long
    If Len(ls) > 0 Then
        If Len(s) > 0 Then
            fcProduct = (31 And Asc(ls)) _
               ^ (15 And Asc(s)) _
               * fcProduct( _
                    Mid$(s, 2), _
                    Mid$(ls, 2) _
               )
        Else
            fcProduct = 1
        End If
    Else
        fcProduct = -1
    End If
End Function

Public Function fcMaxComm(n0 As Long, n1 As Long) As Long
    '''
    ''' fcMaxComm -- Return Greatest Common Factor
    '''
    fcMaxComm = fcProduct(fcCommon( _
        fcPrime(n0), fcPrime(n1) _
    ))
'
'For m = 2 To 49: For n = 1 + m To 49: cf = fcMaxComm((m), (n)): Debug.Print IIf(cf > 1, CStr(n) & "<" & CStr(cf) & ">" & CStr(m) & vbNewLine, "");: Next: Next
'
End Function

Public Function gcfTest() As Long
    '''
    ''' gcfTest -- Test GCF Function fcMaxComm
    '''
    Dim rt As Long
    Dim n0 As Long
    Dim n1 As Long
    Dim nd As Long
    Dim gf As Long
    
    rt = 0
    For n0 = 4 To 49
        nd = n0 - 1
        For n1 = 2 To nd
            gf = fcMaxComm(n0, n1)
            If (n0 Mod gf) _
            + (n1 Mod gf) _
            > 0 Then
                rt = 0
                Debug.Print ;
            End If
        Next
    Next
End Function

Public Function tbPrimesWithSquare( _
    Optional ct As Long = 100000 _
) As LongPtr()
    '''
    ''' tbPrimesWithSquare
    '''
    ''' Generate a table of primes
    ''' and their corresponding squares
    '''
    Dim dbg As Long
    Dim p() As LongPtr
    Dim n As LongPtr
    Dim mp As LongPtr
    Dim dx As Long
    Dim mx As Long
    Dim nx As Long
    
    Dim d0 As Double
    Dim d1 As Double
    
    ReDim p(1, ct)
    mx = UBound(p, 2)
    p(0, 0) = 2: p(1, 0) = 4
    nx = 1
    n = 3
    
    d0 = Timer
    Do
        dx = 0
        mp = 1
        Do
            mp = n Mod p(0, dx)
            If n > p(1, dx) Then
                dx = dx + 1
            Else
                dx = nx
            End If
        Loop While mp * p(0, dx) > 0
        
        If mp > 0 Then
            If p(0, dx) = 0 Then
                p(0, dx) = n
                On Error Resume Next
                p(1, dx) = n * n
                If Err.Number = 0 Then
                    nx = dx + 1
                Else
                    nx = mx + 1
                End If
                On Error GoTo 0
            Else
                Stop
            End If
        End If
        
        n = 1 + n
    Loop Until nx > mx
    d1 = Timer - d0
    
    dbg = 0 'Change to 1 for debug mode
    If dbg Then
        Debug.Print 1000 * d1 '- d0
        Stop
    End If
    
    tbPrimesWithSquare = p
End Function

Public Function d1g1f7() As Long
    Dim d0 As Double
    Dim d1 As Double
    Dim ur As VbMsgBoxResult
    
    d0 = Timer
    ur = MsgBox("", vbOKOnly, "")
    d1 = Timer - d0
    
    Stop
End Function

Public Function bcCtCommFac(dc As Scripting.Dictionary) As Long
    Dim ls As Variant
    Dim rt As Long
    Dim mx As Long
    Dim dx As Long
    
    With dc
        If .Count > 0 Then
            ls = .Keys
            mx = UBound(ls)
            rt = CLng(.Item(ls(0)))
            dx = 1
            
            Do
                rt = fcMaxComm(rt, _
                    CLng(.Item(ls(dx))) _
                )
                If rt = 1 Then
                    dx = 1 + mx
                Else
                    dx = 1 + dx
                End If
            Loop Until dx > mx
        Else
            rt = 1
        End If
    End With

    bcCtCommFac = rt
'
'
End Function

Public Function dcBoltConn1byGCF( _
    dc As Scripting.Dictionary, _
    Optional fc As Long = 0 _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ct As Long
    
    If fc > 0 Then
        Set rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                ct = CLng(.Item(ky))
                rt.Add ky, ct \ fc
            Next
        End With
    Else
        Set rt = dcBoltConn1byGCF(dc, bcCtCommFac(dc))
    End If
    
    Set dcBoltConn1byGCF = rt
'
'Debug.Print dumpLsKeyVal(dcBoltConn1byGCF(dcOfBoltConn02(ad))): Debug.Print
'
End Function

Public Function aiDocProp( _
    AiDoc As Inventor.Document, propName As String, _
    Optional propSet As String = gnCustom _
) As Inventor.Property
    '''
    ''' Proposed Name: aiDocProp
    '''
    Dim rt As Inventor.Property
    
    If AiDoc Is Nothing Then
        Set aiDocProp = Nothing
    Else
        With AiDoc.PropertySets
            If .PropertySetExists(propSet) Then
                '.Item(propSet).GetPropertyInfo()
                Set aiDocProp = aiGetProp( _
                    .Item(propSet), _
                    propName, 0 _
                )
            Else
                Set aiDocProp = Nothing
            End If
        End With
    End If
End Function

Public Function aiDocPropVal( _
    AiDoc As Inventor.Document, propName As String, _
    Optional propSet As String = gnCustom _
) As Variant
    '''
    ''' Proposed Name: aiDocPropVal
    '''
    aiDocPropVal = aiPropVal(aiDocProp( _
        AiDoc, propName, propSet _
    ))
End Function

Public Function d1g2f0() As Variant
    '
    d1g2f0 = ""
End Function

Public Function d1g2f1(AiDoc As Inventor.Document) As Variant
    '''
    '''
    '''
    Dim pt As Long
    Dim sc As Long
    
    pt = 0: sc = 0
    
    If InStr(1, AiDoc.FullFileName, _
        "\Doyle_Vault\Designs\purchased\" _
    ) > 0 Then pt = pt Or 1: sc = sc + 1
    
    If InStr(1, _
        "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
        "|" & CStr(aiDocPropVal( _
            AiDoc, pnFamily, gnDesign _
        )) & "|" _
    ) > 0 Then pt = pt Or 2: sc = sc + 1
    
    d1g2f1 = ""
End Function

Public Function d1g2f2( _
    AiDoc As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    '''
    '''
    '''
    'Dim dc As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim oc As Inventor.ComponentOccurrence
    'Dim rs As ADODB.Recordset
    Dim ob As Inventor.Document
    'Dim ky As Variant
    'Dim pn As String
    'Dim pt As Long
    'Dim sc As Long
    Dim dx As Long
    Dim bs As Inventor.BOMStructureEnum
    
    Set rt = New Scripting.Dictionary
    dx = rt.Count
    'pt = 0: sc = 0
    
    With AiDoc.ComponentDefinition
        For Each oc In .Occurrences
            With oc
                Set ob = aiDocument( _
                    .Definition.Document _
                )
                If .BOMStructure _
                    = kPhantomBOMStructure _
                Then
                    If .DefinitionDocumentType _
                        = kAssemblyDocumentObject _
                    Then
                        With aiDocAssy(ob)
                            '.ComponentDefinition.BOMStructure
                            If .DocumentInterests.HasInterest(guidDesignAccl) Then
                            Else
                            End If
                        End With
                    Else
                    End If
                Else
                End If
                With .Definition
                    With ob
                    End With
                End With
            End With
        Next
    End With
    'If InStr(1, aiDoc.FullFileName, _
        "\Doyle_Vault\Designs\purchased\" _
    ) > 0 Then pt = pt Or 1: sc = sc + 1
    
    'If InStr(1, _
        "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
        "|" & CStr(aiDocPropVal( _
            aiDoc, pnFamily, gnDesign _
        )) & "|" _
    ) > 0 Then pt = pt Or 2: sc = sc + 1
    
    Set d1g2f2 = rt
End Function

Public Function d1g2f3( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    rt.Add AiDoc.DocumentType, AiDoc
    rt.Add AiDoc.DocumentSubType.DocumentSubTypeID, AiDoc
    Set d1g2f3 = rt
End Function

Public Function d1g3f0() As Variant
    '''
    '''
    '''
    d1g3f0 = ""
End Function

Public Function d1g3f1( _
    ad As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    '''
    ''' d1g3f1 --
    '''
    ''' Generate counts of components
    ''' in supplied Assembly, adding
    ''' a sub-Dictionary for any
    ''' "phantom" component recognized
    ''' as either a Bolted Connection,
    ''' or an Assembly of entirely
    ''' Content Center components.
    '''
    ''' (the latter case addresses
    ''' an issue encountered with
    ''' just such an Assembly)
    '''
    Dim rt As Scripting.Dictionary
    Dim oc As Inventor.ComponentOccurrence
    Dim sd As Inventor.Document
    Dim nm As String
    Dim bc As Scripting.Dictionary
    Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    If ad Is Nothing Then 'we got nothing to work with
    Else
        For Each oc In ad.ComponentDefinition.Occurrences
            Set sd = aiDocument(oc.Definition.Document)
            nm = sd.FullDocumentName
            
            With rt
            If .Exists(nm) Then
                ar = .Item(nm)
                Debug.Print ;
                ar(1) = ar(1) + 1
                .Item(nm) = ar
                Debug.Print ;
            Else
                Set bc = Nothing
                If oc.BOMStructure = kPhantomBOMStructure Then
                    With sd.DocumentInterests
                    If .HasInterest(guidDesignAccl) Then
                        Debug.Print "FOUND Design Accelerator"
                        Debug.Print ;
                        Set bc = d1g3f1(sd) 'New Scripting.Dictionary
                    Else
                        Debug.Print "FOUND Phantom Assembly"
                        Debug.Print vbTab & "NOT Design Accelerator"
                        Debug.Print ;
                        Set bc = dcIfDesignAccel(d1g3f1(sd))
                        If bc Is Nothing Then
                        Else
                            Debug.Print vbTab & "but ALL Members ARE Content Center"
                            Debug.Print vbTab & "so WILL Process as Such"
                        End If
                    End If
                    End With
                    
                    Debug.Print vbTab & sd.FullDocumentName
                    Debug.Print vbTab & aiDocPropVal(sd, pnPartNum, gnDesign)
                End If
                rt.Add nm, Array(sd, 1, bc)
            End If
            End With
        Next
    End If
    
    Set d1g3f1 = rt
'For Each oc In sd.ComponentDefinition.Occurrences: Debug.Print oc.Name; "|"; oc.RangeBox.MinPoint.X; "|"; oc.RangeBox.MinPoint.Y; "|"; oc.RangeBox.MinPoint.Z; "|"; oc.RangeBox.MaxPoint.X; "|"; oc.RangeBox.MaxPoint.Y; "|"; oc.RangeBox.MaxPoint.Z: Next
End Function

Public Function d1g3f2( _
    ad As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim sk As Variant
    Dim ar As Variant
    Dim sa As Variant
    Dim ct As Long
    
    Set rt = New Scripting.Dictionary
    With d1g3f1(ad)
    For Each ky In .Keys
        ar = .Item(ky)
        If ar(2) Is Nothing Then
            rt.Add ky, ar '(1)
        Else
            With dcOb(ar(2))
            For Each sk In .Keys
                sa = .Item(sk)
                ct = ar(1) * sa(1)
                With rt
                If .Exists(sk) Then 'some already counted
                    'so need to add to existing total
                    
                    'ct = ct + sa(1) '.Item(sk)
                    sa(1) = ct + .Item(sk)(1)
                    'got type mismatch here, and fixed
                    'but not sure fix is correct
                    
                    .Item(sk) = sa 'ct
                Else 'this is a whole new component
                    'so just add its count to the list
                    
                    sa(1) = ct
                    .Add sk, sa 'ct .Item(sk)
                End If
                End With
            Next
            End With
        End If
    Next
    End With
    
    Set d1g3f2 = rt
End Function

Public Function d1g3f3( _
    ad As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    With d1g3f2(ad)
    For Each ky In .Keys
        ar = .Item(ky)
        With aiDocument(obOf(ar(0))).PropertySets
        With .Item(gnDesign).Item(pnPartNum)
            rt.Add .Value, ar
        End With
        End With
    Next
    End With
    
    Set d1g3f3 = rt
End Function

Public Function d1g3f4( _
    ad As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    With d1g3f3(ad)
    For Each ky In .Keys
        ar = .Item(ky)
        rt.Add ky, ar(1)
    Next
    End With
    
    Set d1g3f4 = rt
End Function

Public Function d1g3f5( _
    ad As Inventor.AssemblyDocument, _
    Optional incTop As Long = 0 _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim sd As Inventor.AssemblyDocument
    Dim ky As Variant
    'Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    With dcRemapByPtNum( _
        dcAiDocComponents(ad, , incTop, 1) _
    )
    For Each ky In .Keys
        Set ad = aiDocAssy(obOf(.Item(ky)))
        If ad Is Nothing Then 'skip it
        Else
            ''  Previous test, just for Bolted Connection
            'With ad.DocumentInterests
            'If .HasInterest(guidDesignAccl) Then
            ''  Replaced with test for ALL Phantom (below)
            
            With ad.ComponentDefinition
            If .BOMStructure = kPhantomBOMStructure Then
                'Phantom -- don't add to Dictionary
                Debug.Print ;
            Else
                rt.Add ky, d1g3f4(ad)
            End If
            End With
        End If
    Next
    End With
    
    Set d1g3f5 = rt
End Function

Public Function d1g3f6( _
    ad As Inventor.AssemblyDocument, _
    Optional incTop As Long = 0 _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    With d1g3f5(ad, incTop)
    For Each ky In .Keys
        rt.Add ky & "|" & ky & "|1", _
            vbNewLine & ky & "|" & dumpLsKeyVal( _
                dcOb(.Item(ky)), "|", _
                vbNewLine & ky & "|" _
            )
        'old key "[" & ky & "]"
        'replaced with pipe-delimited
        'record to fit in better
    Next
    End With
    
    Set d1g3f6 = rt
'Debug.Print dumpLsKeyVal(d1g3f6(aiDocument(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences.Item(1).Definition.Document)), "")
End Function

Public Function d1g3f7( _
    ad As Inventor.AssemblyDocument, _
    Optional incTop As Long = 0 _
) As String
    d1g3f7 = "Product|ItemCode|Qty" & vbNewLine _
    & dumpLsKeyVal(d1g3f6(ad, incTop), "")
'
'Debug.Print d1g3f7(aiDocActive())
'
End Function

Public Function dcOfBoltConnReLabeled( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim rd As Scripting.Dictionary
    Dim wd As Scripting.Dictionary
    Dim ky As Variant
    Dim pn As String
    Dim fn As Variant
    Dim ad As Inventor.Document
    Dim pr As Inventor.Property
    
    Set rt = New Scripting.Dictionary
    
    With dc
        For Each ky In .Keys
            Set wd = dcOb(.Item(ky))
            Set rd = New Scripting.Dictionary
            
            pn = InputBox( _
                Join(Array( _
                    "Part Number proposed", _
                    "  for subassemblies", _
                    Join(wd.Keys, _
                        vbNewLine & "    " _
                    ), _
                    "", _
                    "Modify as necessary,", _
                    "then click OK to confirm.", _
                    "" _
                )), _
                "Verify BC Part Number", CStr(ky) _
            )
            With wd
                For Each fn In .Keys
                    Set ad = aiDocument( _
                        obOf(.Item(fn)) _
                    )
                    Set pr = aiDocProp(ad, _
                        pnPartNum, gnDesign _
                    )
                    If pr Is Nothing Then
                        'nothing to do
                    Else
                        pr.Value = pn
                        Debug.Print ;
                        
                        rd.Add fn, ad
                    End If
                Next
            End With
            
            If rd.Count > 0 Then
                rt.Add pn, rd
            End If
        Next
    End With
    
    Set dcOfBoltConnReLabeled = rt
'
'Debug.Print txDumpLs(dcOfBoltConnReLabeled(dcOfBoltConnIn(aiDocAssy(aiDocActive()))).Keys)
'
End Function

Public Function dcOfBoltConnIn( _
    ad As Inventor.AssemblyDocument, _
    Optional incTop As Long = 0 _
) As Scripting.Dictionary
    With dcAiDocComponents(ad, , incTop, 1)
    End With
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim wd As Scripting.Dictionary
    Dim sd As Inventor.AssemblyDocument
    Dim ky As Variant
    Dim pn As String
    Dim dn As String
    'Dim ar As Variant
    
    Set rt = New Scripting.Dictionary
    With dcAiDocComponents(ad, , incTop, 1)
    For Each ky In .Keys
        Set sd = aiDocAssy(obOf(.Item(ky)))
        pn = pnOfBoltConn(sd)
        
        If Len(pn) > 0 Then
            dn = sd.FullDocumentName
            
            With rt
                If .Exists(pn) Then
                    Set wd = dcOb(.Item(pn))
                Else
                    Set wd = New Scripting.Dictionary
                    .Add pn, wd
                End If
            End With
            
            With wd
                If .Exists(dn) Then
                    If obOf(.Item(dn)) Is sd Then
                        'should be good
                    Else 'not so sure
                        Stop
                    End If
                Else
                    .Add dn, sd
                End If
            End With
        End If
    Next
    End With
    
    Set dcOfBoltConnIn = rt
'
'Debug.Print txDumpLs(dcOfBoltConnIn(aiDocAssy(aiDocActive())).Keys)
'
End Function

Public Function aiDocContentMember( _
    ad As Inventor.PartDocument _
) As Inventor.PartDocument
    If ad Is Nothing Then
        Set aiDocContentMember = ad
    ElseIf ad.ComponentDefinition.IsContentMember Then
        Set aiDocContentMember = ad
    Else
        Set aiDocContentMember = Nothing
    End If
End Function

Public Function dcIfDesignAccel( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcIfDesignAccel
    '''
    ''' Accepting a Dictionary of form
    ''' generated by d1g3f1, verify
    ''' that all Items represent Content
    ''' Center components, and return
    ''' same Dictionary if so.
    '''
    ''' If any Items are NOT Content Center
    ''' components, return Nothing
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ar As Variant
    
    Set rt = dc
    
    With dc
        For Each ky In .Keys
            ar = .Item(ky)
            Debug.Print ;
            If aiDocContentMember( _
                aiDocPart(aiDocument( _
                obOf(ar(0)) _
            ))) Is Nothing Then
                Set rt = Nothing
            End If
        Next
    End With
    
    Set dcIfDesignAccel = rt
End Function

Public Function rsOfBoltConn( _
    ad As Inventor.AssemblyDocument _
) As ADODB.Recordset
    '''
    ''' rsOfBoltConn -- rsOfBoltConn
    '''
    ''' Return Recordset of Components
    ''' of one supplied Assembly Document,
    ''' provided it's a Bolted Connection.
    '''
    ''' Call rsOfBoltConnRedux against this function's
    ''' resulting Recordset rt to condense it
    ''' to definition of a single instance,
    ''' with a count of each member indicating
    ''' the number of instances.
    '''
    ''' (was going to call rsOfBoltConnRedux here and
    ''' return THAT result, but realized
    ''' this function's preprocessed result
    ''' might prove useful in itself, so
    ''' decided to return it directly
    ''' after all)
    '''
    Dim rt As ADODB.Recordset 'Scripting.Dictionary
    Dim pNum As ADODB.Field
    Dim fNam As ADODB.Field
    Dim zPos As ADODB.Field
    Dim xCen As ADODB.Field
    Dim yCen As ADODB.Field
    
    Dim oc As Inventor.ComponentOccurrence
    Dim sd As Inventor.PartDocument
    Dim bc As Scripting.Dictionary
    Dim p0(2) As Double
    Dim p1(2) As Double
    
    Set rt = rsForBoltConn() 'New Scripting.Dictionary
    With rt.Fields
        Set pNum = .Item("pNum")
        Set fNam = .Item("fNam")
        Set zPos = .Item("zPos")
        Set xCen = .Item("xCen")
        Set yCen = .Item("yCen")
    End With
    
    If ad Is Nothing Then 'we got nothing to work with
    Else
        Set bc = Nothing
        If ad.ComponentDefinition.BOMStructure = kPhantomBOMStructure Then
            With ad.DocumentInterests
            If .HasInterest(guidDesignAccl) Then
                Set bc = d1g3f1(ad)
            Else
                Set bc = dcIfDesignAccel(d1g3f1(ad))
            End If
            End With
            
            If bc Is Nothing Then 'do nothing
            Else
                For Each oc In ad.ComponentDefinition.Occurrences
                    With oc
                        Set sd = aiDocument(.Definition.Document)
                        With .RangeBox
                            .MinPoint.GetPointData p0
                            .MaxPoint.GetPointData p1
                        End With
                    End With
                    
                    
                    rt.AddNew
                    pNum.Value = aiDocPropVal(sd, pnPartNum, gnDesign)
                    fNam.Value = sd.FullDocumentName
                    
                    zPos.Value = Round(p0(2), 3)
                    'Debug.Print FormatNumber(p0(2), 3); " ";
                    xCen.Value = Round((p0(0) + p1(0)) / 2, 3)
                    'Debug.Print FormatNumber((p0(0) + p1(0)) / 2, 3); " ";
                    yCen.Value = Round((p0(1) + p1(1)) / 2, 3)
                    'Debug.Print FormatNumber((p0(1) + p1(1)) / 2, 3); " ";
                    'Debug.Print
                    
                    Debug.Print ;
                Next
                
                With rt
                    .Filter = ""
                    If .BOF Then
                        .AddNew
                        pNum.Value = "NONE"
                        fNam.Value = "No Hardware, or Not Bolted Connection!"
                        zPos.Value = 0
                        xCen.Value = 0
                        yCen.Value = 0
                    End If
                    .Sort = "zPos, pNum, yCen, xCen"
                End With
            End If
        End If
    End If
    
    Set rsOfBoltConn = rt 'rsOfBoltConnRedux()
'Debug.Print rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document).GetString
End Function

Public Function dcOfBoltConn( _
    ad As Inventor.AssemblyDocument _
) As Scripting.Dictionary 'ADODB.Recordset 'Scripting.Dictionary
    '''
    ''' dcOfBoltConn
    '''
    ''' Alternate implementation of rsOfBoltConn
    ''' returning a Dictionary instead of
    ''' a Recordset. However, this loses
    ''' the benefit of a Recordset's Sort
    ''' capability, and so is unlikely
    ''' to prove as useful.
    '''
    Dim dc As Scripting.Dictionary
    Dim k0 As String
    Dim ct As Long
    
    Dim rt As ADODB.Recordset 'Scripting.Dictionary
    Dim pNum As ADODB.Field
    Dim fNam As ADODB.Field
    Dim zPos As ADODB.Field
    Dim xCen As ADODB.Field
    Dim yCen As ADODB.Field
    
    Dim oc As Inventor.ComponentOccurrence
    Dim sd As Inventor.PartDocument
    Dim bc As Scripting.Dictionary
    Dim p0(2) As Double
    Dim p1(2) As Double
    
    Set dc = New Scripting.Dictionary
    Set rt = rsForBoltConn()
    With rt.Fields
        Set pNum = .Item("pNum")
        Set fNam = .Item("fNam")
        Set zPos = .Item("zPos")
        Set xCen = .Item("xCen")
        Set yCen = .Item("yCen")
    End With
    
    If ad Is Nothing Then 'we got nothing to work with
    Else
        Set bc = Nothing
        If ad.ComponentDefinition.BOMStructure = kPhantomBOMStructure Then
            With ad.DocumentInterests
            If .HasInterest(guidDesignAccl) Then
                Set bc = d1g3f1(ad)
            Else
                Set bc = dcIfDesignAccel(d1g3f1(ad))
            End If
            End With
            
            If bc Is Nothing Then 'do nothing
            Else
                For Each oc In ad.ComponentDefinition.Occurrences
                    With oc
                        Set sd = aiDocument(.Definition.Document)
                        With .RangeBox
                            .MinPoint.GetPointData p0
                            .MaxPoint.GetPointData p1
                        End With
                    End With
                    
                    k0 = FormatNumber(p0(2), 3) _
                    & "|" & aiDocPropVal( _
                        sd, pnPartNum, gnDesign _
                    )
                    With dc
                        If .Exists(k0) Then
                            ct = 1 + .Item(k0)
                            .Item(k0) = ct
                        Else
                            .Add k0, 1
                        End If
                    End With
                    
                    rt.AddNew
                    pNum.Value = aiDocPropVal(sd, pnPartNum, gnDesign)
                    fNam.Value = sd.FullDocumentName
                    
                    zPos.Value = Round(p0(2), 3)
                    'Debug.Print FormatNumber(p0(2), 3); " ";
                    xCen.Value = Round((p0(0) + p1(0)) / 2, 3)
                    'Debug.Print FormatNumber((p0(0) + p1(0)) / 2, 3); " ";
                    yCen.Value = Round((p0(1) + p1(1)) / 2, 3)
                    'Debug.Print FormatNumber((p0(1) + p1(1)) / 2, 3); " ";
                    'Debug.Print
                    
                    Debug.Print ;
                Next
                
                With rt
                    .Filter = ""
                    If .BOF Then
                        .AddNew
                        pNum.Value = "NONE"
                        fNam.Value = "No Hardware, or Not Bolted Connection!"
                        zPos.Value = 0
                        xCen.Value = 0
                        yCen.Value = 0
                    End If
                    .Sort = "zPos, pNum, yCen, xCen"
                End With
            End If
        End If
    End If
    
    Set dcOfBoltConn = dc 'rt
'Debug.Print dcOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document).GetString
End Function

Public Function rsForBoltConn() As ADODB.Recordset
    '''
    ''' rsForBoltConn -- rsForBoltConn
    '''
    ''' Generate an new, empty Recordset
    ''' to gather data on Bolted Connection
    '''
    Dim rt As ADODB.Recordset
    
    Set rt = New ADODB.Recordset
    With rt
        With .Fields
            .Append "zPos", adDouble
            .Append "pNum", adVarChar, 63
            .Append "fNam", adVarChar, 255
            '.Append "", adVarChar, 63
            .Append "xCen", adDouble
            .Append "yCen", adDouble
            '.Append "", adDouble
            '.Append "", adDouble
        End With
        .Open
    End With
    Set rsForBoltConn = rt
End Function

Public Function rsOfBoltConnRedux( _
    rs As ADODB.Recordset _
) As ADODB.Recordset
    '''
    ''' rsOfBoltConnRedux
    '''
    ''' Condense supplied Recordset
    ''' of Bolted Connection Assembly
    ''' to summary of Components of
    ''' ONE instance.
    '''
    ''' Include count of each member
    ''' Component in Assembly, which
    ''' should be the same for ALL
    ''' Components, and reflect the
    ''' total number of instances
    ''' in the Assembly.
    '''
    ''' In most cases, this count
    ''' should be just one, given
    ''' the way Bolted Connections
    ''' are generated and used here.
    ''' However, some models might
    ''' be found which use patterns
    ''' or multiple holes, thus
    ''' producing one BC Assembly
    ''' defining multiple instances.
    ''' A means to address this might
    ''' therefore be required in future.
    '''
    Dim rt As ADODB.Recordset
    Dim pNumIn As ADODB.Field
    Dim zPosIn As ADODB.Field
    Dim pNumOut As ADODB.Field
    Dim zPosOut As ADODB.Field
    Dim xCenOut As ADODB.Field
    
    Dim dc As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ky As Variant
    
    Dim zp As Double
    Dim pn As String
    
    Set dc = New Scripting.Dictionary
    
    With rs
        With .Fields
            Set pNumIn = .Item("pNum")
            Set zPosIn = .Item("zPos")
        End With
        
        .Sort = "zPos"
        If Not .BOF Then
        Do Until .EOF
            With dc
                zp = zPosIn.Value
                If .Exists(zp) Then
                    Set wk = .Item(zp)
                Else
                    Set wk = New Scripting.Dictionary
                    .Add zp, wk
                End If
            End With
                
            With wk
                pn = pNumIn.Value
                If .Exists(pn) Then
                    .Item(pn) = 1 + .Item(pn)
                Else
                    .Add pn, 1
                End If
            End With
            .MoveNext
        Loop
        End If
    End With
    
    Set rt = rsForBoltConn()
    With rt
        With .Fields
            Set pNumOut = .Item("pNum")
            Set zPosOut = .Item("zPos")
            Set xCenOut = .Item("xCen")
        End With
        
        With dc
            For Each ky In .Keys
                Set wk = .Item(ky)
                
                With wk
                If .Count > 1 Then
                    Stop
                Else
                    rt.AddNew
                    zPosOut.Value = CDbl(ky)
                    pNumOut.Value = .Keys(0)
                    xCenOut.Value = CDbl(.Items(0))
                End If
                End With
            Next
        End With
        
        .Filter = ""
        .Sort = "zPos, xCen"
    End With
    
    Set rsOfBoltConnRedux = rt
'Debug.Print rsOfBoltConnRedux(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)).GetString
End Function

Public Function rsOfBoltConnRedux02( _
    rs As ADODB.Recordset _
) As ADODB.Recordset
    '''
    ''' rsOfBoltConnRedux02
    '''
    ''' Condense supplied Recordset
    ''' of Bolted Connection Assembly
    ''' to summary of Components of
    ''' ONE instance.
    '''
    ''' Include count of each member
    ''' Component in Assembly, which
    ''' should be the same for ALL
    ''' Components, and reflect the
    ''' total number of instances
    ''' in the Assembly.
    '''
    ''' In most cases, this count
    ''' should be just one, given
    ''' the way Bolted Connections
    ''' are generated and used here.
    ''' However, some models might
    ''' be found which use patterns
    ''' or multiple holes, thus
    ''' producing one BC Assembly
    ''' defining multiple instances.
    ''' A means to address this might
    ''' therefore be required in future.
    '''
    Dim rt As ADODB.Recordset
    Dim pNumIn As ADODB.Field
    Dim zPosIn As ADODB.Field
    Dim pNumOut As ADODB.Field
    Dim zPosOut As ADODB.Field
    Dim xCenOut As ADODB.Field
    
    Dim dc As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ky As Variant
    
    Dim zp As Double
    Dim pn As String
    
    Set dc = New Scripting.Dictionary
    
    With rs
        With .Fields
            Set pNumIn = .Item("pNum")
            Set zPosIn = .Item("zPos")
        End With
        
        .Sort = "zPos"
        If Not .BOF Then
        Do Until .EOF
            With dc
                zp = zPosIn.Value
                If .Exists(zp) Then
                    Set wk = .Item(zp)
                Else
                    Set wk = New Scripting.Dictionary
                    .Add zp, wk
                End If
            End With
                
            With wk
                pn = pNumIn.Value
                If .Exists(pn) Then
                    .Item(pn) = 1 + .Item(pn)
                Else
                    .Add pn, 1
                End If
            End With
            .MoveNext
        Loop
        End If
    End With
    
    Set rt = rsForBoltConn()
    With rt
        With .Fields
            Set pNumOut = .Item("pNum")
            Set zPosOut = .Item("zPos")
            Set xCenOut = .Item("xCen")
        End With
        
        With dc
            For Each ky In .Keys
                Set wk = .Item(ky)
                
                With wk
                If .Count > 1 Then
                    Stop
                Else
                    rt.AddNew
                    zPosOut.Value = CDbl(ky)
                    pNumOut.Value = .Keys(0)
                    xCenOut.Value = CDbl(.Items(0))
                End If
                End With
            Next
        End With
        
        .Filter = ""
        .Sort = "zPos, xCen"
    End With
    
    Set rsOfBoltConnRedux02 = rt
'Debug.Print rsOfBoltConnRedux02(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)).GetString
End Function

Public Function bcPtNumFromRS( _
    rs As ADODB.Recordset _
) As String
    bcPtNumFromRS = bcPtNumFromRSv2(rs)
End Function

Public Function bcPtNumFromRSv1( _
    rs As ADODB.Recordset _
) As String
    '''
    ''' bcPtNumFromRSv1
    '''
    ''' Generate a uniquely identifying
    ''' Part Number from supplied Recordset
    ''' Given a "Bolted Connection",
    '''
    Dim pNumIn As ADODB.Field
    Dim xCenIn As ADODB.Field
    Dim rt As String
    Dim pn As String
    Dim ft As Variant
    Dim ct As Long
    
    With rs
        With .Fields
            Set pNumIn = .Item("pNum")
            Set xCenIn = .Item("xCen")
        End With
        
        .Sort = "zPos"
        If .BOF Or .EOF Then
            rt = ""
        Else
            .Sort = "zPos"
            pn = pNumIn.Value
            rt = "BC" & Mid$( _
                pn, 3, Len(pn) - 4 _
            ) & Right$(pn, 2)
            ct = xCenIn.Value
            
            For Each ft In Array( _
                "zPos <= 0", _
                "zPos > 0" _
            )
                rt = rt & "-"
                .Filter = ft
                If Not .BOF Then
                    .Sort = "zPos"
                    Do Until .EOF
                        If ct <> xCenIn.Value Then
                            Stop
                        End If
                        pn = pNumIn.Value
                        rt = rt & Left$(pn, 2)
                        .MoveNext
                    Loop
                End If
            Next
            
            If ct > 1 Then
                rt = rt & Format$(ct, "-X00")
            End If
        End If
    End With
    
    bcPtNumFromRSv1 = rt
'Debug.Print bcPtNumFromRSv1(rsOfBoltConnRedux(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)))
End Function

Public Function bcPtNumFromRSv2( _
    rs As ADODB.Recordset _
) As String
    '''
    ''' bcPtNumFromRSv2
    '''
    ''' Generate a uniquely identifying
    ''' Part Number from supplied Recordset
    ''' Given a "Bolted Connection",
    '''
    Dim pNumIn As ADODB.Field
    Dim xCenIn As ADODB.Field
    Dim rt As String
    Dim pn As String
    Dim ft As Variant
    Dim ct As Long
    
    With rs
        With .Fields
            Set pNumIn = .Item("pNum")
            Set xCenIn = .Item("xCen")
        End With
        
        '.Sort = "zPos"
        .Filter = ""
        If .BOF Or .EOF Then
            rt = ""
        Else
            .Sort = "zPos"
            pn = pNumIn.Value
            rt = "BC" & Right$(pn, 1) & Mid$( _
                pn, 3, Len(pn) - 4 _
            ) '& Right$(pn, 2)
            ct = xCenIn.Value
            
            For Each ft In Array( _
                "zPos <= 0|zPos", _
                "zPos > 0|zPos desc" _
            )
                rt = rt & "-"
                .Filter = Left$(ft, InStr(ft, "|") - 1)
                If Not .BOF Then
                    .Sort = Mid$(ft, InStr(ft, "|") + 1) '"zPos"
                    rt = rt & Left$(pNumIn.Value, 2)
                    .MoveNext
                    Do Until .EOF
                        If ct <> xCenIn.Value Then
                            Stop
                        End If
                        'pn = pNumIn.Value
                        'rt = rt & Left$(pn, 2)
                        rt = rt & Left$(pNumIn.Value, 1)
                        .MoveNext
                    Loop
                End If
            Next
            
            If ct > 1 Then
                rt = rt & Format$(ct, "-X00")
            End If
        End If
    End With
    
    If Len(rt) > 23 Then
        Stop
    End If
    bcPtNumFromRSv2 = rt
'Debug.Print bcPtNumFromRSv2(rsOfBoltConnRedux(rsOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)))
End Function

Public Function pnOfBoltConn( _
    ad As Inventor.AssemblyDocument _
) As String
    pnOfBoltConn = bcPtNumFromRSv1(rsOfBoltConnRedux(rsOfBoltConn(ad)))
'Debug.Print pnOfBoltConn(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document)
End Function

Public Function dcOfBoltConn02( _
    ad As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    '''
    ''' dcOfBoltConn02
    '''
    ''' Second variation on dcOfBoltConn
    ''' returning a Dictionary of Component
    ''' quantities, keyed on Item Number.
    '''
    Dim rt As Scripting.Dictionary
    Dim pNum As String
    Dim ct As Long
    
    Dim oc As Inventor.ComponentOccurrence
    Dim sd As Inventor.PartDocument
    Dim bc As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    
    If ad Is Nothing Then 'we got nothing to work with
    Else
        Set bc = Nothing
        If ad.ComponentDefinition.BOMStructure = kPhantomBOMStructure Then
            With ad.DocumentInterests
            If .HasInterest(guidDesignAccl) Then
                Set bc = d1g3f1(ad)
            Else
                Set bc = dcIfDesignAccel(d1g3f1(ad))
            End If
            End With
            
            If bc Is Nothing Then 'do nothing
            Else
                For Each oc In ad.ComponentDefinition.Occurrences
                    Set sd = aiDocument(oc.Definition.Document)
                    
                    pNum = aiDocPropVal(sd, _
                        pnPartNum, gnDesign _
                    )
                    With rt
                        If .Exists(pNum) Then
                            ct = 1 + .Item(pNum)
                            .Item(pNum) = ct
                        Else
                            .Add pNum, 1
                        End If
                    End With
                    
                    Debug.Print ;
                Next
            End If
        End If
    End If
    
    Set dcOfBoltConn02 = rt
'Debug.Print dcOfBoltConn02(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(5).Definition.Document).GetString
End Function

Public Function rsFiltered( _
    rs As ADODB.Recordset, _
    Optional flText As String = "" _
) As ADODB.Recordset
    rs.Filter = flText
    Set rsFiltered = rs
End Function

Public Function rsFromGnsSql( _
    sqlText As String _
) As ADODB.Recordset
    '''
    '''
    '''
    Dim rt As ADODB.Recordset
    
    With cnGnsDoyle()
        Set rt = .Execute(sqlText)
        If rt Is Nothing Then
            Stop
        End If
        Set rsFromGnsSql = rt
    End With
End Function

Public Function rsAiPurch01fromDict( _
    dc As Scripting.Dictionary _
) As ADODB.Recordset
    '''
    '''
    '''
    Set rsAiPurch01fromDict = rsFromGnsSql( _
        sqlSelAiPurch01fromDict(dc) _
    )
End Function

Public Function rsAiPurch01fromAssy( _
    AiDoc As Inventor.Document _
) As ADODB.Recordset
    '''
    '''
    '''
    Set rsAiPurch01fromAssy _
    = rsFromGnsSql( _
        sqlSelAiPurch01fromAssy(AiDoc) _
    )
End Function

Public Function rsAiPdParts01fromAssy( _
    AiDoc As Inventor.Document _
) As ADODB.Recordset
    '''
    '''
    '''
    Set rsAiPdParts01fromAssy _
    = rsFromGnsSql( _
        sqlSelAiPdParts01fromAssy(AiDoc) _
    )
'Debug.Print rsAiPdParts01fromAssy(aiDocActive()).GetString(adClipString)
End Function

Public Function dcAiPurch01fromAdoRs( _
    rs As ADODB.Recordset _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim fdItem As ADODB.Field
    Dim fdType As ADODB.Field
    Dim fdFmly As ADODB.Field
    
    Set rt = New Scripting.Dictionary
    With rs
        If Not .BOF Then
            .Filter = ""
            
            With .Fields
                Set fdItem = .Item("Item")
                Set fdType = .Item("Type")
                Set fdFmly = .Item("Family")
            End With
            
            Do Until .EOF
                rt.Add fdItem.Value, Array( _
                    fdType.Value, fdFmly.Value _
                )
                .MoveNext
            Loop
            
            .Close
        End If
        Set dcAiPurch01fromAdoRs = rt
    End With
End Function

Public Function dcAiPurch01fromDict( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    '''
    '''
    Set dcAiPurch01fromDict _
    = dcAiPurch01fromAdoRs( _
        rsAiPurch01fromDict(dc) _
    )
End Function

Public Function dcAiPurch01fromAssy( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    '''
    '''
    '''
    Set dcAiPurch01fromAssy _
    = dcAiPurch01fromAdoRs( _
        rsAiPurch01fromAssy(AiDoc) _
    )
End Function

