

Public Function aiDocPartNum( _
    AiDoc As Inventor.Document, _
    Optional ifNone As String = "" _
) As String
    If AiDoc Is Nothing Then
        aiDocPartNum = ifNone
    Else
        aiDocPartNum = AiDoc.PropertySets _
        .Item(gnDesign) _
        .Item(pnPartNum) _
        .Value
    End If
End Function

Public Function dc0g1f0( _
    AiDoc As Inventor.Document, _
    Optional prName As String = "", _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pn As String
    Dim ky As String
    
    If dc Is Nothing Then
        Set dc0g1f0 = dc0g1f0( _
            AiDoc, prName, _
            New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        
        With AiDoc
            pn = .PropertySets.Item(gnDesign).Item(pnPartNum).Value
            ky = prName & vbTab & pn
            With rt
                If .Exists(ky) Then
                    If .Item(ky) Is AiDoc Then
                    Else
                        'Stop
                    End If
                Else
                    .Add ky, AiDoc
                End If
            End With
            
            If .DocumentType = kAssemblyDocumentObject Then
                Set rt = dc0g1f1(AiDoc, rt)
            ElseIf .DocumentType <> kPartDocumentObject Then
                Stop
            Else
            End If
        End With
        
        Set dc0g1f0 = rt
    End If
End Function

Public Function dc0g1f1( _
    AiDoc As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim aiOcc As Inventor.ComponentOccurrence
    Dim pn As String
    
    If dc Is Nothing Then
        Set dc0g1f1 = dc0g1f1(AiDoc, _
        New Scripting.Dictionary)
    Else
        Set rt = dc
        
        With AiDoc
            pn = .PropertySets.Item(gnDesign).Item(pnPartNum).Value
            With .ComponentDefinition
                For Each aiOcc In .Occurrences
                    Set rt = dc0g1f0(aiOcc.Definition.Document, pn, rt)
                Next
            End With
        End With
        
        Set dc0g1f1 = rt
    End If
End Function

Public Function dc0g2f0( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If AiDoc Is Nothing Then
        Set dc0g2f0 = dc0g2f0( _
            ThisApplication.ActiveDocument _
        )
    Else
        Set wk = dcAiDocsByPtNum( _
            dcAiDocComponents(AiDoc) _
        ) 'dcAiDocPartNumbers
        Set rt = New Scripting.Dictionary
        With wk
            For Each ky In .Keys
                Set rt = dc0g2f2(aiDocument(obOf(.Item(ky))), rt)
            Next
        End With
        Set dc0g2f0 = rt
    End If
End Function

Public Function dc0g2f1( _
    AiDoc As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    'Dim rt As Scripting.Dictionary
    Dim aiOcc As Inventor.ComponentOccurrence
    Dim prName As String
    Dim ptName As String
    Dim ky As String
    Dim ct As Long
    
    If dc Is Nothing Then
        Set dc0g2f1 = dc0g2f1(AiDoc, _
        New Scripting.Dictionary)
    Else
        prName = aiDocPartNum(AiDoc)
        With AiDoc.ComponentDefinition
            For Each aiOcc In .Occurrences
                ptName = aiDocPartNum( _
                    aiOcc.Definition.Document _
                )
                ky = prName & vbTab & ptName
                
                With dc
                    If .Exists(ky) Then
                        ct = .Item(ky)
                        .Item(ky) = 1 + ct
                    Else
                        .Add ky, 1
                    End If
                End With
            Next
        End With
        Set dc0g2f1 = dc
    End If
End Function

Public Function dc0g2f2(AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    If AiDoc.DocumentType = kAssemblyDocumentObject Then
        Set dc0g2f2 = dc0g2f1(AiDoc, dc)
    Else
        Set dc0g2f2 = dc
    End If
End Function

Public Function dc0g3f0( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' (just so we don't forget what this is for)
    ''' This function accepts a Dictionary
    ''' of Inventor Documents, and generates
    ''' a new Dictionary of Dictionaries,
    ''' keyed on Genius Family names, and
    ''' containing all Documents in its Family,
    ''' themselves keyed on Part Number.
    '''
    ''' Function dc0g3f1 makes use of this below.
    '''
    Dim rt As Scripting.Dictionary
    Dim fm As Scripting.Dictionary
    Dim ky As Variant
    Dim ad As Inventor.Document
    Dim nm As String
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set ad = aiDocument(obOf(.Item(ky)))
            If ad Is Nothing Then
                'Stop
            Else
                With ad.PropertySets.Item(gnDesign)
                    nm = .Item(pnFamily).Value
                    pn = .Item(pnPartNum).Value
                End With
                
                With rt
                    If .Exists(nm) Then
                        Set fm = .Item(nm)
                    Else
                        Set fm = New Scripting.Dictionary
                        .Add nm, fm
                    End If
                End With
                
                With fm
                    If .Exists(pn) Then
                        Stop
                    Else
                        .Add pn, ad
                    End If
                End With
            End If
        Next
    End With
    
    Set dc0g3f0 = rt
End Function

Public Function dc0g3f1() As Scripting.Dictionary
    '''
    ''' This function calls dc0g3f0 above
    ''' against a Dictionary of Inventor Documents
    ''' generated from the components of the active
    ''' Inventor Document. It then dumps a list of
    ''' the Genius Family names encountered, and if
    ''' any were blank, the list of part numbers
    ''' in the blank Family group is also revealed.
    '''
    With dc0g3f0(dcAssyDocComponents(aiDocAssy(aiDocActive())))
        Debug.Print txDumpLs(.Keys)
        Stop
        If .Exists("") Then
            Debug.Print "NO FAMILY"
            With dcOb(.Item(""))
                Debug.Print txDumpLs(.Keys)
            End With
            Stop
        Else
            Stop
        End If
    End With
End Function

Public Function dc0g4f0(AiDoc As Inventor.AssemblyDocument) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim ki As Variant
    
    Set rt = New Scripting.Dictionary
    With dcAssyComponentsImmediate(AiDoc)
        For Each ky In .Keys
            With dcAssyComponentsImmediate(aiDocAssy(.Item(ky)))
                For Each ki In .Keys
                    If Not rt.Exists(ki) Then
                        rt.Add ki, .Item(ki)
                    End If
                Next
            End With
        Next
    End With
    Set dc0g4f0 = rt
End Function
'Debug.Print txDumpLs(dcAssyComponentsImmediate(ThisApplication.ActiveDocument).Keys)
'Debug.Print txDumpLs(dc0g4f0(ThisApplication.ActiveDocument).Keys)
'Debug.Print txDumpLs(dcAiDocPartNumbers(dc0g4f0(ThisApplication.ActiveDocument)).Keys)

Public Function dc0g4f1( _
    adIn As Inventor.AssemblyDocument, _
    adOut As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    Dim ky As Variant
    Dim tg As Inventor.ComponentOccurrences
    Dim oc As Inventor.ComponentOccurrence
    Dim ps As Inventor.Matrix
    
    Set ps = ThisApplication.TransientGeometry.CreateMatrix()
    
    Set tg = adOut.ComponentDefinition.Occurrences
    With dc0g4f0(adIn)
        For Each ky In .Keys
            Set oc = tg.Add(ky, ps)
        Next
    End With
End Function

Public Function dcBoxDims(bx As Inventor.Box) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim mx As Inventor.Point
    Dim mn As Inventor.Point
    
    Set rt = New Scripting.Dictionary
    
    With bx
        Set mx = .MaxPoint
        Set mn = .MinPoint
    End With
    
    With rt
        .Add "X", (mx.X - mn.X) '/ 2.54
        .Add "Y", (mx.Y - mn.Y) '/ 2.54
        .Add "Z", (mx.Z - mn.Z) '/ 2.54
    End With
    
    Set dcBoxDims = rt
End Function

Public Function dcBoxDimsCm2in(bx As Inventor.Box) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dcBoxDims(bx)
        For Each ky In .Keys
            rt.Add ky, CDbl(.Item(ky)) / cvLenIn2cm
        Next
    End With
    
    Set dcBoxDimsCm2in = rt
End Function

Public Function dcAiPropsInSet( _
    ps As Inventor.PropertySet _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    
    Set rt = New Scripting.Dictionary
    For Each pr In ps
        If rt.Exists(pr.Name) Then
            Stop
        Else
            rt.Add pr.Name, pr
        End If
    Next
    Set dcAiPropsInSet = rt
End Function
'Debug.Print Join(Filter(dcAiPropsInSet(ThisApplication.ActiveDocument.PropertySets(gnDesign)).Keys, "web", , vbTextCompare), vbNewLine)

Public Function dcAiDocParVals( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Set dcAiDocParVals _
    = dcAiParValues( _
    dcAiDocParams(AiDoc))
End Function

Public Function dcAiParValues( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim pr As Inventor.Parameter
    
    Set rt = New Scripting.Dictionary
    
    If dc Is Nothing Then
    Else
        With dc: For Each ky In .Keys
            Set pr = obAiParam(obOf(.Item(ky)))
            If pr Is Nothing Then
            Else
                rt.Add ky, Array(pr.Value, pr.Units)
            End If
        Next: End With
    End If
    
    Set dcAiParValues = rt
End Function

Public Function dcAiDocParams( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Set dcAiDocParams _
    = dcCompDefParams( _
    compDefOf(AiDoc))
End Function
'Debug.Print Join(Filter(dcAiDocParams(ThisApplication.ActiveDocument.PropertySets(gnDesign)).Keys, "web", , vbTextCompare), vbNewLine)

Public Function dcCompDefParams( _
    CpDef As Inventor.ComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    If CpDef Is Nothing Then
        Set dcCompDefParams = New Scripting.Dictionary
    ElseIf TypeOf CpDef Is Inventor.AssemblyComponentDefinition Then
        Set dcCompDefParams = dcCompDefParamsAssy(CpDef)
    ElseIf TypeOf CpDef Is Inventor.PartComponentDefinition Then
        Set dcCompDefParams = dcCompDefParamsPart(CpDef)
    Else
        Set dcCompDefParams = dcCompDefParams(Nothing)
    End If
End Function

Public Function dcCompDefParamsPart( _
    CpDef As Inventor.PartComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim pr As Inventor.Parameters
    
    If CpDef Is Nothing Then
        Set pr = Nothing
    Else
        Set pr = CpDef.Parameters
    End If
    
    Set dcCompDefParamsPart _
    = dcOfAiParameters(pr, dc)
End Function

Public Function dcCompDefParamsAssy( _
    CpDef As Inventor.AssemblyComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim pr As Inventor.Parameters
    
    If CpDef Is Nothing Then
        Set pr = Nothing
    Else
        Set pr = CpDef.Parameters
    End If
    
    Set dcCompDefParamsAssy _
    = dcOfAiParameters(pr, dc)
End Function

Public Function dcOfAiParameters( _
    AiPars As Inventor.Parameters, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Parameter
    
    If dc Is Nothing Then
        Set rt = dcOfAiParameters(AiPars, _
        New Scripting.Dictionary)
    Else
        Set rt = dc
        
        If AiPars Is Nothing Then
        Else
            For Each pr In AiPars
                rt.Add pr.Name, pr
            Next
        End If
    End If
    
    Set dcOfAiParameters = rt
End Function

Public Function dcOfPropsInAiDoc( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ps As Inventor.PropertySet
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    If AiDoc Is Nothing Then
    Else
        With AiDoc: For Each ps In .PropertySets
            Set wk = dcAiPropsInSet(ps)
            
            With dcKeysMissing(wk, rt)
            For Each ky In .Keys
                rt.Add ky, .Item(ky)
                wk.Remove ky
            Next: End With
            
            With wk 'dcKeysInCommon(wk, rt)
                If .Count > 0 Then
                    Debug.Print "=== DUPLICATE PROPERTY NAMES ==="
                    Debug.Print "  Item " & aiProperty(rt.Item(pnPartNum)).Value & " (" & AiDoc.FullDocumentName & ")"
                    Debug.Print dumpLsKeyVal(dcPropVals(wk), ": ")
                    Debug.Print "--- previously found"
                    Debug.Print dumpLsKeyVal(dcPropVals(dcKeysInCommon(wk, rt, 2)), ": ")
                    Debug.Print ; 'Breakpoint Landing
                End If
            End With
        Next: End With
    End If
    
    Set dcOfPropsInAiDoc = rt
End Function

Public Function dcAiPropValsFromDc( _
    dc As Scripting.Dictionary, _
    Optional Flags As Long = 0 _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcNewIfNone(dc): For Each ky In .Keys
        Set pr = aiProperty(obOf(.Item(ky)))
        If pr Is Nothing Then
            If Flags And 1 Then
                'Keep non-Property Items
                rt.Add ky, .Item(ky)
            End If
        Else
            rt.Add ky, aiPropVal(pr, Empty)
        End If
    Next: End With
    
    Debug.Print ; 'Breakpoint Landing
    Set dcAiPropValsFromDc = rt
End Function

Public Function dcForAiDocIType( _
    dc As Scripting.Dictionary, _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ky As String
    
    If TypeOf AiDoc Is Inventor.PartDocument Then
        ky = "Part"
        With aiDocPart(AiDoc).ComponentDefinition
            If .IsContentMember Then ky = "c" & ky
            If .IsiPartFactory Then ky = "f" & ky
            If .IsiPartMember Then ky = "i" & ky
            If .IsModelStateFactory Then 'ky = "msf" & ky
                ''' ???
            End If
            If .IsModelStateMember Then ky = "s" & ky
        End With
    ElseIf TypeOf AiDoc Is Inventor.AssemblyDocument Then
        ky = "Assy"
        With aiDocAssy(AiDoc).ComponentDefinition
            If .IsiAssemblyFactory Then ky = "f" & ky
            If .IsiAssemblyMember Then ky = "i" & ky
            If .IsModelStateFactory Then ky = "msf" & ky
            If .IsModelStateMember Then ky = "s" & ky
        End With
    Else
        ky = ""
    End If
    
    With dc
        If Not .Exists(ky) Then
        .Add ky, New Scripting.Dictionary
        End If
        Set dcForAiDocIType = .Item(ky)
    End With
End Function

Public Function dcAiDocsByIType( _
    dc As Scripting.Dictionary, _
    Optional Flags As Long = 0 _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    'Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcNewIfNone(dc): For Each ky In .Keys
        'Set pr = aiProperty(obOf(.Item(ky)))
        'If pr Is Nothing Then
            'If Flags And 1 Then
                'Keep non-Property Items
                rt.Add ky, .Item(ky)
            'End If
        'Else
            'rt.Add ky, aiPropVal(pr, Empty)
        'End If
    Next: End With
    
    Debug.Print ; 'Breakpoint Landing
    Set dcAiDocsByIType = rt
End Function

Public Function nvmTest01() As Variant
    Dim ad As Inventor.ApplicationAddIn
    Dim il As Object 'Inventor.ApplicationAddIn '
    Dim mp As Inventor.NameValueMap
    Dim md As Inventor.Document
    
    Set ad = ThisApplication.ApplicationAddIns.ItemById(guidILogicAdIn)
    If Not ad.Activated Then ad.Activate
    Set il = ad.Automation
    Set md = ThisApplication.Documents.ItemByName("C:\Doyle_Vault\Designs\Misc\andrewT\dvl\iLogVltSrch_2022-0622_01.ipt")
    Set mp = dc2aiNameValMap( _
        nuDcPopulator().Setting( _
        "PartNumber", "60-" _
    ).Dictionary)  'IN 60- 04-
    
    il.RunRuleWithArguments md, "vlt02", mp 'tst01
    'il.RunRule md, "tst01" ', mp
    
    Debug.Print mp.Value("OUT")
    
    Debug.Print mp.Count
End Function

'''
'''
'''
Public Function dvlDict0() As String
    dvlDict0 = "dvlDict0"
End Function
