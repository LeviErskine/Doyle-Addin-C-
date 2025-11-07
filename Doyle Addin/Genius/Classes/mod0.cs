

Public Function m0g3f1(rs As ADODB.Recordset) As Variant
    Dim rt As Variant
    Dim ar As Variant
    Dim rw As Variant
    Dim mx As Long
    Dim dx As Long
    Dim ct As Long
    Dim fd As Long
    
    With rs
        .Filter = .Filter
        If .BOF Or .EOF Then
            rt = Array(Array("<NODATA>"))
        Else
            ct = .Fields.Count - 1
            ar = Split(.GetString(adClipString, , vbTab, vbVerticalTab), vbVerticalTab)
            mx = UBound(ar) - 1 'Last row is empty/blank
            ReDim rt(mx, ct)
            For dx = 0 To mx
                rw = Split(ar(dx), vbTab)
                For fd = 0 To ct
                    rt(dx, fd) = rw(fd)
                Next
            Next
        End If
    End With
    
    m0g3f1 = rt
End Function

Public Function m0g3f2(AiDoc As Inventor.Document) As String
    Dim ar As Variant
    Dim dx As Long
    Dim ky As Variant
    
    With newFmTest1()
        .AskAbout AiDoc
        With .ItemData
            If .Count > 0 Then
                ar = .Keys
                For dx = 0 To UBound(ar) ' Each ky In ar
                    'Debug.Print ky, .Item(ky)
                    ar(dx) = ar(dx) & "=" & .Item(ar(dx))
                Next
            Else
                ar = Array("<NODATA>")
            End If
            m0g3f2 = Join(ar, vbNewLine)
        End With
    End With
End Function

Public Function m0g2f1(AiDoc As Inventor.Document) As Long
    m0g2f1 = newFmTest0().ft0g0f0(AiDoc.Thumbnail)
End Function

Public Function m0g2f2() As Long
    m0g2f2 = m0g2f1(ThisApplication.ActiveDocument)
End Function

Public Function m0g2f3() As Long
    Dim AiDoc As Inventor.Document
    'Dim dc As Scripting.Dictionary
    Dim ky As Variant
    
    'Set dc = New Scripting.Dictionary
    With New Scripting.Dictionary 'fmTest0
        For Each AiDoc In ThisApplication.Documents
            On Error Resume Next
            '.ft0g0f0 aiDoc.Thumbnail
            If .Exists(AiDoc.FullFileName) Then
                .Item(AiDoc.FullFileName) = 1 + .Item(AiDoc.FullFileName)
            Else
                .Add AiDoc.FullFileName, 1
            End If
            
            For Each ky In .Keys
                If .Item(ky) > 1 Then
                    Debug.Print .Item(ky), ky
                End If
            Next
            
            On Error GoTo 0
        Next
    End With
End Function

Public Function m0g1f0( _
    AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim tp As Inventor.DocumentTypeEnum
    Dim ky As String
    
    If dc Is Nothing Then
        Set rt = m0g1f0(AiDoc, _
        New Scripting.Dictionary)
    Else
        With AiDoc
            tp = .DocumentType
            ky = .FullFileName
        End With
        
        Set rt = dc
        If rt.Exists(ky) Then 'we've been here before
            'Don't do anything more, for now
            
        Else 'we've got a new document to process
            rt.Add ky, AiDoc
            
            If tp = kAssemblyDocumentObject Then
                Set rt = m0g1f0assy(AiDoc, dc)
            ElseIf tp = kPartDocumentObject Then
            Else
            End If
        End If
    End If
    
    Set m0g1f0 = rt
End Function
'Set dc = m0g1f0(ThisApplication.ActiveDocument): For Each ky In dc.Keys: Debug.Print aiDocument(dc.Item(ky)).PropertySets(gnDesign).Item(pnPartNum).Value, aiDocument(dc.Item(ky)).PropertySets(gnDesign).Item(pnMaterial).Value: Next

Public Function m0g1f0part( _
    AiDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If dc Is Nothing Then
        Set rt = m0g1f0part(AiDoc, _
            New Scripting.Dictionary _
        )
    Else
        Set rt = dc
    End If
    
    Set m0g1f0part = rt
End Function

Public Function m0g1f0assy( _
    AiDoc As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim aiOcc As Inventor.ComponentOccurrence
    
    If dc Is Nothing Then
        Set rt = m0g1f0assy(AiDoc, _
            New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        For Each aiOcc In AiDoc.ComponentDefinition.Occurrences
            Set rt = m0g1f0(aiOcc.Definition.Document, rt)
        Next
    End If
    
    Set m0g1f0assy = rt
End Function

Public Function dcAssyComp2A( _
    AiDoc As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim invOcc As Inventor.ComponentOccurrence
    Dim tp As Inventor.ObjectTypeEnum
    
    If dc Is Nothing Then
        Set rt = dcAssyComp2A(AiDoc, New Scripting.Dictionary)
    Else
        Set rt = dc
        For Each invOcc In AiDoc.ComponentDefinition.Occurrences
            With invOcc
                'Remove suppressed and excluded parts from the process
                'Moved out here from inner checks
                If .Visible And Not .Suppressed And Not .Excluded Then
                    tp = .Definition.Type
                    
                    If tp <> kAssemblyComponentDefinitionObject _
                    And tp <> kWeldmentComponentDefinitionObject _
                    Then
                        If tp <> kWeldsComponentDefinitionObject Then
                            Set rt = dcAddAiDoc(.Definition.Document, rt)
                        End If
                    Else 'assembly, check BOM Structure
                        If .BOMStructure = kPurchasedBOMStructure Then 'it's purchased
                            Set rt = dcAddAiDoc(.Definition.Document, rt)
                        ElseIf .BOMStructure = kNormalBOMStructure Then 'we make it
                            Set rt = dcAssyComp2A(.SubOccurrences, _
                                dcAddAiDoc(.Definition.Document, rt) _
                            ) 'NOT forgetting to add THIS document!
                        ElseIf .BOMStructure = kInseparableBOMStructure Then 'maybe weldment?
                            If tp = kWeldmentComponentDefinitionObject Then 'it is
                                Set rt = dcAssyComp2A(.SubOccurrences, _
                                    dcAddAiDoc(.Definition.Document, rt) _
                                )
                            Else 'it's not
                                Stop 'and see if we can figure out what its type is
                            End If
                        ElseIf .BOMStructure = kPhantomBOMStructure Then '"phantom" component
                            'Gather its components, but NOT the document itself
                            Set rt = dcAssyComp2A(.SubOccurrences, rt)
                        Else 'not sure what we've got
                            Stop 'and have a look at it
                        End If
                    End If 'part or assembly
                End If
            End With
        Next
    End If
    
    Set dcAssyComp2A = New Scripting.Dictionary
End Function

Public Function dcAssyCompStops( _
    Occurences As Inventor.ComponentOccurrences, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional dcStops As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    ''' Traverse the assembly,
    ''' including any/all subassemblies,
    ''' and collect all parts to be processed.
    Dim rt As Scripting.Dictionary
    Dim invOcc As Inventor.ComponentOccurrence
    Dim tp As Inventor.ObjectTypeEnum
    
    If dc Is Nothing Then
        Set rt = dcAssyCompStops(Occurences, New Scripting.Dictionary, dcStops)
    Else
        Set rt = dc
        For Each invOcc In Occurences
            With invOcc
                If dcStops.Exists(aiDocument(.Definition.Document _
                ).PropertySets.Item(gnDesign).Item(pnPartNum).Value _
                ) Then
                    'Stop
                End If
                
                'Remove suppressed and excluded parts from the process
                'Moved out here from inner checks
                If .Visible And Not .Suppressed And Not .Excluded Then
                    tp = .Definition.Type
                    
                    'MsgBox Join(Array( _
                        "TYPE: " & tp, _
                        "VISIBLE: " & .Visible, _
                        "NAME: " & .Name, _
                        "Suboccurence: " & .SubOccurrences.Count, _
                        "Occurence Type: " & .Definition.Occurrences.Type, _
                        "BOMStructure: " & .BOMStructure _
                    ), vbNewLine)
                    
                    If tp <> kAssemblyComponentDefinitionObject _
                    And tp <> kWeldmentComponentDefinitionObject _
                    Then
                        '(moved suppression/exclusion check OUTSIDE)
                        If tp <> kWeldsComponentDefinitionObject Then
                            
                            'Set rt = dcAddAiDoc(aiDocument(.Definition.Document), rt)
                            ''' Recasting by aiDocument not likely necessary here.
                            ''' Revised to following:
                            Set rt = dcAddAiDoc(.Definition.Document, rt)
                            
                        End If 'inVisible, suppressed, excluded or Welds
                        
                    Else 'assembly, check BOM Structure
                        If .BOMStructure = kPurchasedBOMStructure Then 'it's purchased
                            'Just add it to the Dictionary
                            Set rt = dcAddAiDoc(.Definition.Document, rt)
                        ElseIf .BOMStructure = kNormalBOMStructure Then 'we make it
                            'Gather its components
                            Set rt = dcAssyCompStops(.SubOccurrences, _
                                dcAddAiDoc(.Definition.Document, rt), _
                                dcStops _
                            ) 'NOT forgetting to add THIS document!
                        ElseIf .BOMStructure = kInseparableBOMStructure Then 'maybe weldment?
                            If tp = kWeldmentComponentDefinitionObject Then 'it is
                                'Treat it like an assembly
                                Set rt = dcAssyCompStops(.SubOccurrences, _
                                    dcAddAiDoc(.Definition.Document, rt), _
                                    dcStops _
                                )
                            Else 'it's not
                                Stop 'and see if we can figure out what its type is
                            End If
                        ElseIf .BOMStructure = kPhantomBOMStructure Then '"phantom" component
                            'Gather its components, but NOT the document itself
                            Set rt = dcAssyCompStops(.SubOccurrences, rt, dcStops)
                        Else 'not sure what we've got
                            Stop 'and have a look at it
                        End If
                    End If 'part or assembly
                End If
            End With
        Next
    End If
    Set dcAssyCompStops = rt
End Function

Public Function dcStopAt(tx As String, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If dc Is Nothing Then
        Set rt = dcStopAt(tx, New Scripting.Dictionary)
    Else
        Set rt = dc
    End If
    
    With rt
        If Not .Exists(tx) Then .Add tx, tx
    End With
    
    Set dcStopAt = rt
End Function

Public Function m0g0f0() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim tx As Variant
    
    Set rt = New Scripting.Dictionary
    For Each tx In Array( _
        "29-197A", "29-197B", "29-633", _
        "29-634", "29-637", "29-638", _
        "29-647", "29-648", "29-650", _
        "29-651", "29-652", "HD182" _
    )
        Set rt = dcStopAt((tx), rt)
    Next
    Set m0g0f0 = rt
End Function

Public Function m0g4f0(dcIn As Scripting.Dictionary) As Scripting.Dictionary
    ''  g4 currently reserved for development
    ''  relating to identification of purchased parts
    ''  by reference to Vault file path, as well as
    ''  BOM Structure and Cost Center settings
    ''
    ''  f0 separates members of incoming Dictionary
    ''  into design subgroups, including "doyle" and "purchased",
    ''  as indicated by the fourth component of each Vault path
    ''
    Dim rt As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    Dim ky As Variant
    Dim ls As Variant
    Dim mx As Long
    Dim dx As Long
    Dim g0 As String
    Dim g1 As String
    Dim fp As String
    Dim fn As String
    Dim rf As String
    
    Set rt = New Scripting.Dictionary
    With dcIn
        For Each ky In .Keys
            ls = Split(ky, "\")
            mx = UBound(ls)
            
            fn = ls(mx)
            g0 = ls(3)
            g1 = ls(4)
            fp = ""
            For dx = 5 To mx - 1
                fp = fp & "\" & ls(dx)
            Next
            fp = Mid(fp, 2)
            rf = Join(Array(fn, g1, fp), vbTab)
            
            With rt
                'Stop
                If .Exists(g0) Then
                    Set sd = .Item(g0)
                Else
                    Set sd = New Scripting.Dictionary
                    .Add g0, sd
                End If
                
                With sd
                    If .Exists(rf) Then
                        Stop
                    Else
                        .Add rf, dcIn.Item(ky)
                    End If
                End With
            End With
        Next
    End With
    Set m0g4f0 = rt
End Function

Public Function m0g4f1(dcIn As Scripting.Dictionary) As Scripting.Dictionary
    ''  g4 currently reserved for development
    ''  relating to identification of purchased parts
    ''  by reference to Vault file path, as well as
    ''  BOM Structure and Cost Center settings
    ''
    ''  f1 scans documents in the "purchased" design group
    ''  for unexpected settings, specifically:
    ''  BOMStructure should be
    ''      kPurchasedBOMStructure (51973 / 0xCB05)
    ''  Design Tracking Property "Cost Center"
    ''      should be either "D-PTS" or "R-PTS"
    ''
    ''  Any documents not matching these parameters
    ''  are dropped into one or both subDictionaries
    ''  in the returned Dictionary, according to issue
    ''
    Dim rt As Scripting.Dictionary
    Dim rtBom As Scripting.Dictionary
    Dim rtFam As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    Dim ivDoc As Inventor.Document
    Dim CpDef As Inventor.ComponentDefinition
    Dim prFam As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    Set rtBom = New Scripting.Dictionary
    Set rtFam = New Scripting.Dictionary
    rt.Add "bom", rtBom
    rt.Add "fam", rtFam
    
    Set sd = dcIn.Item("purchased")
    With sd
        For Each ky In .Keys
            Set ivDoc = .Item(ky)
            Set CpDef = aiCompDefOf(ivDoc)
            
            If CpDef Is Nothing Then
                Stop
            Else
                With CpDef
                    If .BOMStructure <> kPurchasedBOMStructure Then
                        rtBom.Add ky, ivDoc
                    End If
                    
                    Set prFam = aiDocument(.Document _
                        ).PropertySets.Item( _
                        gnDesign).Item(pnFamily _
                    )
                    
                    If prFam.Value <> "D-PTS" Or prFam.Value <> "R-PTS" Then
                        rtFam.Add ky, ivDoc
                    End If
                End With
            End If
        Next
    End With
    Set m0g4f1 = rt
End Function
'Debug.Print Join(m0g4f1(m0g4f0(dcAssyDocComponents(ThisApplication.ActiveDocument))).Item("fam").Keys, vbNewLine)

Public Function m0g4f1fixBom( _
    dcIn As Scripting.Dictionary, _
    Optional RiverView As Long = 0 _
) As Scripting.Dictionary
    ''  g4 currently reserved for development
    ''  relating to identification of purchased parts
    ''  by reference to Vault file path, as well as
    ''  BOM Structure and Cost Center settings
    ''
    ''  f1fixBom purports to fix incorrect BOM
    ''  Structure settings in purchased parts
    ''
    Dim rt As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    'Dim ivDoc As Inventor.Document
    Dim CpDef As Inventor.ComponentDefinition
    Dim ky As Variant
    
    Set sd = dcIn.Item("bom")
    Set rt = New Scripting.Dictionary
    With sd
        For Each ky In .Keys
            Set CpDef = aiCompDefOf(.Item(ky))
            If CpDef Is Nothing Then
                Stop
                rt.Add ky, 0
            Else
                With CpDef
                    .BOMStructure = kPurchasedBOMStructure
                    rt.Add ky, IIf(.BOMStructure _
                        = kPurchasedBOMStructure, _
                    1, 0)
                End With
            End If
        Next
    End With
    Set m0g4f1fixBom = rt
End Function

Public Function m0g4f1fixFam( _
    dcIn As Scripting.Dictionary, _
    Optional RiverView As Long = 0 _
) As Scripting.Dictionary
    ''  g4 currently reserved for development
    ''  relating to identification of purchased parts
    ''  by reference to Vault file path, as well as
    ''  BOM Structure and Cost Center settings
    ''
    ''  f1fixFam purports to fix incorrect
    ''  Family settings in purchased parts
    ''
    Dim rt As Scripting.Dictionary
    Dim sd As Scripting.Dictionary
    Dim ivDoc As Inventor.Document
    Dim CpDef As Inventor.ComponentDefinition
    Dim ky As Variant
    Dim txFam As String
    
    txFam = IIf(RiverView, "R", "D") & "-PTS"
    Set sd = dcIn.Item("fam")
    
    Set rt = New Scripting.Dictionary
    With sd
        For Each ky In .Keys
            Set ivDoc = .Item(ky)
            With ivDoc.PropertySets.Item(gnDesign).Item(pnFamily)
                .Value = txFam
                rt.Add ky, IIf(.Value = txFam, 1, 0)
            End With
            
            Set CpDef = aiCompDefOf(.Item(ky))
            If CpDef Is Nothing Then
                Stop
                rt.Add ky, 0
            Else
                With CpDef
                    .BOMStructure = kPurchasedBOMStructure
                    rt.Add ky, IIf(.BOMStructure _
                        = kPurchasedBOMStructure, _
                    1, 0)
                End With
            End If
        Next
    End With
    Set m0g4f1fixFam = rt
End Function
