

Public Function m1g0f0() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim gt0 As Long
    Dim eq0 As Long
    
    Set rt = New Scripting.Dictionary
    With dcAiSheetMetal(dcAiDocsByType( _
        dcAssyDocComponents(ThisApplication.ActiveDocument) _
    ))
        gt0 = 0
        eq0 = 0
        For Each ky In .Keys
            With aiDocPart(.Item(ky))
            With vcChkFlatPat(aiCompDefShtMetal( _
                .ComponentDefinition _
            ))
                If Abs(.X * .Y * .Z) > 0 Then
                    Debug.Print .X * .Y * .Z, ky
                    'Stop
                    gt0 = 1 + gt0
                Else
                    'Stop
                    eq0 = 1 + eq0
                End If
                'If .HasFlatPattern Then
                    'With .FlatPattern
                        'If .Features.Count > 0 Then
                            'gt0 = 1 + gt0
                            'rt.Add ky, .Document
                        'Else
                            'eq0 = 1 + eq0
                            ''Debug.Print .Width * .Length
                            ''Stop
                        'End If
                    'End With
                'Else
                    ''Debug.Print aiDocument(.Document).FullFileName
                    ''Stop
                'End If
            End With
            End With
        Next
        Debug.Print gt0, eq0
    End With
    Set m1g0f0 = rt
End Function

Public Function m1g0f1() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim gt0 As Long
    Dim eq0 As Long
    
    Set rt = New Scripting.Dictionary
    With dcAssyDocComponents( _
        ThisApplication.ActiveDocument, , 1 _
    )
        gt0 = 0
        eq0 = 0
        For Each ky In .Keys
            With aiDocument(.Item(ky))
                If .DocumentInterests.HasInterest( _
                    guidDesignAccl _
                ) Then
                    Stop
                End If
            End With
        Next
        Debug.Print gt0, eq0
    End With
    Set m1g0f1 = rt
End Function

Public Function vcDiaBox(bx As Inventor.Box) As Inventor.Vector
    ''  Given a Box Object
    ''  containing diametrically opposed
    ''  MinPoint and MaxPoint Objects,
    ''  return the Vector from Min to Max
    With bx
        Set vcDiaBox = .MinPoint.VectorTo(.MaxPoint)
    End With
End Function

Public Function vc2flatPat( _
    df As Inventor.SheetMetalComponentDefinition _
) As Inventor.Vector
    ''  From an Inventor Sheet Metal Component Definition,
    ''  obtain a Vector representing the translation
    ''  of the Folded Model's bounding box diagonal
    ''  to that of the Flat Pattern.
    ''
    Dim rt As Inventor.Vector
    
    With df
        Set rt = vcDiaBox(.RangeBox)
        If .HasFlatPattern Then
            rt.SubtractVector vcDiaBox(.FlatPattern.RangeBox)
        Else
            rt.SubtractVector rt
        End If
    End With
    ''
    ''  If they're the same point, this vector should
    ''  have zero length, however, this is NOT proof
    ''  positive of an invalid Flat Pattern. A valid
    ''  flat piece, with no folds, should produce
    ''  the same result.
    ''
    ''  A good follow-up check would probably be to
    ''  compare the Flat Pattern diagonal vector's
    ''  Z component to the model's Thickness
    
    Set vc2flatPat = rt
End Function
'Debug.Print vc2flatPat(aiCompDefShtMetal(aiDocPart(ThisApplication.Documents(2)).ComponentDefinition)).Length

Public Function vcCubicThickness( _
    df As Inventor.SheetMetalComponentDefinition _
) As Inventor.Vector
    Dim hk As Double
    
    hk = df.Thickness.Value
    With ThisApplication.TransientGeometry
        Set vcCubicThickness = .CreateVector(hk, hk, hk)
    End With
End Function

Public Function vcChkFlatPat( _
    df As Inventor.SheetMetalComponentDefinition _
) As Inventor.Vector
    ''  From an Inventor Sheet Metal Component Definition,
    ''  subtract a vector of cubic thickness from the
    ''  diagonal vector of either its Flat Pattern,
    ''  if available, or otherwise, the Folded Model.
    ''
    Dim rt As Inventor.Vector
    
    With df
        If .HasFlatPattern Then
            Set rt = vcDiaBox(.FlatPattern.RangeBox)
            'With .Thickness
                'rt.SubtractVector ThisApplication.TransientGeometry.CreateVector(.Value, .Value, .Value)
                'With ThisApplication.TransientGeometry.CreateVector(.Value, .Value, .Value)
                    '.SubtractVector rt
                    'If (.X * .Y * .Z) <> 0 Then
                        'Debug.Print .X * .Y * .Z
                        'Stop
                    'End If
                'End With
            'End With
            'rt.SubtractVector vcDiaBox(.RangeBox)
        Else
            Set rt = vcDiaBox(.RangeBox)
        End If
    End With
    rt.SubtractVector vcCubicThickness(df)
    ''
    ''  If the model is a valid sheet metal part,
    ''  one of the dimensions of its flat pattern's
    ''  bounding box diagonal should either equal
    ''  the defined sheet metal thickness,
    ''  or fall very close. At least, in theory.
    ''
    ''  Plan at this point is to try to determine
    ''  just how often this bears out.
    ''  While this HAS failed one pretest, the Flat
    ''  Pattern of that model includes features;
    ''  a relatively infrequent occurrence, and
    ''  quite possibly one that can throw off
    ''  the boundaries.
    
    Set vcChkFlatPat = rt
End Function
'Debug.Print vcChkFlatPat(aiCompDefShtMetal(aiDocPart(ThisApplication.Documents(2)).ComponentDefinition)).Length

Public Function m1tst0() As Variant
    Debug.Print iFacAssy(aiCompDefAssy(aiDocAssy(aiCompDefAssy(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition).Occurrences(1).Definition.Document).ComponentDefinition))
End Function

Public Function m1tst1() As Variant
    Dim ky As Variant
    Dim pr As Inventor.Property
    'Dim bm As Inventor.BOMStructureEnum
    
    For Each ky In Filter(dcAssyCompAndSub( _
        aiDocDefAssy(ThisApplication.ActiveDocument).Occurrences _
    ).Keys, "(DC)")
        With aiDocPart(ThisApplication.Documents.ItemByName((ky)))
            Set pr = .PropertySets("Design Tracking Properties").Item("Cost Center")
            With .ComponentDefinition
                If .BOMStructure <> kPurchasedBOMStructure Then
                    .BOMStructure = kPurchasedBOMStructure
                    pr.Value = "D-PTS"
                    Debug.Print pr.Value, .BOMStructure, ky
                End If
            End With
            '.ComponentDefinition
            '.ComponentDefinition.BOMStructure = kPurchasedBOMStructure) & "|" & ky
        End With
        'Debug.Print (aiDocPart(ThisApplication.Documents.ItemByName((ky))).ComponentDefinition.BOMStructure = kPurchasedBOMStructure) & "|" & ky
    Next
End Function

Public Function iFacAssy( _
    ob As Inventor.AssemblyComponentDefinition _
) As Inventor.iAssemblyFactory
    With ob
        If .iAssemblyFactory Is Nothing Then
            If .iAssemblyMember Is Nothing Then
                Set iFacAssy = Nothing
            Else
                Set iFacAssy = .iAssemblyMember.ParentFactory
            End If
        Else
            Set iFacAssy = .iAssemblyFactory
        End If
    End With
End Function

Public Function aiOccDoc( _
    ob As Inventor.ComponentOccurrence _
) As Inventor.Document
    Set aiOccDoc = ob.Definition.Document
End Function

Public Function aiDocDefAssy( _
    ob As Inventor.AssemblyDocument _
) As Inventor.AssemblyComponentDefinition
    Set aiDocDefAssy = ob.ComponentDefinition
End Function

Public Function dcAssyComponentsImmediate( _
    ob As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim oc As Inventor.ComponentOccurrence
    Dim cp As Inventor.Document
    Dim fn As String
    
    Set rt = New Scripting.Dictionary
    For Each oc In ob.ComponentDefinition.Occurrences
        'aiDocument(oc.Definition.Document).FullFileName
        Set cp = oc.Definition.Document
        fn = cp.FullFileName
        With rt
            If Not .Exists(fn) Then .Add fn, cp
        End With
    Next
    Set dcAssyComponentsImmediate = rt
End Function
'For Each pn In Array(pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnArea, pnLength, pnWidth, pnThickness): Debug.Print pn & "=" & aiDocument(ThisApplication.ActiveDocument).PropertySets.Item(gnCustom).Item((pn)).Value: Next

Public Function iSyncPartFactory(pd As Inventor.PartDocument) As Long
    ''
    ''  Backport iPart Member Properties to parent Factory
    ''
    Dim dcCols As Scripting.Dictionary
    Dim dcRows As Scripting.Dictionary
    Dim tr As Inventor.iPartTableRow
    Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    Dim rt As Long
    Dim ck As VbMsgBoxResult
    
    rt = 0
    With pd.ComponentDefinition
        Set ps = aiDocument(.Document).PropertySets.Item(gnCustom)
        If .iPartMember Is Nothing Then
            'Stop
        Else
            With .iPartMember
                Set dcRows = dcIPartTbRows(.ParentFactory.TableRows)
                If dcRows.Exists(pd.DisplayName) Then
                    Set dcCols = dcIPartTbCols(.ParentFactory.TableColumns)
                    
                    On Error Resume Next
                    Err.Clear
                    Set tr = .Row
                    
                    ''' REV[2022.03.23.1624]
                    '''     adding "pickup" attempt to capture
                    '''     and recover from errors encountered
                    '''     in trying to retrieve .Row directly
                    If Err.Number = 0 Then
                    Else
                        Set tr = dcRows.Item(pd.DisplayName)
                        If tr.MemberName = pd.DisplayName Then
                            Err.Clear
                        ElseIf tr.PartName = pd.DisplayName Then
                            Err.Clear
                        Else
                            Stop
                        End If
                    End If
                    
                    If Err.Number = 0 Then
                    For Each pr In ps
                        If dcCols.Exists(pr.Name) Then
                            Debug.Print pr.Name; "["; pr.Value; "]: "
                            With tr.Item(dcCols.Item(pr.Name))
                                Debug.Print "  "; .Value;
                                If pr.Value = .Value Then
                                    'Stop 'No change necessary
                                    Debug.Print " (NO CHANGE)"
                                Else
                                    On Error Resume Next
                                    .Value = pr.Value
                                    If Err.Number = 0 Then
                                        rt = 1 + rt
                                        
                                        '' The update invalidated the object
                                        '' We'll have to grab it again
                                        With tr.Item(dcCols.Item(pr.Name))
                                            Debug.Print " -> "; .Value
                                        End With
                                    Else
                                        Debug.Print " <!!ERROR!!> Couldn't Change"
                                        'Debug.Print Err.Number, Err.Description
                                    End If
                                    
                                    On Error GoTo 0
                                End If
                            End With
                        Else
                            'Stop
                        End If
                    Next
                    Debug.Print ; 'Breakpoint Landing
                    Else
                        Debug.Print "=== CAN'T SYNC IFACTORY ==="
                        Debug.Print "   Failed to access Row"
                        Debug.Print "   for Member " & pd.DisplayName
                        Debug.Print "   of Factory " & aiDocument(.ParentFactory.Parent).DisplayName
                        Debug.Print "=== PLEASE CHECK PARENT ==="
                        Debug.Print "====== FACTORY TABLE ======"
                        Debug.Print "Error 0x" & Hex$(Err.Number) & "("; CStr(Err.Number) & ")"
                        Debug.Print "    " & Err.Description
                        Debug.Print "==========================="
                        ck = MsgBox(Join(Array("" _
                            & "iPart Member " & pd.DisplayName _
                            , "in Factory" & aiDocument(.ParentFactory.Parent).DisplayName _
                            , "could not be accessed for updates." _
                            , "" _
                            , "Its Row might still be present" _
                            , "but somehow unavailable." _
                            , "" _
                            , "Please review iPart Factory." _
                        ), vbNewLine), vbOKCancel, _
                            "ERROR ACCESSING MEMBER ROW!" _
                        )
                        If ck = vbCancel Then
                            Stop
                        End If
                    End If
                    
                    On Error GoTo 0
                    Debug.Print ; 'Breakpoint Landing
                Else
                    Debug.Print "==== CAN'T FIND MEMBER ===="
                    Debug.Print "   Failed to locate Row"
                    Debug.Print "   for Member " & pd.DisplayName
                    Debug.Print "   of Factory " & aiDocument(.ParentFactory.Parent).DisplayName
                    Debug.Print "=== PLEASE CHECK PARENT ==="
                    Debug.Print "====== FACTORY TABLE ======"
                    'Debug.Print "Error 0x" & Hex$(Err.Number) & "("; CStr(Err.Number) & ")"
                    'Debug.Print "    " & Err.Description
                    'Debug.Print "==========================="
                    ck = MsgBox(Join(Array("" _
                        & "iPart Member " & pd.DisplayName _
                        , "could not be located in Factory" _
                        , aiDocument(.ParentFactory.Parent).DisplayName _
                        , "" _
                        , "Its Row might have been removed" _
                        , "or separated from the main table." _
                        , "" _
                        , "Please review iPart Factory Table." _
                    ), vbNewLine), vbOKCancel, _
                        "WARNING!! MEMBER NOT FOUND!" _
                    )
                    If ck = vbCancel Then
                        Stop
                    End If
                End If
            End With
        End If
    End With
End Function

Public Function iSyncAssyFactory(pd As Inventor.AssemblyDocument) As Long
    ''
    ''  Backport iPart Member Properties to parent Factory
    ''
    Dim dcCols As Scripting.Dictionary
    Dim dcRows As Scripting.Dictionary
    Dim tr As Inventor.iAssemblyTableRow
    Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    Dim rt As Long
    
    rt = 0
    With pd.ComponentDefinition
        Set ps = aiDocument(.Document).PropertySets.Item(gnCustom)
        If .iAssemblyMember Is Nothing Then
            'Stop
        Else
            With .iAssemblyMember
                Set dcRows = dcIAssyTbRows(.ParentFactory.TableRows)
                If dcRows.Exists(pd.DisplayName) Then
                    Set dcCols = dcIAssyTbCols(.ParentFactory.TableColumns)
                    Set tr = .Row
                    For Each pr In ps
                        If dcCols.Exists(pr.Name) Then
                            Debug.Print pr.Name; "["; pr.Value; "]: "
                            With tr.Item(dcCols.Item(pr.Name))
                                Debug.Print "  "; .Value;
                                If pr.Value = .Value Then
                                    'Stop 'No change necessary
                                    Debug.Print " (NO CHANGE)"
                                Else
                                    On Error Resume Next
                                    .Value = pr.Value
                                    If Err.Number = 0 Then
                                        rt = 1 + rt
                                        
                                        '' The update invalidated the object
                                        '' We'll have to grab it again
                                        With tr.Item(dcCols.Item(pr.Name))
                                            Debug.Print " -> "; .Value
                                        End With
                                    Else
                                        Debug.Print " <!!ERROR!!> Couldn't Change"
                                    End If
                                    
                                    On Error GoTo 0
                                End If
                            End With
                        Else
                            'Stop
                        End If
                    Next
                Else
                    Stop
                End If
            End With
        End If
    End With
End Function

Public Function dcColumnsIPart( _
    pd As Inventor.PartDocument _
) As Scripting.Dictionary
    ''  Retrieve Dictionary of iPart Factory Table Columns
    ''  If supplied Part Document is NOT an iPart Factory
    ''  OR Member, returned Dictionary will be empty
    ''
    With pd.ComponentDefinition
        If .iPartMember Is Nothing Then
            If .iPartFactory Is Nothing Then
                Set dcColumnsIPart = New Scripting.Dictionary
            Else
                Set dcColumnsIPart = dcIPartTbCols( _
                    .iPartFactory.TableColumns _
                )
            End If
        Else
            Set dcColumnsIPart = dcIPartTbCols( _
                .iPartMember.ParentFactory.TableColumns _
            )
        End If
    End With
End Function

Public Function dcColumnsIAssy( _
    pd As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    ''  Retrieve Dictionary of iAssembly Factory Table Columns
    ''  If supplied Assembly Document is NOT an iAssembly Factory
    ''  OR Member, returned Dictionary will be empty
    ''
    With pd.ComponentDefinition
        If .iAssemblyMember Is Nothing Then
            If .iAssemblyFactory Is Nothing Then
                Set dcColumnsIAssy = New Scripting.Dictionary
            Else
                Set dcColumnsIAssy = dcIAssyTbCols( _
                    .iAssemblyFactory.TableColumns _
                )
            End If
        Else
            Set dcColumnsIAssy = dcIAssyTbCols( _
                .iAssemblyMember.ParentFactory.TableColumns _
            )
        End If
    End With
End Function

Public Function dcIPartTbCols( _
    ls As Inventor.iPartTableColumns _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim it As Inventor.iPartTableColumn
    
    Set rt = New Scripting.Dictionary
    On Error Resume Next
    For Each it In ls
        rt.Add it.DisplayHeading, it.Index
        If Err.Number Then
            rt.Add it.FormattedHeading, it.Index
            Err.Clear
        End If
    Next
    On Error GoTo 0
    Set dcIPartTbCols = rt
End Function

Public Function dcIPartTbRows( _
    ls As Inventor.iPartTableRows _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim it As Inventor.iPartTableRow
    Dim ck As Long
    
    Set rt = New Scripting.Dictionary
    
    ''' REV[2022.08.05.1003]
    ''' added error trap code to more gracefully handle
    ''' errors accessing iPart/iAssembly factory table
    ''' (replicated changes also to dcIAssyTbRows)
    On Error Resume Next
    ck = ls.Count 'to trigger potential error
    
    If Err.Number = 0 Then 'should be good
        For Each it In ls
            ''' REV[2022.03.23.1618]
            '''     replacing Index of iPartTableRow
            '''     with iPartTableRow itself, so it
            '''     can just be pulled directly out
            '''     of the Dictionary by the client
            '''     process. if it needs the Index,
            '''     it can just get it itself, right?
            rt.Add it.MemberName, it '.Index
            rt.Add it.PartName, it '.Index
            Debug.Print ; 'debug landing
        Next
    Else
        Debug.Print "ERROR " & CStr(Err.Number) & " (" & Hex$(Err.Number) & ")" & vbNewLine & Err.Description
        Debug.Print Join(Array("Could not access Table Rows", "for member of iPart factory."), vbNewLine)
        'Stop
        Debug.Print ; 'Breakpoint Landing
    End If
    
    On Error GoTo 0
    
    Set dcIPartTbRows = rt
'Debug.Print aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows.Count
'Debug.Print txDumpLs(dcIPartTbRows(aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows).Keys)
End Function

Public Function dcIAssyTbCols( _
    ls As Inventor.iAssemblyTableColumns _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim it As Inventor.iAssemblyTableColumn
    
    Set rt = New Scripting.Dictionary
    On Error Resume Next
    For Each it In ls
        rt.Add it.DisplayHeading, it.Index
        If Err.Number Then
            rt.Add it.FormattedHeading, it.Index
            Err.Clear
        End If
    Next
    On Error GoTo 0
    Set dcIAssyTbCols = rt
End Function

Public Function dcIAssyTbRows( _
    ls As Inventor.iAssemblyTableRows _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim it As Inventor.iAssemblyTableRow
    Dim ck As Long
    
    Set rt = New Scripting.Dictionary
    
    ''' REV[2022.08.05.1003]
    ''' added error trap code to more gracefully handle
    ''' errors accessing iPart/iAssembly factory table
    ''' (replicated from dcIPartTbRows)
    On Error Resume Next
    ck = ls.Count 'to trigger potential error
    
    If Err.Number = 0 Then 'should be good
        For Each it In ls
            rt.Add it.MemberName, it.Index
            rt.Add it.DocumentName, it.Index
            Debug.Print ; 'debug landing
        Next
    Else
        Debug.Print "ERROR " & CStr(Err.Number) & " (" & Hex$(Err.Number) & ")" & vbNewLine & Err.Description
        Debug.Print Join(Array("Could not access Table Rows", "for member of iAssembly factory."), vbNewLine)
        Stop
    End If
    
    On Error GoTo 0
    
    Set dcIAssyTbRows = rt
'Debug.Print aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows.Count
'Debug.Print txDumpLs(dcIAssyTbRows(aiDocAssy(aiDocAssy(aiDocActive()).ComponentDefinition.Occurrences(1).Definition.Document).ComponentDefinition.iAssemblyMember.ParentFactory.TableRows).Keys)
End Function

Public Function m1g1f2(vr As Variant) As Inventor.iPartTableColumn
    Set m1g1f2 = vr
End Function

Public Function m1g1f3(ad As Inventor.AssemblyDocument) As Long
    Dim oc As Inventor.ComponentOccurrence
    Dim rt As Long
    
    rt = 0
    For Each oc In ad.ComponentDefinition.Occurrences '(1)
        With oc.Definition
            'Debug.Print Hex$(aiDocument(.Document).Type)
            'Stop
            If aiDocument(.Document).DocumentType = kPartDocumentObject Then
                rt = rt + iSyncPartFactory(aiDocPart(.Document))
            Else
            End If
        End With
    Next
    m1g1f3 = rt
End Function

Public Function m1g1f4() As Long
    m1g1f4 = m1g1f3(aiDocAssy(ThisApplication.ActiveDocument))
End Function

Public Function m1g1f5( _
    pd As Inventor.PartDocument _
) As Scripting.Dictionary
    ''  Retrieve Dictionary of Custom Properties
    ''
    Dim rt As Scripting.Dictionary
    Dim psMember As Inventor.PropertySet
    Dim psFactry As Inventor.PropertySet
    'Dim dcMember As Scripting.Dictionary
    Dim dcFactry As Scripting.Dictionary
    Dim pr As Inventor.Property
    
    Set rt = New Scripting.Dictionary
    With pd
        Set psMember = .PropertySets.Item(gnCustom)
        With .ComponentDefinition
            If Not .iPartMember Is Nothing Then
                With .iPartMember
                    With aiDocument(.ParentFactory.Parent)
                        Set psFactry = .PropertySets.Item(gnCustom)
                        Set dcFactry = dcAiPropsInSet(psFactry)
                    End With
                    
                    For Each pr In psMember
                        If Not dcFactry.Exists(pr.Name) Then
                            'rt.Add pr.Name, pr
                            Set rt = dcWithProp( _
                                psFactry, pr.Name, pr.Value, rt _
                            )
                        End If
                    Next
                End With
            End If
        End With
    End With
    Set m1g1f5 = rt
End Function

Public Function m1g1f5t0() As String
    m1g1f5t0 = Join(m1g1f5(aiDocPart( _
        aiDocAssy(ThisApplication.ActiveDocument _
        ).ComponentDefinition.Occurrences.Item(1 _
        ).Definition.Document _
    )).Keys, vbNewLine)
End Function

