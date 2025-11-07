

Public Function gnsPropSetAll() As Scripting.Dictionary
    Set gnsPropSetAll = nuDcPopulator( _
        ).Setting("Part Number", "Item" _
        ).Setting("Description", "Description1" _
        ).Setting("Cost Center", "Family" _
        ).Setting("GeniusMass", "Weight" _
        ).Setting("Thickness", "Thickness" _
        ).Setting("Extent_Area", "Diameter" _
        ).Setting("Extent_Width", "Width" _
        ).Setting("Extent_Length", "Length" _
        ).Setting("RM", "Stock" _
        ).Setting("RMQTY", "QuantityInConversionUnit" _
        ).Setting("RMUNIT", "ConversionUnit" _
        ).Setting("OFFTHK", "OFFTHK" _
    ).Dictionary
    ''' _
    '''
End Function

Public Function gnsPropSetItem() As Scripting.Dictionary
    Set gnsPropSetItem = nuDcPopulator( _
        ).Setting("Part Number", "Item" _
        ).Setting("Cost Center", "Family" _
        ).Setting("GeniusMass", "Weight" _
        ).Setting("Thickness", "Thickness" _
        ).Setting("Extent_Area", "Diameter" _
        ).Setting("Extent_Width", "Width" _
        ).Setting("Extent_Length", "Length" _
        ).Setting("OFFTHK", "OFFTHK" _
    ).Dictionary
    ''' _
        ).Setting("Description", "Description1" _
        ).Setting("RM", "" _
        ).Setting("RMQTY", "" _
        ).Setting("RMUNIT", "" _
    '''
End Function

Public Function gnsPropSetBomRaw() As Scripting.Dictionary
    '''
    ''' gnsPropSetBomRaw -- Property Names for Genius BOM
    ''' !!!NOT READY!!! Just dup'd from gnsPropSetItem
    ''' Needs adjustment to BOM Column/Field names
    '''
    Set gnsPropSetBomRaw = nuDcPopulator( _
        ).Setting("Part Number", "Product" _
        ).Setting("RM", "Item" _
        ).Setting("RMQTY", "QuantityInConversionUnit" _
        ).Setting("RMUNIT", "ConversionUnit" _
    ).Dictionary
    ''' _
        ).Setting("Cost Center", "Family" _
        ).Setting("GeniusMass", "Weight" _
        ).Setting("Thickness", "Thickness" _
        ).Setting("Extent_Area", "Diameter" _
        ).Setting("Extent_Width", "Width" _
        ).Setting("Extent_Length", "Length" _
        ).Setting("OFFTHK", "OFFTHK" _
        ).Setting("Description", "Description1" _
    '''
End Function

Public Function gnsPropsCurrent( _
    Optional AiDoc As Inventor.Document = Nothing, _
    Optional dcProps As Scripting.Dictionary = Nothing, _
    Optional incTop As Long = 0, _
    Optional inclPhantom As Long = 0 _
) As Scripting.Dictionary
    'Dim rf As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    'Dim ActiveDoc As Document
    Dim ky As Variant
    'Dim dx As Long
    
    If AiDoc Is Nothing Then
        Set gnsPropsCurrent = gnsPropsCurrent( _
            aiDocActive(), dcProps, incTop, inclPhantom _
        )
    ElseIf dcProps Is Nothing Then
        Set gnsPropsCurrent = gnsPropsCurrent( _
            AiDoc, gnsPropSetAll(), incTop, inclPhantom _
        ) 'gnsPropSetItem
    Else
        'Set rf = dcProps 'gnsPropSetItem() 'dcOfIdent(Array( _
            "Part Number", "Cost Center", _
            "GeniusMass", "Extent_Area", _
            "Extent_Width", "Extent_Length", _
            "RM", "RMQTY", "RMUNIT", _
            "OFFTHK" _
        ))
        
        ''' Collect Components for Processing
        ''' (retained from Sub UpdateGeniusProperties_2023_0406_pre)
        ''' NOTE[2021.08.09]:
        '''     Function dcRemapByPtNum previously
        '''     revised to address Key collisions
        '''     in a crude manner. See that function
        '''     for details.
        '''
        Set rt = dcRemapByPtNum( _
            dcAiDocComponents( _
                AiDoc, , incTop _
            ) _
        )
        'NOTE: incTop here is used to indicate
        '   whether to include top level assembly.
        '   This decision probably still needs
        '   to be made, but is not really
        '   addressed at the moment.
        
        ''' Retrieve the full Component Collection
        With rt
            For Each ky In .Keys
                ''' Get Genius Properties for Item
                ''' Probably DON'T want to replace
                ''' Property Objects with their
                ''' values at this point, so that
                ''' they can be updated without
                ''' having to retrieve them again.
                Set .Item(ky) = dcKeysInCommon( _
                    dcOfPropsInAiDoc(.Item(ky)) _
                , dcProps, 1) 'dcProps replaces rf
            Next
            
            Debug.Print ; 'Breakpoint Landing
        End With
        
        Set gnsPropsCurrent = rt
        ''' Don't index here.
        ''' Let the client worry about that.
    End If
End Function

Public Sub testGnsPropsCurrent()
    Dim md As Inventor.Document
    Dim mPrpGnsAll As Scripting.Dictionary
    Dim mPrpGnsItm As Scripting.Dictionary
    Dim mPrpBomRaw As Scripting.Dictionary
    Dim dcPr As Scripting.Dictionary
    Dim dcVl As Scripting.Dictionary
    Dim dcGn As Scripting.Dictionary
    Dim dcGnDx As Scripting.Dictionary
    Dim dcBomDx As Scripting.Dictionary
    Dim nd As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim goAhead As VbMsgBoxResult
    Dim txOut As String
    Dim ky As Variant
    
    Set md = aiDocActive()
    Set mPrpGnsAll = gnsPropSetAll() 'gnsPropSetItem
    'Set mPrpBomRaw = gnsPropSetBomRaw()
    
    ''' Retrieve all Documents'
    ''' Genius Property Objects...
    Set dcPr = gnsPropsCurrent(md, mPrpGnsAll)
    
    ''' ...and their current Values
    Set dcVl = New Scripting.Dictionary
    With dcPr: For Each ky In .Keys
        'Set .Item(ky) = dcPropVals(dcOb(.Item(ky)))
        dcVl.Add ky, dcPropVals(dcOb(.Item(ky)))
    Next: End With
    Debug.Print ; 'Breakpoint Landing
    If False Then
        send2clipBdWin10 ConvertToJson(dcVl, vbTab): Stop
    End If
    
    Set dcGnDx = dcRecSetDcDx4json(dcDxFromRecSetDc( _
        dcFromAdoRS(cnGnsDoyle().Execute(q1g1x2(md))) _
    ))
    With dcGnDx
        Set dcGn = dcOb(.Item("Item"))
        With dcOb(.Item(""))
            For Each ky In dcGn.Keys
                Set dcGn.Item(ky) = .Item(dcGn.Item(ky)(0))
            Next
        End With
    End With
    Debug.Print ; 'Breakpoint Landing
    If False Then
        send2clipBdWin10 ConvertToJson(dcGn, vbTab): Stop
    End If
    
    ''' This is to extract BOM from Assembly
    'bomInfoBkDn(bomViewStruct(aiDocAssy(
    
    Set dcBomDx = dcRecSetDcDx4json(dcDxFromRecSetDc( _
        dcFromAdoRS(cnGnsDoyle().Execute(q1g2x2(md))) _
    ))
    Debug.Print ; 'Breakpoint Landing
    If False Then
        send2clipBdWin10 ConvertToJson(dcBomDx, vbTab): Stop
    End If
    
    Set nd = New Scripting.Dictionary
    With dcPr: For Each ky In .Keys
        Set ck = dcKeysMissing(mPrpGnsAll, dcOb(.Item(ky)))
        If ck.Count > 0 Then nd.Add ky, ck
    Next: End With
    'Debug.Print ConvertToJson(nd, vbTab)
    Debug.Print ; 'Breakpoint Landing
    
    ''' Index the Dictionary here
    ''' (might be temporary)
    'Set dcPr = dcRecSetDcDx4json( _
        dcDxFromRecSetDc(dcPr) _
    )
    
    ''' Dump to JSON text format
    txOut = ConvertToJson( _
        dcRecSetDcDx4json( _
        dcDxFromRecSetDc(dcVl _
    )), vbTab)  'dcPr
    'Debug.Print txOut
    
    goAhead = MsgBox( _
        Join(Array( _
            "Assembly Name:", _
            md.DisplayName, _
            "Process Completed", _
            "", _
            "Copy report text", _
            "(JSON format)", _
            "to Clipboard?", _
            "", _
            "(Cancel for Debug)" _
        ), vbNewLine _
        ), vbYesNoCancel, "Update Complete" _
    )
    If goAhead = vbCancel Then
        Stop
    ElseIf goAhead = vbYes Then
        send2clipBdWin10 txOut
    End If
End Sub

Public Sub ExposeAllSheetMetalThicknesses()
    Dim pt As Inventor.PartDocument
    Dim tk As Inventor.Parameter
    Dim ky As Variant
    Dim ct As Long
    Dim xp As Long
    Dim nc As Long
    
    With dcAiSheetMetal(dcAiPartDocs(dcAiDocComponents( _
        ThisApplication.ActiveDocument _
    )))
        ct = 0
        xp = 0
        nc = 0
        For Each ky In .Keys
            With aiCompDefShtMetal(aiDocPart(.Item(ky)).ComponentDefinition)
                If .BOMStructure = kNormalBOMStructure Then
                    ct = 1 + ct
                    
                    On Error Resume Next
                    Err.Clear
                    
                    ''' REV[2023.01.18.1626]
                    ''' added check for iPart member
                    ''' to check exposure of Thickness
                    ''' parameter of its Parent Factory
                    ''' rather than its own.
                    '''
                    ''' this seeks to avoid errors
                    ''' that now seem to arise when
                    ''' attempting to set exposure
                    ''' on the members themselves.
                    '''
                    If .IsiPartMember Then
                        'Stop
                        With aiCompDefShtMetal(aiDocPart( _
                        .iPartMember.ParentFactory.Parent))
                            Set tk = .Thickness
                        End With
                    Else
                        Set tk = .Thickness
                    End If
                    
                    If Err.Number Then
                        Debug.Print "!ERROR!: " & ky
                        Set tk = Nothing
                        nc = 1 + nc
                        Err.Clear
                    Else
                        With tk
                            If .ExposedAsProperty Then
                                Debug.Print "NOCHNGE: " & ky
                            ElseIf aiCompDefShtMetal( _
                                .Parent.Parent _
                            ).IsiPartMember Then
                                'Stop
                                If aiDocPart( _
                                    aiCompDefShtMetal(.Parent.Parent _
                                    ).iPartMember.ParentFactory.Parent _
                                ).ComponentDefinition.Parameters.Item( _
                                    pnThickness _
                                ).ExposedAsProperty Then
                                    Debug.Print "NOCHNGE: " & ky
                                Else
                                    nc = 1 + nc
                                    Debug.Print "FAILED!: " & ky
                                End If
                            Else
                                .ExposedAsProperty = True
                                If .ExposedAsProperty Then
                                    xp = 1 + xp
                                    Debug.Print "EXPOSED: " & ky
                                Else
                                    nc = 1 + nc
                                    Debug.Print "FAILED!: " & ky
                                End If
                            End If
                        End With
                    End If
                    On Error GoTo 0
                End If
            End With
        Next
        If xp + nc > 0 Then
            MsgBox Join(Array( _
                "Found " & CStr(ct) & " components.", _
                "Thickness already exposed on " & CStr(ct - xp - nc), _
                "   Exposed additional " & CStr(xp), _
                "   Failed to expose " & CStr(nc) _
            ), vbNewLine), vbOKOnly, "Sheet Metal Processed"
        Else
            MsgBox Join(Array( _
                "Thickness already exposed", _
                "on " & CStr(ct) & " components." _
            ), vbNewLine), vbOKOnly, "No Change Required"
        End If
    End With
End Sub

Public Sub AddProps4Genius()
    Dim ky As Variant
    Dim pr As Inventor.Property
    
    With dcProps4genius(ThisApplication.ActiveDocument)
        For Each ky In .Keys
            With aiProperty(obOf(.Item(ky)))
                Debug.Print .Parent.Name & ":" & .Name & "=" & .Value
            End With
        Next
    End With
End Sub
'For Each itm In ActiveDocsComponents(ThisApplication): Debug.Print aiDocument(obOf(itm)).FullFileName: Next

Public Sub MakeViewImageFiles()
    Debug.Print d0g9f0()
End Sub

Sub iParts_GenerateAll()
    Dim oDoc As Inventor.PartDocument
    Dim oFactory As Inventor.iPartFactory
    Dim sFile As String
    Dim iCount As Long
    Dim i As Long
    Dim mx As Long
    Dim dx As Long
    Dim bk As Long
    
    'Set oDoc = AskUser4aiDoc(, dcOf_iAll_Factories( _
        ThisApplication.Documents.VisibleDocuments _
    ))
    
    Set oDoc = AskUser4aiDoc( _
    , dcOf_iPartFactories())
    
    If oDoc Is Nothing Then
        'do nothing
    ElseIf oDoc.ComponentDefinition.IsiPartFactory = True Then
        sFile = oDoc.FullFileName
        
        Set oFactory = oDoc.ComponentDefinition.iPartFactory
        
        'With oFactory
        mx = oFactory.TableRows.Count
        dx = 1
        Do 'For i = 1 To mx
            bk = 1 + mx - dx
            If bk > 10 Then bk = 10
            iCount = 0
            Do
                ThisApplication.StatusBarText _
                    = CStr(dx) & "/" & CStr(mx) & ": " _
                    & oFactory.TableRows.Item(dx).MemberName
                'Member File creation
                '.CreateMember dx
                'disabled for testing
                MsgBox oFactory.TableRows.Item(dx).MemberName, _
                    vbOKOnly, "Member " & CStr(dx) & "/" & CStr(mx)
                
                dx = dx + 1
                iCount = iCount + 1
            Loop While iCount < bk
            
            If dx <= mx Then 'iCount = 10
                oDoc.Close
                Set oDoc = ThisApplication.Documents.Open(sFile)
                Set oFactory = oDoc.ComponentDefinition.iPartFactory
                iCount = 0
            End If
        Loop Until dx > mx 'Next
        'End With
    End If
End Sub
