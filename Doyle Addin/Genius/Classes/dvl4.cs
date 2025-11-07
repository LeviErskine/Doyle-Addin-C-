

Public Function d4g2f1( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Scripting.Dictionary
    '''
    ''' d4g2f1 -- not sure what
    '''     but looks like something to do
    '''     with grouping items by family
    '''
    Dim rt As Scripting.Dictionary
    
    If AiDoc Is Nothing Then
        Set rt = d4g2f1(aiDocActive())
    Else
        With dcAiDocCompSetsByPtNum(AiDoc)
            'Stop
            If .Exists(1) Then
                Set rt = dcFrom2Fields(cnGnsDoyle().Execute( _
                    "select ls.it, ISNULL(i.Family, '') fm from " _
                    & sqlValuesFromDc(dcOb(.Item(1)), "ls", "it") _
                    & " left join vgMfiItems i on ls.it = i.Item" _
                ), "it", "fm")
            Else
                Set rt = New Scripting.Dictionary
                Stop
            End If
        End With
    End If
    
    Set d4g2f1 = rt
'send2clipBdWin10 ConvertToJson(dcTransGrouped(d4g2f1(aiDocActive())), vbTab)
End Function

Public Function d4g0f1( _
    AiDoc As Inventor.Document, _
    Optional incTop As Long = 0 _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    'Dim gn As Scripting.Dictionary
    Dim ky As Variant
    Dim pt As Inventor.Document '.PartDocument
    
    Set rt = New Scripting.Dictionary
    With dcOfDcOfDcByPlurality(dcAiDocSetsByPtNum( _
        dcAiDocComponents( _
        AiDoc, , incTop _
    )))
        If .Exists(2) Then
            'ky = TypeName( _
                userChoiceFromDc( _
                dcNewIfNone( _
                dcOb(.Item(2)) _
            )))
            ky = nuSelFromDict( _
                dcNewIfNone( _
                dcOb(.Item(2)) _
            )).GetReply()
            'Stop
        End If
        
        If .Exists("") Then
            Stop
        End If
        
        Stop
        'Debug.Print txDumpLs(dcNewIfNone(dcOb(obOf(.Item(1)))).Keys)
        Set gp = dcNewIfNone(dcOb(obOf(.Item(1))))
        With dcNewIfNone(dcOb(obOf(.Item(1)))) 'dcAiDocComponents(AiDoc, , incTop)
            For Each ky In .Keys
                Set pt = aiDocument(.Item(ky)) 'aiDocPart()
                If pt Is Nothing Then
                Else
                rt.Add ky, dcGnsPtProps_Rev20220830_inProg(pt)
                'dcAiPropValsFromDc()
                'dcOfGnsProps
                End If
            Next
        End With
    End With
    
    'Set gn = dcDxFromRecSetDc(dcFromAdoRS( _
    '    cnGnsDoyle().Execute(q1g1x2(AiDoc) _
    ')))
    'With gn
    'End With
    
    
    Set d4g0f1 = rt
'send2clipBdWin10 ConvertToJson(d4g0f1(aiDocActive()), vbTab)
End Function

Public Function dcOfGnsProps( _
    invDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcOfGnsProps
    '''
    Dim rt As Scripting.Dictionary
    Dim dcPr As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = dcNewIfNone(dc)
    
    Set dcPr = dcOfPropsInAiDoc(invDoc)
    With dcPr
        For Each ky In Array( _
            pnPartNum, pnFamily, pnDesc, _
            pnMaterial, pnStockNum, pnCatWebLink, _
            pnMass, pnRawMaterial, pnRmQty, pnRmUnit, _
            pnThickness, pnLength, pnWidth, pnArea _
        )
            If .Exists(ky) Then
                rt.Add ky, .Item(ky)
            Else
                rt.Add ky, Nothing
            End If
        Next
        
        ''' NOTE[2022.09.16.1024]
        ''' extraction of Categories XML text
        ''' expected to move to content center
        ''' processing in dcGnsValFromContentCtr
        'If .Exists("Categories") Then
            If Len(obAiProp(.Item("Categories")).Value) > 0 Then
                rt.Add "Categories", .Item("Categories")
                'rt.Add "Parameters", dcAiDocParVals(invDoc)
                Debug.Print ; 'Breakpoint Landing: Content Center
            Else
                'Stop
            End If
        'Else
        '    Stop
        'End If
    End With
    
    Set dcOfGnsProps = rt
End Function

Public Function dcGnsValFromContentCtr( _
    CpDef As Inventor.PartComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsValFromContentCtr
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ky As Variant
    
    Dim catXml As String
    Dim pr As Inventor.Parameter
    
    Set rt = dcNewIfNone(dc)
    
    If CpDef Is Nothing Then
    Else
        With CpDef
            With aiDocPart(.Document).PropertySets
                catXml = .Item( _
                gnDesign).Item( _
                "Categories").Value
                
                Set wk = dcAiPropsInSet( _
                    .Item(guidPrSetCLib) _
                )
            End With
            
            ''' NOTE[2022.09.16.1022]
            ''' Categories XML processing
            ''' with Parameter mapping
            ''' will likely be addressed
            ''' in a separate function
            ''' to be called from here.
            'For Each pr In .Parameters
            'Next
        End With
        
        With wk: For Each ky In .Keys 'Array( _
            "Member FileName", "Family", _
            "Standard", "Size Designation", _
            "Categories" _
        )
            'If .Exists(ky) Then
            rt.Add ky, obAiProp(.Item(ky)).Value
            'End If
        Next: End With
    End If
    
    Set dcGnsValFromContentCtr = rt
End Function

Public Function dcGnsValFromPartCompDef( _
    CpDef As Inventor.PartComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsValFromPartCompDef
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcNewIfNone(dc)
    
    If CpDef Is Nothing Then
    Else
        With CpDef
            'part of general ComponentDefinition
            'rt.Add "bomStruct", .BOMStructure
            
            If .IsContentMember Then
                Debug.Print ; 'Breakpoint Landing: Content Member
                'Stop 'and look at this one
                'rt.Add "ContentMem", 1
                rt.Add "ccPropVals", dcGnsValFromContentCtr(CpDef)
                'want to develop this further
                'possible to get library
                'or collection?
            End If
            
            With .MassProperties
                rt.Add pnMass, Round(cvMassKg2LbM * .Mass, 4)
            End With
            
            If .IsiPartMember Then
                rt.Add "isIPartMem", 1
                rt.Add "iPartFactory", .iPartMember.ReferencedDocumentDescriptor.FullDocumentName
            End If
            
            'part of general ComponentDefinition
            'With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
            '    rt.Add "dimsModel", .Dictionary()
            'End With
            
            'part of general ComponentDefinition
            'MIGHT have some use
            'for this one in future
            'With .BOMQuantity
            '    '.BaseUnits
            'End With
            
            'also possible, but unsure
            With .Parameters
            End With
        End With
        
        With dcFlatPatVals( _
            aiCompDefShtMetal(CpDef), _
            dcDotted() _
        )
            If .Count > 2 Then
                rt.Add "flatPat", _
                dcUnDotted(.Item("."))
            Else
                'Debug.Print "KEYS = {" _
                & Join(.Keys, ", ") _
                & "}"
                'Stop 'and make sure
                'all that's in the Dictionary
                'are the self- and back-links
                '(check Immediate pane)
            End If
        End With
    End If
    
    Set dcGnsValFromPartCompDef = rt
End Function

Public Function dcGnsValFromAssyCompDef( _
    CpDef As Inventor.AssemblyComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsValFromAssyCompDef
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcNewIfNone(dc)
    
    If CpDef Is Nothing Then
    Else
        With CpDef
            'part of general ComponentDefinition
            'rt.Add "bomStruct", .BOMStructure
            
            With .MassProperties
                rt.Add pnMass, Round(cvMassKg2LbM * .Mass, 4)
            End With
            
            If .IsiAssemblyMember Then
                rt.Add "isIAssyMem", 1
                rt.Add "iAssyFactory", .iAssemblyMember.ReferencedDocumentDescriptor.FullDocumentName
            End If
            
            'part of general ComponentDefinition
            'With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
            '    rt.Add "dimsModel", .Dictionary()
            'End With
            
            'part of general ComponentDefinition
            'MIGHT have some use
            'for this one in future
            'With .BOMQuantity
            '    '.BaseUnits
            'End With
            
            'also possible, but unsure
            With .Parameters
            End With
        End With
        
        With dcFlatPatVals( _
            aiCompDefShtMetal(CpDef), _
            dcDotted() _
        )
            If .Count > 2 Then
                rt.Add "flatPat", _
                dcUnDotted(.Item("."))
            Else
                'Debug.Print "KEYS = {" _
                & Join(.Keys, ", ") _
                & "}"
                'Stop 'and make sure
                'all that's in the Dictionary
                'are the self- and back-links
                '(check Immediate pane)
            End If
        End With
    End If
    
    Set dcGnsValFromAssyCompDef = rt
End Function

Public Function dcGnsValFromGenCompDef( _
    CpDef As Inventor.ComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsValFromGenCompDef
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcNewIfNone(dc)
    
    If CpDef Is Nothing Then
    Else
        With CpDef
            rt.Add "bomStruct", .BOMStructure
            
            With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
                rt.Add "dimsModel", .Dictionary()
            End With
            
            'MIGHT have some use
            'for this one in future
            'With .BOMQuantity
            '    '.BaseUnits
            'End With
        End With
    End If
    
    Set dcGnsValFromGenCompDef = _
        dcGnsValFromAssyCompDef(aiCompDefAssy(CpDef), _
        dcGnsValFromPartCompDef(aiCompDefPart(CpDef), rt _
    ))
End Function

Public Function dcGnsValGeneral( _
    AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsValGeneral
    '''
    Dim rt As Scripting.Dictionary
    'Dim dcPr As Scripting.Dictionary
    'Dim ky As Variant
    
    Set rt = dcNewIfNone(dc)
    
    If AiDoc Is Nothing Then
    Else
        With AiDoc
            With .PropertySets.Item(gnDesign)
                rt.Add pnPartNum, .Item(pnPartNum).Value
            End With
            
            rt.Add "subType", .SubType
        End With
        
        Set rt = dcGnsValFromGenCompDef( _
            aiCompDefOf(AiDoc), rt _
        )
    End If
    
    'Set dcPr = dcOfPropsInAiDoc(AiDoc)
    'With dcPr
    '    For Each ky In Array( _
    '        pnPartNum, pnFamily, pnDesc, _
    '        pnMaterial, pnStockNum, pnCatWebLink, _
    '        pnMass, pnRawMaterial, pnRmQty, pnRmUnit, _
    '        pnThickness, pnLength, pnWidth, pnArea _
    '    )
    '        If .Exists(ky) Then
    '            rt.Add ky, .Item(ky)
    '        Else
    '            rt.Add ky, Nothing
    '        End If
    '    Next
    'End With
    
    Set dcGnsValGeneral = rt
End Function

Public Function dcGnsPtProps_Rev20220830_inProg( _
    AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary '.PartDocument
    '''
    ''' dcGnsPtProps_Rev20220830_inProg
    '''
    Dim rt As Scripting.Dictionary
    'Dim dcPr As Scripting.Dictionary
    'Dim dcVl As Scripting.Dictionary
    'Dim ky As Variant
    
    'Set rt = dcNewIfNone(dc)
    Set rt = dcGnsValGeneral( _
        AiDoc, dcNewIfNone(dc) _
    )
    
    With dcOfGnsProps( _
    AiDoc, dcDotted())
        If .Count > 2 Then
            rt.Add "props", _
            dcUnDotted(.Item("."))
            
            rt.Add "propVals", _
            dcPropVals(rt.Item("props"))
            
            If .Exists("Categories") Then
                rt.Add "Parameters", _
                dcAiDocParVals(AiDoc)
            End If
        End If
    End With
    
    If AiDoc Is Nothing Then
    Else
        'Set dcPr = dcOfGnsProps(AiDoc) 'dcOfPropsInAiDoc
        'Set dcVl = dcAiPropValsFromDc(dcPr)
        
        With AiDoc
            With .PropertySets.Item(gnDesign)
                'rt.Add pnPartNum, .Item(pnPartNum).Value
                'only want this one to tie
                'back to other Dictionaries
                
                'rt.Add pnFamily, .Item(pnFamily).Value
                'rt.Add pnDesc, .Item(pnDesc).Value
                'rt.Add pnMaterial, .Item(pnMaterial).Value
                'rt.Add pnStockNum, .Item(pnStockNum).Value
                'rt.Add pnCatWebLink, .Item(pnCatWebLink).Value
                
                'rt.Add pnThickness, .Item(pnThickness).Value
                'this one's a Custom property
                'whether Inventor handles
                'itself or not is uncertain.
                
                'pnMass, pnRawMaterial, pnRmQty, pnRmUnit,
                'pnLength, pnWidth, pnArea
            End With
            
            '.ActiveMaterial
            '.FullDocumentName
            '.UnitsOfMeasure
            '.NeedsMigrating
            '.RequiresUpdate
        End With
    End If 'AiDoc Is Nothing
    
    'Stop
    'Call iSyncPartFactory(AiDoc)
    'Set rt = dcVl

    Set dcGnsPtProps_Rev20220830_inProg = rt
End Function

Public Function sqlValuesFromDc( _
    dc As Scripting.Dictionary, _
    Optional vw As String = "ls", _
    Optional it As String = "it" _
) As String
    '''
    ''' sqlValuesFromDc - generate SQL
    '''     "VALUES" clause from Keys
    '''     of supplied Dictionary.
    '''     result is a relation
    '''     of one attribute
    '''     '
    '''     VALUES clause must end
    '''     with an AS phrase naming
    '''     the relation and all
    '''     attributes.
    '''     '
    '''     in this function, the
    '''     names default to "ls"
    '''     (list) for the relation,
    '''     and "it" (item) for
    '''     its one attribute
    '''
    If dc Is Nothing Then
        sqlValuesFromDc = sqlValuesFromDc( _
        New Scripting.Dictionary)
    Else
        sqlValuesFromDc = "(values ('" _
        & Join(dc.Keys, "'), ('") & _
        "')) as ls(it)"
    End If
End Function

Public Function dcAiDocCompSetsByPtNum( _
    AiDoc As Inventor.Document, _
    Optional incTop As Long = 0 _
) As Scripting.Dictionary
    '''
    ''' dcAiDocCompSetsByPtNum -- formerly d4g3f1
    '''
    '''
    '''
    Dim dc As Scripting.Dictionary
    'Dim ct As Long
    
    'ct = 1 'to include main assembly (for now)
    'now disabled in favor of input parameter incTop
    
    'dcAiDocSetsByPtNum replaces dcRemapByPtNum
    Set dc = dcOfDcOfDcByPlurality( _
        dcAiDocSetsByPtNum( _
        dcAiDocComponents( _
            AiDoc, , incTop _
        ) _
    )) 'incTop replaces ct
    'dcOfDcOfDcByItemCount() removed from original lineup
    'of dcOfDcOfDcByPlurality(dcOfDcOfDcByItemCount(dcAiDocSetsByPtNum(
    'dcOfDcOfDcByPlurality calls dcOfDcOfDcByItemCount internally
    'as part of its normal processing.
    
    Set dcAiDocCompSetsByPtNum = dc
    'WAS going to simply replace assignment of
    'dc above with direct assignment of dcAiDocCompSetsByPtNum,
    'but further processing here may be desired.
End Function

Public Function dcAiDocSetsByPtNum( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcAiDocSetsByPtNum -- formerly d4g3f2
    '''
    ''' Returns Dictionary of Dictionaries
    ''' of Inventor Documents keyed on
    ''' associated Part Numbers.
    '''
    ''' Derived from dcRemapByPtNum, this
    ''' variation collects all models of a
    ''' given Part Number into a secondary
    ''' Dictionary, under each Document's
    ''' file name. Ideally, Part Numbers
    ''' should map one-to-one to Documents,
    ''' so each sub Dictionary should
    ''' contain only one entry.
    '''
    ''' However, as it IS possible for more
    ''' than one model to represent the same
    ''' Part, more than one Document might
    ''' in fact have the same Part Number.
    '''
    ''' Therefore, it may sometimes prove
    ''' necessary to take additional steps
    ''' in properly identifying which model
    ''' (or models) to process in preparation
    ''' for Genius.
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim pt As Inventor.Document
    Dim ky As Variant
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set pt = aiDocument(.Item(ky))
            pn = CStr(aiDocPropVal( _
                pt, pnPartNum, gnDesign _
            ))
            
            ''' REV[2022.05.17.1536]
            ''' removing check/handling
            ''' for blank/empty part number.
            ''' client process can deal with that.
            With rt
                If .Exists(pn) Then 'do nothing
                Else
                    .Add pn, New Scripting.Dictionary
                End If
                
                Set gp = .Item(pn)
            End With
            
            With gp
                If .Exists(ky) Then
                    Stop 'this should NOT happen
                Else
                    .Add ky, pt
                End If
            End With
        Next
    End With
    
    Set dcAiDocSetsByPtNum = rt
End Function

Public Function dcOfDcOfDcByItemCount( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcOfDcOfDcByItemCount -- formerly d4g3f3
    '''
    ''' subdivide supplied Dictionary
    ''' of Dictionaries into groups
    ''' by Count of members.
    '''
    ''' result is a 3rd-order Dictionary,
    ''' that is, a Dictionary (1)
    '''     keyed by member count
    ''' of Dictionaries (2)
    '''     keyed by a shared key
    ''' of yet more Dictionaries (3)
    '''     keyed to some unique value
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim xp As Scripting.Dictionary
    Dim ky As Variant
    Dim ct As Long
    
    Set rt = New Scripting.Dictionary
    
    With dc: For Each ky In .Keys
        Set gp = .Item(ky)
        ct = gp.Count
        
        With rt
            If Not .Exists(ct) Then
            .Add ct, New Scripting.Dictionary
            End If
            
            Set xp = .Item(ct)
        End With
        
        With xp
        If .Exists(ky) Then
            Stop 'this should NOT happen
        Else
            .Add ky, gp
        End If: End With
    Next: End With
    
    Set dcOfDcOfDcByItemCount = rt
End Function

Public Function dcOfDcOfDcByPlurality( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcOfDcOfDcByPlurality -- formerly d4g3f4
    '''
    ''' given a 2nd-order Dictionary (NOT 3rd)
    ''' as supplied by dcAiDocSetsByPtNum (NOT dcOfDcOfDcByItemCount),
    ''' return a reorganized version as follows:
    '''
    ''' under key 1: a Dictionary of all
    '''     Dictionaries having only one member.
    '''     this should form the bulk of the
    '''     supplied Dictionary's content.
    '''
    ''' under key 2: a Dictionary of Dictionaries
    '''     having more than one member. these
    '''     "plurals" might require additional
    '''     review and/or processing to resolve
    '''     ambiguities, conflicts, etc.
    '''
    ''' under key "" (blank string): the Dictionary,
    '''     if present, of members with no assigned
    '''     part or item number. this should almost
    '''     NEVER arise, but again, might require
    '''     special processing to resolve issues.
    '''
    ''' '''
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim xp As Scripting.Dictionary
    Dim ky As Variant
    Dim ct As Long
    
    Set rt = New Scripting.Dictionary
    
    ''' to avoid modifying supplied Dictionary,
    ''' generate a copy to work with directly.
    Set gp = dcCopy(dc)
    With gp
    If .Exists("") Then 'get the set
        'of blank "numbered"
        'items moved over...
        
        rt.Add "", .Item("")
        .Remove ""
    End If: End With
    'before grouping by member counts:
    
    Set gp = dcOfDcOfDcByItemCount(gp)
    With gp
        'prep the "singles" return Dictionary
        Set xp = New Scripting.Dictionary
        rt.Add 1, xp
        
        If .Exists(1) Then 'proceed to move
            'the (one, single) member of each
            'Dictionary under this one
            With dcOb(.Item(1))
            For Each ky In .Keys
                With dcOb(.Item(ky))
                xp.Add ky, .Items(0) '.Item(.Keys(0))
                End With
            Next: End With
            .Remove 1
        Else 'we've got nothing to move.
            'this is actually a problem,
            'but that's for the client
            'process to handle.
        End If
        'at this point, any remaining members
        'should be "plural" Dictionaries
        'containing more than one member.
        
        'THESE are to be combined into one
        '"plural" Dictionary to be returned.
        Set xp = New Scripting.Dictionary
        'DO NOT add to return Dictionary yet!
        
        For Each ky In .Keys
            'this step generates a NEW
            'Dictionary at each iteration.
            Set xp = dcKeysCombined( _
                xp, dcOb(.Item(ky)), 1 _
            )
            'it does NOT update the original!
        Next
        
        'NOW, we can add the final result
        'to the return Dictionary...
        If xp.Count > 0 Then
            '...ASSUMING any are
            'left to add, of course!
            rt.Add 2, xp
        End If
    End With
    
    'Stop 'because NOT sure this thing
         'is ready for prime time...
    
    'disabling the following section completely
    'it should not be needed, as all three parts
    'of the return Dictionary should be in place
    'at the end of the preceding With block.
    
    'this SHOULD add Dictionaries for all
    'Part Numbers with more than one Document,
    'in a single Dictionary of "plurals",
    'but need to check it out for sure yet.
    'Hence the preceding Stop
    
    'rt.Add 2, dcKeysMissing(gp, rt.Item(1))
    
    'With dc: For Each ky In .Keys
    '    Set gp = .Item(ky)
    '    ct = gp.Count
    '
    '    With rt
    '        If Not .Exists(ct) Then
    '        .Add ct, New Scripting.Dictionary
    '        End If
    '
    '        Set xp = .Item(ct)
    '    End With
    '
    '    With xp
    '    If .Exists(ky) Then
    '        Stop 'this should NOT happen
    '    Else
    '        .Add ky, gp
    '    End If: End With
    'Next: End With
    
    Set dcOfDcOfDcByPlurality = rt
End Function

Public Function d4g3f5from2( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' d4g3f5from2
    '''
    ''' Returns Dictionary of Dictionaries
    ''' of Inventor Documents keyed on
    ''' associated Part Numbers.
    '''
    ''' Derived from dcRemapByPtNum, this
    ''' variation collects all models of a
    ''' given Part Number into a secondary
    ''' Dictionary, under each Document's
    ''' file name. Ideally, Part Numbers
    ''' should map one-to-one to Documents,
    ''' so each sub Dictionary should
    ''' contain only one entry.
    '''
    ''' However, as it IS possible for more
    ''' than one model to represent the same
    ''' Part, more than one Document might
    ''' in fact have the same Part Number.
    '''
    ''' Therefore, it may sometimes prove
    ''' necessary to take additional steps
    ''' in properly identifying which model
    ''' (or models) to process in preparation
    ''' for Genius.
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim pt As Inventor.Document
    Dim ky As Variant
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set pt = aiDocument(.Item(ky))
            
            'pt.co
            
            pn = CStr(aiDocPropVal( _
                pt, pnPartNum, gnDesign _
            ))
            
            ''' REV[2022.05.17.1536]
            ''' removing check/handling
            ''' for blank/empty part number.
            ''' client process can deal with that.
            With rt
                If .Exists(pn) Then 'do nothing
                Else
                    .Add pn, New Scripting.Dictionary
                End If
                
                Set gp = .Item(pn)
            End With
            
            With gp
                If .Exists(ky) Then
                    Stop 'this should NOT happen
                Else
                    .Add ky, pt
                End If
            End With
        Next
    End With
    
    Set d4g3f5from2 = rt
End Function

Public Function d4g3f5b( _
    cd As Inventor.ComponentDefinition _
) As Variant
    
End Function

Public Function d4g3f5a( _
    cd As Inventor.ComponentDefinition _
) As Object 'Inventor.iAssemblyTableCell
    'aiCompDefOf
    
    If cd Is Nothing Then
        Set d4g3f5a = Nothing
    ElseIf TypeOf cd Is Inventor.AssemblyComponentDefinition Then
        With aiCompDefAssy(cd)
            If .IsiAssemblyMember Then
            ElseIf .IsiAssemblyFactory Then
            ElseIf .IsModelStateMember Then
            ElseIf .IsModelStateFactory Then
            End If
        End With
    ElseIf TypeOf cd Is Inventor.PartComponentDefinition Then
    End If
End Function

Public Function gnsUpdtAll_iFact( _
    cd As Inventor.ComponentDefinition _
) As Scripting.Dictionary
    If TypeOf cd Is Inventor.PartComponentDefinition Then
        Set gnsUpdtAll_iFact = gnsUpdtAll_iPart(cd)
    ElseIf TypeOf cd Is Inventor.AssemblyComponentDefinition Then
        Set gnsUpdtAll_iFact = gnsUpdtAll_iAssy(cd)
    Else
        Set gnsUpdtAll_iFact = New Scripting.Dictionary
    End If
End Function

Public Function gnsUpdtAll_iAssy( _
    cd As Inventor.AssemblyComponentDefinition _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim md As Inventor.AssemblyDocument
    Dim fc As Inventor.iAssemblyFactory
    Dim rw As Inventor.iAssemblyTableRow
    Dim r0 As Inventor.iAssemblyTableRow
    
    Set rt = New Scripting.Dictionary
    
    If cd Is Nothing Then
    ElseIf cd.IsiAssemblyFactory Then
        With cd.iAssemblyFactory
            Set md = .Parent.Parent
            
            With .TableColumns '.Item()
                '"GeniusMass [Custom]"
            End With
            
            'note initial DefaultRow
            Set r0 = .DefaultRow
            
            For Each rw In .TableRows
                .DefaultRow = rw
                rt.Add .DefaultRow.MemberName, _
                dcOfDcAiPropVals(dcGeniusProps(md))
                'md.Save
            Next
            
            'restore initial DefaultRow
            .DefaultRow = r0
        End With
    Else
    End If
    
    Set gnsUpdtAll_iAssy = rt
'Debug.Print dcOfDcAiPropVals(dcGeniusProps(aiDocActive())).Count
'send2clipBdWin10 ConvertToJson(gnsUpdtAll_iAssy(aiDocAssy(aiDocActive()).ComponentDefinition), vbTab)
End Function

Public Function gnsUpdtAll_iPart( _
    cd As Inventor.PartComponentDefinition _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim md As Inventor.PartDocument
    Dim fc As Inventor.iPartFactory
    Dim rw As Inventor.iPartTableRow
    Dim r0 As Inventor.iPartTableRow
    
    Set rt = New Scripting.Dictionary
    
    If cd Is Nothing Then
    ElseIf cd.IsiPartFactory Then
        With cd.iPartFactory
            Set md = .Parent '.Parent
            
            With .TableColumns '.Item()
                '"GeniusMass [Custom]"
            End With
            
            'note initial DefaultRow
            Set r0 = .DefaultRow
            
            For Each rw In .TableRows
                .DefaultRow = rw
                rt.Add .DefaultRow.MemberName, _
                dcAiPropValsFromDc(dcGeniusProps(md)) 'dcOfDcAiPropVals
                DoEvents
                md.Save
            Next
            
            'restore initial DefaultRow
            .DefaultRow = r0
        End With
    Else
    End If
    
    Set gnsUpdtAll_iPart = rt
'send2clipBdWin10 ConvertToJson(gnsUpdtAll_iPart(aiDocPart(aiDocActive()).ComponentDefinition), vbTab)
End Function

Public Function d4g3f6pt( _
    cd As Inventor.PartComponentDefinition _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim mx As Long
    Dim dx As Long
    Dim ck As Inventor.iPartTableColumn
    
    Set rt = New Scripting.Dictionary
    If cd Is Nothing Then
    ElseIf cd.IsiPartFactory Then
        Stop
        With cd.iPartFactory.TableColumns
            mx = .Count
            Stop
            For dx = 1 To mx
                Set ck = .Item(dx)
                With ck
                    rt.Add .Heading, nuDcPopulator( _
                    ).Setting("dx", dx _
                    ).Setting("dh", .DisplayHeading _
                    ).Setting("fh", .FormattedHeading _
                    ).Setting("dt", .ReferencedDataType _
                    ).Setting("ob", .ReferencedObject _
                    ).Setting("ot", TypeName(.ReferencedObject) _
                    ).Dictionary
                    '
                    ').Setting("hd", .Heading _
                    '
                    Stop
                End With
                'ck.ReferencedObject
                'Debug.Print obAiProp(ck.ReferencedObject).Name
                'rt.Add dx, nuDcPopulator().Setting("hd", ck.Heading).Setting("dh", ck.DisplayHeading).Setting("fh", ck.FormattedHeading).Dictionary
                'If ck.ReferencedDataType = kMemberNameColumn Then
                'End If
            Next
        End With
    ElseIf cd.IsiPartMember Then
        Stop
        Set rt = d4g3f6pt(aiDocPart(cd.iPartMember.ParentFactory.Parent).ComponentDefinition)
        'cd.iPartMember.ParentFactory.Parent
    Else
    End If
    
    Set d4g3f6pt = rt
End Function

Public Function d4g3f6as( _
    cd As Inventor.AssemblyComponentDefinition _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim mx As Long
    Dim dx As Long
    Dim ck As Inventor.iAssemblyTableColumn
    
    Set rt = New Scripting.Dictionary
    If cd Is Nothing Then
    ElseIf cd.IsiAssemblyFactory Then
        Stop
        With cd.iAssemblyFactory.TableColumns
            mx = .Count
            Stop
            For dx = 1 To mx
                Set ck = .Item(dx)
                Stop
            Next
        End With
    ElseIf cd.IsiAssemblyMember Then
        Stop
        'cd.iAssemblyMember.ParentFactory
    Else
    End If
    
    Set d4g3f6as = rt
End Function

Public Function d4g3f7pt( _
    cd As Inventor.PartComponentDefinition _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim hd As Scripting.Dictionary
    'Dim md As Inventor.PartDocument
    Dim fc As Inventor.iPartFactory
    Dim co As Inventor.iPartTableColumn
    Dim rw As Inventor.iPartTableRow
    'Dim r0 As Inventor.iPartTableRow
    'Dim df As Long
    
    Set rt = New Scripting.Dictionary
    rt.Add "", New Scripting.Dictionary
    Set hd = rt.Item("")

    
    If cd Is Nothing Then
        Set fc = Nothing
    ElseIf cd.IsiPartFactory Then
        Set fc = cd.iPartFactory
    ElseIf cd.IsiPartMember Then
        Set fc = cd.iPartMember.ParentFactory
    Else
        Set fc = Nothing
    End If
    
    If Not fc Is Nothing Then
    With fc
        'Set md = .Parent '.Parent
        
        hd.Add "", Array( _
            "Index", _
            "Key", _
            "CustomColumn", _
            "DisplayHeading", _
            "FormattedHeading", _
            "ReferencedDataType" _
        )
        For Each co In .TableColumns: With co
            '.Item()
            With co
            hd.Add .Heading, Array( _
                .Index, _
                .Key, _
                .CustomColumn, _
                .DisplayHeading, _
                .FormattedHeading, _
                .ReferencedDataType _
            )
            End With
            'Stop
        End With: Next
        
        'note initial DefaultRow
        'Set r0 = .DefaultRow
        
        For Each rw In .TableRows
            With rw
                'df = rw Is r0
                rt.Add .Index, Array(.Index, .MemberName, .PartName, rw) ', df
            End With
            '.DefaultRow = rw
            'rt.Add .DefaultRow.MemberName, _
            dcAiPropValsFromDc(dcGeniusProps(md)) 'dcOfDcAiPropVals
            'DoEvents
            'md.Save
        Next
        
        'restore initial DefaultRow
        '.DefaultRow = r0
    End With: End If
    
    Set d4g3f7pt = rt
'send2clipBdWin10 ConvertToJson(d4g3f7pt(aiDocPart(aiDocActive()).ComponentDefinition), vbTab)
End Function

Public Function d4g4f0( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Scripting.Dictionary
    '''
    ''' d4g4f0 -- rebuilding Sub Update_Genius_Properties
    '''           more or less from the ground up
    '''
    'Dim invProgressBar As Inventor.ProgressBar
    
    'Dim fc As gnsIfcAiDoc
    Dim dc As Scripting.Dictionary
    Dim mt As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    'Dim goAhead As VbMsgBoxResult
    'Dim ActiveDoc As Document
    'Dim txOut As String
    Dim ky As Variant
    Dim k2 As Variant
    'Dim kyPt As Variant
    Dim ct As Long
    
    'Dim dx As Long
    
    'Dim fm As fmIfcTest05A
    
    ''' NOTE[2022.06.01.1441]
    ''' adding check for supplied Document
    ''' call for selection if none
    If AiDoc Is Nothing Then
        'Set rt = d4g4f0(userChoiceFromDc(dcAiDocsVisible(), aiDocActive()))
        Set rt = d4g4f0(aiDocActive())
    Else
        ''' NOTE[2022.06.01.1442]
        ''' disabling/skipping user checks for now
        ''' this isn't the purpose of this whole mess.
        ''' Confirm User Request
        ''' to process active Document
        'goAhead = MsgBox( _
        '    Join(Array( _
        '        "Are you sure you want to process this document?", _
        '        "The process may require a few minutes depending on assembly size.", _
        '        "Suppressed and excluded parts will not be processed." _
        '    ), " "), _
        '    vbYesNo + vbQuestion, _
        '    "Process Document Custom iProperties" _
        ')
        'If goAhead = vbYes Then
        '
        'Else
        'End If
        
        ''' NOTE[2022.06.01.1444]
        ''' see Update_Genius_Properties REV[2022.05.24.0956]
        With dcAiDocCompSetsByPtNum(AiDoc, ct) 'ActiveDoc
            If .Exists("") Then
                Stop 'for now
                'don't expect this situation
                'to occur frequently, so won't
                'worry about a handler just yet
            End If
            
            If .Exists(2) Then
                'THIS situation IS known to occur,
                'if not TERRIBLY frequently, so a
                'handler here is a good idea.
                '
                With nuDcPopulator(.Item(2)) 'd4g4f4(dcOb(.Item(2)))
                    Debug.Print MsgBox( _
                        msg_2022_0603_1127(.Dictionary), _
                        vbOKOnly Or vbInformation, _
                        "Duplicate Part Numbers!" _
                    ) 'with just a slight modification,
                End With
                
                'With d4g4f4(dcOb(.Item(2)))
                '    Stop
                '
                '    'fortunately, we have one ready made
                '    'in the dcRemapByPtNum function this
                '    'section is replacing (see above).
                '    Debug.Print MsgBox( _
                '        Join(Array( _
                '            "The following Part Numbers are", _
                '            "assigned to more than one Model:", _
                '            "", _
                '            vbTab & Join(.Keys, vbNewLine & vbTab), _
                '            "" _
                '        ), vbNewLine), _
                '        vbOKOnly Or vbInformation, _
                '        "Duplicate Part Numbers!" _
                '    ) 'with just a slight modification,
                '    'this serves to notify the user
                '    'just as dcRemapByPtNum did before.
                '    'a more sophisticated response may
                '    'eventually be called for, but for
                '    'now, this will do.
                'End With
            End If
            
            'and HERE is the step which ACTUALLY
            'replaces the prior version above.
            'Key 1 is guaranteed to be present
            'in the Dictionary returned, so no
            'need to check for it here.
            Set dc = dcOb(.Item(1))
        End With
        '''
        '''
        '''
        
        '''
        ''' NOTE[2022.06.01.1502]
        ''' this section expected to be
        ''' exported to its own function
        ''' NOTE[2022.06.02.0906]
        ''' (follow-up) original code
        ''' extracted to functions
        ''' dcOfKeys2match and d4g4f1
        Set mt = dcOfKeys2match(Array( _
            pnFamily, pnMass, _
            pnRawMaterial, _
            pnRmQty, pnRmUnit, _
            pnWidth, pnLength, _
            pnArea, pnThickness _
        ))
        'pnFamily       replaces "Cost Center"
        'pnMass         replaces "GeniusMass"
        'pnRawMaterial  replaces "RM"
        'pnRmQty        replaces "RMQTY"
        'pnRmUnit       replaces "RMUNIT"
        'pnWidth        replaces "Extent_Width"
        'pnLength       replaces "Extent_Length"
        'pnArea         replaces "Extent_Area"
        'pnThickness    replaces "Thickness"
        
        'Set rt = d4g4f1(dc, mt)
        
        Set rt = New Scripting.Dictionary
        With d4g4f1(dc, mt)
        For Each ky In .Keys
            With rt
                If Not .Exists(ky) Then
                .Add ky, New Scripting.Dictionary
                End If
                
                Set wk = .Item(ky)
            End With
            
            With dcOb(.Item(ky))
                For Each k2 In .Keys
                wk.Add k2, obAiProp(.Item(k2)).Value
                Next
            End With
        Next: End With
        
        '''
        '''
        '''
        'Set rt = dcCopy(dc)
    End If
    '''
    Set d4g4f0 = rt
'send2clipBdWin10 ConvertToJson(d4g4f0(), vbTab)
End Function

Public Function d4g4f1( _
    dc As Scripting.Dictionary, _
    rf As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' d4g4f1 -- returns a Dictionary of Dictionaries
    '''     copied from supplied Dictionary dc,
    '''     but with only those Keys matching those
    '''     found in supplied 'reference' Dictionary rf
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        rt.Add ky, dcKeysInCommon( _
            dcOfPropsInAiDoc( _
            aiDocument(.Item(ky)) _
            ), rf, 1 _
        )
    Next: End With
    
    Set d4g4f1 = rt
End Function

Public Function d4g4f2( _
    dc As Scripting.Dictionary _
) As Inventor.PartDocument
    '''
    ''' d4g4f2 -- given a Dictionary of Part Documents
    '''     return the first Content Center Member found
    '''     (if none found, return Nothing)
    '''
    Dim ky As Variant
    Dim pt As Inventor.PartDocument
    
    With dc: For Each ky In .Keys
    If pt Is Nothing Then
        Set pt = aiDocPart(aiDocument(.Item(ky)))
        If Not pt Is Nothing Then
            If Not pt.ComponentDefinition( _
            ).IsContentMember Then Set pt = Nothing
        End If
    End If
    Next: End With
    
    Set d4g4f2 = pt
End Function

Public Function d4g4f3( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' d4g4f3 -- given a Dictionary of Part Document
    '''     Dictionaries, return a subset containing
    '''     only Content Center Members
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim pt As Inventor.PartDocument
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        Set pt = d4g4f2(.Item(ky))
        If Not pt Is Nothing Then
            rt.Add ky, pt
        End If
    Next: End With
    
    Set d4g4f3 = rt
End Function

Public Function d4g4f4( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' d4g4f4 -- given a Dictionary of Part Document
    '''     Dictionaries, return a subset dropping
    '''     any with Content Center Members
    '''
    Set d4g4f4 = dcKeysMissing(dc, d4g4f3(dc))
End Function

Public Function d4g4f5() As Scripting.Dictionary ' _
    dc As Scripting.Dictionary _
'''
    '''
    ''' d4g4f5 -- given a Dictionary of Part Document
    '''     Dictionaries, return a subset dropping
    '''     any with Content Center Members
    '''
    Dim rt As Scripting.Dictionary
    Dim dcRb As Scripting.Dictionary
    Dim dcTb As Scripting.Dictionary
    Dim dcRp As Scripting.Dictionary
    'Dim dcRb As Scripting.Dictionary
    Dim rb As Inventor.Ribbon
    Dim tb As Inventor.RibbonTab
    Dim rp As Inventor.RibbonPanel
    Dim ob As Object
    
    Set rt = New Scripting.Dictionary
    'ThisApplication.UserInterfaceManager.RibbonState
    With ThisApplication.UserInterfaceManager
    For Each rb In .Ribbons
        Set dcRb = New Scripting.Dictionary
        With rb
            rt.Add .InternalName, dcRb
            For Each tb In .RibbonTabs '.QuickAccessControls
                Set dcTb = New Scripting.Dictionary
                With tb
                    dcRb.Add .InternalName, dcTb
                    For Each rp In .RibbonPanels
                        Set dcRp = New Scripting.Dictionary
                        With rp
                            dcTb.Add .InternalName, dcRp
                            Stop
                        End With
                    Next
                End With
            Next
        End With
    Next: End With
    Set d4g4f5 = rt
End Function

Public Function compDefOfPart( _
    AiDoc As Inventor.PartDocument _
) As Inventor.ComponentDefinition
    If AiDoc Is Nothing Then
        Set compDefOfPart = Nothing
    Else
        Set compDefOfPart = AiDoc.ComponentDefinition
    End If
End Function

Public Function compDefOfAssy( _
    AiDoc As Inventor.AssemblyDocument _
) As Inventor.ComponentDefinition
    If AiDoc Is Nothing Then
        Set compDefOfAssy = Nothing
    Else
        Set compDefOfAssy = AiDoc.ComponentDefinition
    End If
End Function

Public Function compDefOf( _
    AiDoc As Inventor.Document _
) As Inventor.ComponentDefinition
    Dim rt As Inventor.ComponentDefinition
    
    Set rt = compDefOfPart(aiDocPart(AiDoc))
    If rt Is Nothing Then
    Set rt = compDefOfAssy(aiDocAssy(AiDoc))
    End If
    
    'AiDoc.FullFileName
    
    Set compDefOf = rt
End Function

Public Function famOfAiDoc( _
    AiDoc As Inventor.Document _
) As String
    '''
    Dim itNum As String
    Dim mdFam As String
    Dim gnFam As String
    
    Dim pf As String
    Dim sf As String
    
    Dim ck As VbMsgBoxResult
    
    If AiDoc Is Nothing Then
        famOfAiDoc = ""
    Else
        With AiDoc
            ''' NOTE!!! ONLY use this for Assemblies!
            ''' will disable until better set up
            'With nuDcPopulator( _
            '    ).Setting("doyle", "D" _
            '    ).Setting("riverview", "R" _
            ').Matching(Split(.FullDocumentName, "\"))
            '    If .Count = 1 Then
            '        pf = .Item(.Keys(0))
            '    Else
            '        pf = ""
            '    End If
            'End With
            
            With .PropertySets.Item(gnDesign)
                mdFam = .Item(pnFamily).Value
                itNum = .Item(pnPartNum).Value
                gnFam = famInGenius(itNum)
            End With
        End With
        
        With compDefOf(AiDoc)
            If .BOMStructure = kPurchasedBOMStructure Then
                sf = "PTS"
            End If
        End With
    End If
End Function

Public Function famInGenius(itNum As String) As String
    Dim gnFam As String
    
    With cnGnsDoyle().Execute( _
        "select Family from vgMfiItems where Item = '" _
        & itNum & "';" _
    )
        If .BOF Or .EOF Then
            gnFam = ""
        Else
            gnFam = Split(.GetString( _
            adClipString, , "", "", "" _
            ), vbCr)(0)
        End If
    End With
    
    famInGenius = gnFam
End Function

Public Function famIfValid(mdFam As String) As String
    With cnGnsDoyle().Execute(Join(Array( _
        "select ISNULL(f.Family, '') Family", _
        "from (values ('" & mdFam & "')) as i(f)", _
        "left join vgMfiFamilies f on i.f = f.Family" _
    )))
        If .BOF Or .EOF Then
            famIfValid = ""
        Else
            famIfValid = .Fields("Family").Value
        End If
    End With
End Function

Public Function famVsGenius(itNum As String, _
    Optional mdFam As String = "" _
) As String
    Dim ckFam As String
    Dim gnFam As String
    Dim ck As VbMsgBoxResult
    
    ''' get current family from
    ''' Genius, if it has one
    gnFam = famInGenius(itNum)
    
    If Len(gnFam) = 0 Then 'no family in Genius
        famVsGenius = mdFam 'so just use the model's
    Else 'need to check the model against it
        ''' first, verify model family
        ckFam = famIfValid(mdFam)
        ''' if not in Genius...
        
        If Len(ckFam) = 0 Then 'no need to ask
            famVsGenius = gnFam
        ElseIf gnFam = ckFam Then 'it's good
            famVsGenius = ckFam
        Else 'check with user
            ck = MsgBox(Join(Array( _
                "Item " & itNum, _
                "Model Part Family " & ckFam & " differs", _
                "from Genius Part Family " & gnFam, "", _
                "Change Model to match Genius?", "", _
                "(click [CANCEL] to debug)" _
            ), vbNewLine), _
                vbYesNoCancel + vbQuestion, _
                "Use Genius Family?" _
            )
            
            If ck = vbCancel Then
                Stop 'to debug
            ElseIf ck = vbYes Then 'match Genius
                famVsGenius = gnFam
            Else 'keep model Family
                famVsGenius = ckFam
            End If
        End If
    End If
    '''
End Function

Public Function msg_2022_0603_1127( _
    dc As Scripting.Dictionary _
) As String
    '''
    ''' msg_2022_0603_1127
    '''
    Dim cc As Scripting.Dictionary
    Dim rm As Scripting.Dictionary
    Dim rt As String
    
    rt = ""
    
    Set cc = d4g4f3(dc)
    Set rm = dcKeysMissing(dc, cc)
    
    If rm.Count > 0 Then
        rt = Join(Array(rt, _
            "The following Part Numbers are", _
            "assigned to more than one Model:", _
            "", _
            vbTab & Join(rm.Keys, vbNewLine & vbTab), _
            "" _
        ), vbNewLine)
    End If
    
    If cc.Count > 0 Then
        rt = Join(Array(rt, _
            "These duplicated Part Numbers are", _
            "associated with at least one Content", _
            "Center Member, which cannot be modified:", _
            "", _
            vbTab & Join(cc.Keys, vbNewLine & vbTab), _
            "" _
        ), vbNewLine)
    End If
    
    rt = Join(Array(rt, _
        "These will not be processed.", _
        "" _
    ), vbNewLine)
    
    msg_2022_0603_1127 = rt 'Join(Array( _
        "The following Part Numbers are", _
        "assigned to more than one Model:", _
        "", _
        vbTab & Join(d4g4f4(.Dictionary).Keys, vbNewLine & vbTab), _
        "", _
        "These duplicated Part Numbers are", _
        "associated with at least one Content", _
        "Center Member, which cannot be modified:", _
        "", _
        vbTab & Join(cc.Keys, vbNewLine & vbTab), _
        "", _
        "These will not be processed.", _
        "" _
    ), vbNewLine)
End Function

Public Function askUserForPartMatl( _
    AiDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' askUserForPartMatl -- Prompt User
    '''     for Part Family and Material
    '''     Selection, returning result
    '''     in Dictionary
    '''
    ''' REV[2022.08.29.1621]
    ''' add optional Dictionary parameter to which
    ''' data from this function can be added.
    ''' (see also askUserForMatlQty)
    '''
    Dim rt As Scripting.Dictionary
    Dim ck As VbMsgBoxResult
    Dim bd As aiBoxData
    Dim tx As String
    
    If dc Is Nothing Then
        Set rt = askUserForPartMatl( _
        AiDoc, New Scripting.Dictionary)
    Else
        Set rt = dc 'New Scripting.Dictionary
        
        With rt
            If .Exists(pnFamily) Then .Remove pnFamily
            If .Exists(pnRawMaterial) Then .Remove pnRawMaterial
            
            If AiDoc Is Nothing Then
                .Add pnFamily, ""
                .Add pnRawMaterial, ""
            Else
                With newFmTest1()
                    Set bd = nuAiBoxData().UsingInches.SortingDims( _
                        AiDoc.ComponentDefinition.RangeBox _
                    )
                    ck = .AskAbout(AiDoc, _
                        "No Stock Found! Please Review" _
                        & vbNewLine & vbNewLine & bd.Dump(0) _
                    )
                    
                    If ck = vbYes Then
                        'Stop 'because this will
                        'override supplied Dictionary!
                        'Set rt =
                        With .ItemData()
                            rt.Add pnFamily, .Item(pnFamily)
                            rt.Add pnRawMaterial, .Item(pnRawMaterial)
                        End With
                    Else
                        'Set rt = New Scripting.Dictionary
                        Stop
                        
                        With AiDoc.PropertySets
                            tx = .Item(gnDesign _
                                ).Item(pnFamily _
                            ).Value
                            rt.Add pnFamily, tx
                            
                            On Error Resume Next
                            Err.Clear
                            tx = .Item(gnCustom _
                                ).Item(pnRawMaterial _
                            ).Value
                            If Err.Number Then tx = ""
                            On Error GoTo 0
                            rt.Add pnRawMaterial, tx
                        End With
                    End If
                End With
            End If
        End With
    End If
    
    Set askUserForPartMatl = rt
End Function

Public Function askUserForMatlQty( _
    AiDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' askUserForMatlQty -- Prompt User
    '''     for Material Quantity and Units,
    '''     returning result in Dictionary
    '''
    ''' REV[2022.08.29.1624]
    ''' add optional Dictionary parameter to which
    ''' data from this function can be added.
    ''' (see also askUserForPartMatl)
    '''
    Dim rt As Scripting.Dictionary
    Dim ck As VbMsgBoxResult
    Dim bd As aiBoxData
    Dim tx As String
    
    If dc Is Nothing Then
        Set rt = askUserForMatlQty( _
        AiDoc, New Scripting.Dictionary)
    Else
        Set rt = dc 'New Scripting.Dictionary
        With rt
            If .Exists(pnRmQty) Then .Remove pnRmQty
            If .Exists(pnRmUnit) Then .Remove pnRmUnit
        End With
        
        If AiDoc Is Nothing Then
            rt.Add pnRmQty, 0
            rt.Add pnRmUnit, ""
        Else
            With nu_fmIfcMatlQty01().SeeUserWithPart(AiDoc)
                ''' following copied from dcGeniusPropsPartRev20180530 line 1632~?
                If .Exists(pnRmQty) Then
                    rt.Add pnRmQty, .Item(pnRmQty)
                    ''' REV[2022.08.29.1459]
                    ''' removing extraneous comments left over
                    ''' from dcGeniusPropsPartRev20180530
                Else
                    'Stop
                End If
                
                If .Exists(pnRmUnit) Then
                    rt.Add pnRmUnit, .Item(pnRmUnit)
                Else
                    'Stop
                End If
            End With
        End If
    End If
    
    Set askUserForMatlQty = rt
End Function

Public Function askUserForPartMatlUpdate( _
    AiDoc As Inventor.PartDocument _
) As Scripting.Dictionary
    '''
    ''' askUserForPartMatlUpdate --
    '''     Attempt to update Part Document
    '''     material Properties from results
    '''     of askUserForPartMatl
    '''         (Family and Material Selection)
    '''     and askUserForMatlQty
    '''         (Material Quantity and Units)
    '''     Return Dictionary of results
    '''
    ''' NOTE[2022.08.29.1627]
    ''' want to separate user data collection
    ''' from property updates in this function.
    ''' further review/development called for.
    '''
    Dim dcPr As Scripting.Dictionary
    Dim dcWk As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    Dim ck As VbMsgBoxResult
    Dim ky As Variant
    
    Set dcPr = dcOfPropsInAiDoc(AiDoc)
    ck = vbOK
    If Not dcPr.Exists(pnRawMaterial) Then
        ck = MsgBox(Join(Array( _
            "Custom Property " & pnRawMaterial & ",", _
            "used to identify Raw Material,", _
            "is not yet present in this model.", _
            "", _
            "Go ahead and create it?" _
        ), vbNewLine), _
            vbYesNo + vbQuestion, _
            "Required Property Missing!" _
        )
        
        If ck = vbYes Then
            With AiDoc.PropertySets.Item(gnCustom)
                On Error Resume Next
                
                For Each ky In Array( _
                    pnRawMaterial, pnRmQty, pnRmUnit _
                )
                Next
                
                Err.Clear
                Set pr = .Add("", ky) 'pnRawMaterial
                If Err.Number = 0 Then
                    dcPr.Add ky, pr 'pnRawMaterial
                    ck = vbOK
                Else
                    ck = vbAbort
                End If
                On Error GoTo 0
            End With
        Else
            ck = vbOK
        End If
    End If
    
    If ck <> vbOK Then
        ck = MsgBox(Join(Array( _
            "Custom Property " & pnRawMaterial & ",", _
            "was not created! Raw Material", _
            "will not be saved!" _
        ), vbNewLine), _
            vbOKCancel + vbExclamation, _
            "Property Not Created!" _
        )
    End If
    
    Set rt = New Scripting.Dictionary
    
    If ck = vbOK Then
        ''' REV[2022.08.29.1616]
        ''' condense two nearly identical With blocks
        ''' into one, combining results of part material
        ''' and material quantity data collections.
        ''' NOTE: this required additional REVs
        ''' to askUserForPartMatl (nee d4g1f1)
        ''' and askUserForMatlQty (nee d4g1f3)
        ''' to accept optional Dictionary to receive
        ''' data points collected by each function.
        With askUserForMatlQty(AiDoc, _
             askUserForPartMatl(AiDoc _
        ))
            For Each ky In .Keys
                With dcPr
                If .Exists(ky) Then
                    Set pr = .Item(ky)
                Else
                    Set pr = Nothing
                End If
                End With
                
                If Not pr Is Nothing Then
                    If Len(Trim(.Item(ky))) > 0 Then
                    If pr.Value <> .Item(ky) Then
                        On Error Resume Next
                        Err.Clear
                        'Stop 'so we can make sure this works
                        pr.Value = .Item(ky)
                        Debug.Print ; 'Breakpoint Landing
                        'DON'T try to step at pr.Value
                        If Err.Number Then
                            Stop
                        End If
                        On Error GoTo 0
                    End If
                    End If
                    rt.Add ky, pr.Value
                End If
            Next
        End With
    End If
    
    Set askUserForPartMatlUpdate = rt
End Function

Public Function dcGeniusPropsPartRev20180530_ck( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    ''' '''
    ''' NOTICE TO DEVELOPER [2021.11.12]
    ''' '''
    '''
    ''' This function definition was restored
    ''' from a prior copy of this project
    ''' (VB-000-1002_2021-1001.ipt)
    ''' to restore current "normal" operation
    ''' of the Genius Properties Update macro.
    ''' The prior development version was
    ''' retained for reference, renamed to
    ''' dcGeniusPropsPartRev20180530_ck_broken
    '''
    ''' One minor revision was made to this
    ''' restored version to retain improved
    ''' generation of Genius Mass data.
    ''' Additional changes should be kept
    ''' to a MINIMUM to maintain correct
    ''' operation going forward, and any
    ''' desired changes implemented through
    ''' some form of "shim"
    '''
    ''' '''
    Dim rt As Scripting.Dictionary
    ''' REV[2022.01.21.1351]
    ''' Added following two Dictionaries
    Dim dcIn As Scripting.Dictionary
    ''' to collect settings already in Genius
    Dim dcFP As Scripting.Dictionary
    ''' to add a layer of separation
    ''' to FlatPattern data collection
    ''' (might not want to use for Properties
    '''  so don't update immediately)
    
    ''
    Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
    ''
    Dim prPartNum   As Inventor.Property 'pnPartNum
    ''' ADDED[2021.03.11] to simplify access
    ''' to Part Number of Model, since it's
    ''' requested several times in function
    Dim prFamily    As Inventor.Property
    Dim prRawMatl   As Inventor.Property 'pnRawMaterial
    Dim prRmUnit    As Inventor.Property 'pnRmUnit
    Dim prRmQty     As Inventor.Property 'pnRmQty
    ''
    Dim pnModel     As String
    ''' ADDED[2021.03.11] to further
    ''' simplify access to Part Number
    Dim nmFamily As String
    Dim mtFamily As String
    ''' UPDATE[2018.05.30]:
    '''     Rename variable Family to nmFamily
    '''     to minimize confusion between code
    '''     and comment text in searches.
    '''     Also add variable mtFamily
    '''     for raw material Family name
    Dim pnStock As String
    Dim qtUnit As String
    Dim bomStruct As Inventor.BOMStructureEnum
    Dim ck As VbMsgBoxResult
    Dim bd As aiBoxData
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartRev20180530_ck = _
        dcGeniusPropsPartRev20180530_ck( _
            invDoc, New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        
        With invDoc
            ' Get Property Sets
            With .PropertySets
                Set aiPropsUser = .Item(gnCustom)
                Set aiPropsDesign = .Item(gnDesign)
            End With
            
            ' Get Custom Properties
            Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
            Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
            Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
            
            ' Part Number and Family properties
            ' are from Design, NOT Custom set
            Set prPartNum = aiGetProp( _
                aiPropsDesign, pnPartNum)
                'ADDED 2021.03.11
            pnModel = prPartNum.Value
            Set prFamily = aiGetProp( _
                aiPropsDesign, pnFamily)
            
            ''' We should check HERE for possibly misidentified purchased parts
            ''' UPDATE[2018.02.06]: Using new UserForm; see below
            With .ComponentDefinition
                ''' Request #1: Get the Mass in Pounds
                ''' and add to Custom Property GeniusMass
                With .MassProperties
                    ''' Update [2021.11.12]
                    '''     Round mass to nearest ten-thousandth
                    '''     to try to match expected Genius value.
                    '''     This should reduce or minimize reported
                    '''     discrepancies during ETM process.
                    Set rt = dcWithProp(aiPropsUser, pnMass, _
                        Round(cvMassKg2LbM * .Mass, 4), rt _
                    )
                End With
                
                '''
                ''' Get BOM Structure type, correcting if appropriate,
                ''' and prepare Family value for part, if purchased.
                '''
                ck = vbNo
                ''' UPDATE[2018.05.31]: Combined both InStr checks
                ''' by addition to generate a single test for > 0
                ''' If EITHER string match succeeds, the total
                ''' SHOULD exceed zero, so this SHOULD work.
                If InStr(1, invDoc.FullFileName, _
                    "\Doyle_Vault\Designs\purchased\" _
                ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                    "|" & prFamily.Value & "|" _
                ) > 0 Then
                    ''' UPDATE[2018.02.06]: Using same
                    '''     new UserForm as noted above.
                    ck = newFmTest2().AskAbout(invDoc, , _
                        "Is this a Purchased Part?" _
                        & vbNewLine & "(Cancel to debug)" _
                    )
                End If
                
                ''' Check process below replaces duplicate check/responses above.
                If ck = vbCancel Then
                    Stop
                ElseIf ck = vbYes Then
                    If .BOMStructure <> kPurchasedBOMStructure Then
                        On Error Resume Next
                        .BOMStructure = kPurchasedBOMStructure
                        If Err.Number = 0 Then
                            bomStruct = .BOMStructure
                        Else
                            bomStruct = kPurchasedBOMStructure
                            ''' WARNING: NOT a good way to go about this
                            '''     but will go with it for now
                        End If
                        On Error GoTo 0
                    Else
                        bomStruct = .BOMStructure 'to make sure this is captured
                    End If
                Else
                    bomStruct = .BOMStructure 'to make sure this is captured
                End If
                
                'Request #2: Change Cost Center iProperty.
                'If BOMStructure = Purchased and not content center,
                'then Family = D-PTS, else Family = D-HDWR.
                '
                'UPDATE[2018-05-30]: Value produced here
                'will now be held for later processing,
                'more toward the end of this function.
                If bomStruct = kPurchasedBOMStructure Then
                    If .IsContentMember Then
                        nmFamily = "D-HDWR"
                    Else
                        nmFamily = "D-PTS"
                        'NOTE: NON Content Center members
                        '       might still be D-HDWR
                        '       Additional checks might
                        '       be recommended
                    End If
                Else
                    nmFamily = ""
                End If
            End With
            ''' At this point, nmFamily SHOULD be set
            ''' to a non-blank value if Item is purchased.
            ''' We should be able to check this later on,
            ''' if Item BOMStructure is NOT Normal
            
            'Request #4: Change Cost Center iProperty.
            'If BOMStructure = Normal, then Family = D-MTO,
            'else if BOMStructure = Purchased then Family = D-PTS.
            If bomStruct = kNormalBOMStructure Then
                
                ''' REV[2022.01.28.1014]
                ''' Added initial raw material capture
                ''' to check against Genius
                ''' HOLD![2022.01.28.1046]
                ''' commenting out again
                ''' probably best below, still
                pnStock = prRawMatl.Value
                ''' REV[2022.02.08.1304]
                ''' restored, to obtain any
                ''' value already defined.
                ''' MIGHT need moved further down,
                ''' but hold off on that for now.
                
                ''' REV[2022.01.17.1123]
                ''' Start adding code to capture
                ''' any raw material items for
                ''' part already in Genius.
                ''' REV[2022.01.21.1357]
                ''' Separated capture from With statement
                ''' into new Dictionary object in order
                ''' to check and use it further down,
                ''' as well as passing it to nuSelFromDict
                ''' to handle multiple line items
                ''' REV[2022.01.31.1008]
                ''' Restored assignment of dcFromAdoRS
                ''' result to Dictionary Object dcIn,
                ''' in order to pass it to other
                ''' functions, as needed.
                '''
                Set dcIn = dcFromAdoRS(cnGnsDoyle().Execute( _
                    sqlOf_GnsPartMatl(pnModel) _
                ))
                'Debug.Print ConvertToJson(dcDxFromRecSetDc(dcIn), vbTab)
                'Stop
                'Set dcIn = dcOb(dcDxFromRecSetDc(dcIn).Item(pnRawMaterial))
                If dcIn.Count > 0 Then 'Genius found something
                    With dcOb(dcDxFromRecSetDc(dcIn).Item(pnRawMaterial))
                        ''' REV[2022.01.28.1336]
                        ''' Added code to collect captured
                        'Set dcIn = New Scripting.Dictionary
                        
                        
                        ''' REV[2022.01.28.0857]
                        ''' Added code to collect captured
                        ''' material item number, asking user
                        ''' to select from list if more than one.
                        If .Count > 0 Then 'Genius found something
                            If Len(pnStock) > 0 Then
                                'some material already assigned
                                If .Exists(pnStock) Then 'do nothing
                                    'it's a valid option; stick with it
                                    '
                                Else 'probably going to need an update
                                    'so forget current value (for now)
                                    pnStock = ""
                                End If
                            End If
                            
                            If Len(pnStock) = 0 Then
                                'grab first material item found
                                'Stop
                                'pnStock = dcOb(.Item(.Keys(0))).Item(pnRawMaterial)
                                pnStock = .Keys(0)
                                ''' REV[2022.02.08.1336]
                                ''' switched back to pulling first Key.
                                ''' since the With statement pulls the
                                ''' Dictionary keyed on raw materials,
                                ''' this SHOULD be reliably correct.
                                '''
                                ''' that the nuSelector below is called
                                ''' with the list of Keys would seem
                                ''' to support this expectation.
                            End If
                            
                            'and use it for the default...
                            If .Count > 1 Then
                                Stop 'because selection is going
                                'to be a lot more complicated.
                                '(just look at that pnStock
                                ' assignment up there!)
                                
                                pnStock = nuSelector( _
                                ).GetReply(.Keys, pnStock)
                                
                                Stop 'to make sure things are okay
                            End If
                        Else 'do nothing
                            ''' REV[2022.02.08.1353]
                            ''' disabled pnStock assignment
                            ''' from prRawMatl here, since
                            ''' this is now done automatically
                            ''' (see REV[2022.02.08.1304] above)
                            'pnStock = prRawMatl.Value
                        End If
                        
                        ''' REV[2022.01.28.0903]
                        ''' Separated Dictionary capture
                        ''' from Count check
                        If Len(pnStock) > 0 Then
                            If Len(CStr(prRawMatl.Value)) = 0 Then 'don't worry.
                                'it'll be taken care of further down
                            ElseIf pnStock = prRawMatl.Value Then 'don't worry.
                                'should only be minor quantity changes
                                'Stop 'and make sure we want to do this.
                                
                                'Set dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                'Deactivated, moved down and out of this If-Then nest.
                                'Search below for active copy
                                
                                'Debug.Print ; 'Breakpoint Landing
                            Else 'need to ask User what to go with
                                Debug.Print "=== CURRENT GENIUS MATERIAL DATA ==="
                                'Debug.Print dumpLsKeyVal(dcIn, ":" & vbTab)
                                ck = newFmTest2().AskAbout(invDoc, _
                                    "Raw Material " & prRawMatl.Value _
                                    & vbNewLine & " for Item" _
                                    , _
                                    "does not match " & pnStock _
                                    & vbNewLine & "indicated in Genius." _
                                    & vbNewLine & vbNewLine _
                                    & "Change to match Genius?" _
                                    & vbNewLine & "(Cancel to debug)" _
                                )
                                If ck = vbCancel Then
                                    Stop 'to check things out
                                ElseIf ck = vbNo Then
                                    ''' NOTE[2022.02.08.1359]
                                    ''' DO NOT DISABLE this instance
                                    ''' of the pnStock assignment!
                                    pnStock = prRawMatl.Value
                                    ''' this one implements the user's
                                    ''' decision NOT to change the
                                    ''' current material assignment,
                                    ''' by forcing pnStock back to the
                                    ''' original Value from prRawMatl
                                End If
                                
                                'Stop 'to grab current raw material item
                                ''' NOTE: Since material data already
                                ''' in Genius is now captured in dcIn,
                                ''' the following assignments are NOT
                                ''' immediately necessary.
                                'prRawMatl.Value = dcOb(.Item(.Keys(0))).Item(pnRawMaterial) 'pnStock
                                'prRmQty.Value = CStr(dcOb(.Item(.Keys(0))).Item(pnRmQty))
                                'prRmUnit.Value = dcOb(.Item(.Keys(0))).Item(pnRmUnit)
                                '= dcOb(.Item(.Keys(0))).Item("MtFamily")
                                ''' In fact, these Properties should
                                ''' first be checked against Genius
                                ''' data, and the user prompted
                                ''' to go ahead with updates.
                            End If
                            
                            ''' REV[2022.01.28.1448]
                            ''' Changed data extraction process here
                            ''' to work with form returned from dcFromAdoRS
                            '''
                            ''' NOTE! This is !!!TEMPORARY!!!
                            ''' Implemented during run time,
                            ''' some truly insane acrobatics were required
                            ''' to make it work without resetting the run.
                            ''' This code, including the With statement
                            ''' above, MUST be rewritten as soon as feasible!
                            '''
                            'Stop 'because we're doing to need to do something different
                            'Debug.Print ConvertToJson(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial), vbTab)
                            'Debug.Print ConvertToJson(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock), vbTab)
                            'Debug.Print ConvertToJson(dcOb(.Item(dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0))), vbTab)
                                'dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0)
                            'Stop
                            
                            If .Exists(pnStock) Then
                                Set dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                'This is DEFINITELY going to need a rework!
                                'But that will need a new function, most likely
                                
                                'deactivated the version below
                                'to be superceded by the one above
                                'Set dcIn = dcOb(.Item(dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0)))
                                
                                'original version, also deactivated
                                'for obvious reasons
                                'Set dcIn = .Item(pnStock)
                                
                                Debug.Print ; 'Breakpoint Landing
                                
                                'Debug.Print ConvertToJson(dcIn, vbTab)
                                'Stop
                            Else
                                Stop 'because we've got a REAL problem here!
                                ''' THINK pnStock should ALWAYS be in the
                                ''' With contextual Dictionary at this point,
                                ''' so if it isn't, something probably went
                                ''' seriously wrong
                            End If
                        Else
                            Set dcIn = New Scripting.Dictionary
                        End If
                    End With
                End If
                
                With dcIn
                If .Count = 0 Then
                    .Add "Ord", 0
                    .Add "RM", ""
                    .Add "MtFamily", ""
                    .Add "RMQTY", 0
                    .Add "RMUNIT", ""
                    '.Add "", ""
                End If: End With
                
                '----------------------------------------------------'
                If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                '----------------------------------------------------'
                '''
                ''' NOTE[2018-05-31]: At this point, we MAY wish
                ''' to check for a valid flat pattern,
                ''' and otherwise attempt to verify
                ''' an actual sheet metal design.
                '''
                    
                    ''' REV[2022.01.28.0903]
                    ''' HERE is where things start to get interesting
                    ''' Before processing Part as sheet metal,
                    ''' want to make sure it's supposed to be.
                    '''
                    ''' FIRST, check what Genius had to say
                    With dcIn
                        If .Exists("MtFamily") Then
                            mtFamily = .Item("MtFamily")
                        Else
                            mtFamily = ""
                        End If
                    End With
                    
                    If Len(mtFamily) = 0 Then 'need more info
                        ck = vbRetry
                    ElseIf mtFamily = "DSHEET" Then 'Genius says...
                        ck = vbYes
                    Else
                        ck = vbNo
                    End If
                    
                    ''' REV[2022.01.31.1335]
                    ''' Move flat pattern collection out here
                    ''' from inside the next If-Then block
                    If ck = vbNo Then 'we don't want/need flat pattern
                        Set dcFP = New Scripting.Dictionary
                        ''' we're going to want to do
                        ''' something different here
                    Else 'we MIGHT, so lets get it
                        Set dcFP = dcFlatPatVals(.ComponentDefinition)
                        ''' try to get flat pattern data
                        ''' WITHOUT mucking up Properties!
                        ''' Want to avoid dirtying file with
                        ''' changes until absolutely necessary)
                        
                        If dcFP.Exists(pnThickness) Then
                            pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                            dcFP.Add pnRawMaterial, pnStock
                            'need to change this to use the following more directly
                            'sqlSheetMetal(.ActiveMaterial.DisplayName,CDbl(dcfp.Item(pnThickness)))
                        End If
                    End If
                    Debug.Print ; 'Breakpoint Landing
                    If False Then
                        Debug.Print ConvertToJson(Array(dcIn, dcFP), vbTab)
                    End If
                    
                    If ck = vbRetry Then 'Genius check is inconclusive
                        ''' so let's see what the flat pattern can tell us
                        
                        If dcFP.Exists("mtFamily") Then
                            If dcFP.Item("mtFamily") = "DSHEET" Then
                                If dcFP.Exists("OFFTHK") Then
                                    Stop
                                    ck = newFmTest2().AskAbout(invDoc, "This Part: ", "might not be sheet metal. " & vbNewLine & vbNewLine & "Is it in fact sheet metal?")
                                    If ck = vbCancel Then
                                        ck = vbRetry
                                        Stop 'to debug
                                    End If
                                Else
                                    ck = vbYes
                                End If
                            ElseIf dcFP.Item("mtFamily") = "D-BAR" Then
                                ck = vbNo
                            Else
                                ck = vbRetry
                            End If
                        Else
                            ck = vbRetry
                        End If
                    End If
                    
                    If ck = vbRetry Then
                        Debug.Print ConvertToJson(Array(dcIn, dcFP), vbTab)
                        Stop 'so we can figure out what to do next.
                             'for now, most likely just press [F5]
                             'to continue
                    End If
                    
                    'Request #3:
                    '   Get sheet metal extent area
                    '   and add to custom property "RMQTY"
                    
                    ''' REV[2022.01.28.1556]
                    ''' change if-then-else sequence
                    ''' to check ck instead of dcIn
                    If ck = vbYes Then
                        Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                        ''' NOTE[2022.01.28.1551]:
                        ''' THIS call should change to move
                        ''' values from dcFP into Properties.
                        ''' Probably a new function, which might
                        ''' ALSO be called from dcFlatPatProps
                    ElseIf ck = vbRetry Then
                        Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                    ElseIf ck = vbNo Then 'probably
                        'don't do anything here
                    Else 'we got a problem
                        'material type detection SHOULD produce
                        'one of the three preceding values
                        
                        Stop 'and check it out
                    End If
                    
                    ''' NOTE[2018-05-30]:
                    '''     Raw Material Quantity value
                    '''     SHOULD be set upon return
                    '''     We may need to review the process
                    '''     to find an appropriate place
                    '''     to set for NON sheet metal
                    
                    'Moved to start of block to check for NON sheet metal
                    
                    'NOTE: THIS call might best be combined somehow
                    '   with the flat pattern prop pickup above.
                    '   Note especially that if dcFlatPatProps
                    '   FINDS NO .FlatPattern, then there should
                    '   BE NO sheet metal part number!
                    If prRawMatl Is Nothing Then
                        If rt.Exists("OFFTHK") Then
                        ''' NOTE[2021.12.10]:
                        '''     Believe this OFFTHK property is meant
                        '''     to capture "Sheet Metal" Parts that
                        '''     aren't actually Sheet Metal.
                        '''     This check might be needed further down.
                            ''' UPDATE[2018.05.30]:
                            '''     Restoring original key check
                            '''     and adding code for debug
                            '''     Previously changed to "~OFFTHK"
                            '''     to avoid this block and its issues.
                            '''     (Might re-revert if not prepped to fix now)
                            Debug.Print aiProperty(rt.Item("OFFTHK")).Value
                            Stop 'because we're going to need to do something with this.
                            
                            pnStock = "" 'Originally the ONLY line in this block.
                            ' A more substantial response is required here.
                            
                            If 0 Then Stop '(just a skipover)
                        Else
                            Stop 'because we don't know IF this is sheet metal yet
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                        End If
                    Else
                        ''  ACTION ADVISED[2018.09.14]:
                        ''  pnStock can probably be set
                        ''  to prRawMatl.Value and THEN
                        ''  checked for length to see
                        ''  if lookup needed.
                        ''  This might also allow us to check
                        ''  for machined or other non-sheet
                        ''  metal parts.
                        
                        ''' REV[2021.12.17]: sanity check
                        '''     Add sanity check to make sure
                        '''     any existing sheet metal stock
                        '''     number matches model specs
                        If Len(prRawMatl.Value) > 0 Then
                            ' we need to check it
                            
                            If Len(pnStock) = 0 Then
                                ''' REV[2022.01.28.1445]:
                                ''' Placed this pnStock stock assignment
                                ''' inside this If-Then block to prevent
                                ''' overriding value from Genius
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                            End If
                            ''' NOTE[2021.12.17@15:32]:
                            '''     copied this up from
                            '''     NOTE[2021.12.17@15:32]
                            '''     for use in sanity check
                            
                            ''' NOTE[2021.12.17]:
                            '''     This section simply warns the user
                            '''     that the current raw material does
                            '''     not match the recommended default,
                            '''     and offers an opportunity to fix it.
                            '''
                            '''     This is yet another quick and dirty
                            '''     "solution" that should be revised
                            ''' NOTE[2022.01.05]:
                            '''     Adding check for empty recommendation.
                            '''     Do NOT believe user should be offered
                            '''     opportunity to overwrite any current
                            '''     part number with a BLANK one. Believe
                            '''     the option to CLEAR is somewhere below.
                            If Len(pnStock) > 0 Then
                                If pnStock <> prRawMatl.Value Then
                                    'Stop
                                    
                                    ''' NOTE[2022.01.03]:
                                    '''     Following text SHOULD no longer
                                    '''     be needed. Verify function of
                                    '''     fmTest2 following, and when good,
                                    '''     disable and/or remove this block.
                                    Debug.Print "!!! NOTICE !!!"
                                    Debug.Print "Recommended Sheet Metal Stock (" & pnStock & ")"
                                    Debug.Print "does not match current Stock (" & prRawMatl.Value & ")"
                                    Debug.Print
                                    Debug.Print "To continue with no change, just press [F5]. Otherwise,"
                                    Debug.Print "press [ENTER] on the following line first to change:"
                                    Debug.Print "prRawMatl.Value = """ & pnStock & """"
                                    Debug.Print
                                    
                                    ''' NOTE[2022.01.03]:
                                    '''     Now using fmTest2(?) to prompt
                                    '''     user as in other checks (above?)
                                    ck = newFmTest2().AskAbout(invDoc, _
                                        "Suggest Sheet Metal change from" _
                                        & vbNewLine & prRawMatl.Value & " to" _
                                        & vbNewLine & pnStock & " for", _
                                        "Change it?" _
                                    )
                                    If ck = vbCancel Then
                                        Stop 'to check things out
                                    ElseIf ck = vbYes Then
                                        'Stop
                                        prRawMatl.Value = pnStock
                                    End If
                                    'Stop
                                End If
                            End If
                        ElseIf Len(pnStock) > 0 Then
                            'go ahead and assign material
                            prRawMatl.Value = pnStock
                            ''' REV[2022.02.08.1406]
                            ''' added new branch to assign pnStock,
                            ''' if not blank, to prRawMatl, if it is.
                        End If
                        
                        If Len(prRawMatl.Value) > 0 Then
                            If rt.Exists("OFFTHK") Then
                                'Stop 'and verify raw material item
                                ''' NOTE[2021.12.13]:
                                '''     OFFTHK property check added
                                '''     to catch sheet metal already
                                '''     assigned by accident.
                                ck = newFmTest2().AskAbout(invDoc, _
                                    "Assigned Raw Material " & prRawMatl.Value _
                                    & vbNewLine & " might be incorrect for ", _
                                    "Clear it?" _
                                )
                                If ck = vbCancel Then
                                    Stop 'to check things out
                                ElseIf ck = vbYes Then
                                    prRawMatl.Value = ""
                                End If
                                'Stop
                            End If
                            
                            
                            If pnStock = prRawMatl.Value Then
                                'no need to assign it again
                                Debug.Print ; 'Breakpoint Landing
                            Else 'need to check things out...
                                Debug.Print ConvertToJson(Array(pnStock, prRawMatl.Value)) 'and...
                                Stop 'before we do something stupid!
                                pnStock = prRawMatl.Value
                            End If

                            ''' The following With block copied and modified [2021.03.11]
                            ''' from elsewhere in this function as a temporary measure
                            ''' to address a stopping situation later in the function.
                            ''' See comment below for details.
                            '''
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    If pnStock <> "0" Then
                                        Stop 'because Material value likely invalid
                                        ''' REV[2022.03.01.1553]
                                        ''' embedded in check
                                        ''' for string value "0"
                                        ''' as this seems to come
                                        ''' up as a legacy issue,
                                        ''' and is readily remedied
                                        ''' in this section. No stop
                                        ''' is needed in that case.
                                    End If
                                    ''' REV[2022.02.08.1413]
                                    ''' reinstated interruption here
                                    ''' because at this point, pnStock
                                    ''' has likely already been assigned
                                    ''' to prRawMatl, so changing it here
                                    ''' is NOT likely to be productive.
                                    ''' this section will likely need
                                    ''' reconsideration, revision,
                                    ''' and/or possibly removal.
                                    ''' UPDATE[2021.12.10]:
                                    '''     added this check for OFFTHK
                                    '''     to avoid blindly adding sheet
                                    '''     metal stock to a "sheet metal"
                                    '''     part that isn't actually meant
                                    '''     to be made of sheet metal.
                                    If rt.Exists("OFFTHK") Then 'likely NOT
                                        'actual Sheet Metal, so just clear this:
                                        pnStock = ""
                                    Else
                                        pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                                        Debug.Print ; 'Breakpoint Landing
                                        ''' UPDATE[2021.12.10]:
                                        '''     embedded this call in the OFFTHK
                                        '''     check noted above. see that note
                                        '''     for details
                                        '''
                                        ''  ACTION TAKEN[2021.03.11]:
                                        ''  temporary measure to try to ensure
                                        ''  recovered material Item number is valid,
                                        ''  and if not, to fix it automatically.
                                        ''  This seeks to address a stop situation
                                        ''  later in this function, encountered
                                        ''  when the Part Number property is neither
                                        ''  blank NOR valid (typically "0"), likely
                                        ''  as a result of an uninitialized iPart property.
                                        ''  (see ACTION ADVISED[2018.09.14] elsewhere)
                                    End If
                                Else
                                    ''  This section retained from source,
                                    ''  but disabled to avoid potential issues
                                    ''  with subsequent operations, just in case
                                    ''  anything depends on mtFamily remaining
                                    ''  uninitialized up to that point.
                                    ''
                                    ''  With .Fields
                                    ''      mtFamily = .Item("Family").Value
                                    ''  End With
                                End If
                            End With
'''
''' This section likely should be removed when primary issue resolved.
'''
                        ElseIf rt.Exists("OFFTHK") Then
                        ''' UPDATE[2021.12.10]:
                        '''     another OFFTHK check added to avoid
                        '''     adding sheet metal stock by mistake.
                            pnStock = ""
                            ''' NOTE[2021.12.10]:
                            ''' by keeping this value blank, it is hoped
                            ''' to force the User to select the appropriate
                            ''' raw material item, rather than assigning
                            ''' a sheet metal item by mistake, just because
                            ''' it matches the Part's defined Thickness.
                            '''
                            ''' if the measured height of the flat pattern
                            ''' doesn't closely match the defined Thickness,
                            ''' the part is most likely NOT sheet metal.
                            '''
                            ''' it might be most appropriate to move the
                            ''' OFFTHK check outside and above the others
                            ''' in this sequence, as it likely determines
                            ''' whether ANY so-called sheet metal part
                            ''' should actually be treated as such.
                            '''
                        Else
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                            ''' UPDATE[2021.12.10]:
                            '''     just as before, embedded this call
                            '''     in another OFFTHK check for the same
                            '''     reason noted above
                            ''' NOTE[2021.12.17@15:32]:
                            '''     copying this up to ...
                        End If
                        
                        If Len(pnStock) = 0 Then
                            ''' UPDATE[2018.05.30]:
                            '''     Pulling ALL code/text from this section
                            '''     to get rid of excessive cruft.
                            '''
                            '''     In fact, reversing logic to go directly
                            '''     to User Prompt if no stock identified
                            '''
                            '''     IN DOUBLE FACT, hauling this WHOLE MESS
                            '''     RIGHT UP after initial pnStock assignment
                            '''     to prompt user IMMEDIATELY if no stock found
                            With newFmTest1()
                                If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                                
                                Set bd = nuAiBoxData().UsingInches.SortingDims( _
                                    invDoc.ComponentDefinition.RangeBox _
                                )
                                ck = .AskAbout(invDoc, _
                                    "No Stock Found! Please Review" _
                                    & vbNewLine & vbNewLine & bd.Dump(0) _
                                )
                                
                                If ck = vbYes Then
                                ''' UPDATE[2018.05.30]:
                                '''     Pulling some extraneous commented code
                                '''     from here and beginning of block
                                    With .ItemData
                                        If .Exists(pnFamily) Then
                                            nmFamily = .Item(pnFamily)
                                            Debug.Print pnFamily & "=" & nmFamily
                                        End If
                                        
                                        If .Exists(pnRawMaterial) Then
                                            pnStock = .Item(pnRawMaterial)
                                            Debug.Print pnRawMaterial & "=" & pnStock
                                        End If
                                    End With
                                    If 0 Then Stop 'Use this for a debugging shim
                                End If
                            End With
                        ElseIf Left$(pnStock, 2) = "LG" Then 'it's probably lagging
                            Debug.Print pnModel & ": PROBABLE LAGGING [" & pnStock & "]"
                            Debug.Print "  TRY TO VERIFY. IF CHANGE REQUIRED,"
                            Debug.Print "  FILL IN NEW VALUE FOR pnStock BELOW, "
                            Debug.Print "  AND PRESS ENTER ON THE LINE. WHEN "
                            Debug.Print "  READY, PRESS [F5] TO CONTINUE."
                            Debug.Print "  pnStock = """ & pnStock & """"
                            Stop
                        End If
                        
                        If Len(pnStock) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                            
                            With cnGnsDoyle().Execute( _
                                "select Family, Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16 " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]:
                                    ''  Will need to address this situation
                                    ''  in a more robust manner.
                                    ''  A more thorough query above
                                    ''  might also be called for.
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    
                                    ''' UPDATE[2021.06.18]:
                                    '''     New pre-check for Material Item
                                    '''     in Purchased Parts Family.
                                    '''     VERY basic handler simply
                                    '''     maps Material Family to D-BAR
                                    '''     to force extra processing below.
                                    '''     Further refinement VERY much needed!
                                    If mtFamily Like "?-MT*" Then
                                        'Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        Debug.Print pnModel & "[" & prRmQty.Value & qtUnit & "*" & pnStock & ": " & aiPropsDesign(pnDesc).Value & "]" ' prRawMatl.Value
                                        Stop 'FULL Stop!
                                    ElseIf mtFamily = "D-PTS" Then
                                        nmFamily = "D-RMT"
                                        Stop 'NOT SO FAST!
                                        mtFamily = "D-BAR"
                                    ElseIf mtFamily = "R-PTS" Then
                                        nmFamily = "R-RMT"
                                        Stop 'NOT SO FAST!
                                        mtFamily = "D-BAR"
                                    End If
                                    
                                    If mtFamily = "DSHEET" Then
                                        'We should be okay. This is sheet metal stock
                                        nmFamily = "D-RMT"
                                        qtUnit = "FT2"
                                        ''' UPDATE[2018.05.30]:
                                        '''     Moving part family assignment
                                        '''     to this section for better mapping
                                        '''     and updating to new Family names
                                        '''     as well as pulling up qtUnit assignment
                                    ElseIf mtFamily = "D-BAR" Then
                                        ''' UPDATE[2021.06.18]:
                                        '''     Added check for Part Family already set
                                        '''     to more properly handle new situation (above)
                                        If Len(nmFamily) = 0 Then
                                            nmFamily = "R-RMT"
                                        Else
                                            Debug.Print ; 'Breakpoint Landing
                                            'Stop
                                        End If
                                        
                                        ''' UPDATE[2022.01.11]:
                                        '''     Adding Do..Loop Until to following section
                                        '''     to allow user to retry setting material
                                        '''     quantity and units. This change made in
                                        '''     conjunction with new prompt form (below).
                                        ''' NOTE! This is FIRST instance of revision
                                        '''     Search on UPDATE text above to locate
                                        '''     the other in this function
                                        qtUnit = prRmUnit.Value '"IN"
                                        ck = vbCancel
                                        Do
                                        
                                        ''may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        '''     Will keep Stop for now
                                        '''     pending further review,
                                        '''     hopefully soon
                                        'Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        'Debug.Print CDbl(dcIn.Item(pnRmQty))
                                        ''' UPDATE[2021.03.11]: Replaced
                                        ''' aiPropsDesign.Item(pnPartNum)
                                        ''' with prPartNum (and now pnModel)
                                        ''' since it's used in several places
                                        
                                        'Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                        'Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                        'Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
                                        
                                        ''' REV[2022.02.08.1511]
                                        ''' replaced boilerplate above with new version below
                                        ''' in hopes of better presenting change options
                                        ''' in a more compact and accessible form.
                                        
                                        Debug.Print "===== CHECK AND VERIFY RAW MATERIAL QUANTITY ====="
                                        Debug.Print "  If change required, place new values at end"
                                        Debug.Print "  of lines below for prRmQty.Value and qtUnit."
                                        Debug.Print "  Press [ENTER] on each line to be changed."
                                        Debug.Print "  Press [F5] when ready to continue."
                                        Debug.Print "----- " & pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value & " -----"
                                        'Debug.Print ""
                                        
                                        ''' REV[2022.02.09.0923]
                                        ''' replication of REV[2022.02.09.0919]
                                        ''' from section below: prep to replace
                                        ''' old dimension dump operation with more
                                        ''' compact call to aiBoxData's Dump method
                                        If True Then 'go ahead and run old dump
                                        Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                                        With invDoc.ComponentDefinition.RangeBox
                                            Debug.Print _
                                            Round((.MaxPoint.X - .MinPoint.X) / cvLenIn2cm, 4); " '", _
                                            Round((.MaxPoint.Y - .MinPoint.Y) / cvLenIn2cm, 4); " '", _
                                            Round((.MaxPoint.Z - .MinPoint.Z) / cvLenIn2cm, 4); " '"
                                        End With
                                        End If
                                        
                                        With nuAiBoxData().UsingInches().UsingBox( _
                                            invDoc.ComponentDefinition.RangeBox _
                                        )
                                            Debug.Print .Dump(0)
                                        End With
                                        'Stop 'and check output against prior version
                                        
                                        ''' REV[2022.02.08.1446]
                                        ''' removed block of Debug.Print lines
                                        ''' disabled now for some time, as they
                                        ''' do not seem to have been missed.
                                        Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value); " 'in model. ";
                                        If dcIn.Exists(pnRmQty) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmQty));
                                        Debug.Print
                                        Debug.Print "qtUnit = """; qtUnit; """ 'in model. ";
                                        If dcIn.Exists(pnRmUnit) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmUnit));
                                        If dcIn.Item(pnRmUnit) <> "IN" Then Debug.Print " ( or try IN )";
                                        Debug.Print
                                        'Debug.Print "qtUnit = ""IN"""
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        Stop 'because we might want a D-BAR handler
                                        ''' Actually, we might NOT need to stop here
                                        ''' if bar stock is already selected,
                                        ''' because quantities would presumably
                                        ''' have been established already.
                                        ''' Any D-BAR handler probably needs
                                        ''' to be implemented in prior section(s)
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                                        ck = newFmTest2().AskAbout(invDoc, _
                                            "Raw Material Quantity is now " _
                                            & CStr(prRmQty.Value) & qtUnit & " for", _
                                            "If this is okay, click [YES]. Otherwise," _
                                            & vbNewLine & "click [NO] or [CANCEL] to fix." _
                                        )
                                        'Stop
                                        Loop Until ck = vbYes
                                        ''' UPDATE[2022.01.11]:
                                        '''     This is the terminal end of the
                                        '''     Do..Loop Until block noted above
                                        
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        '''     However, might need more handling here.
                                        Stop 'because we don't know WHAT to do with it
                                    End If
                                End If
                            End With
                        Else
                            If 0 Then Stop 'and regroup
                            ''' Things are looking a right royal mess
                            ''' at the moment I'm writing this comment.
                        End If
                    End If
                    
                    With prRawMatl
                        If Len(Trim$(.Value)) > 0 Then
                            If pnStock <> .Value Then
                                'Debug.Print "Raw Stock Selection"
                                'Debug.Print "  Current : " & prRawMatl.Value
                                'Debug.Print "  Proposed: " & pnStock
                                'Stop 'because we might not want to change existing stock setting
                                'if
                                ck = MsgBox( _
                                    Join(Array( _
                                        "Raw Stock Change Suggested", _
                                        "  for Item " & pnModel, _
                                        "", _
                                        "  Current : " & prRawMatl.Value, _
                                        "  Proposed: " & pnStock, _
                                        "", "Change It?", "" _
                                    ), vbNewLine), _
                                    vbYesNo, pnModel & " Stock" _
                                )
                                '"Change Raw Material?"
                                '"Suggested Sheet Metal"
                                If ck = vbYes Then .Value = pnStock
                            End If
                        Else
                            .Value = pnStock
                        End If
                    End With
                    Set rt = dcAddProp(prRawMatl, rt)
                    
                    With prRmUnit
                        If Len(.Value) > 0 Then
                            If Len(qtUnit) > 0 Then
                                If .Value <> qtUnit Then
                                    'Stop 'and check both so we DON'T
                                    'automatically "fix" the RMUNIT value
                                    
                                    ck = newFmTest2().AskAbout(invDoc, , _
                                        "Raw Material " & prRawMatl.Value _
                                        & vbNewLine & "Unit of Measure currently " _
                                        & .Value & vbNewLine & vbNewLine _
                                        & "Change to " & qtUnit & "?" _
                                        & vbNewLine & " " _
                                    )
                                    If ck = vbCancel Then
                                        Stop
                                    ElseIf ck = vbYes Then
                                        .Value = qtUnit
                                    End If
                                    If 0 Then Stop 'Ctrl-9 here to skip changing
                                End If
                            End If
                        Else 'we're setting a new quantity unit
                            .Value = qtUnit
                        End If
                    End With
                    Set rt = dcAddProp(prRmUnit, rt)
                    'Set rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                    ''' Plan to remove commented line above,
                    ''' superceded by the one above that
                    Debug.Print ; 'Another landing line
                    
                '--------------------------------------------'
                Else 'for standard Part (NOT Sheet Metal) ---'
                '--------------------------------------------'
                            ''' [2018.07.31 by AT]
                            ''' Duped following block from above
                            ''' to mod for material assignment
                            ''' to non-sheet metal part.
                            '''
                            ''' Except, this isn't enough.
                            ''' Also need the code to add
                            ''' Stock PN to Attribute RM.
                            ''' That's a whole 'nother
                            ''' block of code, and likely
                            ''' best consolidated.
                            With newFmTest1()
                                If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                                
                                ''' [2018.07.31 by AT]
                                ''' Added the following to try to
                                ''' preselect non-sheet metal stock
                                '.dbFamily.Value = "D-BAR"
                                '.lbxFamily.Value = "D-BAR"
                                ''' Doesn't quite do it.
                                'With New aiBoxData
                                'Set bd = nuAiBoxData().UsingInches.UsingBox( _
                                    invDoc.ComponentDefinition.RangeBox _
                                )
                                Set bd = nuAiBoxData().UsingInches.SortingDims( _
                                    invDoc.ComponentDefinition.RangeBox _
                                )
                                'End With
                                
                                ck = .AskAbout(invDoc, _
                                    "Please Select Stock for Machined Part" _
                                    & vbNewLine & vbNewLine & bd.Dump(0) _
                                )
                                
                                If ck = vbYes Then
                                ''' UPDATE[2018.05.30]:
                                '''     Pulling some extraneous commented code
                                '''     from here and beginning of block
                                    With .ItemData
                                        If .Exists(pnFamily) Then
                                            nmFamily = .Item(pnFamily)
                                            Debug.Print pnFamily & "=" & nmFamily
                                        End If
                                        
                                        If .Exists(pnRawMaterial) Then
                                            pnStock = .Item(pnRawMaterial)
                                            Debug.Print pnRawMaterial & "=" & pnStock
                                        End If
                                    End With
                                    If 0 Then Stop 'Use this for a debugging shim
                                    ''  We're going to need something here
                                    ''  to make sure raw material gets added
                                    ''  for non sheet metal parts, as well
                                    ''  What we're going to need to do
                                    ''  is refactor this whole bloody thing.
                                End If
                            End With
                            '''
                            '''
                            '''
'''
''' The following If block is copied
''' wholesale from sheet metal section above.
''' Some changes (to be) made to accommodate
''' machined or other non-sheet metal stock.
'''
''' Ultimately, whole mess to require refactor.
'''
                        If Len(pnStock) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                            
''' This enclosing With block should NOT be necessary
''' since the newFmTest1 above takes care of collecting
''' the Stock Family along with the Stock itself
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]:
                                    ''  Will need to address this situation
                                    ''  in a more robust manner.
                                    ''  A more thorough query above
                                    ''  might also be called for.
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
'''
''' Content formerly here moved BELOW and OUT of this section
''' as it should only require results of newFmTest1 exchange above
                                End If
                            End With
''' These closing statements moved up from below following If block
'''

'mtFamily = nmFamily 'to force "correct" behavior of following section
                                    If mtFamily = "DSHEET" Then
Stop 'because we should NOT be doing Sheet Metal in this section.
' This might require further investigation and/or development, if encountered.
                                        'We should be okay. This is sheet metal stock
                                        nmFamily = "D-RMT"
                                        qtUnit = "FT2"
                                        ''' UPDATE[2018.05.30]:
                                        '''     Moving part family assignment
                                        '''     to this section for better mapping
                                        '''     and updating to new Family names
                                        '''     as well as pulling up qtUnit assignment
                                    ElseIf mtFamily = "D-BAR" Then
                                        ''' UPDATE[2022.01.11]:
                                        '''     Adding Do..Loop Until to following section
                                        '''     to allow user to retry setting material
                                        '''     quantity and units. This change made in
                                        '''     conjunction with new prompt form (below).
                                        ''' NOTE! This is SECOND instance of revision
                                        '''     Search on UPDATE text above to locate
                                        '''     the other in this function
                                        nmFamily = "R-RMT"
                                        qtUnit = prRmUnit.Value '"IN"
                                        ck = vbCancel
                                        Do
                                        'Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                                        ''' UPDATE[2021.03.11]: Replaced
                                        ''' aiPropsDesign.Item(pnPartNum)
                                        ''' as noted above
                                        'Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                        'Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                        'Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
                                        
                                        ''' REV[2022.02.08.1521]
                                        ''' replaced boilerplate above with new version below
                                        ''' as per REV[2022.02.08.1511]
                                        
                                        Debug.Print "===== CHECK AND VERIFY RAW MATERIAL QUANTITY ====="
                                        Debug.Print "  If change required, place new values at end"
                                        Debug.Print "  of lines below for prRmQty.Value and qtUnit."
                                        Debug.Print "  Press [ENTER] on each line to be changed."
                                        Debug.Print "  Press [F5] when ready to continue."
                                        Debug.Print "----- " & pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value & " -----"
                                        'Debug.Print ""
                                        
                                        ''' REV[2022.02.09.0919]
                                        ''' prep to replace old dimension dump
                                        ''' operation with more compact call
                                        ''' to aiBoxData's Dump method
                                        If True Then 'go ahead and run old dump
                                        Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                                        ''' REV[2022.02.09.0904]
                                        ''' replicated With block from other section
                                        ''' to replace original "sprawled out" version
                                        ''' of Print statement hastily generated
                                        ''' during run time.
                                        With invDoc.ComponentDefinition.RangeBox
                                            Debug.Print _
                                            Round((.MaxPoint.X - .MinPoint.X) / cvLenIn2cm, 4); " '", _
                                            Round((.MaxPoint.Y - .MinPoint.Y) / cvLenIn2cm, 4); " '", _
                                            Round((.MaxPoint.Z - .MinPoint.Z) / cvLenIn2cm, 4); " '"
                                        End With
                                        End If
                                        
                                        With nuAiBoxData().UsingInches().UsingBox( _
                                            invDoc.ComponentDefinition.RangeBox _
                                        )
                                            Debug.Print .Dump(0)
                                        End With
                                        'Stop 'and check output against prior version
                                        
                                        ''' REV[2022.02.08.1446]
                                        ''' removed block of Debug.Print lines
                                        ''' disabled now for some time, as they
                                        ''' do not seem to have been missed.
                                        Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value); " 'in model. ";
                                        If dcIn.Exists(pnRmQty) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmQty));
                                        Debug.Print
                                        Debug.Print "qtUnit = """; qtUnit; """ 'in model.";
                                        If dcIn.Exists(pnRmUnit) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmUnit));
                                        Debug.Print " ( or try IN )"
                                        
                                        ''' REV[2022.02.08.1525]
                                        ''' replaced boilerplate below with new version
                                        ''' above in like manner to REV[2022.02.08.1446]
                                        ''' and also per REV[2022.02.08.1511]
                                        
                                        'Debug.Print "qtUnit = ""IN"""
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        'Debug.Print "PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED."
                                        'Debug.Print "PRESS ENTER/RETURN TWICE. THEN CONTINUE."
                                        'Debug.Print ""
                                        'Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                                        'Debug.Print "qtUnit = ""IN"""
                                        Debug.Print ""
                                        Stop 'because we might want a D-BAR handler
                                        ''' Actually, we might NOT need to stop here
                                        ''' if bar stock is already selected,
                                        ''' because quantities would presumably
                                        ''' have been established already.
                                        ''' Any D-BAR handler probably needs
                                        ''' to be implemented in prior section(s)
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                                        ck = newFmTest2().AskAbout(invDoc, _
                                            "Raw Material Quantity is now " _
                                            & CStr(prRmQty.Value) & qtUnit & " for", _
                                            "If this is okay, click [YES]. Otherwise," _
                                            & vbNewLine & "click [NO] or [CANCEL] to fix." _
                                        )
                                        'Stop
                                        Loop Until ck = vbYes
                                        ''' UPDATE[2022.01.11]:
                                        '''     This is the terminal end of the
                                        '''     Do..Loop Until block noted above
                                        
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        '''     However, might need more handling here.
                                        Stop 'because we don't know WHAT to do with it
                                    End If
                        Else
                            If 0 Then Stop 'and regroup
                            ''' Things are looking a right royal mess
                            ''' at the moment I'm writing this comment.
                        End If
                            
                    
                    ''' NOTE[2022.01.07.1004]:
                    '''     Another check for empty recommendation.
                    '''     (SEE NOTE[2022.01.05] elsewhere in this function)
                    '''     Again, don't want user accidentally
                    '''     clearing an existing part number.
                    If Len(pnStock) > 0 Then
                    With prRawMatl
                        If Len(Trim$(.Value)) > 0 Then
                            If pnStock <> .Value Then
                                'Debug.Print "Raw Stock Selection"
                                'Debug.Print "  Current : " & prRawMatl.Value
                                'Debug.Print "  Proposed: " & pnStock
                                'Stop 'because we might not want to change existing stock setting
                                'if
                                ck = MsgBox( _
                                    Join(Array( _
                                        "Raw Stock Change Suggested", _
                                        "  Current : " & prRawMatl.Value, _
                                        "  Proposed: " & pnStock, _
                                        "", "Change It?", "" _
                                    ), vbNewLine), _
                                    vbYesNo, "Change Raw Material?" _
                                )
                                '"Suggested Sheet Metal"
                                If ck = vbCancel Then
                                    Stop
                                ElseIf ck = vbYes Then
                                    .Value = pnStock
                                End If
                            End If
                        Else
                            .Value = pnStock
                        End If
                    End With
                    End If
                    Set rt = dcAddProp(prRawMatl, rt)
                    
                    With prRmUnit
                        If Len(.Value) > 0 Then
                            If Len(qtUnit) > 0 Then
                                If .Value <> qtUnit Then
                                    'Stop 'and check both so we DON'T
                                    'automatically "fix" the RMUNIT value
                                    
                                    ck = newFmTest2().AskAbout(invDoc, , _
                                        "Raw Material " & prRawMatl.Value _
                                        & vbNewLine & "Unit of Measure currently " _
                                        & .Value & vbNewLine & vbNewLine _
                                        & "Change to " & qtUnit & "?" _
                                        & vbNewLine & " " _
                                    )
                                    If ck = vbCancel Then
                                        Stop
                                    ElseIf ck = vbYes Then
                                        .Value = qtUnit
                                    End If
                                    If 0 Then Stop 'Ctrl-9 here to skip changing
                                End If
                            End If
                        Else 'we're setting a new quantity unit
                            .Value = qtUnit
                        End If
                    End With
                    Set rt = dcAddProp(prRmUnit, rt)
                            '''
                            '''
                            '''
                End If 'Sheetmetal vs Part
            ElseIf bomStruct = kPurchasedBOMStructure Then
                ''' As mentioned above, nmFamily
                ''' SHOULD be set at this point
                If Len(nmFamily) = 0 Then
                    If 1 Then Stop 'because we might
                    'need to check out the situation
                    nmFamily = "D-PTS" 'by default
                End If
            ElseIf bomStruct = kPhantomBOMStructure Then
                ''' REV[2022.01.17.1135]
                '''     Adding a crude handler for Phantom
                '''     Part Documents. Since they shouldn't
                '''     have subcomponents to promote, they
                '''     shouldn't have that BOM structure.
                '''     User intervention might be required.
                ck = newFmTest2().AskAbout(invDoc, _
                    "For some reason, THIS Item is marked Phantom:", _
                    "Is this okay? (Click [NO] OR [CANCEL] if not)" _
                )
                If ck = vbYes Then
                    'just let it go
                Else
                    Stop
                End If
            Else
                ''' REV[2022.01.17.1138]
                '''     Adding another handler to catch Part
                '''     Documents with an unexpected BOM Structure. Since they shouldn't
                '''     have subcomponents to promote, they
                '''     shouldn't have that BOM structure.
                '''     User intervention might be required.
                ck = newFmTest2().AskAbout(invDoc, _
                    "The following Item has an unhandled BOM Structure:", _
                    "Skip it? (Click [NO] OR [CANCEL] to review)" _
                )
                If ck = vbYes Then
                    'just let it go
                Else
                    Stop 'and let User decide what to do with it.
                    'NOTE to USER: See 'bomStruct' in the 'Locals'
                    'window ('Locals Window' under View menu)
                    'to see name of current BOM structure.
                End If
                Stop '(extraneous; disable/remove whenever)
            End If
            
            ' Get the design tracking property set,
            ' and update the Cost Center Property
            If invDoc.ComponentDefinition.IsContentMember Then
                ' Don't muck around with the Family!
            Else
                If Len(nmFamily) > 0 Then
                    On Error Resume Next
                    prFamily.Value = nmFamily
                    If Err.Number Then
                        Debug.Print "CHGFAIL[FAMILY]{'" _
                            & prFamily.Value & "' -> '" & nmFamily & "'}: " _
                            & invDoc.DisplayName & " (" & invDoc.FullDocumentName & ")"
                        If MsgBox( _
                            "Couldn't Change Family" & vbNewLine _
                            & "for Item " & invDoc.DisplayName & vbNewLine _
                            & vbNewLine & "(" & invDoc.FullDocumentName & ")" _
                            & vbNewLine & vbNewLine & "Stop to Review?", _
                            vbYesNo Or vbDefaultButton2, _
                            invDoc.DisplayName _
                        ) = vbYes Then
                            Stop
                        End If
                    Else
                    End If
                    On Error GoTo 0
                    Set rt = dcAddProp(prFamily, rt)
                    'Set rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt)
                End If
            End If
        End With
        
        Call iSyncPartFactory(invDoc) 'Backport Properties to iPart Factory
        Set dcGeniusPropsPartRev20180530_ck = rt
    End If
End Function
