

Public Function dcGeniusPropsPartRev20180530( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' WAYPOINTS (search on phrase)
    '''     (NOT Sheet Metal)
    '''
    ' (removed lines 11-41)
    Dim rt As Scripting.Dictionary
    ''' REV[2022.01.21.1351] (lines 43-44)
    Dim dcIn As Scripting.Dictionary
    ''' to collect settings already in Genius
    Dim dcFP As Scripting.Dictionary
    ' (removed lines 48-51)
    
    ''
    Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
    ''
    ''' ADDED[2021.03.11] (lines 58-60)
    Dim prPartNum   As Inventor.Property 'pnPartNum
    Dim prFamily    As Inventor.Property
    '''
    ''' REV[2023.05.23.1134]
    ''' introduce break between built-in
    ''' and user-defined properties to
    ''' better distinguish the two groups
    '''
    ''' also move ADDED note (above)
    ''' before prPartNum to pull
    ''' the two built-ins together
    '''
    Dim prRawMatl   As Inventor.Property 'pnRawMaterial
    Dim prRmUnit    As Inventor.Property 'pnRmUnit
    Dim prRmQty     As Inventor.Property 'pnRmQty
    ''
    Dim pnModel     As String
    ''' ADDED[2021.03.11] (lines 67-68)
    Dim nmFamily As String
    Dim mtFamily As String
    ''' REV[2022.05.05.1110] (lines 71-85)
    Dim qtRawMatl   As Double
    Dim pnStock     As String
    Dim qtUnit      As String
    Dim bomStruct   As Inventor.BOMStructureEnum
    Dim ck          As VbMsgBoxResult
    Dim bd          As aiBoxData
    ''' REV[2022.09.29.1448]
    ''' added String variable txTmp
    ''' as temporary text holder
    ''' initially for lagging assignment (see below)
    ''' but potentially useful in other places
    Dim txTmp As String
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartRev20180530 = _
        dcGeniusPropsPartRev20180530( _
            invDoc, New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        
        With invDoc
            ''' REV[2022.05.06.1113] (lines 102-112)
            If .ComponentDefinition.IsContentMember Then
                'Stop
            End If
            
            ' Get Property Sets
            With .PropertySets
                Set aiPropsUser = .Item(gnCustom)
                Set aiPropsDesign = .Item(gnDesign)
            End With
            
            ' Get Custom Properties...
            ''' REV[2022.05.06.1124] (lines 124-130)
            If .ComponentDefinition.IsContentMember Then
                pnStock = ""
                qtRawMatl = 0#
                qtUnit = ""
            Else
                ''' REV[2023.05.23.1148]
                ''' move collection of user-defined properties
                ''' inside If block for normal BOM structure.
                ''' this SHOULD prevent unecessary creation
                ''' of these properties in purchased parts,
                ''' in particular.
                'Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                'Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                'Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                
                '' ...and their initial Values
                'If prRawMatl Is Nothing Then
                '''' REV[2022.05.10.1427] (lines 142-144)
                '    pnStock = ""
                'Else
                '    pnStock = prRawMatl.Value
                'End If
                
                '''' REV[2022.05.05.1517] (lines 149-152)
                'If prRmQty Is Nothing Then
                '    qtRawMatl = 0#
                'ElseIf IsNumeric(prRmQty.Value) Then
                '    qtRawMatl = Round(prRmQty.Value, 4)
                '    ''' REV[2023.01.16.1605]
                '    ''' added Round function to raw
                '    ''' material quantity extraction
                'Else
                '    qtRawMatl = 0#
                'End If
                
                'If prRmUnit Is Nothing Then
                '    qtUnit = ""
                'Else
                '    qtUnit = prRmUnit.Value
                'End If
                ''' REV[2022.05.05.1128]
                ''' added initial Value collection
                ''' for custom Raw Material Properties
            End If
            
            ' Part Number and Family properties
            ' are from Design, NOT Custom set
            Set prPartNum = aiGetProp( _
                aiPropsDesign, pnPartNum)
                'ADDED 2021.03.11
            pnModel = prPartNum.Value
            Set prFamily = aiGetProp( _
                aiPropsDesign, pnFamily)
            ''' REV[2022.05.05.1551] (lines 179-185)
            nmFamily = famVsGenius(pnModel, prFamily.Value)
            
            ''' REV[2022.06.29.1351] (lines 188-225)
            
            ''' We should check HERE for possibly misidentified purchased parts
            ''' UPDATE[2018.02.06]: Using new UserForm; see below
            With .ComponentDefinition
                ''' Request #1: Get the Mass in Pounds
                ''' and add to Custom Property GeniusMass
                With .MassProperties
                    ''' REV[2021.11.12] (lines 233-241)
                    On Error Resume Next
                    Set rt = dcWithProp(aiPropsUser, pnMass, _
                        Round(cvMassKg2LbM * .Mass, 4), rt _
                    )
                    If Err.Number Then
                        ' (removed lines 247-260)
                        Stop
                    End If
                    On Error GoTo 0
                End With
                
                '''
                ''' Get BOM Structure type, correcting if appropriate,
                ''' and prepare Family value for part, if purchased.
                '''
                ck = vbNo
                ''' REV[2022.05.06.1118] (lines 271-273)
                If .IsContentMember Then
                    ck = vbYes
                ''' REV[2023.01.16.1615]
                ''' added D-BAR and DSHEET to matching lists
                ''' to categorize those as purchased parts
                ElseIf InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|D-BAR|DSHEET|", _
                    "|" & nmFamily & "|" _
                ) > 0 Then
                ''' REV[2022.06.29.1416] (lines 277-281)
                    ck = vbYes
                ElseIf InStr(1, invDoc.FullFileName, _
                    "\Doyle_Vault\Designs\purchased\" _
                ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|D-BAR|DSHEET|", _
                    "|" & nmFamily & "|" _
                ) > 0 Then
                ''' REV[2022.05.06.1118] (lines 288-299)
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
                'REV[2018.05.30] (lines 331-334)
                ''' REV[2023.01.16.1618]
                ''' embedded default nmFamily assignment
                ''' in length check on nmFamily to avoid
                ''' overwriting nonblank value,
                ''' such as from Genius
                If Len(nmFamily) = 0 Then
                    If bomStruct = kPurchasedBOMStructure Then
                        If .IsContentMember Then
                        ''' NOTE[2022.05.06.1130] (lines 337-340)
                            nmFamily = "D-HDWR"
                        Else
                            nmFamily = "D-PTS"
                            ' (removed lines 344-347)
                        End If
                    Else
                        ''' REV[2020.05.05.1559] (lines 350-355)
                    End If
                End If
            End With
            ' (remove lines 358-361)
            
            'Request #4: Change Cost Center iProperty.
            'If BOMStructure = Normal, then Family = D-MTO,
            'else if BOMStructure = Purchased then Family = D-PTS.
            If bomStruct = kNormalBOMStructure Then
                ''' REV[2023.05.23.1148]
                '''
                ''' move collection of user-defined properties into
                ''' start of If block for bomStruct = kNormalBOMStructure
                ''' to avoid unecessary creation of these properties
                ''' where not needed; specifically in purchased parts.
                '''
                ''' search on REV tag above to find original location
                ''' source there remains in commented form, pending removal
                '''
                Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                
                ' ...and their initial Values
                If prRawMatl Is Nothing Then
                ''' REV[2022.05.10.1427] (lines 142-144)
                    pnStock = ""
                Else
                    pnStock = prRawMatl.Value
                End If
                
                ''' REV[2022.05.05.1517] (lines 149-152)
                If prRmQty Is Nothing Then
                    qtRawMatl = 0#
                ElseIf IsNumeric(prRmQty.Value) Then
                    qtRawMatl = Round(prRmQty.Value, 4)
                    ''' REV[2023.01.16.1605]
                    ''' added Round function to raw
                    ''' material quantity extraction
                Else
                    qtRawMatl = 0#
                End If
                
                If prRmUnit Is Nothing Then
                    qtUnit = ""
                Else
                    qtUnit = prRmUnit.Value
                End If
                ''' REV[2022.05.05.1128]
                ''' added initial Value collection
                ''' for custom Raw Material Properties
                ''''''
                ''''''
                ''' END of REV[2023.05.23.1148]
                ''''''
                ''''''
                
                ''' REV[2022.01.28.1014] (lines 368-373)
                pnStock = prRawMatl.Value
                ''' REV[2022.02.08.1304] (lines 375-396)
                Set dcIn = dcFromAdoRS(cnGnsDoyle().Execute( _
                    sqlOf_GnsPartMatl(pnModel) _
                ))
                ' (removed lines 400-402)
                If dcIn.Count > 0 Then 'Genius found something
                    With dcOb(dcDxFromRecSetDc(dcIn).Item(pnRawMaterial))
                        ''' REV[2022.01.28.1336] (lines 405-413)
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
                                ' (removed lines 427-429)
                                pnStock = .Keys(0)
                                ''' REV[2022.02.08.1336] (lines 431-439)
                            End If
                            
                            'and use it for the default...
                            If .Count > 1 Then
                                ''' REV[2022.04.14.1131] (lines 444-447)
                                Debug.Print pnModel & vbNewLine & vbTab & Join(.Keys, vbNewLine & vbTab) 'prPartNum.Value
                                ' (removed lines 449-452)
                                
                                pnStock = nuSelector( _
                                ).GetReply(.Keys, pnStock)
                                
                                Debug.Print "Selected " & IIf(Len(pnStock) > 0, pnStock, "(nothing)")
                                Stop 'to make sure things are okay
                            End If
                        Else 'do nothing
                            ''' REV[2022.02.08.1353] (lines 461-466)
                        End If
                        
                        ''' REV[2022.01.28.0903]
                        ''' Separated Dictionary capture
                        ''' from Count check
                        If Len(pnStock) > 0 Then
                            If Len(CStr(prRawMatl.Value)) = 0 Then 'don't worry.
                                'it'll be taken care of further down
                                Debug.Print ; 'Breakpoint Landing
                            ElseIf pnStock = prRawMatl.Value Then 'don't worry.
                                ' (removed lines 477-483)
                                Debug.Print ; 'Breakpoint Landing
                            Else 'need to ask User what to go with
                                ' (removed lines 486-506)
                                ck = vbOK
                                If ck = vbCancel Then
                                    Stop 'to check things out
                                ElseIf ck = vbNo Then
                                    ''' NOTE[2022.02.08.1359]
                                    ''' DO NOT DISABLE this instance
                                    ''' of the pnStock assignment!
                                    pnStock = prRawMatl.Value
                                    ' (removed lines 515-519)
                                End If
                                ' (removed lines 521-534)
                            End If
                            ''' REV[2022.01.28.1448] (lines 536-554)
                            If .Exists(pnStock) Then
                                Set dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                ' (removed lines 557-567)
                                Debug.Print ; 'Breakpoint Landing
                                ' (removed lines 569-571)
                            Else
                                Stop 'because we've got a REAL problem here!
                                ' (removed lines 574-581)
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
                ''' NOTE[2018-05-31] (602-608)
                    ''' REV[2022.01.28.0903] (609-614)
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
                        ck = vbRetry 'vbNo
                        ''' REV[2023.02.15.1442]
                        ''' change vbNo to vbRetry
                        ''' to prevent short-circuit
                        ''' of sheet metal verification
                        ''' see next If ck = vbNo (below)
                    End If
                    
                    ''' REV[2022.01.31.1335] (lines 631-633)
                    If ck = vbNo Then 'we don't want/need flat pattern
                        Set dcFP = New Scripting.Dictionary
                        ''' we're going to want to do
                        ''' something different here
                    Else 'we MIGHT, so lets get it
                        ''' NOTE[2022.04.12.1157] (lines 639-644)
                        Set dcFP = dcFlatPatVals(.ComponentDefinition)
                        ' (removed lines 646-649)
                        
                        ''' REV[2023.02.15.1503]
                        ''' added duplication of pnStock to txTmp
                        ''' in order to check for changes.
                        ''' user prompt added to this segment
                        ''' immediately disabled as redundant.
                        ''' most changes to this segment can
                        ''' PROBABLY be reverted
                        ''' REV[2023.04.21.1503]
                        ''' re-enabling user prompt (see below)
                        ''' REV[2023.04.24.0928]
                        ''' overhauling -- see below
                        If dcFP.Exists(pnThickness) Then
                            ''' REV[2023.04.24.0929]
                            ''' switched assignment of projected
                            ''' sheet metal item directly to txTmp
                            ''' instead of using it as a placeholder
                            ''' for pnStock. that way, only one
                            ''' assignment is required at this stage
                            txTmp = ptNumShtMetal(invDoc.ComponentDefinition) 'pnStock
                            'pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                            
                            ''' NOTE[2022.05.31.1158] (lines 653-662)
                            If Len(pnStock) = 0 Then
                                ''' REV[2023.04.24.0949]
                                ''' added check for blank pnStock
                                ''' with automatic assignment from
                                ''' txTmp in that case
                                '''
                                pnStock = txTmp
                            ElseIf Len(txTmp) = 0 Then
                                Stop
                                'pnStock = InputBox("", "Need material for " & pnModel, pnStock)
                            ElseIf pnStock <> txTmp Then
                                'change confirmation code
                                'duplicated with modification
                                'from NOTE[2022.01.03] (below)
                                '
                                Debug.Print ; 'Breakpoint Landing
                                ''' REV[2023.04.21.1503]
                                ''' re-enabling user prompt to fix
                                ''' issue with potential changes not
                                ''' getting picked up
                                ''' also switching pnStock with txTmp
                                ''' based on assignments above, they
                                ''' were in the wrong order in the prompt
                                ''' REV[2023.04.21.1526] disabled AGAIN
                                ''' and changed check [REV:1527] below
                                ck = newFmTest2().AskAbout(invDoc, _
                                    "Suggest Material change from" _
                                    & vbNewLine & pnStock & " to" _
                                    & vbNewLine & txTmp & " for", _
                                    "Change it?" _
                                ) 'vbYes '
                                If ck = vbCancel Then
                                    'Debug.Print ConvertToJson(nuDcPopulator.Setting(pnModel, nuDcPopulator.Setting("from", prRawMatl.Value).Setting("into", pnStock).Dictionary).Dictionary, vbTab)
                                    Stop 'to check things out
                                    'send2clipBdWin10 ConvertToJson(nuDcPopulator.Setting(pnModel,nuDcPopulator.Setting("from",prRawMatl.Value).Setting("into",pnStock).Dictionary).Dictionary,vbTab)
                                ElseIf ck = vbYes Then 'vbNo '
                                ''' REV[2023.04.21.1527] switch check
                                ''' from YES to NO to force new stock
                                ''' number. it SHOULD get picked up
                                ''' and prompted toward the end.
                                    pnStock = txTmp
                                End If
                                
                                'Stop
                            End If
                            txTmp = ""
                            dcFP.Add pnRawMaterial, pnStock
                            ' (removed lines 668-671)
                        Else 'can't be sure about existing raw stock
                            'so clear it for now
                            pnStock = ""
                            'this doesn't seem to help
                            'it gets set again further down
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
                                    'Stop
                                    ck = newFmTest2().AskAbout(invDoc, _
                                        "This Part: ", "might not be sheet metal. " _
                                        & vbNewLine & vbNewLine _
                                        & "Is it in fact sheet metal?" _
                                    )
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
                        Debug.Print ConvertToJson(Array(dcIn, dcFP, pnModel), vbTab) 'prPartNum.Value
                        Stop 'so we can figure out what to do next.
                             'for now, most likely just press [F5]
                             'to continue
                    End If
                    
                    'Request #3:
                    '   Get sheet metal extent area
                    '   and add to custom property "RMQTY"
                    
                    ''' REV[2022.01.28.1556] (lines 724-726)
                    If ck = vbYes Then
                        Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                        ''' NOTE[2022.01.28.1551] (lines 729-733)
                    ElseIf ck = vbRetry Then
                        Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                    ElseIf ck = vbNo Then 'probably
                        'don't do anything here
                        ''' #NOTE[2022.04.12.1202] (lines 738-741)
                    Else 'we got a problem
                        '(removed lines 743-745)
                        Stop 'and check it out
                    End If
                    
                    ''' NOTE[2018-05-30] (lines 749-762)
                    If prRawMatl Is Nothing Then
                        If rt.Exists("OFFTHK") Then
                        ''' NOTE[2021.12.10] (lines 765-769)
                            ''' UPDATE[2018.05.30] (lines 770-775)
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
                        ''  ACTION ADVISED[2018.09.14] (lines 788-796)
                        ''' REV[2021.12.17] (lines 797-800)
                        If Len(prRawMatl.Value) > 0 Then
                            ' we need to check it
                            
                            If Len(pnStock) = 0 Then
                                ''' REV[2022.01.28.1445] (lines 805-808)
                                pnStock = ptNumShtMetal(.ComponentDefinition)
                                ''' NOTE[2022.05.31.1149] (lines 810-816)
                                ''' NOTE[2022.04.12.1258] (lines 817-821)
                            End If
                            ''' NOTE[2021.12.17@15:32] (lines 823-827)
                            ''' NOTE[2021.12.17] (lines 828-835)
                            ''' NOTE[2022.01.05] (lines 836-841)
                            If Len(pnStock) > 0 Then
                                If pnStock <> prRawMatl.Value Then
                                    'Stop
                                    ''' REV[2022.04.19.0945] (lines 845-861)
                                    ''' REV[2022.04.19.0944] (lines 862-865)
                                    If UCase$(prRawMatl.Value) = pnStock Then
                                        ck = vbYes
                                    Else
                                    ''' NOTE[2022.01.03] (lines 869-871)
                                        ck = newFmTest2().AskAbout(invDoc, _
                                            "Suggest Sheet Metal change" _
                                            & vbNewLine & "from " & prRawMatl.Value _
                                            & vbNewLine & "  to " & pnStock & " for", _
                                            "Change it?" _
                                        )
                                    End If
                                    
                                    If ck = vbCancel Then
                                        Debug.Print ConvertToJson(nuDcPopulator.Setting(pnModel, nuDcPopulator.Setting("from", prRawMatl.Value).Setting("into", pnStock).Dictionary).Dictionary, vbTab)
                                        Stop 'to check things out
                                        'send2clipBdWin10 ConvertToJson(nuDcPopulator.Setting(pnModel,nuDcPopulator.Setting("from",prRawMatl.Value).Setting("into",pnStock).Dictionary).Dictionary,vbTab)
                                    ElseIf ck = vbYes Then
                                        On Error Resume Next
                                        Err.Clear
                                        prRawMatl.Value = pnStock
                                        If Err.Number Then
                                            Stop 'and check for Member not Found
                                            'could indicate a problem assigning
                                            'a value to an iPart Table Cell,
                                            'like, say, it's presently set
                                            'to use a formula
                                        End If
                                        On Error GoTo 0
                                    End If
                                    'Stop
                                End If
                            End If
                        ElseIf Len(pnStock) > 0 Then
                            'go ahead and assign material
                            On Error Resume Next
                            Err.Clear
                            prRawMatl.Value = pnStock
                            If Err.Number Then
                                Stop
                            End If
                            On Error GoTo 0
                            ''' REV[2023.01.19.0906]
                            ''' embedded raw material Property
                            ''' assignment in error trap block
                            ''' REV[2022.02.08.1406]
                            ''' added new branch to assign pnStock,
                            ''' if not blank, to prRawMatl, if it is.
                        End If
                        
                        If Len(prRawMatl.Value) > 0 Then
                            If rt.Exists("OFFTHK") Then
                                'Stop 'and verify raw material item
                                ''' NOTE[2021.12.13] (lines 902-905)
                                ck = newFmTest2().AskAbout(invDoc, _
                                    "Assigned Raw Material " & prRawMatl.Value _
                                    & vbNewLine & " might be incorrect for ", _
                                    "Clear it?" _
                                )
                                If ck = vbCancel Then
                                    Stop 'to check things out
                                ElseIf ck = vbYes Then
                                    prRawMatl.Value = ""
                                    pnStock = prRawMatl.Value
                                    ''' REV[2022.03.16.1555] (lines 916-920)
                                End If
                                'Stop
                            End If
                            
                            
                            If pnStock = prRawMatl.Value Then
                                'no need to assign it again
                                Debug.Print ; 'Breakpoint Landing
                            Else 'need to check things out...
                                Debug.Print ConvertToJson(Array(pnStock, prRawMatl.Value)) 'and...
                                'Stop 'before we do something stupid!
                                pnStock = prRawMatl.Value
                            End If

                            ''' The following With block copied and modified [2021.03.11]
                            ''' from elsewhere in this function as a temporary measure
                            ''' to address a stopping situation later in the function.
                            ''' See comment below for details.
                            '''
                            With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
                                "vgMfiItems", "Family", "Item", pnStock _
                            ))
                            'With cnGnsDoyle().Execute( _
                                "select Family from vgMfiItems where Item='" _
                                & Replace(pnStock, "'", "''") & "';" _
                            )
                            ''' REV[2022.08.26.1055] (lines 947-950)
                                If .BOF Or .EOF Then
                                    If pnStock <> "0" Then
                                    ''' REV[2022.03.01.1553] (lines 953-960)
                                        If Len(pnStock) > 0 Then
                                        ''' REV[2022.07.07.1340]
                                        ''' added secondary check for string length.
                                        ''' an empty string requires no user attention.
                                            Stop 'because Material value likely invalid
                                        End If
                                    End If
                                    ''' REV[2022.02.08.1413] (lines 968-976)
                                    ''' UPDATE[2021.12.10] (lines 977-982)
                                    If rt.Exists("OFFTHK") Then 'likely NOT
                                        'actual Sheet Metal, so just clear this:
                                        pnStock = ""
                                    Else
                                        pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                                        Debug.Print ; 'Breakpoint Landing
                                        ''' UPDATE[2021.12.10] (lines 989-993)
                                        ''  ACTION TAKEN[2021.03.11] (lines 994-1003)
                                    End If
                                Else
                                    ' (removed lines 1006-1014)
                                End If
                            End With
                            ' WAYPOINT ALERT (lines 1017-1019)
                        ElseIf rt.Exists("OFFTHK") Then
                        ''' UPDATE[2021.12.10] (lines 1021-1023)
                            pnStock = ""
                            ''' NOTE[2021.12.10] (lines 1025-1040)
                            '''
                        Else
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                            ''' UPDATE[2021.12.10] (lines 1044-1049)
                        End If
                        
                        If Len(pnStock) = 0 Then
                            ''' UPDATE[2018.05.30] (lines 1053-1062)
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
                                ''' UPDATE[2018.05.30] (lines 1075-1077)
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
                            ''' NOTE[2022.05.10.1559] (lines 1093-1097)
                            Debug.Print pnModel & ": PROBABLE LAGGING [" & pnStock & "]"
                            Debug.Print "  TRY TO VERIFY. IF CHANGE REQUIRED,"
                            Debug.Print "  FILL IN NEW VALUE FOR pnStock BELOW, "
                            Debug.Print "  AND PRESS ENTER ON THE LINE. WHEN "
                            Debug.Print "  READY, PRESS [F5] TO CONTINUE."
                            'Debug.Print "  pnStock = """ & pnStock & """"
                            
                            ck = vbNo
                            Do
                                txTmp = Trim(InputBox( _
                                    Join(Array( _
                                        "Item " & pnModel & " appears", _
                                        "to be lagging, likely " & pnStock & ".", _
                                        "Try to verify, and if not correct,", _
                                        "fill in correct material item below.", _
                                        "", _
                                        "(WARNING! update NOT working yet!", _
                                        " Program will stop when entry complete", _
                                        " to permit manual update)" _
                                    ), vbNewLine), _
                                    "Verify Lagging " & pnStock & " for " & pnModel, _
                                    pnStock _
                                ))
                                Debug.Print "  pnStock = """ & txTmp & """"
                                If Len(txTmp) > 0 Then
                                    If txTmp = pnStock Then
                                        ck = MsgBox("Go ahead with " & txTmp & "?", vbYesNoCancel, "Confirm Same Material")
                                        If ck = vbNo Then ck = vbRetry
                                    Else
                                        ck = MsgBox( _
                                            Join(Array( _
                                                "Change Lagging Material ", _
                                                pnStock & " to " & txTmp & "?" _
                                            ), vbNewLine), _
                                            vbYesNoCancel, _
                                            "Confirm Material Change" _
                                        )
                                    End If
                                Else
                                    'ck = MsgBox(Join(Array("Input appears to have been cleared.", "Are you sure you want to remove the", "current material," & pnStock & "?", "", "([Cancel] to debug)"), vbNewLine), vbYesNoCancel, "Remove Material?")
                                    '"Select [No] to keep it.",
                                    '
                                    ck = MsgBox( _
                                        Join(Array( _
                                            "No material entered.", _
                                            "(perhaps entry was canceled?)", _
                                            "", "Do you wish to remove the current", _
                                            "material, " & pnStock & ", without replacement?" _
                                        ), vbNewLine), _
                                        vbYesNoCancel, _
                                        "No Material!" _
                                    )
                                    ''' WARNING!!![2022.10.17.1434]
                                    ''' there's something screwy going on here.
                                    ''' Pressing [F8] in debug mode on the line above
                                    ''' SHOULD stop on the next If statement, below.
                                    ''' Instead, execution continues straight through
                                    ''' to the Stop statement further down (which
                                    ''' is SUPPOSED to go away, eventually!)
                                    '''
                                    ''' It's not clear what might be causing this,
                                    ''' or what it might take to regain expected behavior.
                                    ''' For now, have added a Breakpoint Landing
                                    ''' in a crude attempt to address the matter.
                                    '''
                                    Debug.Print ; 'Breakpoint Landing
                                    '''
                                    ''' and THAT doesn't seem to be helping.
                                    ''' will have to look into this more later.
                                    
                                    If ck = vbNo Then
                                        ck = MsgBox( _
                                            "Do you want to keep " & pnStock & "?", _
                                            vbYesNoCancel, "Keep Current?" _
                                        )
                                        If ck = vbNo Then
                                            ck = vbRetry 'to bypass debug below
                                        ElseIf ck = vbYes Then
                                            ck = vbNo
                                        End If
                                    End If
                                End If
                                'Stop
                                
                                If ck = vbCancel Then
                                    Stop
                                ElseIf ck = vbRetry Then
                                    ck = vbCancel 'to force retry
                                ElseIf ck = vbYes Then
                                    pnStock = txTmp
                                End If
                            Loop While ck = vbCancel
                        End If
                        
                        If Len(pnStock) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                            Debug.Print ; 'Breakpoint Landing'Stop 'WAYPOINT to check block WITH@1764
                            ''' REV[2022.08.26.1001]
                            ''' placing temporary Stops at start
                            ''' and end of following With block
                            ''' to check use of fields normally
                            ''' requested in SQL select statement.
                            '''
                            With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
                                "vgMfiItems", "Family", "Item", pnStock _
                            ))
                            'preceding (disabled) With statement
                            'to replace the following, assuming
                            'tests prove successful. if so, it
                            'might permit further streamlining
                            'With cnGnsDoyle().Execute( _
                                "select Family from vgMfiItems where Item='" & _
                                Replace(pnStock, "'", "''") & "';" _
                            ) ', Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16
                            ''' REV[2022.08.26.1059]
                            ''' (duping REV[2022.08.26.1055] above)
                            ''' replaced direct ref to pnStock
                            ''' with Replace operation to "escape"
                            ''' it, re REV[2022.08.19.1416] (below)
                            ''' REV[2022.08.26.1001] NOTE
                            ''' it is known that field Family
                            ''' is used directly below, however,
                            ''' usage of other fields is unclear.
                            ''' '
                            ''' to check their necessity, they
                            ''' have been removed from the SQL
                            ''' source string to a commend after
                            ''' the SQL call statement, to be
                            ''' recovered as needed.
                            ''' '
                            ''' Stops have been placed just before
                            ''' this With block (above), and just
                            ''' before its End (below) to mark both
                            ''' entry and exit from this block.
                            ''' in this way, it is hoped the critical
                            ''' period of execution may be delineated.
                            ''' '
                            ''' assuming no errors are encountered
                            ''' between entry and exit from this block,
                            ''' it may be assumed that no other fields
                            ''' but Family are required here, and they
                            ''' can likely be removed without harm.
                            ''' '
                            ''' this should permit replacement of the
                            ''' "hard-coded" SQL statement with a call
                            ''' to the new Function sqlOf_simpleSelWhere
                            '''
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14] (lines 1162-1166)
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    
                                    ''' UPDATE[2021.06.18] (lines 1172-1178)
                                    ''' REV[2022.04.15.1035] (lines 1179-1185)
                                    If mtFamily Like "?-MT*" Then
                                        ' (removed lines 1187-1188)
                                        Debug.Print pnModel & "[" _
                                            & qtRawMatl & qtUnit _
                                            & " of " & pnStock & ": " _
                                            & aiPropsDesign(pnDesc).Value _
                                            & "]" 'prRmQty.Value prRawMatl.Value
                                        Stop 'FULL Stop!
                                    ''' NOTE[2022.05.05.1603]
                                    ''' new ElseIf branch called for here
                                    ''' see corresponding block under
                                    ''' Standard Part branch.
                                    ElseIf mtFamily = "D-PTS" Then
                                        'Stop 'NOT SO FAST!
                                        mtFamily = "D-BAR"
                                        'nmFamily = "D-RMT"
                                    ElseIf mtFamily = "R-PTS" Then
                                        Stop 'NOT SO FAST!
                                        mtFamily = "D-BAR"
                                        'nmFamily = "R-RMT"
                                    End If
                                    
                                    If mtFamily = "DSHEET" Then
                                        'We should be okay. This is sheet metal stock
                                        nmFamily = "D-RMT"
                                        qtUnit = "FT2"
                                        ''' UPDATE[2018.05.30] (lines 1213-1217)
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
                                        
                                        ''' UPDATE[2022.01.11] (lines 1229-1236)
                                        qtUnit = prRmUnit.Value '"IN"
                                        ck = vbCancel
                                        Do
                                        
                                        ''may want function here
                                        ''' UPDATE[2018.05.30]: (lines 1242-1272)
                                        ''' REV[2022.02.09.0923] (lines 1273-1277)
                                        If False Then 'go ahead and run old dump
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
                                        
                                        ''' REV[2022.02.08.1446] (lines 1295-1299)
                                        Debug.Print "qtRawMatl = "; CStr(qtRawMatl); " 'in model. ";
                                        If dcIn.Exists(pnRmQty) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmQty));
                                        Debug.Print
                                        Debug.Print "qtUnit = """; qtUnit; """ 'in model. ";
                                        If dcIn.Exists(pnRmUnit) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmUnit));
                                        If dcIn.Item(pnRmUnit) <> "IN" Then Debug.Print " ( or try IN )";
                                        Debug.Print
                                        ' (removed lines 1307-1314)
                                        With nu_fmIfcMatlQty01().SeeUser(invDoc) '.Result()
                                            If .Exists(pnRmQty) Then
                                                ''' REV[2022.04.04.1404] (lines 1317-1323)
                                                If CDbl("0" & CStr(qtRawMatl)) _
                                                 = CDbl(.Item(pnRmQty)) Then 'prRmQty.Value => qtRawMatl
                                                    'no change -- don't change
                                                Else
                                                    'Debug.Print "prRmQty.Value FROM " & prRmQty.Value & " TO " & .Item(pnRmQty)
                                                    Debug.Print "qtRawMatl FROM " & qtRawMatl & " TO " & .Item(pnRmQty)
                                                    
                                                    'Stop 'and double-check
                                                    'might still be equivalent
                                                    qtRawMatl = .Item(pnRmQty) 'prRmQty.Value
                                                End If
                                            Else
                                                Stop
                                            End If
                                            
                                            If .Exists(pnRmUnit) Then
                                                If qtUnit = .Item(pnRmUnit) Then
                                                    'no change -- don't change
                                                Else
                                                    Debug.Print "qtUnit FROM " & qtUnit & " TO " & .Item(pnRmUnit)
                                                    'Stop 'and double-check
                                                    'might still be equivalent
                                                    qtUnit = .Item(pnRmUnit)
                                                End If
                                            Else
                                                Stop
                                            End If
                                        End With
                                        ' (removed lines 1352-1358)
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(qtRawMatl); qtUnit; ". IF OKAY, CONTINUE." 'prRmQty.Value
                                        ck = newFmTest2().AskAbout(invDoc, _
                                            "Raw Material Quantity is now " _
                                            & CStr(qtRawMatl) & qtUnit & " for", _
                                            "If this is okay, click [YES]." _
                                            & vbNewLine & "Otherwise, click [NO] to review." _
                                            & vbNewLine & "" _
                                            & vbNewLine & "( for debug, click [CANCEL] )" _
                                        ) 'prRmQty.Value
                                        If ck = vbCancel Then
                                            Stop 'to debug
                                        End If
                                        Loop Until ck = vbYes
                                        ''' UPDATE[2022.01.11]:
                                        '''     This is the terminal end of the
                                        '''     Do..Loop Until block noted above
                                        
                                        ''' REV[2023.02.22.1325]
                                        ''' add error trap code
                                        ''' and precheck for equality
                                        ''' to reduce error potential here
                                        On Error Resume Next
                                        Err.Clear
                                        ''' REV[2023.04.21.1517] on RMQTY
                                        ''' added rounding of raw material
                                        ''' quantity to four digits
                                        ''' immediately ahead of assignment
                                        ''' to RMQTY property to ensure
                                        ''' the assigned value IS rounded
                                        qtRawMatl = Round(qtRawMatl, 4)
                                        If prRmQty.Value <> qtRawMatl Then
                                            prRmQty.Value = qtRawMatl
                                        End If
                                        
                                        If Err.Number Then
                                            Stop
                                            If False Then
                                            ''' REV[2023.05.10.1650]
                                            ''' attempt to work around problems
                                            ''' with iPart members by going to
                                            ''' the factory table cell itself.
                                            ''' '
                                            ''' doomed to fail if that column
                                            ''' isn't actually in the table...
                                            ''' '
                                            ''' not that it's likely
                                            ''' to succeed anyway.
                                            ''' '
                                            aiDocPart(prRmQty.Parent.Parent.Parent).ComponentDefinition.iPartMember.Row.Item(dcIPartTbCols(aiDocPart(prRmQty.Parent.Parent.Parent).ComponentDefinition.iPartMember.ParentFactory.TableColumns).Item(pnRmQty)).Value = qtRawMatl
                                            End If
                                        End If
                                        On Error GoTo 0
                                        
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        ''' REV[2022.09.20.1038]
                                        ''' added step to notify user of situation,
                                        ''' and offer opportunity to collect part
                                        ''' and material numbers for later review.
                                        'Debug.Print "== DONTKNOW =="
                                        'Debug.Print "Item: " & pnModel
                                        'Debug.Print "Matl: " & pnStock
                                        Debug.Print nu_FmGetList().AskUser(Join(Array("Unable to process", "the current Item.", "", "Copy the following ", "for later reference:", "", "Item: " & pnModel, "Matl: " & pnStock, ""), vbNewLine))
                                        
                                        ''' REV[2022.09.20.1042]
                                        ''' in conjunction with REV[2022.09.20.1038]
                                        ''' (above), disabled the following breakpoint,
                                        ''' as the new User notification effectively
                                        ''' supplants it
                                        'Stop 'because we don't know WHAT to do with it
                                             'and we do NOT want to clear anything
                                             'until we know what's going on!
                                        
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        '''     However, might need more handling here.
                                    End If
                                End If
                            'Stop 'WAYPOINT at end of block WITH@1764
                            End With
                        Else
                            If 0 Then Stop 'and regroup
                            ''' Things are looking a right royal mess
                            ''' at the moment I'm writing this comment.
                        End If
                    End If
                '--------------------------------------------'
                Else 'for standard Part (NOT Sheet Metal) ---'
                '--------------------------------------------'
                    ''' REV[2022.05.04.1501] (lines 1400-1406)
                    If .DocumentInterests.HasInterest( _
                        guidPipingSgmt _
                    ) Then 'for Piping'
                        'Stop
                        ck = newFmTest2().AskAbout(invDoc, _
                            "", _
                            Join(Array("" _
                                , "appears to be Hose or Tubing," _
                                , "presently " & IIf( _
                                    Len(pnStock) > 0, _
                                    pnStock, "unset" _
                                ) & ".", "" _
                                , "Would you like to " & IIf( _
                                    Len(pnStock) > 0, _
                                    "change", "set" _
                                ) & " it?" _
                            ), vbNewLine) _
                        )
                        ' (removed lines 1425-1431)
                        If ck = vbCancel Then
                            Stop
                        ElseIf ck = vbYes Then
                            ' (removed lines 1435-1440)
                            pnStock = userChoiceFromDc( _
                                dcFrom2Fields(cnGnsDoyle().Execute(sqlOf_GnsTubeHose( _
                                    .ComponentDefinition.Parameters.Item( _
                                        "Size_Designation" _
                                    ).Value _
                                )), "Description", "Item"), _
                                pnStock _
                            )
                            qtUnit = Trim$(UCase$(aiPropsUser.Item("ROPL").Value))
                            qtRawMatl = Round(Val(Split(qtUnit & " ", " ")(0)), 4)
                            qtUnit = Split(qtUnit & " ", " ")(1)
                            
                            ck = newFmTest2().AskAbout(invDoc, _
                                Join(Array("Stock Quantity of " _
                                    , qtRawMatl & qtUnit _
                                    & " of " & pnStock _
                                    , "selected for Item " _
                                ), vbNewLine), _
                                Join(Array( _
                                    "If this is okay, click [YES]" _
                                    , "(CANCEL to debug)" _
                                ), vbNewLine) _
                            )
                            If ck = vbCancel Then
                            ElseIf ck = vbYes Then
                                prRawMatl.Value = pnStock
                                prRmQty.Value = qtRawMatl
                                prRmUnit.Value = qtUnit
                                Debug.Print ; 'Breakpoint Landing
                            Else
                                Stop
                            End If
                            Debug.Print ; 'Breakpoint Landing
                        End If
                    End If
                    ''' REV[2022.05.04.1501] ENDS HERE
                    ''' NOTE[2022.05.04.1638] (lines 1477-1487)
                            ''' [2018.07.31 by AT] (lines 1488-1498)
                            With newFmTest1()
                                If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                                
                                ''' [2018.07.31 by AT] (lines 1502-1511)
                                Set bd = nuAiBoxData().UsingInches.SortingDims( _
                                    invDoc.ComponentDefinition.RangeBox _
                                )
                                'End With
                                
                                ck = .AskAbout(invDoc, _
                                    "Please Select Stock for Machined Part" _
                                    & vbNewLine & vbNewLine & bd.Dump(0) _
                                )
                                
                                If ck = vbYes Then
                                ''' UPDATE[2018.05.30]: (lines 1523-1525)
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
                                    ' (removed lines 1538-1542)
                                Else
                                    Debug.Print ; 'Breakpoint Landing
                                End If
                            End With
                            ' (removed lines 1547-1557) WAYPOINT ALERT!
                        If Len(pnStock) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                        '(removed lines 1560-1563)
                            Debug.Print ; 'Breakpoint Landing
                            'With cnGnsDoyle().Execute( _
                                "select Family from vgMfiItems where Item='" & _
                                Replace(pnStock, "'", "''") & "';" _
                            )
                            'Replace(pnStock, "'", "''")
                            ''' REV[2022.08.19.1416]
                            ''' temporarily replacing direct use
                            ''' of pnStock with Replace operation
                            ''' on single quotes in string
                            ''' '
                            ''' have already noted need for a 'handler',
                            ''' or 'preprocessor' to prepare values
                            ''' for SQL to avoid errors.
                            ''' see REV[2022.08.19.1359]
                            ''' '
                            With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
                                "vgMfiItems", "Family", "Item", pnStock _
                            ))
                            ''' REV[2022.08.26.1104]
                            ''' re 'handler' per REVS[
                            '''     2022.08.19.1359
                            '''     2022.08.19.1416
                            ''' ]
                            ''' new calls to sqlOf_simpleSelWhere
                            ''' added in disabled (commented) form
                            ''' to ultimately replace use of "hard
                            ''' coded" SQL statements nearby.
                            ''' '
                            ''' search this Function for sqlOf_simpleSelWhere
                            ''' to locate other instances of REV
                            ''' '
                            ''' new function sqlOf_simpleSelWhere
                            ''' automatically escapes single quotes
                            ''' in any String values supplied for
                            ''' matching, eliminating the need for
                            ''' this in the calling procedure.
                            ''' '
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]: (lines 1603-1607)
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    
                                    ''' UPDATE[2022.04.29.0852]
                                    ''' replicating code from UPDATE[2021.06.18]
                                    ''' above, noting also REV[2022.04.15.1035]
                                    If mtFamily Like "?-MT*" Then
                                        'Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        'Debug.Print pnModel & "[" & prRmQty.Value & qtUnit & "*" & pnStock & ": " & aiPropsDesign(pnDesc).Value & "]" ' prRawMatl.Value
                                        Debug.Print pnModel & "[" _
                                            & qtRawMatl & qtUnit _
                                            & " of " & pnStock & ": " _
                                            & aiPropsDesign(pnDesc).Value _
                                            & "]" 'prRmQty.Value prRawMatl.Value
                                        Stop 'FULL Stop!
                                    ElseIf mtFamily Like "?-PT*" Then
                                    ''' REV[2022.05.05.1343] (lines 1626-1639)
                                        If nmFamily Like "?-RM*" Then
                                            'ck = vbNo
                                            Debug.Print ; 'Breakpoint Landing
                                        Else
                                            ck = MsgBox(Join(Array( _
                                                "Part " & pnModel & " uses " & pnStock, _
                                                "which is not sheet metal.", _
                                                "", _
                                                "These parts are usually assigned", _
                                                "to the Riverview family, R-RMT.", _
                                                "", _
                                                "Do you want to use this Family?", _
                                                "Click [NO] to see other options.", _
                                                "(CANCEL to debug)" _
                                            ), vbNewLine), _
                                                vbYesNoCancel + vbQuestion, _
                                                "Select Part Family?" _
                                            )
                                            ' (removed lines 1658-1659)
                                            Debug.Print ; 'Breakpoint Landing
                                            If ck = vbCancel Then
                                                Stop 'to debug. (developers only!)
                                            ElseIf ck = vbYes Then
                                                nmFamily = "R-RMT"
                                            Else
                                                If Len(nmFamily) = 0 Then
                                                    nmFamily = "R-RMT"
                                                End If
                                                
                                                With nuDcPopulator( _
                                                    ).Setting("D-RMT", "Doyle (typ. sheet metal)" _
                                                    ).Setting("R-RMT", "Riverview (most others)" _
                                                )
                                                    If Not .Exists(nmFamily) Then
                                                        .Setting nmFamily, "Current (" & nmFamily & ")"
                                                    End If
                                                    
                                                    nmFamily = userChoiceFromDc( _
                                                        dcTransposed(.Dictionary()), _
                                                        nmFamily _
                                                    )
                                                End With
                                            End If
                                        End If
                                        
                                        mtFamily = "D-BAR"
                                    ElseIf mtFamily = "D-PTS" Then
                                        mtFamily = "D-BAR"
                                        Stop 'NOT SO FAST!
                                        'nmFamily = "D-RMT"
                                    ElseIf mtFamily = "R-PTS" Then
                                        mtFamily = "D-BAR"
                                        Stop 'NOT SO FAST!
                                        'nmFamily = "R-RMT"
                                    End If
                                    ' (removed lines 1696-1719)
                                End If
                            End With
                            '(removed lines 1722-1725) WAYPOINT ALERT!
                                    If mtFamily = "DSHEET" Then
                                        Stop 'because we should NOT be doing Sheet Metal in this section.
                                        ' This might require further investigation and/or development, if encountered.
                                        'We should be okay. This is sheet metal stock
                                        nmFamily = "D-RMT"
                                        qtUnit = "FT2"
                                        ''' UPDATE[2018.05.30]: (lines 1732-1736)
                                    ElseIf mtFamily = "D-BAR" Then
                                        ''' UPDATE[2022.01.11]: (lines 1738-1745)
                                        nmFamily = "R-RMT"
                                        qtUnit = prRmUnit.Value '"IN"
                                        ck = vbCancel
                                        Do
                                        ''' UPDATE[2021.03.11] (lines 1750-1776)
                                        If True Then 'go ahead and run old dump
                                        Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                                        ''' REV[2022.02.09.0904] (lines 1779-1783)
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
                                        
                                        ''' REV[2022.02.08.1446] (lines 1799-1803)
                                        Debug.Print "qtRawMatl = "; CStr(qtRawMatl); " 'in model. ";
                                        If dcIn.Exists(pnRmQty) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmQty));
                                        Debug.Print
                                        Debug.Print "qtUnit = """; qtUnit; """ 'in model.";
                                        If dcIn.Exists(pnRmUnit) Then Debug.Print "In Genius: "; CStr(dcIn.Item(pnRmUnit));
                                        Debug.Print " ( or try IN )"
                                        
                                        ''' REV[2022.02.08.1525] (lines 1811-1825)
                                        Debug.Print ""
                                        ''' REV[2022.03.11.1112] (lines 1827-1830)
                                        With nu_fmIfcMatlQty01().SeeUser(invDoc) '.Result()
                                            If .Exists(pnRmQty) Then
                                                ''' REV[2022.04.04.1404] (lines 1833-1839)
                                                If CDbl("0" & CStr(qtRawMatl)) _
                                                 = CDbl(.Item(pnRmQty)) Then 'prRmQty.Value => qtRawMatl
                                                    'no change -- don't change
                                                Else
                                                    'Debug.Print "prRmQty.Value FROM " & prRmQty.Value & " TO " & .Item(pnRmQty)
                                                    Debug.Print "qtRawMatl FROM " & qtRawMatl & " TO " & .Item(pnRmQty)
                                                    
                                                    'Stop 'and double-check
                                                    'might still be equivalent
                                                    qtRawMatl = .Item(pnRmQty) 'prRmQty.Value
                                                End If
                                            Else
                                                Stop
                                            End If
                                            
                                            If .Exists(pnRmUnit) Then
                                                If qtUnit = .Item(pnRmUnit) Then
                                                    'no change -- don't change
                                                Else
                                                    Debug.Print "qtUnit FROM " & qtUnit & " TO " & .Item(pnRmUnit)
                                                    'Stop 'and double-check
                                                    'might still be equivalent
                                                    qtUnit = .Item(pnRmUnit)
                                                End If
                                            Else
                                                Stop
                                            End If
                                        End With
                                        '(removed lines 1868-1874)
                                        ''' REV[2022.10.17.1504] CANCELED
                                        ''' disabled following confirmation
                                        ''' UserForm prompt as redundant
                                        ''' REV[2022.10.17.1511] undid 1504
                                        ''' this prompt might not be quite
                                        ''' so redundant as presumed
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(qtRawMatl); qtUnit; ". IF OKAY, CONTINUE." 'prRmQty.Value
                                        ck = newFmTest2().AskAbout(invDoc, _
                                            "Raw Material Quantity is now " _
                                            & CStr(qtRawMatl) & qtUnit & " for", _
                                            "If this is okay, click [YES]." _
                                            & vbNewLine & "Otherwise, click [NO] to review." _
                                            & vbNewLine & "" _
                                            & vbNewLine & "( for debug, click [CANCEL] )" _
                                        ) 'prRmQty.Value
                                        If ck = vbCancel Then
                                            Stop 'to debug.
                                        End If
                                        ''' REV[2022.10.17.1504]
                                        
                                        Loop Until ck = vbYes
                                        ''' UPDATE[2022.01.11]:
                                        '''     This is the terminal end of the
                                        '''     Do..Loop Until block noted above
                                        
                                        prRmQty.Value = qtRawMatl
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        Stop 'because we don't know WHAT to do with it
                                        ''' REV[2022.04.29.0755]
                                        ''' moved Stop AHEAD of the following assignments to
                                        ''' avoid clearing any potentially essential values.
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        '''     However, might need more handling here.
                                    End If
                        Else
                            If 0 Then Stop 'and regroup
                            ''' Things are looking a right royal mess
                            ''' at the moment I'm writing this comment.
                        End If
                '--------------------------------------------'
                End If 'Sheetmetal vs Part
                '--------------------------------------------'
                ''' REV[2022.05.05.1257] (lines 1914-1922)
                If Len(pnStock) > 0 Then
                With prRawMatl
                    If Len(Trim$(.Value)) > 0 Then
                        If pnStock <> .Value Then
                            '(removed comment lines 1927-1931)
                            ck = MsgBox(Join(Array( _
                                "Raw Stock Change Suggested", _
                                "  for Item " & pnModel, _
                                "", _
                                "  Current : " & prRawMatl.Value, _
                                "  Proposed: " & pnStock, _
                                "", "Change It?", "" _
                            ), vbNewLine), _
                                vbYesNo, pnModel & " Stock" _
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
                'Set rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                ''' Plan to remove commented line above,
                ''' superceded by the one above that
                Debug.Print ; 'Another landing line
                            '''
                            '''
                            '''
            ElseIf bomStruct = kPurchasedBOMStructure Then
                ''' As mentioned above, nmFamily
                ''' SHOULD be set at this point
                If Len(nmFamily) = 0 Then
                    If 1 Then Stop 'because we might
                    'need to check out the situation
                    nmFamily = "D-PTS" 'by default
                End If
            ElseIf bomStruct = kPhantomBOMStructure Then
                ''' REV[2022.01.17.1135] (lines 1999-2004)
                ck = newFmTest2().AskAbout(invDoc, _
                    "For some reason, THIS Item is marked Phantom:", _
                    "Is this okay? (Click [NO] OR [CANCEL] if not)" _
                )
                If ck = vbYes Then
                    'just let it go
                Else
                    Stop
                End If
            ElseIf bomStruct = kInseparableBOMStructure Then 'WT#?!?!!
                'How the HECK does a PART get marked Inseparable?!?
                ck = newFmTest2().AskAbout(invDoc, "This Item is marked Inseperable:", Join(Array("This is likely not correct,", "and should be fixed ASAP.", "Would you like to copy the Part", "Number for later review?", "", vbNewLine & vbNewLine & "([CANCEL] to debug)"), " "))
                If ck = vbYes Then
                'InputBox Join(Array("Copy this Part Number and paste it into another document or memo for review later."), vbNewLine), "Copy Part Number " & pnModel, pnModel
                InputBox Join(Array("Copy this Part Number, and paste", "it into another document or memo", "for later review."), vbNewLine), "Copy Part Number " & pnModel, pnModel
                ElseIf ck = vbCancel Then
                    Stop 'to debug. (developers only)
                    'press [F5] when ready to continue.
                End If

                Stop 'really, just STOP!
            Else
                ''' REV[2022.01.17.1138] (lines 2027-2032)
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
                    ''' REV[2022.04.15.1044]
                    ''' add check against current value.
                    ''' why try to fix what ain't broken?
                    If prFamily.Value <> nmFamily Then
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
                    End If
                    Set rt = dcAddProp(prFamily, rt)
                    'Set rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt)
                End If
            End If
        End With
        
        Call iSyncPartFactory(invDoc) 'Backport Properties to iPart Factory
        Set dcGeniusPropsPartRev20180530 = rt
    End If
End Function

Public Function dcAiDocComponents(AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional incTop As Long = 0, _
    Optional inclPhantom As Long = 0 _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If dc Is Nothing Then
        Set rt = dcAiDocComponents( _
            AiDoc, New Scripting.Dictionary _
            , incTop, inclPhantom _
        )
    Else
        Set rt = dc
    End If
    
    If AiDoc Is Nothing Then
    ElseIf AiDoc.DocumentType = kAssemblyDocumentObject Then
        If incTop Then
            With rt
                If .Exists(AiDoc.FullFileName) Then
                    If .Item(AiDoc.FullFileName) Is AiDoc Then
                    Else
                        Stop 'because somethin' ain't right.
                    End If
                Else
                    .Add AiDoc.FullFileName, AiDoc
                End If
            End With
        End If
        ''' NOTE[2023.01.27.1207]
        ''' not sure following call is correct.
        ''' it's a self-referential call passing
        ''' the received 'include phantom' flag
        ''' as the 'include top' argument, with
        ''' no clear reason why.
        '''
        ''' no, wait; this is NOT a self-referential call
        '''
        Set rt = dcAssyDocComponents( _
            AiDoc, rt, inclPhantom _
        )
    ElseIf AiDoc.DocumentType = kPartDocumentObject Then
        ''' REV[2022.04.12.1130]
        ''' add guard code to catch key collision
        ''' and check if matching Item is already
        ''' filed under that key. If not, manual
        ''' intervention may be required
        If rt.Exists(AiDoc.FullFileName) Then
            If rt.Item(AiDoc.FullFileName) Is AiDoc Then
                'it's already in. no action required
            Else
                Stop 'because somethin' just ain't right.
            End If
        Else
            rt.Add AiDoc.FullFileName, AiDoc
        End If
    Else
    End If
    
    Set dcAiDocComponents = rt
End Function

Public Function dcAssyDocComponents( _
    Assy As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional inclPhantom As Long = 0 _
) As Scripting.Dictionary
    Set dcAssyDocComponents = dcAssyCompAndSub( _
        Assy.ComponentDefinition.Occurrences, _
        dc, inclPhantom _
    )
End Function

Public Function dcAssyCompAndSub( _
    Occurences As Inventor.ComponentOccurrences, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional inclPhantom As Long = 0 _
) As Scripting.Dictionary
    ''' Traverse the assembly,
    ''' including any/all subassemblies,
    ''' and collect all parts to be processed.
    Dim rt As Scripting.Dictionary
    Dim invOcc As Inventor.ComponentOccurrence
    Dim tp As Inventor.ObjectTypeEnum
    Dim ocDef As Inventor.ComponentDefinition
    
    If dc Is Nothing Then
        Set rt = dcAssyCompAndSub( _
            Occurences, _
            New Scripting.Dictionary, _
            inclPhantom _
        )
    Else
        Set rt = dc
        For Each invOcc In Occurences
            With compOccFromProxy(invOcc) '(instead of just invOcc) [AT:2018.09.28]
                If .[_IsSimulationOccurrence] Then
                    Stop
                Else
                '''!!!WARNING!!!'''
                ''  The latest modification above
                ''  attempts to get around an issue with
                ''  ComponentOccurrenceProxy Occurences.
                ''  These seem to fail on attempts
                ''  to retrieve their Definition,
                ''  and its associated Document.
                ''
                ''  It is hoped the ContainingOccurrence
                ''  will supply the correct objects.
                ''  However, we DO NOT KNOW if this
                ''  is what we actually get.
                ''
                ''  Function compOccFromProxy includes a Stop
                ''  that occurs whenever a ComponentOccurrenceProxy
                ''  is discovered. In these instances, the process
                ''  should be carefully stepped through and traced
                ''  for any indication of the actual relationship
                ''  between a ComponentOccurrenceProxy
                ''  and its ContainingOccurrence.
                
                'Remove suppressed and excluded parts from the process
                'Moved out here from inner checks
                If .Visible And Not .Suppressed And Not .Excluded Then
                    ''' UPDATE[2018.08.20,AT]
                    ''' Error encountered on line noted.
                    ''' Adding Error trap with code to try alternative
                    On Error Resume Next
                    ''' when stopped under REV[2023.01.27.1329] below
                    ''' set next statement (Ctrl-F9) at On Error above
                    ''' and continue (F5)
                    
                    Set ocDef = .Definition
                    If Err.Number <> 0 Then
                        Stop
                        If .[_IsSimulationOccurrence] Then 'TypeOf invOcc Is Inventor.ComponentOccurrenceProxy Then
                        '    Err.Clear
                        '    Set ocDef = .ContainingOccurrence.Definition
                        '    If Err.Number <> 0 Then
                        '        Stop
                        '    End If
                            Stop
                        Else
                            Stop
                        End If
                        'Set ocDef = .ContextDefinition 'NO!! This will NOT WORK!
                    End If
                    On Error GoTo 0
                    
                    If ocDef Is Nothing Then
                        Stop
                    Else
                    '''''
                        'tp = .ContextDefinition.Type
                        tp = ocDef.Type
                        
                        If tp <> kAssemblyComponentDefinitionObject _
                        And tp <> kWeldmentComponentDefinitionObject _
                        Then
                            '(moved suppression/exclusion check OUTSIDE)
                            If tp <> kWeldsComponentDefinitionObject Then
                                
                                'Set rt = dcAddAiDoc(aiDocument(ocDef.Document), rt)
                                ''' Recasting by aiDocument not likely necessary here.
                                ''' Revised to following:
                                Set rt = dcAddAiDoc(ocDef.Document, rt)
                                
                            End If 'inVisible, suppressed, excluded or Welds
                            
                        Else 'assembly, check BOM Structure
                            If .BOMStructure = kPurchasedBOMStructure Then 'it's purchased
                                'Just add it to the Dictionary
                                Set rt = dcAddAiDoc(ocDef.Document, rt)
                            ElseIf .BOMStructure = kNormalBOMStructure Then 'we make it
                                'Gather its components
                                Set rt = dcAssyCompAndSub(.SubOccurrences, _
                                    dcAddAiDoc(ocDef.Document, rt), inclPhantom _
                                ) 'NOT forgetting to add THIS document!
                            ElseIf .BOMStructure = kInseparableBOMStructure Then 'maybe weldment?
                                If tp = kWeldmentComponentDefinitionObject Then 'it is
                                    'Treat it like an assembly
                                    Set rt = dcAssyCompAndSub(.SubOccurrences, _
                                        dcAddAiDoc(ocDef.Document, rt), inclPhantom _
                                    )
                                    'Except, we MIGHT want to make this
                                    'a NOT phantom assembly
                                ElseIf tp = kAssemblyComponentDefinitionObject Then
                                    'just an ordinary Assembly.
                                    'Same handling as above,
                                    'but use own branch, just in case.
                                    Set rt = dcAssyCompAndSub(.SubOccurrences, _
                                        dcAddAiDoc(ocDef.Document, rt), inclPhantom _
                                    )
                                Else 'it's not
                                    Stop 'and see if we can figure out what its type is
                                End If
                            ElseIf .BOMStructure = kPhantomBOMStructure Then '"phantom" component
                                If (inclPhantom And 1) = 1 Then
                                    'Get the Document as well as its components
                                    '(this is mainly for debugging/development)
                                    Set rt = dcAssyCompAndSub(.SubOccurrences, _
                                        dcAddAiDoc(ocDef.Document, rt), inclPhantom _
                                    )
                                Else
                                    'Gather its components, but NOT the document itself
                                    Set rt = dcAssyCompAndSub(.SubOccurrences, rt, inclPhantom)
                                End If
                            ElseIf .BOMStructure = kReferenceBOMStructure Then '"reference" component
                                Debug.Print newFmTest2.AskAbout(ocDef.Document, _
                                    "Reference Component will not be processed.", _
                                    "Click any button to acknowledge and continue." _
                                )
                            Else 'not sure what we've got
                                Debug.Print newFmTest2.AskAbout(ocDef.Document, _
                                    "Unhandled Condition on this component.", _
                                    "Going to Debug -- Click any button." _
                                )
                                Stop 'and have a look at it
                            End If
                        End If 'part or assembly
                    '''''
                    End If
                Else 'in case of improper skips
                    ''' REV[2023.01.27.1329]
                    '''
                    ''' add Else branch to deal with missed items
                    ''' under certain circumstances, for example,
                    ''' when instances are brought from an iAssembly
                    ''' (or iPart) factory in a Suppresed or Excluded state
                    '''
                    ''' USAGE: set breakpoint on following line when needed
                    Debug.Print ; 'and then, on Break here, search within
                    ''' THIS procedure on REV[2023.01.27.1329]
                    ''' and follow instructions at that point
                End If
                End If '(SimulationOccurrence)
            End With
            Set ocDef = Nothing
        Next
    End If
    Set dcAssyCompAndSub = rt
End Function

Public Function dcAddAiDoc( _
    AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim fn As String
    
    If dc Is Nothing Then
        Set dcAddAiDoc = dcAddAiDoc( _
            AiDoc, New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        fn = AiDoc.FullFileName
        With rt
            If Not .Exists(fn) Then .Add fn, AiDoc
        End With
        Set dcAddAiDoc = rt
    End If
End Function

Public Function dcPropVals( _
    dc As Scripting.Dictionary, _
    Optional nonProp As Long = 0 _
) As Scripting.Dictionary
    '''
    ''' dcPropVals -- Extract Values from
    '''     Properties in supplied Dictionary
    '''     non Property Items are processed
    '''     according to nonProp:
    '''     0 - Key/Item NOT added
    '''     1 - Key/Item added as is
    '''     >1 - Key/"blank" added
    '''
    ''' NOTE: similar functions may be
    '''     due for deprecation:
    '''         dcAiPropValsFromDc
    '''         dcMapAiProps2vals
    '''         '
    '''
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcNewIfNone(dc): For Each ky In .Keys
        Set pr = aiProperty(obOf(.Item(ky)))
        If pr Is Nothing Then
            If nonProp > 0 Then
                If nonProp > 1 Then
                    'Generate "blank" Item
                    rt.Add ky, noVal( _
                    VarType(.Item(ky)))
                Else
                    'Keep non-Property Item
                    rt.Add ky, .Item(ky)
                End If
            End If
        Else
            rt.Add ky, pr.Value
        End If
    Next: End With
    
    Debug.Print ; 'Breakpoint Landing
    Set dcPropVals = rt
End Function

Public Function dcGeniusProps(invDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    With invDoc
        '''
        Debug.Print "== Item " & .PropertySets(gnDesign).Item(pnPartNum).Value & " <" & .FullDocumentName & ">"
        If .DocumentType = kPartDocumentObject Then
            Set dcGeniusProps = dcGeniusPropsPart(invDoc, dc)
        ElseIf .DocumentType = kAssemblyDocumentObject Then
            Set dcGeniusProps = dcGeniusPropsAssy(invDoc, dc)
        Else
            Stop 'cuz we don't know WHAT to do with it
        End If
    End With
End Function

Public Function dcGeniusPropsAssy( _
    AiDoc As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim aiPropsUser     As Inventor.PropertySet
    Dim aiPropsDesign   As Inventor.PropertySet
    
    Dim prPartNum   As Inventor.Property 'pnPartNum
    Dim prFamily    As Inventor.Property
    Dim pnModel     As String
    Dim nmFamily    As String

    Dim rt As Scripting.Dictionary
    Dim ck As VbMsgBoxResult
    'Dim fm As String
    
    If dc Is Nothing Then
        Set rt = dcGeniusPropsAssy( _
        AiDoc, New Scripting.Dictionary)
    Else
        Set rt = dc
        
        If AiDoc Is Nothing Then
        Else
            With AiDoc
                ' Get the custom property set.
                With .PropertySets
                    Set aiPropsUser = .Item(gnCustom)
                    Set aiPropsDesign = .Item(gnDesign)
                End With
                
                ''' REV[2022.06.30.1546]
                ''' duplicated prPartNum and prFamily
                ''' from dcGeniusPropsPartRev20180530
                ''' along with some related processes.
                With aiPropsDesign
                    Set prPartNum = .Item(pnPartNum)
                    'aiGetProp( aiPropsDesign, pnPartNum)
                    Set prFamily = .Item(pnFamily)
                    'aiGetProp(aiPropsDesign, pnFamily)
                End With
                
                ' For now, we'll just assume all assemblies are made here.
                'fm = "D-MTO"
                
                ''' REV[2022.06.30.1550]
                ''' replaced above assigment with the two below
                ''' for more robust Family assignment for assemblies.
                '''
                ''' by checking model and Genius for an existing Family,
                ''' one hopes to avoid indiscriminately overriding
                ''' established Families, particularly in Genius.
                '''
                pnModel = prPartNum.Value
                
                ''' REV[2023.0106.1623]
                ''' due to replacement of Level of Detail
                ''' with Model States, which use a different
                ''' name for the default level/state, original
                ''' check is demoted to a secondary check
                ''' for the new default name.
                If .LevelOfDetailName <> "[Primary]" Then
                If .LevelOfDetailName <> "Master" Then
                'If .ModelStateName <> "[Primary]" Then
                    Stop
                'End If
                End If
                End If
                
                With .ComponentDefinition
                    ''' REV[2023.0113.1624]
                    ''' Family assignment moved into With statement
                    ''' and modified to collect existing value, if set,
                    ''' or otherwise generate one based on BOM structure,
                    ''' so purchased assemblies might be identified
                    
                    nmFamily = prFamily.Value
                    
                    If Len(nmFamily) = 0 Then
                        If .BOMStructure = kPurchasedBOMStructure Then
                            nmFamily = "D-PTS"
                        Else
                            nmFamily = "D-MTO"
                        End If
                    End If
                    
                    ''' REV[2023.0113.1625]
                    ''' simplified Family check against Genius
                    ''' to simply use current value of nmFamily,
                    ''' now it SHOULD be set to a non-blank value
                    nmFamily = famVsGenius(pnModel, nmFamily)
                    
                    Set rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt)
                    
                    With .MassProperties
                        On Error Resume Next
                        Err.Clear
                        Set rt = dcWithProp(aiPropsUser, pnMass, Round(.Mass * cvMassKg2LbM, 4), rt)
                        If Err.Number Then
                            Debug.Print Join(Array( _
                                "NOMASS", AiDoc.FullFileName _
                            ), ":")
                            ck = MsgBox(Join(Array("" _
                                & "An Error occurred while collecting" _
                                , "or updating Mass Property information" _
                                , "for " & AiDoc.DisplayName & "." _
                                , "" _
                                , "Click [Cancel] to enter debug mode" _
                                , "and attempt to review and correct." _
                                , "" _
                                , "Otherwise click [OK] to continue." _
                                , "(Mass will probably be incorrect)" _
                            ), vbNewLine), vbOKCancel, _
                                "ERROR(" & AiDoc.DisplayName & ")!" _
                            )
                            If ck = vbCancel Then
                                Stop
                            End If
                        End If
                        On Error GoTo 0
                    End With
                End With
            End With
        End If
        
        Call iSyncAssyFactory(AiDoc) 'Backport Properties to iAssembly Factory
    End If
    
    Set dcGeniusPropsAssy = rt
End Function

Public Function dcGeniusPropsPart( _
    AiDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    If dc Is Nothing Then
        Set dcGeniusPropsPart = _
        dcGeniusPropsPart(AiDoc, _
        New Scripting.Dictionary)
    ElseIf AiDoc Is Nothing Then
        Set dcGeniusPropsPart = dc
    Else
        Set dcGeniusPropsPart = _
        dcGeniusPropsPartRev20180530( _
        AiDoc, dc)
    End If
End Function

Public Function dcFlatPatVals( _
    invSheetMetalComp As Inventor.SheetMetalComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    ''
    'Dim aiPropSet As Inventor.PropertySet
    Dim prThickness As Inventor.Parameter
    ''
    Dim dLength As Double
    Dim dWidth As Double
    Dim dArea As Double
    Dim strWidth As String
    Dim strLength As String
    Dim strArea As String
    Dim mtFamily As String
    Dim mtType As String
    
    Dim dHeight As Double
    Dim dfHtThk As Double
    Dim strDVNs As String
    
    Dim ck As Double
    Dim gn As VbMsgBoxResult
    
    'If dc Is Nothing Then
    '    Set dcFlatPatVals = dcFlatPatVals( _
    '        invSheetMetalComp, _
    '        New Scripting.Dictionary _
    '    )
    'Else
        Set rt = dcNewIfNone(dc)
        
        'Request #3: Get sheet metal extent area and add to custom property "RMQTY"
        'Check to see if flat exists
        If invSheetMetalComp Is Nothing Then
            '''
            'Stop
        Else
            With invSheetMetalComp
                On Error Resume Next
                Err.Clear
                Set prThickness = .Thickness
                If Err.Number Then 'we've got a problem
                    If InStr(1, _
                        aiDocument(.Document).FullFileName, _
                        "\Doyle_Vault\Designs\purchased\" _
                    ) Then
                        Stop 'we probably got a purchased part, here
                        ''' NOTE[2018-05-31]: Don't recall hitting this stop recently,
                        ''' likely because parts matching this file path
                        ''' are getting caught in the calling function now.
                        ''' May be cable to remove this section,
                        ''' but retain for now pending further notice.
                        '''     '
                        ''' CHANGE NEEDED[2021.11.05]:
                        '''     do NOT want to go changing ANYTHING
                        '''     in the model inside THIS function,
                        '''     unless it's unavoidable; for example,
                        '''     if a flat pattern is not available.
                        '.BOMStructure = kPurchasedBOMStructure
                        'dc.Add "bomStruc", kPurchasedBOMStructure
                    Else
                        Stop 'and look into it
                    End If
                Else 'For now, we're going to assume
                '''     anything that fails to yield a valid Thickness object
                '''     probably can't be processed as a Sheet Metal component.
                '''     Therefore, the REST of the process should only proceed
                '''     if the retrieval succeeded.
                    If .HasFlatPattern Then
                        mtFamily = "DSHEET"
                        With .FlatPattern
                            If .Body Is Nothing Then 'there's gonna be a problem
                            ''' UPDATE[2021.06.11] Implementing check for Body
                            '''     to try to avoid raising an Error
                            '''     by diving blind into the With block
                            '''     and handle the missing Body situation
                            '''     in a more appropriate fashion.
                            '''     This comment supercedes [2019.12.13],
                            '''     now removed to notes_2021-0611_general-01.txt
                                With newFmTest2()
                                    If .AskAbout( _
                                        invSheetMetalComp.Document, _
                                        Join(Array( _
                                            "ISSUE WITH FLAT PATTERN:", _
                                            "   NO BODY FOUND" _
                                        ), vbNewLine), _
                                        Join(Array( _
                                            "Please consider reviewing model,", _
                                            "and rebuilding its flat pattern.", _
                                            "", _
                                            "Pause for review? (not necessary)" _
                                        ), vbNewLine) _
                                    ) = vbYes Then
                                    '(please check part for outdated flat pattern)
                                        Stop
                                    Else
                                        'Stop
                                    End If
                                End With
                                
                                Stop 'BKPT-2021-1105-1256
                                ''' CHANGE NEEDED[2021.11.05]:
                                '''     Not sure what, exactly
                                ''' NOTE[2021.12.13]
                                '''     These values are converted to inches
                                '''     from centimeters below in ...
                                '''     They should NOT be converted HERE!
                                '''     (don't think so, anyway)
                                '''     disabling conversion operations
                                '''     pending review on debug
                                dfHtThk = .Parameters.Item("Thickness").Value '/ cvLenIn2cm
                                dLength = .length '/ cvLenIn2cm
                                dWidth = .Width '/ cvLenIn2cm
                                dArea = dLength * dWidth
                                
                                'Stop 'just to keep the developer informed
                            Else
                                'Stop 'BKPT-2021-1105-1250
                                ''' CHANGE NEEDED[2021.11.05]:
                                '''     Actually have a function
                                '''     to collect X, Y, and Z spans
                                '''     which is used in the main
                                '''     function. Might be usable here.
                                '''     '
                                '''     (one might think the note would
                                '''      IDENTIFY the aforementioned
                                '''      function here, wouldn't one?
                                '''      Well... one would be WRONG!!!)
                                '''     '
                                '''     Here we go: nuAiBoxData()
                                With nuAiBoxData().UsingInches(0 _
                                ).UsingBox(.Body.RangeBox) '.SortingDims
                                    ''' UPDATE[2021.12.13]
                                    '''     Changed UsingInches argument
                                    '''     to zero (DON'T use) because
                                    '''     conversion is performed below.
                                    ' Check height against thickness
                                    ' Valid flat pattern should return
                                    ' zero or VERY minimal difference
                                    dHeight = Round(.SpanZ, 6) '(.MaxPoint.Z - .MinPoint.Z)
                                    
                                    ' Get the extent of the face.
                                    ' Extract the width, length and area from the range.
                                    dLength = Round(.SpanX, 6) '(.MaxPoint.X - .MinPoint.X)
                                    dWidth = Round(.SpanY, 6) '(.MaxPoint.Y - .MinPoint.Y)
                                End With
                                
                                With .Body.RangeBox
                                    ''' CHECKPOINT[2021.12.07]:
                                    '''     not actually stopping here
                                    '''     but running a quick check
                                    '''     to make sure revised code
                                    '''     above works correctly.
                                    '''     If so, this section SHOULD
                                    '''     be good to remove or disable
                                    ck = 0
                                    ck = ck + Abs(.MaxPoint.Z - .MinPoint.Z - dHeight) '* cvLenIn2cm
                                    ck = ck + Abs(.MaxPoint.X - .MinPoint.X - dLength) '* cvLenIn2cm
                                    ck = ck + Abs(.MaxPoint.Y - .MinPoint.Y - dWidth) '* cvLenIn2cm
                                    ''' UPDATE[2021.12.13]
                                    '''     conversion operations removed
                                    '''     since no longer converting
                                    '''     before end stage
                                    If Round(ck, 5) > 0 Then Stop 'BKPT-2021-1207-1158
                                    ''' need to go back to ck = 0
                                    ''' and step through to see
                                    ''' where the discrepancy lies
                                End With
                                
                                '''UPDATE[2021.11.05]:
                                '''     Moved derived calculations
                                '''     outside of With block above.
                                '''     Might or might not prove useful.
                                dfHtThk = Round(dHeight - prThickness.Value, 6) '/ cvLenIn2cm
                                ''' UPDATE[2021.12.13]
                                '''     conversion operation removed since
                                '''     no longer converting before end stage
                                ''' UPDATE[2022.01.28.1512]
                                '''     moved rounding operation to
                                '''     end stage, AFTER conversion
                                dArea = dLength * dWidth 'Round(, 6)
                                
                                Debug.Print ; 'Breakpoint Landing
                            End If
                            
                            If dArea = 0 Then 'let's try something else...
                            '   [2021.06.11] Moved alternate calculation
                            '       sequence to new If-Then block preceding
                            '       this one, and removed previous comment
                            '       as no longer relevant.
                                Stop 'and note when this branch taken
                                     'if it proves a frequent occurence,
                                     'might be appropriate to make this
                                     'the normal/default process
                            End If
                            
                            If dArea > 0 Then 'this one's a longshot, BUT!
                                ''' an invalid flat pattern SHOULD have no geometry,
                                ''' which means it SHOULD have no area to speak of.
                                ''' '
                                ''' One would think this obvious, in retrospect,
                                ''' but one would not be surprised to be proven wrong.
                                ''' Again.
                                Debug.Print ; 'Breakpoint Landing
                            Else 'we don't have a valid FlatPattern
                                If MsgBox(Join(Array( _
                                    "The flat pattern for this", _
                                    "part has no features,", _
                                    "and is likely not valid.", _
                                    "", _
                                    "Pause here to review?", _
                                    "(Click 'NO' to just keep going)" _
                                ), vbNewLine), vbYesNo, _
                                    "Invalid Flat Pattern" _
                                ) = vbYes Then
                                    Stop 'and let the user look into it
                                End If
                                Debug.Print aiDocument(.Document).FullDocumentName
                                
                                Stop
                                mtFamily = "D-BAR"
                            End If
                        End With
                        
''''''
''''''  The following section should be moved OUTSIDE this branch!
''''''
                        With aiDocPart(.Document)
                            'Set aiPropSet = .PropertySets.Item(gnCustom)
                            'prOffThknss
                            
                            ' Convert values into document units.
                            ' This will result in strings that are identical
                            ' to the strings shown in the Extent dialog.
                            '''     '
                            ''' NOTE[2021.11.09]
                            '''     If UsingInches is set as shown above,
                            '''     this section might not work properly.
                            '''     Might be better to NOT use inches,
                            '''     and simply let things take care of
                            '''     themselves, here.
                            ''' UPDATE[2021.12.13]
                            '''     Changed UsingInches argument to zero
                            '''     in two calls above.
                            With .UnitsOfMeasure
                                strWidth = .GetStringFromValue(dWidth, _
                                    .GetStringFromType(.LengthUnits))
                                strLength = .GetStringFromValue(dLength, _
                                    .GetStringFromType(.LengthUnits))
                                strArea = .GetStringFromValue(dArea, _
                                    .GetStringFromType(.LengthUnits) & "^2")
                                
                                If dfHtThk > 0.01 Then
                                    strDVNs = .GetStringFromValue( _
                                        dfHtThk, .GetStringFromType(.LengthUnits))
                                    'Debug.Print Join(Array("OFFTHK", _
                                        aiDocument(.Document).FullFileName, _
                                        Format$(dHeight, "0.0000"), _
                                        Format$(prThickness.Value, "0.0000"), _
                                        Format$(dHeight - prThickness.Value, "0.0000") _
                                    ), ":")
                                    'Stop
                                Else
                                    strDVNs = ""
                                End If
                            End With
                        End With
                        
                        'Stop 'BKPT-2021-1105-1304
                        ''' CHANGE NEEDED[2021.11.05]:
                        '''     This is where Properties are set.
                        '''     Want to change this to simply collect
                        '''     the generated values, and pass them
                        '''     back to the client for processing.
                        '''     '
                        '''     A separate process can then
                        '''     assign them to Properties.
                        '''     '
                        ' Add area to custom property set
                        'Set rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt)
                        rt.Add pnRmQty, Round(dArea * cvArSqCm2SqFt, 4)
                        '
                        ' 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
                        '
                        ' /  1ft | 1in    \2     2                2
                        '( ------+-------- ) * cm  = 0.00107639 ft
                        ' \ 12in | 2.54cm /
                        '
                        ' That value really needs to go into a constant
                        ' and so it HAS: cvArSqCm2SqFt (noted 2022.01.28)
                        ''' REV[2022.01.28.1516]
                        ''' add Raw Material Unit Quantity to output
                        rt.Add pnRmUnit, "FT2"
                        ''' Thickness, too:
                        
                        ' Add Thickness to returned values
                        rt.Add pnThickness, .Thickness.Value / cvLenIn2cm
                        
                        ' Add Width to custom property set
                        'Set rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt)
                        rt.Add pnWidth, strWidth
                        
                        ' Add Length to custom property set
                        'Set rt = dcWithProp(aiPropSet, pnLength, strLength, rt)
                        rt.Add pnLength, strLength
                        
                        ' Add AreaDescription to custom property set
                        'Set rt = dcWithProp(aiPropSet, pnArea, strArea, rt)
                        rt.Add pnArea, strArea
                        
                        If Len(strDVNs) > 0 Then
                            rt.Add "OFFTHK", strDVNs
                        End If
                        
                        Debug.Print ; 'Breakpoint Landing
                    Else 'we have no flat pattern!
                        mtFamily = "D-BAR"
                    End If
                    
                    rt.Add "mtFamily", mtFamily
                    ''' NOTE[2022.01.28.1524]:
                    ''' might want to make "mtFamily" a constant
                    ''' identifying name of Raw Material Family
                    ''' Maybe something like RMFAM or RMTYPE
                    
                    '''
                    '''
                End If
                On Error GoTo 0
            End With
        End If
        Set dcFlatPatVals = rt
    'End If 'dc Is Nothing
End Function

Public Function dcFlatPatProps( _
    invSheetMetalComp As Inventor.SheetMetalComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    ''
    Dim aiPropSet As Inventor.PropertySet
    Dim prThickness As Inventor.Parameter
    Dim prOffThknss As Inventor.Parameter
    ''
    Dim dLength As Double
    Dim dWidth As Double
    Dim dArea As Double
    Dim strWidth As String
    Dim strLength As String
    Dim strArea As String
    
    Dim dHeight As Double
    Dim dfHtThk As Double
    Dim strDVNs As String
    
    Dim ck As Double
    Dim gn As VbMsgBoxResult
    
    If dc Is Nothing Then
        Set dcFlatPatProps = dcFlatPatProps( _
            invSheetMetalComp, _
            New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        
        'Request #3: Get sheet metal extent area and add to custom property "RMQTY"
        'Check to see if flat exists
        If invSheetMetalComp Is Nothing Then
            '''
            Stop
        Else
            With invSheetMetalComp
                On Error Resume Next
                Err.Clear
                Set prThickness = .Thickness
                If Err.Number Then 'we've got a problem
                    If InStr(1, _
                        aiDocument(.Document).FullFileName, _
                        "\Doyle_Vault\Designs\purchased\" _
                    ) Then
                        Stop 'we probably got a purchased part, here
                        ''' NOTE[2018-05-31]: Don't recall hitting this stop recently,
                        ''' likely because parts matching this file path
                        ''' are getting caught in the calling function now.
                        ''' May be cable to remove this section,
                        ''' but retain for now pending further notice.
                        '''     '
                        ''' CHANGE NEEDED[2021.11.05]:
                        '''     do NOT want to go changing ANYTHING
                        '''     in the model inside THIS function,
                        '''     unless it's unavoidable; for example,
                        '''     if a flat pattern is not available.
                        .BOMStructure = kPurchasedBOMStructure
                    Else
                        Stop 'and look into it
                    End If
                Else 'For now, we're going to assume
                '''     anything that fails to yield a valid Thickness object
                '''     probably can't be processed as a Sheet Metal component.
                '''     Therefore, the REST of the process should only proceed
                '''     if the retrieval succeeded.
                    If Not .HasFlatPattern Then
                        Stop 'BKPT-2021-1105-1213
                        ''' CHANGE NEEDED[2021.11.05]:
                        '''     A new Flat Pattern should NOT
                        '''     be generated in this section!
                        '''     This should be handled in a new
                        '''     check function to determine
                        '''     whether identified Sheet Metal
                        '''     Part is in fact sheet metal
                        ''' UPDATE[2018.02.06]: New UserForm Available!
                        ''' Modify code here to use UserForm fmTest2
                        ''' to prompt user with part image and data
                        ''' while asking about Flat Pattern
                        ''' UPDATE[2021.12.13]: Bug Fix (hopefully)
                        '''     Bug encountered, wherein a "No"
                        '''     response is mistaken for a "Yes".
                        '''     Exact cause unclear, but suspect
                        '''     issue with terminated UserForm
                        '''     leaving behind undefined result.
                        '''
                        '''     This fix is intended to keep
                        '''     UserForm fmTest2 active while
                        '''     retrieving result, and thereby
                        '''     ensure correct result received.
                        '''
                        '''     Note that it depends on a new method
                        '''     function added to UserForm fmTest2:
                        '''     Using takes a supplied Inventor
                        '''     Document and locks it in for use
                        '''     on the next call to AskAbout, for
                        '''     which the Document is now Optional.
                        '''
                        With newFmTest2().Using(.Document)
                            gn = .AskAbout(, , "NO FLAT PATTERN!" _
                                & vbNewLine & "Try to generate one?" _
                            )
                        End With
                        If gn = vbYes Then
                        ''' UPDATE[2018-05-31]: Removing comment-disabled
                        ''' legacy code from switch to new UserForm above.
                        ''' Successful use since update noted above
                        ''' indicates no further need.
                            On Error Resume Next
                            Err.Clear
                            .Unfold
                            If Err.Number Then 'another problem
                                Debug.Print Join(Array("FPFAIL", _
                                    aiDocument(.Document).FullFileName _
                                ), ":")
                                'Stop 'and check it out
                            Else
                                If .HasFlatPattern Then .FlatPattern.ExitEdit
                            End If
                            On Error GoTo 0
                            
                            'We'll want to do something else instead of the following
                            'to make sure any document openened by the Unfold operation
                            'gets closed before we move on.
                            With ThisApplication.Documents.VisibleDocuments
                                If .Item(.Count) Is invSheetMetalComp.Document Then
                                    .Item(.Count).Close True
                                End If
                            End With
                        Else 'user elected not to create missing flat pattern
                            '''
                        End If
                    End If
                    
                    If .HasFlatPattern Then
                        With .FlatPattern
                            'First, make sure it's VALID
                            'If .Features.Count > 0 Then 'should be good? NOPE!!!
                            ''' Turns out, most flat patterns don't HAVE features.
                            ''' Not sure how they work, but they're not typical elements.
                            
                            'If .BaseFace Is Nothing Then 'This is an alternate test
                            'Keep on hand in case primary doesn't work out.
                            'Changeover will require switching Then and Else blocks.
                            
                            If .Body Is Nothing Then 'there's gonna be a problem
                            '   [2021.06.11] Implementing check for Body
                            '       to try to avoid raising an Error
                            '       by diving blind into the With block
                            '       and handle the missing Body situation
                            '       in a more appropriate fashion.
                            '       This comment supercedes [2019.12.13],
                            '       now removed to notes_2021-0611_general-01.txt
                                With newFmTest2()
                                    'If newFmTest2().AskAbout(.Document, , _
                                        "NO FLAT PATTERN!" & vbNewLine & _
                                        "Try to generate one?" _
                                    ) = vbYes Then
                                    If .AskAbout( _
                                        invSheetMetalComp.Document, _
                                        Join(Array( _
                                            "ISSUE WITH FLAT PATTERN:", _
                                            "   NO BODY FOUND" _
                                        ), vbNewLine), _
                                        Join(Array( _
                                            "Please consider reviewing model,", _
                                            "and rebuilding its flat pattern.", _
                                            "", _
                                            "Pause for review? (not necessary)" _
                                        ), vbNewLine) _
                                    ) = vbYes Then
                                    '(please check part for outdated flat pattern)
                                        Stop
                                    Else
                                        'Stop
                                    End If
                                End With
                                
                                Stop 'BKPT-2021-1105-1256
                                ''' CHANGE NEEDED[2021.11.05]:
                                '''     Not sure what, exactly
                                ''' NOTE[2021.12.13]
                                '''     These values are converted to inches
                                '''     from centimeters below in ...
                                '''     They should NOT be converted HERE!
                                '''     (don't think so, anyway)
                                '''     disabling conversion operations
                                '''     pending review on debug
                                dfHtThk = .Parameters.Item("Thickness").Value '/ cvLenIn2cm
                                dLength = .length '/ cvLenIn2cm
                                dWidth = .Width '/ cvLenIn2cm
                                dArea = dLength * dWidth
                                
                                'Stop 'just to keep the developer informed
                            Else
                                'Stop 'BKPT-2021-1105-1250
                                ''' CHANGE NEEDED[2021.11.05]:
                                '''     Actually have a function
                                '''     to collect X, Y, and Z spans
                                '''     which is used in the main
                                '''     function. Might be usable here.
                                '''     '
                                '''     (one might think the note would
                                '''      IDENTIFY the aforementioned
                                '''      function here, wouldn't one?
                                '''      Well... one would be WRONG!!!)
                                '''     '
                                '''     Here we go: nuAiBoxData()
                                With nuAiBoxData().UsingInches(0 _
                                ).UsingBox(.Body.RangeBox) '.SortingDims
                                    ''' UPDATE[2021.12.13]
                                    '''     Changed UsingInches argument
                                    '''     to zero (DON'T use) because
                                    '''     conversion is performed below.
                                    ' Check height against thickness
                                    ' Valid flat pattern should return
                                    ' zero or VERY minimal difference
                                    dHeight = Round(.SpanZ, 6) '(.MaxPoint.Z - .MinPoint.Z)
                                    
                                    ' Get the extent of the face.
                                    ' Extract the width, length and area from the range.
                                    dLength = Round(.SpanX, 6) '(.MaxPoint.X - .MinPoint.X)
                                    dWidth = Round(.SpanY, 6) '(.MaxPoint.Y - .MinPoint.Y)
                                End With
                                
                                With .Body.RangeBox
                                    ''' CHECKPOINT[2021.12.07]:
                                    '''     not actually stopping here
                                    '''     but running a quick check
                                    '''     to make sure revised code
                                    '''     above works correctly.
                                    '''     If so, this section SHOULD
                                    '''     be good to remove or disable
                                    ck = 0
                                    ck = ck + Abs(.MaxPoint.Z - .MinPoint.Z - dHeight) '* cvLenIn2cm
                                    ck = ck + Abs(.MaxPoint.X - .MinPoint.X - dLength) '* cvLenIn2cm
                                    ck = ck + Abs(.MaxPoint.Y - .MinPoint.Y - dWidth) '* cvLenIn2cm
                                    ''' UPDATE[2021.12.13]
                                    '''     conversion operations removed
                                    '''     since no longer converting
                                    '''     before end stage
                                    If Round(ck, 5) > 0 Then Stop 'BKPT-2021-1207-1158
                                    ''' need to go back to ck = 0
                                    ''' and step through to see
                                    ''' where the discrepancy lies
                                End With
                                
                                '''UPDATE[2021.11.05]:
                                '''     Moved derived calculations
                                '''     outside of With block above.
                                '''     Might or might not prove useful.
                                dfHtThk = Round(dHeight - prThickness.Value, 6) '/ cvLenIn2cm
                                ''' UPDATE[2021.12.13]
                                '''     conversion operation removed since
                                '''     no longer converting before end stage
                                dArea = Round(dLength * dWidth, 6)
                                
                                Debug.Print ; 'Breakpoint Landing
                            End If
                            
                            If dArea = 0 Then 'let's try something else...
                            '   [2021.06.11] Moved alternate calculation
                            '       sequence to new If-Then block preceding
                            '       this one, and removed previous comment
                            '       as no longer relevant.
                                Stop 'and note when this branch taken.
                                     'if it proves a frequent occurence,
                                     'might be appropriate to make this
                                     'the normal/default process
                            End If
                            
                            If dArea > 0 Then 'this one's a longshot, BUT!
                                ''' an invalid flat pattern SHOULD have no geometry,
                                ''' which means it SHOULD have no area to speak of.
                                ''' '
                                ''' One would think this obvious, in retrospect,
                                ''' but one would not be surprised to be proven wrong.
                                ''' Again.
                            Else 'we don't have a valid FlatPattern
                                If MsgBox(Join(Array( _
                                    "The flat pattern for this", _
                                    "part has no features,", _
                                    "and is likely not valid.", _
                                    "", _
                                    "Pause here to review?", _
                                    "(Click 'NO' to just keep going)" _
                                ), vbNewLine), vbYesNo, _
                                    "Invalid Flat Pattern" _
                                ) = vbYes Then
                                    Stop 'and let the user look into it
                                End If
                                Debug.Print aiDocument(.Document).FullDocumentName
                            End If
                        End With
                        
                        If dfHtThk > 0 Then
                            With dcFlatPatSpansByVertices(.FlatPattern)
                                ck = Round(.Item("Z") - prThickness.Value, 6)
                                If dfHtThk > ck Then
                                    dHeight = .Item("Z")
                                    dfHtThk = Round(dHeight - prThickness.Value, 6)
                                    Debug.Print Round(.Item("X") - dLength, 6)
                                    Debug.Print Round(.Item("Y") - dWidth, 6)
                                    Debug.Print ; 'Breakpoint Landing
                                End If
                            End With
                        End If
                    Else 'we have no flat pattern!
                        'aiDocPart(.Document).
                        With nuAiBoxData().UsingInches(0 _
                        ).UsingBox(.RangeBox) '.SortingDims
                            ''' UPDATE[2021.12.13]
                            '''     Changed UsingInches argument
                            '''     to zero (DON'T use) because
                            '''     conversion is performed below.
                            ' Check height against thickness
                            ' Valid flat pattern should return
                            ' zero or VERY minimal difference
                            dHeight = Round(.SpanZ, 6) '(.MaxPoint.Z - .MinPoint.Z)
                            
                            ' Get the extent of the face.
                            ' Extract the width, length and area from the range.
                            dLength = Round(.SpanX, 6) '(.MaxPoint.X - .MinPoint.X)
                            dWidth = Round(.SpanY, 6) '(.MaxPoint.Y - .MinPoint.Y)
                            
                            dArea = Round(dLength * dWidth, 6)
                            'not really valid, here
                            'but it's used below
                            dfHtThk = Round(dHeight - prThickness.Value, 6) '/ cvLenIn2cm
                            'as is this
                            ''' both of these are set in the valid
                            ''' sheet branch as well as this one.
                            ''' maybe they can be consolidated below?
                        End With
                        
                        Debug.Print Join(Array("NOFLAT", _
                            aiDocument(.Document).FullFileName _
                        ), ":")
                        'Stop
                        'May have to stop and do something here
                        'Could we generate the flat pattern on the fly?
                    End If
''''''
''''''  The following section should be moved OUTSIDE this branch!
''''''
                    With aiDocPart(.Document)
                        Set aiPropSet = .PropertySets.Item(gnCustom)
                        'prOffThknss
                        
                        ' Convert values into document units.
                        ' This will result in strings that are identical
                        ' to the strings shown in the Extent dialog.
                        '''     '
                        ''' NOTE[2021.11.09]
                        '''     If UsingInches is set as shown above,
                        '''     this section might not work properly.
                        '''     Might be better to NOT use inches,
                        '''     and simply let things take care of
                        '''     themselves, here.
                        ''' UPDATE[2021.12.13]
                        '''     Changed UsingInches argument to zero
                        '''     in two calls above.
                        With .UnitsOfMeasure
                            strWidth = .GetStringFromValue(dWidth, _
                                .GetStringFromType(.LengthUnits))
                            strLength = .GetStringFromValue(dLength, _
                                .GetStringFromType(.LengthUnits))
                            strArea = .GetStringFromValue(dArea, _
                                .GetStringFromType(.LengthUnits) & "^2")
                            
                            If dfHtThk > 0.01 Then
                                strDVNs = .GetStringFromValue( _
                                    dfHtThk, .GetStringFromType(.LengthUnits))
                                'Debug.Print Join(Array("OFFTHK", _
                                    aiDocument(.Document).FullFileName, _
                                    Format$(dHeight, "0.0000"), _
                                    Format$(prThickness.Value, "0.0000"), _
                                    Format$(dHeight - prThickness.Value, "0.0000") _
                                ), ":")
                                'Stop
                            Else
                                strDVNs = ""
                            End If
                        End With
                    End With
                    
                    'Stop 'BKPT-2021-1105-1304
                    ''' CHANGE NEEDED[2021.11.05]:
                    '''     This is where Properties are set.
                    '''     Want to change this to simply collect
                    '''     the generated values, and pass them
                    '''     back to the client for processing.
                    '''     '
                    '''     A separate process can then
                    '''     assign them to Properties.
                    '''     '
                    ' Add area to custom property set
                    Set rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt)
                    '
                    ' 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
                    '
                    ' /  1ft | 1in    \2     2                2
                    '( ------+-------- ) * cm  = 0.00107639 ft
                    ' \ 12in | 2.54cm /
                    '
                    ' That value really needs to go into a constant
                    '
                    
                    ' Add Width to custom property set
                    Set rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt)
                    
                    ' Add Length to custom property set
                    Set rt = dcWithProp(aiPropSet, pnLength, strLength, rt)
                    
                    ' Add AreaDescription to custom property set
                    Set rt = dcWithProp(aiPropSet, pnArea, strArea, rt)
                    
                    If Len(strDVNs) > 0 Then
                        Set rt = dcWithProp(aiPropSet, "OFFTHK", strDVNs, rt)
                        'WAS assigning strWidth to this property, but think
                        'that's just because it was (probably) copied
                        'from the first instance, for width (above)
                    Else
                    End If
                    
                    Debug.Print ; 'Breakpoint Landing
                    
                    '''
                    '''
                End If
                On Error GoTo 0
            End With
        End If
        Set dcFlatPatProps = rt
    End If
End Function

Public Function dcWithProp( _
    propSet As PropertySet, _
    Name As String, Value As Variant, _
    Optional rt As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Set dcWithProp = dcAddProp(aiSetProp(propSet, Name, Value), rt)
End Function

Public Function dcAddProp( _
    aiProp As Inventor.Property, _
    Optional dcIn As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim dcOut As Scripting.Dictionary
    Dim nm As String
    
    If dcIn Is Nothing Then
        Set dcAddProp = dcAddProp( _
            aiProp, New Scripting.Dictionary _
        )
    Else
        Set dcOut = dcIn
        If Not aiProp Is Nothing Then
            nm = aiProp.Name
            With dcOut
                If .Exists(nm) Then .Remove nm
                .Add nm, aiProp
            End With
        End If
        Set dcAddProp = dcOut
    End If
End Function

Public Function aiGetProp( _
    aiPropSet As PropertySet, _
    aiPropName As String, _
    Optional Create As Long = 0 _
) As Inventor.Property
    Dim aiProp As Inventor.Property
    
    On Error Resume Next
    
    ''' FORDEBUG[2021.08.09] -- disable when not debugging
    ''' report names of Property Set and desired Property
    'Debug.Print "PROPSET[" & aiPropSet.Name & "].ITEM[" & aiPropName & "]";
    
    ' Attempt to get an existing custom property named aiPropName.
    Set aiProp = aiPropSet.Item(aiPropName)
    
    If Err.Number = 0 Then 'got it!
        ''' FORDEBUG[2021.08.09] -- disable when not debugging
        ''' report successful acuisition
        'Debug.Print " FOUND"
    Else 'failed to get property, which means it doesn't exist
        If Create = 0 Then 'not going to try creating
            ''' FORDEBUG[2021.08.09] -- disable when not debugging
            'Debug.Print " NOTFOUND"
            
            Set aiProp = Nothing
        Else 'try to create it:
            Err.Clear
            Set aiProp = aiPropSet.Add("", aiPropName)
            If Err.Number = 0 Then 'made it!
                ''' FORDEBUG[2021.08.09] -- disable when not debugging
                'Debug.Print " CREATED"
            Else 'creation also failed.
                ''' FORDEBUG[2021.08.09] -- disable when not debugging
                'Debug.Print " CANTMAKE"
                
                Set aiProp = Nothing
            End If
        End If
    End If
    
    On Error GoTo 0
    Set aiGetProp = aiProp
End Function

Public Function aiSetProp( _
    aiPropSet As PropertySet, _
    aiPropName As String, _
    aiPropValue As Variant _
) As Inventor.Property
    Dim aiProp As Inventor.Property
    
    'Try to acquire Property Object
    Set aiProp = aiGetProp(aiPropSet, aiPropName, 1)
    
    If aiProp Is Nothing Then 'we couldn't create it
        'Might have to handle that here
        Debug.Print "    NOPROP/CANTMAKE -- couldn't create Property " & aiPropName; 'Breakpoint Landing
        
    Else 'we've got it, so let's try to set it
        On Error Resume Next
        '''
        ''' All Property Object acquisition code
        ''' moved to aiGetProp (called above)
        ''' remains in comment-disabled form below.
        ''' Remove once functionality verified
        '''
        
        ' Attempt to get an existing custom property named aiPropName.
        'Set aiProp = aiPropSet.Item(aiPropName)
        
        'If Err.Number <> 0 Then 'failed to get property, which means it doesn't exist
            
            'Try to create it:
            'Err.Clear
            'Set aiProp = aiPropSet.Add((aiPropValue), aiPropName)
            ''  NOTE: aiPropValue apparently needs to be in parentheses
            ''      for some values. Specifically, number/unit strings
            ''      like "1.500 in" seem to trigger VBA Error 51: Internal Error
            ''      Embedding the variable in parentheses forces VBA
            ''      to resolve the Variant to a string, maybe?
            ''      In any case, that seems to fix the problem.
            
            'If Err.Number <> 0 Then 'creation also failed.
                'Set aiProp = Nothing
            'End If
        'Else 'Got the property so update the value:
            'Let's check if it's different, first
            If aiPropValue = aiProp.Value Then
                'Stop 'to verify they're the same
                'Debug.Print "    SAMEVAL(" & aiPropValue & ")";
                Debug.Print "    SAMEVAL(" & aiProp.Name & "): " & aiProp.Value;
            ElseIf CStr(aiPropValue) = CStr(aiProp.Value) Then
                'Stop 'BKPT-2021-1105-1419
                ''' CHANGE NEEDED[2021.11.05]:
                '''     Need to make sure values really ARE different!
                '''     RMQTY especially seems to have trouble with this.
                '''     '
                '''     Example: CHGVAL(RMQTY): 0.20833313172 ==> 0.20833313172
                '''     'VERY minor difference in values:
                '''     '  Debug.Print aiPropValue - aiProp.Value
                '''     '  -2.77555756156289E-17
                '''     '
                '''     Need to include a check between converted copies.
                '''     Believe implemented before, but got lost in crash.
                '''     '
                ''' UPDATE[SAME_DAY]:
                '''     This ElseIf clause adds confirmative test
                '''     by checking String conversions of each
                '''     against each other.
                '''     '
                Debug.Print "    EQUIVAL(" & aiProp.Name & "): " & aiProp.Value;
            Else
                ''' CHANGE NEEDED[2021.11.05]:
                '''     Need to make sure values really ARE different!
                ''' UPDATE[SAME_DAY]:
                '''     Added confirmative test; see ElseIf above
                '''     (hopefully no issue with failed CStr calls)
                Debug.Print "    CHGVAL(" & aiProp.Name & "): " & aiProp.Value & " => " & aiPropValue;
                'Stop 'and make sure it really IS different
                aiProp.Value = (aiPropValue)
                ''  See note above on setting Property.
                ''  Assuming parentheses also required here.
                
                If Err.Number = 0 Then 'set/update successful
                    'Debug.Print " ==> (" & aiProp.Value & ")";
                    'Debug.Print " ==> " & aiProp.Value;
                    Debug.Print "    OK!"; 'Breakpoint Landing
                Else 'couldn't set/update
                    'Not much else we can do at this point
                    'Set aiProp = Nothing
                    Debug.Print "    FAILED:CANTCHG"; 'Breakpoint Landing
                End If
            End If
        'End If
        
        On Error GoTo 0
    End If
    
    ''' FORDEBUG[2021.08.09] -- disable when not debugging
    Debug.Print 'forcing newline
    
    Set aiSetProp = aiProp
End Function

Public Function ptNumShtMetal( _
    aiSMdef As Inventor.SheetMetalComponentDefinition _
) As String
    'Request #2: Get Genius SheetMetal
    'by matching Style Name and Material.
    'Add to Custom Property RM
    Dim invGeniusMaterial As String 'Return value
    '
    Dim prThickness As Inventor.Parameter
    Dim rs As ADODB.Recordset
    Dim dc As Scripting.Dictionary
    Dim invSheetMetalMaterial As String
    Dim invSheetMetalThickness As Double
    Dim sqlText As String
    Dim docName As String
    
    With aiSMdef
        On Error Resume Next
        Err.Clear
        Set prThickness = .Thickness
        If Err.Number = 0 Then 'we got valid Thickness
        '''
        ''' For now, we must assume we can only proceed
        ''' if a valid Thickness parameter is retrieved
        
        docName = aiDocPart(.Document).FullDocumentName
        
        invSheetMetalMaterial = aiDocPart(.Document _
            ).PropertySets.Item(gnDesign).Item(pnMaterial).Value
        invSheetMetalThickness = prThickness.Value / cvLenIn2cm 'Internal Units in cm???
        'invSheetMetalThickness = .Thickness.Value / cvLenIn2cm 'Internal Units in cm???
        sqlText = sqlSheetMetal(invSheetMetalMaterial, invSheetMetalThickness)
        
        With cnGnsDoyle()
            Set rs = .Execute(sqlText)
            With rs
                If (.BOF And .EOF) Then 'it wasn't found in Genius
                    .Close
                    
                    'Here's where we resort to the HARD way.
                    'Debug.Print Val(aiSMdef.ActiveSheetMetalStyle.Thickness) - invSheetMetalThickness < 0.0001
                    If Val(aiSMdef.ActiveSheetMetalStyle.Thickness) - invSheetMetalThickness < 0.0001 Then
                    ''' UPDATE[2022.01.12.1314]:
                    '''     Add check for matching Thickness
                    '''     between Part Property and its
                    '''     active Sheet Metal Style.
                    '''     If they DON'T match, then it's
                    '''     probably NOT a Sheet Metal Part.
                    '''     Will probably need a better check
                    '''     moving forward, but this SHOULD do
                    '''     for now.
                    invGeniusMaterial = pnShtMetalHardCoded( _
                        invSheetMetalMaterial, _
                        aiSMdef.ActiveSheetMetalStyle.Name _
                    )
                    Else
                        Debug.Print docName
                        Debug.Print aiSMdef.ActiveSheetMetalStyle.Name
                        Debug.Print aiSMdef.ActiveSheetMetalStyle.Thickness
                        Stop
                        invGeniusMaterial = ""
                    End If
                    
                    If Len(invGeniusMaterial) > 0 Then
                    'something might be missing from Genius
                        If Left$(invGeniusMaterial, 2) = "LG" Then
                        'might actually be lagging
                            Debug.Print "POSSIBLE LAGGING ITEM"
                        Else
                        End If
                        Stop 'and review the situation
                    Else 'it's not hardcoded, either
                        invGeniusMaterial = ""
                        'The caller will have to deal with this.
                    End If
                Else 'Genius found it!
                    '(or SOMETHING, anyway) '''REV[2023.05.17.1211]
                    ''' added User Prompt to pick from multiple options
                    ''' when more than one material option returned.
                    ''' details noted below
                    'Stop
                    Set dc = dcFrom2Fields(rs, "Item", "Item") 'dcFromAdoRS(rs)
                    With dc
                    If .Count > 0 Then
                        invGeniusMaterial = .Keys(0) ' .Fields(0).Value
                        If .Count > 1 Then
                            invGeniusMaterial = userChoiceFromDc( _
                                dc, invGeniusMaterial _
                            ) 'nuSelFromDict '.Fields(0).Value
                        End If
                    Else 'ElseIf dc.Count < 1 Then
                        Stop
                        invGeniusMaterial = ""
                    End If: End With
                    
                    .Close
                End If
            End With
            
            .Close
        End With
        
        ''' The preceding block only ran if a
        ''' valid Thickness parameter was retrieved
        Else 'we've got a problem with our Sheet Metal component
            Stop 'and review the situation
        End If
        On Error GoTo 0
    End With
    
    ptNumShtMetal = invGeniusMaterial
End Function

Public Function sqlSheetMetal( _
    Optional mtName As String = "", _
    Optional thk As Double = 0 _
) As String
    Dim hdr2match As String
    Dim thk2match As String
    Dim mtl2match As String
    ''' REV[2022.04.13.0939]
    ''' modified to replace header match
    ''' (FM-, FS-, etc.) with match against
    ''' metal/material type (MS/SS) and
    ''' thus be able to catch expanded
    ''' metal options, assuming the sheet
    ''' metal thickness is correctly set.
    ''' NOTE: this is HIGHLY experimental.
    ''' It SHOULD still work under most
    ''' circumstances, but be aware of
    ''' potential issues.
    
    If mtName = "Stainless Steel" Then
        hdr2match = "FS"
        mtl2match = "SS"
    ElseIf mtName = "Stainless Steel, Austenitic" Then
        hdr2match = "FS"
        mtl2match = "S4" 'for 409
    ElseIf mtName = "Stainless Steel 304" Then
        hdr2match = "FS"
        mtl2match = "SS"
    ElseIf mtName = "304SS" Then
        hdr2match = "FS"
        mtl2match = "SS"
    ElseIf mtName = "Steel, Mild" Then
        hdr2match = "FM"
        mtl2match = "MS"
    ElseIf mtName = "Rubber" Then
        hdr2match = "LG"
        mtl2match = "" 'not metal
    ElseIf mtName = "Rubber, Silicone" Then
        hdr2match = "LG"
        mtl2match = "" 'not metal
    ElseIf mtName = "Lagging" Then
        hdr2match = "LG"
        mtl2match = "" 'not metal
    ElseIf mtName = "UHMW, White" Then
        hdr2match = "UH"
        mtl2match = "" 'not metal
    Else
        Debug.Print mtName
        'Stop
        hdr2match = "XX"
        mtl2match = "" 'not metal
    End If
    
    thk2match = Format$(thk, "0.000")
            '"Item Like '" & hdr2match & "%'"
    sqlSheetMetal = Trim$(Join(Array( _
        "SELECT Item, Description1", _
        "FROM vgMfiItems", _
        "WHERE " & IIf(Len(mtName) > 0, _
            "Specification6 = '" & mtl2match & "'", _
            "1=1" _
        ), _
        "  AND Family = " & IIf(hdr2match = "LG", _
            "'D-PTS'", "'DSHEET'" _
        ), _
        IIf(hdr2match = "LG", "  AND Item LIKE 'LG%'", _
            "  AND Specification1 = 'STANDARDSHEET'" _
        ), _
        IIf(thk > 0, _
            "  AND Abs(Thickness - " & thk2match & ") < 0.007", _
            "" _
        ), _
        ";" _
    ), vbNewLine))
End Function

Public Function pnShtMetalHardCoded( _
        invSheetMetalMaterial As String, _
        invSheetMetalName As String _
) As String
    Dim invGeniusMaterial As String
    
    'Stop 'because this function should
        'not be getting called anymore
    
    'Map combination to corresponding Genius Part Number
    If invSheetMetalMaterial = "Stainless Steel" Then
        If invSheetMetalName = "18 GA" Then
            invGeniusMaterial = "FS-48x96x0.048"
        ElseIf invSheetMetalName = "14 GA" Then
            invGeniusMaterial = "FS-60x120x0.075"
        ElseIf invSheetMetalName = "13 GA" Then
            invGeniusMaterial = "FS-60x97x0.09"
        ElseIf invSheetMetalName = "12 GA" Then
            invGeniusMaterial = "FS-60x120x0.105"
        ElseIf invSheetMetalName = "10 GA" Then
            invGeniusMaterial = "FS-60x144x0.135"
        ElseIf invSheetMetalName = "3/16""" Then
            invGeniusMaterial = "FS-60x144x0.188"
        ElseIf invSheetMetalName = "1/4""" Then
            invGeniusMaterial = "FS-60x144x0.25"
        ElseIf invSheetMetalName = "5/16""" Then
            invGeniusMaterial = "FS-60x144x0.313"
        ElseIf invSheetMetalName = "3/8""" Then
            invGeniusMaterial = "FS-60x144x0.375"
        ElseIf invSheetMetalName = "1/2""" Then
            invGeniusMaterial = "FS-60x144x0.5"
        Else
            invGeniusMaterial = ""
        End If
    ElseIf invSheetMetalMaterial = "Steel, Mild" Then
        If invSheetMetalName = "14 GA" Then
            invGeniusMaterial = "FM-60x144x0.075"
        ElseIf invSheetMetalName = "12 GA" Then
            invGeniusMaterial = "FM-60x144x0.105"
        ElseIf invSheetMetalName = "10 GA" Then
            invGeniusMaterial = "FM-60x144x0.135"
        ElseIf invSheetMetalName = "3/16""" Then
            invGeniusMaterial = "FM-60x144x0.188"
        ElseIf invSheetMetalName = "1/4""" Then
            invGeniusMaterial = "FM-60x144x0.25"
        ElseIf invSheetMetalName = "5/16""" Then
            invGeniusMaterial = "FM-60x144x0.313"
        ElseIf invSheetMetalName = "3/8""" Then
            invGeniusMaterial = "FM-60x144x0.375"
        ElseIf invSheetMetalName = "1/2""" Then
            invGeniusMaterial = "FM-60x144x0.5"
        ElseIf invSheetMetalName = "5/8""" Then
            invGeniusMaterial = "FM-60x144x0.625"
        ElseIf invSheetMetalName = "3/4""" Then
            invGeniusMaterial = "FM-60x120x0.75"
        ElseIf invSheetMetalName = "1""" Then
            invGeniusMaterial = "FM-48x120x1"
        Else
            invGeniusMaterial = ""
        End If
    ElseIf invSheetMetalMaterial = "Rubber" Then
        'Debug.Print "POSSIBLE LAGGING ITEM"
        'Debug.Print "invGeniusMaterial = ""LG"""
        Stop
        invGeniusMaterial = "LG"
        'If invSheetMetalName = "14 GA" Then
        '    invGeniusMaterial = ""
        'Else
        '    invGeniusMaterial = ""
        'End If
    Else
        invGeniusMaterial = ""
    End If 'Mapping of material
    
    pnShtMetalHardCoded = invGeniusMaterial
End Function

Public Function dcAddDocPtNum( _
    dcIn As Scripting.Dictionary, _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pn As String
    Dim fn As String
    
    Set rt = dcIn
    If AiDoc Is Nothing Then
    Else
        pn = AiDoc.PropertySets.Item(gnDesign).Item(pnPartNum).Value
        fn = AiDoc.FullFileName
        With rt
            If .Exists(pn) Then
                .Item(pn) = .Item(pn) & "|" & fn
            Else
                .Add pn, fn
            End If
        End With
    End If
    Set dcAddDocPtNum = rt
End Function

Public Function dcAiDocPartNumbers( _
    dcIn As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    For Each ky In dcIn
        Set rt = dcAddDocPtNum(rt, _
            aiDocument(obOf(dcIn.Item(ky))) _
        )
    Next
    Set dcAiDocPartNumbers = rt
End Function

