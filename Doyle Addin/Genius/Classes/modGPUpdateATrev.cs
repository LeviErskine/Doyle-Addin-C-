

Public Function dcGeniusPropsPartRev20180530_withComments( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' WAYPOINTS (search on phrase)
    '''     (NOT Sheet Metal)
    '''
    '''
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
    ''' dcGeniusPropsPartRev20180530_withComments_broken
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
    ''' NOTE: predecessor variant function
    ''' dcGeniusPropsPartPre20180530 moved
    ''' to module modGPUpdateATrev to be
    ''' retained for potential reference,
    ''' prior to eventual deprecation
    ''' and, presumably, removal.
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
    ''' REV[2022.05.05.1110]
    '''     add variable qtRawMatl to store
    '''     material quantity BEFORE applying
    '''     it to Property prRmQty
    '''     Ultimate goal is to separate
    '''     value changes from collection,
    '''     moving the former as far down
    '''     the process as possible, and
    '''     ultimately, to the end.
    Dim qtRawMatl   As Double
    Dim pnStock     As String
    Dim qtUnit      As String
    Dim bomStruct   As Inventor.BOMStructureEnum
    Dim ck          As VbMsgBoxResult
    Dim bd          As aiBoxData
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartRev20180530_withComments = _
        dcGeniusPropsPartRev20180530_withComments( _
            invDoc, New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        
        With invDoc
            ''' REV[2022.05.06.1113]
            ''' add trap here for Content Center Items
            ''' new ones likely won't, indeed CAN'T
            ''' have custom properties, so attempts
            ''' to read them will throw errors.
            ''' '
            ''' this is a stopgap to deal with a run
            ''' in progress. a more thorough revision
            ''' to properly address Content Center
            ''' members (and other purchased parts)
            ''' will be needed when possible.
            If .ComponentDefinition.IsContentMember Then
                'Stop
            End If
            
            ' Get Property Sets
            With .PropertySets
                Set aiPropsUser = .Item(gnCustom)
                Set aiPropsDesign = .Item(gnDesign)
            End With
            
            ' Get Custom Properties...
            ''' REV[2022.05.06.1124]
            ''' embedded Custom Property collection
            ''' process in Else branch of new check
            ''' for Content Center Item.
            ''' HOPEFULLY, this will help bypass
            ''' error triggers when encountring
            ''' Content Center member Items.
            If .ComponentDefinition.IsContentMember Then
                pnStock = ""
                qtRawMatl = 0#
                qtUnit = ""
            Else
                Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                
                ' ...and their initial Values
                If prRawMatl Is Nothing Then
                ''' REV[2022.05.10.1427]
                ''' add check for successful collection
                ''' of custom properties
                    pnStock = ""
                Else
                    pnStock = prRawMatl.Value
                End If
                ''' REV[2022.05.05.1517]
                ''' added trap to catch non-numeric values
                ''' in current Raw Material Quantity Property
                ''' and replace them with zero when encountered.
                If prRmQty Is Nothing Then
                    qtRawMatl = 0#
                ElseIf IsNumeric(prRmQty.Value) Then
                    qtRawMatl = prRmQty.Value
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
            End If
            
            ' Part Number and Family properties
            ' are from Design, NOT Custom set
            Set prPartNum = aiGetProp( _
                aiPropsDesign, pnPartNum)
                'ADDED 2021.03.11
            pnModel = prPartNum.Value
            Set prFamily = aiGetProp( _
                aiPropsDesign, pnFamily)
            ''' REV[2022.05.05.1551]
            ''' added initial Value collection
            ''' for Part Family as well as Number
            
            ''' REV[2022.06.30.1531]
            ''' exported Part Family Value collection
            ''' and Genius check to function famVsGenius
            nmFamily = famVsGenius(pnModel, prFamily.Value)
            
            ''' REV[2022.06.29.1351]
            ''' changed Part Family Value collection
            ''' to check against Genius
            ''' UPDATE: superceded by
            ''' REV[2022.06.30.1531]
            'nmFamily = Split(cnGnsDoyle().Execute( _
                "select Family from vgMfiItems where Item = '" & pnModel & "';" _
            ).GetString(adClipString, , "", "", ""), vbCr)(0)
            'If Len(nmFamily) = 0 Then
            '    nmFamily = prFamily.Value
            'ElseIf Len(prFamily.Value) > 0 Then
            '    If nmFamily <> prFamily.Value Then
            '        ck = MsgBox(Join(Array( _
            '            "Current Model Part Family " & prFamily.Value, _
            '            "differs from Part Family " & nmFamily, _
            '            "reported by Genius.", "", _
            '            "Change to match Genius?", "", _
            '            "(click [CANCEL] to debug)" _
            '        ), vbNewLine), _
            '            vbYesNoCancel + vbQuestion, _
            '            "Match Genius Family?" _
            '        )
            '
            '        If ck = vbCancel Then
            '            Stop 'to debug
            '        ElseIf ck = vbNo Then
            '            nmFamily = prFamily.Value
            '            'to retain model value
            '        Else 'do nothing, and Genius
            '            'Family should prevail
            '        End If
            '    End If
            'Else 'DO NOT SET IT HERE!
            '     'that's supposed to be done below
            'End If
            ''' END of REV[2022.06.29.1351]
            ''' (want to make sure the extent
            ''' of this block is noted)
            
            ''' We should check HERE for possibly misidentified purchased parts
            ''' UPDATE[2018.02.06]: Using new UserForm; see below
            With .ComponentDefinition
                ''' Request #1: Get the Mass in Pounds
                ''' and add to Custom Property GeniusMass
                With .MassProperties
                    ''' REV[2021.11.12]
                    '''     Round mass to nearest ten-thousandth
                    '''     to try to match expected Genius value.
                    '''     This should reduce or minimize reported
                    '''     discrepancies during ETM process.
                    ''' REV[2022.05.06.1349]
                    ''' adding (HOPEFULLY temporary) error trap here
                    ''' to address issue with Application Error
                    ''' when attempting to retrieve Mass.
                    On Error Resume Next
                    Set rt = dcWithProp(aiPropsUser, pnMass, _
                        Round(cvMassKg2LbM * .Mass, 4), rt _
                    )
                    If Err.Number Then
                        ''' suspect it's just an issue with a
                        ''' particular Part Document (for Item SP344)
                        '''
                        ''' however, there may be some indication
                        ''' of an issue relating to a protected
                        ''' Excel worksheet
                        '''
                        ''' see https://docs.microsoft.com/en-us/office/troubleshoot/excel/run-time-error-2147467259-80004005
                        '''
                        ''' Error Number
                        '''     -2147467259
                        '''     (0x80004005)
                        ''' Automation error
                        ''' Unspecified error
                        Stop
                    End If
                    On Error GoTo 0
                End With
                
                '''
                ''' Get BOM Structure type, correcting if appropriate,
                ''' and prepare Family value for part, if purchased.
                '''
                ck = vbNo
                ''' REV[2022.05.06.1118]
                ''' added separate check for Content Center Item.
                ''' (using code from REV[2022.05.06.1113] above)
                If .IsContentMember Then
                    ck = vbYes
                ElseIf InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" & nmFamily & "|") > 0 Then
                ''' REV[2022.06.29.1416]
                ''' added this ElseIf to check against the
                ''' Family collected from Genius into nmFamily.
                ''' (see REV[2022.06.29.1351] above)
                ''' if Genius says it's purchased, it should be.
                    ck = vbYes
                ElseIf InStr(1, invDoc.FullFileName, _
                    "\Doyle_Vault\Designs\purchased\" _
                ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                    "|" & prFamily.Value & "|" _
                ) > 0 Then
                'when possible, change prFamily.Value to nmFamily
                ''' REV[2022.05.06.1118]
                ''' changed If to ElseIf here to "chain" it
                ''' to Content Center check preceding. No need
                ''' to dig deeper if already have that, right?
                
                ''' REV[2018.05.31]: Combined both InStr checks
                ''' by addition to generate a single test for > 0
                ''' If EITHER string match succeeds, the total
                ''' SHOULD exceed zero, so this SHOULD work.
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
                'REV[2018.05.30]: Value produced here
                'will now be held for later processing,
                'more toward the end of this function.
                If bomStruct = kPurchasedBOMStructure Then
                    If .IsContentMember Then
                    ''' NOTE[2022.05.06.1130]
                    ''' just noting this check has been here
                    ''' for some time already. Probably since
                    ''' 2018, noting REV[2018.05.30] above.
                        nmFamily = "D-HDWR"
                    Else
                        nmFamily = "D-PTS"
                        'NOTE: NON Content Center members
                        '       might still be D-HDWR
                        '       Additional checks might
                        '       be recommended
                    End If
                Else
                    'nmFamily = ""
                    ''' REV[2020.05.05.1559]
                    ''' disabled this assignment to avoid
                    ''' clearing an existing Family setting.
                    '''
                    ''' this entire Else branch can probably go.
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
                                ''' REV[2022.04.14.1131]
                                ''' added print statements to inform user
                                ''' of current part number and members
                                ''' of its BOM currently in Genius
                                Debug.Print prPartNum.Value & vbNewLine & vbTab & Join(.Keys, vbNewLine & vbTab)
                                'Stop 'because selection is going
                                'to be a lot more complicated.
                                '(just look at that pnStock
                                ' assignment up there!)
                                
                                pnStock = nuSelector( _
                                ).GetReply(.Keys, pnStock)
                                
                                Debug.Print "Selected " & IIf(Len(pnStock) > 0, pnStock, "(nothing)")
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
                                Debug.Print ; 'Breakpoint Landing
                            ElseIf pnStock = prRawMatl.Value Then 'don't worry.
                                'should only be minor quantity changes
                                'Stop 'and make sure we want to do this.
                                
                                'Set dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                'Deactivated, moved down and out of this If-Then nest.
                                'Search below for active copy
                                
                                Debug.Print ; 'Breakpoint Landing
                            Else 'need to ask User what to go with
                                Debug.Print "=== CURRENT GENIUS MATERIAL DATA ==="
                                'Debug.Print dumpLsKeyVal(dcIn, ":" & vbTab)
                                'ck = newFmTest2().AskAbout(invDoc, _
                                    "Raw Material " & prRawMatl.Value _
                                    & vbNewLine & " for Item" _
                                    , _
                                    "does not match " & pnStock _
                                    & vbNewLine & "indicated in Genius." _
                                    & vbNewLine & vbNewLine _
                                    & "Change to match Genius?" _
                                    & vbNewLine & "(Cancel to debug)" _
                                )
                                ''' REV[2022.04.01.1443]
                                ''' short-circuiting this prompt
                                ''' and assuming automatic material
                                ''' change confirmation at this stage.
                                ''' '
                                ''' user gets another opportunity
                                ''' to confirm below. that should
                                ''' make this one redundant
                                '''
                                ck = vbOK
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
                                ''' FLAG[2022.03.16.1343]: DIALOG NEEDED
                                ''' need to replace this Stop with a dialog
                                ''' to alert user to issue here, assuming
                                ''' there is an issue to address.
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
                        ''' NOTE[2022.04.12.1157]
                        ''' this section might want refinement
                        ''' seems to be trying to determine
                        ''' whether part is clearly sheet metal
                        ''' might want to add something to further
                        ''' determine NON sheet metal status
                        Set dcFP = dcFlatPatVals(.ComponentDefinition)
                        ''' try to get flat pattern data
                        ''' WITHOUT mucking up Properties!
                        ''' Want to avoid dirtying file with
                        ''' changes until absolutely necessary)
                        
                        If dcFP.Exists(pnThickness) Then
                            pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                            ''' NOTE[2022.05.31.1158]
                            ''' this attempt to capture sheet metal item
                            ''' might NOT be appropriate! it appears to be
                            ''' repeated below
                            ''' NOTE[2022.05.31.1146]
                            ''' need to better address failed capture
                            ''' of sheet metal item number. material
                            ''' selection dialog SHOULD be invoked
                            ''' somewhere to address this!
                            ''' see also NOTE[2022.05.31.1149] below
                            If Len(pnStock) = 0 Then
                                Stop
                                'pnStock = InputBox("", "Need material for " & pnModel, pnStock)
                            End If
                            dcFP.Add pnRawMaterial, pnStock
                            'need to change this to use the following more directly
                            'sqlSheetMetal(.ActiveMaterial.DisplayName,CDbl(dcfp.Item(pnThickness)))
                        'ElseIf dcFP.Exists("mtFamily") Then
                            'disabled; won't mess with this one just now
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
                        Debug.Print ConvertToJson(Array(dcIn, dcFP, prPartNum.Value), vbTab)
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
                        ''' #NOTE[2022.04.12.1202]
                        ''' probably NEED to do something here
                        ''' to make sure NON sheet metal part
                        ''' is checked for structural material
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
                                ''' NOTE[2022.05.31.1149]
                                ''' again, sheet metal capture failure
                                ''' needs to be handled, if possible.
                                ''' material selection dialog SHOULD
                                ''' be invoked, but isn't.
                                ''' this might be the place to do it..
                                ''' see also NOTE[2022.05.31.1146] above
                                ''' NOTE[2022.04.12.1258]
                                ''' this "correction" is preventing
                                ''' detection of incorrect material
                                ''' a rewrite will be needed, or else
                                ''' something needs done further up
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
                                    
                                    ''' REV[2022.04.19.0945]
                                    '''     made to following official:
                                    ''' NOTE[2022.01.03]:
                                    '''     Following text SHOULD no longer
                                    '''     be needed. Verify function of
                                    '''     fmTest2 following, and when good,
                                    '''     disable and/or remove this block.
                                    'Debug.Print "!!! NOTICE !!!"
                                    'Debug.Print "Recommended Sheet Metal Stock (" & pnStock & ")"
                                    'Debug.Print "does not match current Stock (" & prRawMatl.Value & ")"
                                    'Debug.Print
                                    'Debug.Print "To continue with no change, just press [F5]. Otherwise,"
                                    'Debug.Print "press [ENTER] on the following line first to change:"
                                    'Debug.Print "prRawMatl.Value = """ & pnStock & """"
                                    'Debug.Print
                                    
                                    ''' REV[2022.04.19.0944]
                                    '''     added check for case mismatch.
                                    '''     if that's the only difference,
                                    '''     no need to bother the user.
                                    If UCase$(prRawMatl.Value) = pnStock Then
                                        ck = vbYes
                                    Else
                                    ''' NOTE[2022.01.03]:
                                    '''     Now using fmTest2(?) to prompt
                                    '''     user as in other checks (above?)
                                        ck = newFmTest2().AskAbout(invDoc, _
                                            "Suggest Sheet Metal change from" _
                                            & vbNewLine & prRawMatl.Value & " to" _
                                            & vbNewLine & pnStock & " for", _
                                            "Change it?" _
                                        )
                                    End If
                                    
                                    If ck = vbCancel Then
                                        Debug.Print ConvertToJson(nuDcPopulator.Setting(pnModel, nuDcPopulator.Setting("from", prRawMatl.Value).Setting("into", pnStock).Dictionary).Dictionary, vbTab)
                                        Stop 'to check things out
                                        'send2clipBdWin10 ConvertToJson(nuDcPopulator.Setting(pnModel,nuDcPopulator.Setting("from",prRawMatl.Value).Setting("into",pnStock).Dictionary).Dictionary,vbTab)
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
                                    pnStock = prRawMatl.Value
                                    ''' REV[2022.03.16.1555]
                                    ''' clearing pnStock along
                                    ''' with prRawMatl.Value
                                    ''' to avoid breaking
                                    ''' at Stop point below
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
                            'With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
                                "vgMfiItems", "Family", "Item", pnStock _
                            ))
                            With cnGnsDoyle().Execute( _
                                "select Family from vgMfiItems where Item='" _
                                & Replace(pnStock, "'", "''") & "';" _
                            )
                            ''' REV[2022.08.26.1055]
                            ''' replaced direct ref to pnStock
                            ''' with Replace operation to "escape"
                            ''' it, re REV[2022.08.19.1416] (below)
                                If .BOF Or .EOF Then
                                    If pnStock <> "0" Then
                                    ''' REV[2022.03.01.1553]
                                    ''' embedded in check
                                    ''' for string value "0"
                                    ''' as this seems to come
                                    ''' up as a legacy issue,
                                    ''' and is readily remedied
                                    ''' in this section. No stop
                                    ''' is needed in that case.
                                        If Len(pnStock) > 0 Then
                                        ''' REV[2022.07.07.1340]
                                        ''' added secondary check for string length.
                                        ''' an empty string requires no user attention.
                                            Stop 'because Material value likely invalid
                                        End If
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
                            ''' NOTE[2022.05.10.1559]
                            ''' see NOTE[2022.05.10.1558]
                            ''' on HOSE AND PIPING (search on this
                            ''' to find in this function) for more
                            ''' robust approach to PROBABLE LAGGING
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
                            Stop 'to check block WITH@1764
                            ''' REV[2022.08.26.1001]
                            ''' placing temporary Stops at start
                            ''' and end of following With block
                            ''' to check use of fields normally
                            ''' requested in SQL select statement.
                            '''
                            'With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
                                "vgMfiItems", "Family", "Item", pnStock _
                            ))
                            'preceding (disabled) With statement
                            'to replace the following, assuming
                            'tests prove successful. if so, it
                            'might permit further streamlining
                            With cnGnsDoyle().Execute( _
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
                                    ''' REV[2022.04.15.1035]
                                    '''     moved Stop statements to head
                                    '''     of their respective branches.
                                    '''     anticipate need to come up
                                    '''     with much better mechanism
                                    '''     to handle "special" raw stock
                                    '''     (read: D/R-PTS stock family)
                                    If mtFamily Like "?-MT*" Then
                                        'Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        'Debug.Print pnModel & "[" & prRmQty.Value & qtUnit & "*" & pnStock & ": " & aiPropsDesign(pnDesc).Value & "]" ' prRawMatl.Value
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
                                        Stop 'NOT SO FAST!
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
                                        
                                        ''' REV[2022.04.04.1350]
                                        ''' disabled immediate mode boilerplate text dump
                                        ''' as user prompt appears to be functioning properly
                                        'Debug.Print "===== CHECK AND VERIFY RAW MATERIAL QUANTITY ====="
                                        'Debug.Print "  If change required, place new values at end"
                                        'Debug.Print "  of lines below for prRmQty.Value and qtUnit."
                                        'Debug.Print "  Press [ENTER] on each line to be changed."
                                        'Debug.Print "  Press [F5] when ready to continue."
                                        'Debug.Print "----- " & pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value & " -----"
                                        'Debug.Print ""
                                        
                                        ''' REV[2022.02.09.0923]
                                        ''' replication of REV[2022.02.09.0919]
                                        ''' from section below: prep to replace
                                        ''' old dimension dump operation with more
                                        ''' compact call to aiBoxData's Dump method
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
                                        
                                        ''' REV[2022.02.08.1446]
                                        ''' removed block of Debug.Print lines
                                        ''' disabled now for some time, as they
                                        ''' do not seem to have been missed.
                                        'Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value); " 'in model. ";
                                        Debug.Print "qtRawMatl = "; CStr(qtRawMatl); " 'in model. ";
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
                                        ''' REV[2022.03.11.1125]
                                        '''     now invoking new UserForm Interface
                                        '''     for Material Quantity determination.
                                        '''     as in REV[2022.03.11.1112] (above)
                                        With nu_fmIfcMatlQty01().SeeUser(invDoc) '.Result()
                                            If .Exists(pnRmQty) Then
                                                ''' REV[2022.04.04.1404]
                                                ''' add checks for value difference
                                                ''' here and to units (below)
                                                ''' REV[2022.04.11.1007]
                                                ''' added additional "guard code"
                                                ''' to avoid error condition resulting
                                                ''' from blank value of property RMQTY
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
                                        'Stop 'because we might want a D-BAR handler
                                        ''' Actually, we might NOT need to stop here
                                        ''' if bar stock is already selected,
                                        ''' because quantities would presumably
                                        ''' have been established already.
                                        ''' Any D-BAR handler probably needs
                                        ''' to be implemented in prior section(s)
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
                                        
                                        prRmQty.Value = qtRawMatl
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        Stop 'because we don't know WHAT to do with it
                                             'and we do NOT want to clear anything
                                             'until we know what's going on!
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        '''     However, might need more handling here.
                                    End If
                                End If
                            Stop 'at end of block WITH@1764
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
                    ''' REV[2022.05.04.1501]
                    ''' adding an option to try to handle
                    ''' hose and piping elements
                    ''' NOTE[2022.05.10.1558]
                    ''' a similar process might be invoked
                    ''' to address PROBABLE LAGGING (search
                    ''' on that to find in this function)
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
                        ''' _
                            & "Raw Material " & prRawMatl.Value _
                            & vbNewLine & "Unit of Measure currently " _
                            & .Value & vbNewLine & vbNewLine _
                            & "Change to " & qtUnit & "?" _
                            & vbNewLine & " " _
                        '''
                        If ck = vbCancel Then
                            Stop
                        ElseIf ck = vbYes Then
                            'pnStock = userChoiceFromDc(nuDcPopulator( _
                                ).Setting("(" & pnStock & ")", pnStock _
                                ).Setting("5/16"" OD HOSE (GR16)", "GR16" _
                                ).Dictionary(), _
                            pnStock)
                            
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
                    ''' NOTE[2022.05.04.1638]
                    ''' originally an alternate branch to both
                    ''' sheet metal and "standard" part handlers,
                    ''' it was decided to move it to the start of
                    ''' the "standard" handler to take advantage
                    ''' of the property setting code there.
                    '''
                    ''' ultimately, things are going to have to be refactored
                    ''' to better manage data gathering and assignment overall.
                    
                    
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
                                Else
                                    Debug.Print ; 'Breakpoint Landing
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
                            'With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
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
                                    ''  ACTION ADVISED[2018.09.14]:
                                    ''  Will need to address this situation
                                    ''  in a more robust manner.
                                    ''  A more thorough query above
                                    ''  might also be called for.
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
                                    ''' REV[2022.05.05.1343]
                                    ''' inserted new check for purchased
                                    ''' (PTS) "material" items. This SHOULD
                                    ''' ultimately replace the following two
                                    ''' ElseIf statements, and consolidate
                                    ''' determination of Part Family.
                                        
                                        ''' REV[2022.05.05.1610]
                                        ''' added preliminary check for RMT
                                        ''' material family, bypassing User
                                        ''' prompt if encountered.
                                        ''' likely want to build on this
                                        ''' to confirm User wants to keep
                                        ''' existing Family setting.
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
                                            '"Part " & pnModel & " uses " & pnStock & vbNewLine & "which is " & mtFamily & " Material."
                                            '"Part " & pnModel & " uses " & pnStock & vbNewLine & "which is not sheet metal." & vbNewLine & "" & vbNewLine & "These parts are usually assigned" & vbNewLine & "to the Riverview family, R-RMT." & vbNewLine & "" & vbNewLine & "Do you want to use this Family?" & vbNewLine & "Click [NO] to see other options." & vbNewLine & "(CANCEL to debug)"
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
                                    ''' note that If and With blocks are
                                    ''' terminated immediately below, here,
                                    ''' whereas the block above extends
                                    ''' right over the following If.
                                    ''' Remember, ultimate goal is to
                                    ''' consolidate duplicate sections, so
                                    ''' don't get too attached to any of this
                                    
                                    ''' REV[2022.05.05.1331]
                                    ''' add second mtFamily check, for D-BAR
                                    ''' not doing anything with it yet, but
                                    ''' might want to check with User about
                                    ''' making part R-RMT or D-RMT.
                                    ''' Note that we don't want to check
                                    ''' EVERY time...
                                    'If mtFamily = "D-BAR" Then
                                    '    'Debug.Print pnModel & " family " & IIf(Len(nmFamily) > 0, nmFamily, "unset")
                                    '    'Debug.Print nmFamily
                                    '    Debug.Print ; 'Breakpoint Landing
                                    '    'nmFamily = "R-RMT"
                                    'End If
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
                                        
                                        ''' REV[2022.04.04.1350]
                                        ''' disabled immediate mode boilerplate text dump
                                        ''' as user prompt appears to be functioning properly
                                        'Debug.Print "===== CHECK AND VERIFY RAW MATERIAL QUANTITY ====="
                                        'Debug.Print "  If change required, place new values at end"
                                        'Debug.Print "  of lines below for prRmQty.Value and qtUnit."
                                        'Debug.Print "  Press [ENTER] on each line to be changed."
                                        'Debug.Print "  Press [F5] when ready to continue."
                                        'Debug.Print "----- " & pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value & " -----"
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
                                        'Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value); " 'in model. ";
                                        Debug.Print "qtRawMatl = "; CStr(qtRawMatl); " 'in model. ";
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
                                        ''' REV[2022.03.11.1112]
                                        '''     now invoking new UserForm Interface
                                        '''     for Material Quantity determination.
                                        '''     see also REV[2022.03.11.1125] (below)
                                        With nu_fmIfcMatlQty01().SeeUser(invDoc) '.Result()
                                            If .Exists(pnRmQty) Then
                                                ''' REV[2022.04.04.1404]
                                                ''' add checks for value difference
                                                ''' here and to units (below)
                                                ''' REV[2022.04.11.1007]
                                                ''' added additional "guard code"
                                                ''' to avoid error condition resulting
                                                ''' from blank value of property RMQTY
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
                                        'Stop 'because we might want a D-BAR handler
                                        ''' Actually, we might NOT need to stop here
                                        ''' if bar stock is already selected,
                                        ''' because quantities would presumably
                                        ''' have been established already.
                                        ''' Any D-BAR handler probably needs
                                        ''' to be implemented in prior section(s)
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
                ''' REV[2022.05.05.1257]
                ''' begin consolidating common steps from end
                ''' of both Sheet Metal and Standard branches.
                
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
        Set dcGeniusPropsPartRev20180530_withComments = rt
    End If
End Function

Public Function dcGeniusPropsPartRev20200409( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGeniusPropsPartRev20200409
    ''' [2020.04.09] begin new revision
    '''
    Dim rt As Scripting.Dictionary
    ''
    Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
    ''
    Dim prFamily    As Inventor.Property
    Dim prRawMatl   As Inventor.Property 'pnRawMaterial
    Dim prRmUnit    As Inventor.Property 'pnRmUnit
    Dim prRmQty     As Inventor.Property 'pnRmQty
    ''
    Dim nmFamily As String
    Dim mtFamily As String
    ''' UPDATE[2018.05.30.01]
    Dim pnStock As String
    Dim qtUnit As String
    Dim bomStruct As Inventor.BOMStructureEnum
    Dim ck As VbMsgBoxResult
    Dim bd As aiBoxData
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartRev20200409 = _
        dcGeniusPropsPartRev20200409( _
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
            
            ' Family property is from Design, NOT Custom set
            Set prFamily = aiGetProp(aiPropsDesign, pnFamily)
            
            ''' We should check HERE for possibly misidentified purchased parts
            ''' UPDATE[2018.02.06.01]: Using new UserForm; see below
            With .ComponentDefinition
                ''' Request #1: Get the Mass in Pounds
                ''' and add to Custom Property GeniusMass
                With .MassProperties
                    Set rt = dcWithProp( _
                        aiPropsUser, pnMass, _
                        Round(cvMassKg2LbM * .Mass, 4), rt _
                    )
                End With
                
                '''
                ''' Get BOM Structure type, correcting if appropriate,
                ''' and prepare Family value for part, if purchased.
                '''
                ''' NOTE[2020.04.09.01]
                ck = vbNo
                ''' UPDATE[2018.05.31.01]
                If InStr(1, invDoc.FullFileName, _
                    "\Doyle_Vault\Designs\purchased\" _
                ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                    "|" & prFamily.Value & "|" _
                ) > 0 Then
                ''' UPDATE[2020.04.09.02]
                    ck = newFmTest2().AskAbout(invDoc, , _
                        "Is this a Purchased Part?" _
                    )
                End If
                
                ''' Check process below replaces duplicate check/responses above.
                ''' NOTE[2020.04.09.02]
                If ck = vbYes Then
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
                'UPDATE[2018.05.30.02]
                If bomStruct = kPurchasedBOMStructure Then
                    If .IsContentMember Then
                        nmFamily = "D-HDWR"
                    Else
                        nmFamily = "D-PTS"
                        ''' NOTE[2020.04.09.03]
                    End If
                Else
                    nmFamily = ""
                End If
            End With
            ''' At this point, nmFamily SHOULD be set
            ''' to a non-blank value if Item is purchased.
            ''' We should be able to check this later on,
            ''' if Item BOMStructure is NOT Normal
            
            ''' UPDATE[2020.04.09.03]
            If bomStruct = kNormalBOMStructure Then
                '----------------------------------------------------'
                If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                '----------------------------------------------------'
                '''
                ''' NOTE[2018.05.31.01]
                    'Request #3:
                    '   Get sheet metal extent area
                    '   and add to custom property "RMQTY"
                    Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                    ''' NOTE[2018.05.30.01]
                    
                    'Moved to start of block to check for NON sheet metal
                    
                    'NOTE: THIS call might best be combined somehow
                    '   with the flat pattern prop pickup above.
                    '   Note especially that if dcFlatPatProps
                    '   FINDS NO .FlatPattern, then there should
                    '   BE NO sheet metal part number!
                    If prRawMatl Is Nothing Then
                        Stop ''' UPDATE[2020.04.09.04]
                        If rt.Exists("OFFTHK") Then
                            ''' UPDATE[2018.05.30.05]
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
                        ''' NOTE[2018.09.14.01]: ACTION ADVISED
                        If Len(prRawMatl.Value) > 0 Then
                            pnStock = prRawMatl.Value
                        Else
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                        End If
                        
                        If Len(pnStock) = 0 Then
                            ''' UPDATE[2018.05.30.03]
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
                                ''' UPDATE[2018.05.30.04]
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
                        End If
                        
                        If Len(pnStock) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                            
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''' NOTE[2018.09.14.02]: ACTION ADVISED
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    
                                    If mtFamily = "DSHEET" Then
                                        'We should be okay. This is sheet metal stock
                                        nmFamily = "D-RMT"
                                        qtUnit = "FT2"
                                        ''' UPDATE[2018.05.30.06]
                                    ElseIf mtFamily = "D-BAR" Then
                                        nmFamily = "R-RMT"
                                        qtUnit = prRmUnit.Value '"IN"
                                        ''may want function here
                                        ''' UPDATE[2018.05.30.07]
                                        Debug.Print aiPropsDesign.Item(pnPartNum).Value; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                        Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                        Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
                                        Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                                        With invDoc.ComponentDefinition.RangeBox
                                            Debug.Print _
                                            (.MaxPoint.X - .MinPoint.X) / 2.54, _
                                            (.MaxPoint.Y - .MinPoint.Y) / 2.54, _
                                            (.MaxPoint.Z - .MinPoint.Z) / 2.54
                                        End With
                                        ''' UPDATE[2020.04.09.05]
                                        Debug.Print ""
                                        Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                                        ''' UPDATE[2020.04.09.05]
                                        Debug.Print "qtUnit = ""IN"""
                                        ''' UPDATE[2020.04.09.05]
                                        Stop 'because we might want a D-BAR handler
                                        ''' UPDATE[2020.04.09.05]
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                                        Stop
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30.08]
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
                                ''' UPDATE[2020.04.09.06]
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
                                    Stop 'and check both so we DON'T
                                    'automatically "fix" the RMUNIT value
                                    
                                    .Value = qtUnit
                                    
                                    If 0 Then Stop 'Ctrl-9 here to skip changing
                                End If
                            End If
                        Else 'we're setting a new quantity unit
                            .Value = qtUnit
                        End If
                    End With
                    Set rt = dcAddProp(prRmUnit, rt)
                    ''' UPDATE[2020.04.09.07]
                    Debug.Print ; 'Another landing line
                    
                '--------------------------------------------'
                Else 'for standard Part (NOT Sheet Metal) ---'
                '--------------------------------------------'
                        ''' NOTE[2018.07.31.01]
                        With newFmTest1()
                            If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                            
                            ''' [2018.07.31.02][by AT]
                            Set bd = nuAiBoxData().UsingInches.SortingDims( _
                                invDoc.ComponentDefinition.RangeBox _
                            )
                            
                            ck = .AskAbout(invDoc, _
                                "Please Select Stock for Machined Part" _
                                & vbNewLine & vbNewLine & bd.Dump(0) _
                            )
                            
                            If ck = vbYes Then
                            ''' UPDATE[2018.05.30.09]
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
                                ''' NOTE[2020.04.09.04]
                            End If
                        End With
                        '''
                        '''
                        '''
                        
                        ''' NOTE[2020.04.09.05]
                        If Len(pnStock) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                            
                            ''' NOTE[2020.04.09.06]
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
                                ''' NOTE[2020.04.09.07]
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
                                        ''' UPDATE[2018.05.30.06]
                                    ElseIf mtFamily = "D-BAR" Then
                                        nmFamily = "R-RMT"
                                        qtUnit = prRmUnit.Value '"IN"
                                        ''may want function here
                                        ''' UPDATE[2018.05.30.07]
                                        Debug.Print aiPropsDesign.Item(pnPartNum).Value; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                        Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                        Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
                                        Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                                        Debug.Print (invDoc.ComponentDefinition.RangeBox.MaxPoint.X - invDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (invDoc.ComponentDefinition.RangeBox.MaxPoint.Y - invDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (invDoc.ComponentDefinition.RangeBox.MaxPoint.Z - invDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54
                                        Debug.Print ""
                                        Debug.Print "PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED."
                                        Debug.Print "PRESS ENTER/RETURN TWICE. THEN CONTINUE."
                                        Debug.Print ""
                                        Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                                        Debug.Print "qtUnit = ""IN"""
                                        Debug.Print ""
                                        Stop 'because we might want a D-BAR handler
                                        ''' UPDATE[2020.04.09.05]
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                                        Stop
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30.08]
                                        Stop 'because we don't know WHAT to do with it
                                    End If
                        Else
                            If 0 Then Stop 'and regroup
                            ''' Things are looking a right royal mess
                            ''' at the moment I'm writing this comment.
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
                                        "  Current : " & prRawMatl.Value, _
                                        "  Proposed: " & pnStock, _
                                        "", "Change It?", "" _
                                    ), vbNewLine), _
                                    vbYesNo, "Change Raw Material?" _
                                )
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
                                    Stop 'and check both so we DON'T
                                    'automatically "fix" the RMUNIT value
                                    
                                    .Value = qtUnit
                                    
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
            Else
                Stop 'because we might need
                    'to do something else
                    'based on an unexpected
                    'BOM Structure
            End If
            
            ' Get the design tracking property set,
            ' and update the Cost Center Property
            If invDoc.ComponentDefinition.IsContentMember Then
                ' Don't muck around with the Family!
            Else
                If Len(nmFamily) > 0 Then
                    prFamily.Value = nmFamily
                    Set rt = dcAddProp(prFamily, rt)
                    'Set rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt)
                End If
            End If
        End With
        
        Call iSyncPartFactory(invDoc) 'Backport Properties to iPart Factory
        Set dcGeniusPropsPartRev20200409 = rt
    End If
End Function
''' NOTE[2018.05.30.01]:
'''     Raw Material Quantity value
'''     SHOULD be set upon return
'''     We may need to review the process
'''     to find an appropriate place
'''     to set for NON sheet metal
''' NOTE[2018.05.31.01]:
'''     At this point, we MAY wish
'''     to check for a valid flat pattern,
'''     and otherwise attempt to verify
'''     an actual sheet metal design.
''' NOTE[2018.07.31.01][by AT]
'''     Duped following block from above
'''     to mod for material assignment
'''     to non-sheet metal part.
'''
'''     Except, this isn't enough.
'''     Also need the code to add
'''     Stock PN to Attribute RM.
'''     That's a whole 'nother
'''     block of code, and likely
'''     best consolidated.
''' [2018.07.31.02][by AT]
'''     Added the following to try to
'''     preselect non-sheet metal stock
'''     [and then disabled the following]
                                '.dbFamily.Value = "D-BAR"
                                '.lbxFamily.Value = "D-BAR"
                                ''' Doesn't quite do it.
                                'With New aiBoxData
                                'Set bd = nuAiBoxData().UsingInches.UsingBox( _
                                    invDoc.ComponentDefinition.RangeBox _
                                )
'
                                'End With
''' NOTE[2018.09.14.01]: ACTION ADVISED
'''     pnStock can probably be set to prRawMatl.Value
'''     and THEN checked for length to see if lookup needed.
'''     This might also allow us to check for machined
'''     or other non-sheet metal parts.
''' NOTE[2018.09.14.02]: ACTION ADVISED
'''     Will need to address this situation
'''     in a more robust manner.
'''     A more thorough query above
'''     might also be called for.
''' NOTE[2020.04.09.01]: This section should check
'''     for Purchased Part status in Genius, as well
'''     as the checks below. BOM Structure should also
'''     be checked, but SETTING it eventually needs
'''     to be shifted to a subsequent operation.
''' NOTE[2020.04.09.02]:
'''     this is where Document's BOMStructure
'''     is set. should be moved to a later stage
''' NOTE[2020.04.09.03]:
'''     [original date unknown]
'''     NON Content Center members
'''     might still be D-HDWR
'''     Additional checks might
'''     be recommended
''' NOTE[2020.04.09.04]
'''     [original date unknown]
'''     We're going to need something here
'''     to make sure raw material gets added
'''     for non sheet metal parts, as well
'''     What we're going to need to do
'''     is refactor this whole bloody thing.
''' NOTE[2020.04.09.05]
'''     [original date unknown]
'''
'''     The following If block is copied
'''     wholesale from sheet metal section above.
'''     Some changes (to be) made to accommodate
'''     machined or other non-sheet metal stock.
'''
'''     Ultimately, whole mess to require refactor.
'''
''' NOTE[2020.04.09.06]
'''     [original date unknown]
'''     This enclosing With block should NOT be necessary
'''     since the newFmTest1 above takes care of collecting
'''     the Stock Family along with the Stock itself
'''
''' NOTE[2020.04.09.07]
'''     [original date unknown]
'''
''' Content formerly here moved BELOW and OUT of this section
''' as it should only require results of newFmTest1 exchange above
'''
''''''
''''''
''''''
'''
''' UPDATE[2018.05.30.01]:
'''     Rename variable Family to nmFamily
'''     to minimize confusion between code
'''     and comment text in searches.
'''     Also add variable mtFamily
'''     for raw material Family name
''' UPDATE[2018.05.30.02]:
'''     Value produced here
'''     will now be held for later processing,
'''     more toward the end of this function.
''' UPDATE[2018.05.30.03]:
'''     Pulling ALL code/text from this section
'''     to get rid of excessive cruft.
'''
'''     In fact, reversing logic to go directly
'''     to User Prompt if no stock identified
'''
'''     IN DOUBLE FACT, hauling this WHOLE MESS
'''     RIGHT UP after initial pnStock assignment
'''     to prompt user IMMEDIATELY if no stock found
''' UPDATE[2018.05.30.04]:
'''     Pulling some extraneous commented code
'''     from here and beginning of block
''' UPDATE[2018.05.30.05]:
'''     Restoring original key check
'''     and adding code for debug
'''     Previously changed to "~OFFTHK"
'''     to avoid this block and its issues.
'''     (Might re-revert if not prepped to fix now)
''' UPDATE[2018.05.30.06]: (two locations)
'''     Moving part family assignment
'''     to this section for better mapping
'''     and updating to new Family names
'''     as well as pulling up qtUnit assignment
''' UPDATE[2018.05.30.07]: (two locations)
'''     As noted above
'''     Will keep Stop for now
'''     pending further review,
'''     hopefully soon
''' UPDATE[2018.05.30.08]: As noted above
'''     However, might need more handling here.
''' UPDATE[2018.05.30.09]:
'''     Pulling some extraneous commented code
'''     from here and beginning of block
''' UPDATE[2018.05.31.01]:
'''     Combined both InStr checks
'''     by addition to generate a single test for > 0
'''     If EITHER string match succeeds, the total
'''     SHOULD exceed zero, so this SHOULD work.
''' UPDATE[2018.02.06.01]:
'''     Using new UserForm
''' UPDATE[2020.04.09.02]:
'''     Remove disabled/outdated code as follows
                    ''' UPDATE[2018.02.06]: Using same
                    '''     new UserForm as noted above.
                    'ck = newFmTest2().AskAbout(invDoc, , _
                        "Is this a Purchased Part?" _
                    )
                'ElseIf InStr(1, _
                    "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                    "|" & prFamily.Value & "|" _
                ) > 0 Then
                    ''' UPDATE[2018.02.06]: Using same
                    '''     new UserForm as noted above.
''' UPDATE[2020.04.09.03]:
'''     Removed disabled/outdated code as follows
            'Request #4: Change Cost Center iProperty.
            'If BOMStructure = Normal, then Family = D-MTO,
            'else if BOMStructure = Purchased then Family = D-PTS.
''' UPDATE[2020.04.09.04]:
'''     Adding Stop here to see if prRawMatl
'''     ever comes up missing inside a sheet metal part
''' UPDATE[2020.04.09.05]: (multiple points)
'''     Removing disabled/obsolete code as follows
                                        'Debug.Print "CURRENT RAW MATERIAL QUANTITY (";
                                        'Debug.Print CStr(prRmQty.Value); ") IS SHOWN BELOW."
                                        'Debug.Print "IF NOT CORRECT, YOU MAY TYPE A NEW VALUE"
                                        'Debug.Print "IN ITS PLACE, AND PRESS ENTER TO CHANGE IT."
                                        'Debug.Print "SOME SUGGESTED VALUES INCLUDE X, Y, AND Z"
                                        'Debug.Print "EXTENTS (ABOVE) OR YOU MAY SUPPLY YOUR OWN."
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        'Debug.Print "YOU MAY ALSO CHANGE THE UNIT OF MEASURE BELOW,"
                                        'Debug.Print "IF DESIRED. BE SURE TO PRESS ENTER/RETURN"
                                        'Debug.Print "AFTER CHANGING EITHER LINE. WHEN FINISHED, "
                                        'Debug.Print "PRESS [F5] TO CONTINUE."
'''
                                        'Debug.Print "qtUnit = """; qtUnit; """"
'''
                                        'Debug.Print ""
                                        'Debug.Print ""
                                        'Debug.Print ""
'''
                                        ''' Actually, we might NOT need to stop here
                                        ''' if bar stock is already selected,
                                        ''' because quantities would presumably
                                        ''' have been established already.
                                        ''' Any D-BAR handler probably needs
                                        ''' to be implemented in prior section(s)
'''
''' UPDATE[2020.04.09.06]:
'''     Removing disabled/obsolete code as follows
                                'Debug.Print "Raw Stock Selection"
                                'Debug.Print "  Current : " & prRawMatl.Value
                                'Debug.Print "  Proposed: " & pnStock
                                'Stop 'because we might not want to change existing stock setting
                                'if
''' UPDATE[2020.04.09.07]:
'''     Removing disabled/obsolete code as follows
                    'Set rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                    ''' Plan to remove commented line above,
                    ''' superceded by the one above that

Public Function dcGnsPartProps( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    ''' NOTES[2021.03.12]
    ''' Don't recall when this function block was created.
    ''' Probably around 2020.04.09, with the generation
    ''' of function dcGeniusPropsPartRev20200409, above.
    '''
    ''' As of this writing, no code present, so will
    ''' use this to rebuild the Part Properties retrieval
    ''' function more or less from the ground up.
    '''
    ''' Primary goal: reconstruct the basic process
    ''' as faithfully as possible, but in a NONdestructive
    ''' manner. That is, avoid changing the Part Document
    ''' in any way, but simply collect as much information
    ''' as is available, and generate whatever else is needed,
    ''' and possible WITHOUT altering the Document.
    '''
    ''' NOTES[2021.03.22]
    ''' Following review of process in functions
    ''' dcGeniusPropsPartRev20180530 and dcFlatPatProps,
    ''' added calls to aiGetProp to retrieve all Property
    ''' items checked and/or set by those functions.
    '''
    ''' Again, this function should NOT attempt to create
    ''' any missing/nonexistent Property items, in order
    ''' to avoid altering the source Document at this stage.
    '''
    
    '''
    '''
    Dim rt As Scripting.Dictionary
    ''
    ''  Property Sets
    Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
    ''
    ''
    ''  Properties
    'Dim prPartNum   As Inventor.Property 'pnPartNum
    'Dim prFamily    As Inventor.Property 'pnFamily
    'Dim prRawMatl   As Inventor.Property 'pnRawMaterial
    'Dim prRmUnit    As Inventor.Property 'pnRmUnit
    'Dim prRmQty     As Inventor.Property 'pnRmQty
    ''
    ''
    ''  Property Values
    Dim pnModel     As String
    Dim nmFamily    As String
    Dim pnStock     As String
    Dim mtFamily    As String
    Dim qtUnit      As String
    ''
    ''
    ''
    Dim bomStruct As Inventor.BOMStructureEnum
    Dim ck As VbMsgBoxResult
    Dim bd As aiBoxData
    
    Set rt = New Scripting.Dictionary
    '''
    
    With invDoc
        ' Get Property Sets
        With .PropertySets
            Set aiPropsUser = .Item(gnCustom)
            Set aiPropsDesign = .Item(gnDesign)
        End With
    End With
        
    With rt
        ' Get Part Number and Family
        ' Properties from Design set
        .Add pnPartNum, aiGetProp(aiPropsDesign, pnPartNum) 'prPartNum
        .Add pnFamily, aiGetProp(aiPropsDesign, pnFamily) 'prFamily
        
        ' Get Custom Properties
        .Add pnRawMaterial, aiGetProp(aiPropsUser, pnRawMaterial) 'prRawMatl
        .Add pnRmUnit, aiGetProp(aiPropsUser, pnRmUnit) 'prRmUnit
        .Add pnRmQty, aiGetProp(aiPropsUser, pnRmQty) 'prRmQty
        'NOTE[2021.03.12]: Removed 'create' flag
        'from these function calls to prevent
        'creation of nonexistent Properties,
        'which would alter the source Document.
        'NOTE ALSO: should try to obtain all other
        'custom Properties intended to generate,
        'in case they're already present.
        
        ' Get Custom Mass/Dimensional Properties
        .Add pnMass, aiGetProp(aiPropsUser, pnMass) '<prMass>
        '.Add pnRmQty, aiGetProp(aiPropsUser, pnRmQty) 'prRmQty
        '   this one already called above
        .Add pnWidth, aiGetProp(aiPropsUser, pnWidth) '<prWidth>
        .Add pnLength, aiGetProp(aiPropsUser, pnLength) '<prLength>
        .Add pnArea, aiGetProp(aiPropsUser, pnArea) '<prArea>
        '.Add "OFFTHK", aiGetProp(aiPropsUser, "OFFTHK") '<prOffThk>
        '   disabled -- not sure if needed any longer
        '   and results in many fewer Prop Dicts
        '   with 'NoVal' Properties
        
        'Set prPartNum = .Item(pnPartNum)
            'pnModel = prPartNum.Value
        'Set prFamily = .Item(pnFamily)
        'Set prRawMatl = .Item(pnRawMaterial)
        'Set prRmUnit = .Item(pnRmUnit)
        'Set prRmQty = .Item(pnRmQty)
        
        Debug.Print ; 'Breakpoint Landing
        'Debug.Print dumpLsKeyVal(mGr1g0f1(rt), ":", ",")
        'Debug.Print dumpLsKeyVal(mGr1g0f1(rt))
        'Stop 'Hard
    End With
    
    '''
    Set dcGnsPartProps = rt
    '''
    '''
    '''
End Function

Public Function dcGnsPartsWithProps( _
    invDoc As Inventor.Document _
) As Scripting.Dictionary
    '''
    ''' function dcGnsPartsWithProps
    '''
    ''' returns Dictionary of Dictionaries
    ''' containing Genius-related Properties
    ''' for each Component of supplied
    ''' Inventor Document, be it Part
    ''' or Assembly.
    '''
    ''' NOTE: actual Dictionary processing
    ''' removed to separate function
    ''' dcGnsPartsWithPropsFromDc
    ''' in order to support invocation
    ''' from other functions w/o need
    ''' for actual source Document
    '''
    'Dim dc As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Dim it As Inventor.PartDocument
    
    Set rt = dcGnsPartsWithPropsFromDc( _
        dcAiDocComponents(invDoc, , 0) _
    )
    
    Debug.Print ; 'Breakpoint Landing
    Set dcGnsPartsWithProps = rt
'Debug.Print dumpLsKeyVal(mGr1g0f2(dcGnsPartsWithProps(ThisApplication.ActiveDocument)), vbNewLine & vbTab, vbNewLine & vbNewLine)
'send2clipBd "{" & dumpLsKeyVal(mGr1g0f2(dcGnsPartsWithProps(ThisApplication.ActiveDocument), ": ", "," & vbNewLine & vbTab), "," & vbNewLine & vbTab, vbNewLine & "}," & vbNewLine & "{") & vbNewLine & "}" & vbNewLine
'send2clipBd ConvertToJson(dcGnsPartsWithProps(ThisApplication.ActiveDocument), " ")
End Function

Public Function dcGnsPartsWithPropsFromDc( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' function dcGnsPartsWithPropsFromDc
    '''
    ''' returns Dictionary of Dictionaries
    ''' containing Genius-related Properties
    ''' for each Inventor Document in supplied
    ''' Dictionary. Intended for invocation
    ''' against a Dictionary of Inventor
    ''' Documents generated by and/or within
    ''' a separate function or procedure.
    '''
    ''' Initial creation intended to support
    ''' companion function dcGnsPartsWithProps
    ''' along with any others which might
    ''' require it
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Dim it As Inventor.PartDocument
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        Set it = aiDocPart(.Item(ky))
        If it Is Nothing Then
            'Stop
        Else
            rt.Add ky, dcGnsPartProps(.Item(ky))
        End If
    Next: End With
    
    Debug.Print ; 'Breakpoint Landing
    Set dcGnsPartsWithPropsFromDc = rt
End Function

Public Function dcOfDcAiPropVals( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    If dc Is Nothing Then
    Else
        With dc: For Each ky In .Keys
            rt.Add ky, dcAiPropValsFromDc( _
                dcOb(.Item(ky)) _
            )
        Next: End With
    End If
    
    Debug.Print ; 'Breakpoint Landing
    Set dcOfDcAiPropVals = rt
End Function

Public Function dcSansNoVals( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    Dim ob As Object
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        it = .Item(ky): With rt
            If IsObject(it) Then
                'Set ob = obOf(it)
                If obOf(it) Is Nothing Then 'don't keep
                    Debug.Print ; 'Breakpoint Landing
                Else
                    .Add ky, it
                    Debug.Print ; 'Breakpoint Landing
                End If
            ElseIf IsNull(it) Then 'don't keep
                Debug.Print ; 'Breakpoint Landing
            ElseIf IsEmpty(it) Then 'don't keep
                Debug.Print ; 'Breakpoint Landing
            Else
                .Add ky, it
                Debug.Print ; 'Breakpoint Landing
            End If
        End With
    Next: End With
    Set dcSansNoVals = rt
End Function

Public Function dcOfOnlyNoVals( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    Dim ob As Object
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        it = .Item(ky): With rt
            If IsObject(it) Then
                'Set ob = obOf(it)
                If obOf(it) Is Nothing Then 'don't keep
                    .Add ky, it
                    Debug.Print ; 'Breakpoint Landing
                Else
                    Debug.Print ; 'Breakpoint Landing
                End If
            ElseIf IsNull(it) Then 'don't keep
                .Add ky, it
                Debug.Print ; 'Breakpoint Landing
            ElseIf IsEmpty(it) Then 'don't keep
                .Add ky, it
                Debug.Print ; 'Breakpoint Landing
            Else
                Debug.Print ; 'Breakpoint Landing
            End If
        End With
    Next: End With
    Set dcOfOnlyNoVals = rt
End Function

Public Function dc4noValStatus(it As Variant, _
    hasVal As Scripting.Dictionary, _
    noVal As Scripting.Dictionary _
) As Scripting.Dictionary
    If IsObject(it) Then
        If obOf(it) Is Nothing Then
            Set dc4noValStatus = noVal
            Debug.Print ; 'Breakpoint Landing
        Else
            Set dc4noValStatus = hasVal
            Debug.Print ; 'Breakpoint Landing
        End If
    ElseIf IsNull(it) Then
        Set dc4noValStatus = noVal
        Debug.Print ; 'Breakpoint Landing
    ElseIf IsEmpty(it) Then
        Set dc4noValStatus = noVal
        Debug.Print ; 'Breakpoint Landing
    Else
        Set dc4noValStatus = hasVal
        Debug.Print ; 'Breakpoint Landing
    End If
End Function

Public Function dcOfNoValStatus( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim nv As Scripting.Dictionary
    Dim hv As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    Dim ob As Object
    Dim ck As Long
    
    Set rt = New Scripting.Dictionary
    Set hv = New Scripting.Dictionary
    Set nv = New Scripting.Dictionary
    rt.Add "HASVAL", hv
    rt.Add "NOVAL", nv
    
    With dc: For Each ky In .Keys
        dc4noValStatus( _
            .Item(ky), hv, nv _
        ).Add ky, .Item(ky)
    Next: End With
    Set dcOfNoValStatus = rt
End Function

Public Function dcOfDcNoValStatus( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcOfDcNoValStatus
    '''
    ''' Given a Dictionary of Dictionaries,
    ''' return a Dictionary of "No Value Status"
    ''' Dictionaries for each Item in the
    ''' source Dictionary
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        rt.Add ky, dcOfNoValStatus( _
            dcOb(.Item(ky)) _
        )
    Next: End With
    Set dcOfDcNoValStatus = rt
End Function

Public Function dcOfDcWithNoVals( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcOfDcWithNoVals
    '''
    ''' Given a Dictionary of Dictionaries,
    ''' return a sub Dictionary of those
    ''' with at least one "No Value" Item
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcOfDcNoValStatus(dc): For Each ky In .Keys
        Set wk = dcOb(dcOb(.Item(ky)).Item("NOVAL"))
        If wk.Count > 0 Then rt.Add ky, wk
    Next: End With
    Set dcOfDcWithNoVals = rt
'Debug.Print txDumpLs(dcOfDcWithNoVals(dcGnsPartsWithProps(ThisApplication.ActiveDocument)).Keys)
End Function

Public Function mGr1g0f1( _
    ob As Inventor.PartDocument, _
    dcIfIs As Scripting.Dictionary, _
    dcIfNot As Scripting.Dictionary _
) As Scripting.Dictionary 'Object
    '''
    '''
    '''
    If ob Is Nothing Then
        Stop
    Else
        If ob.ComponentDefinition.IsContentMember Then
            Set mGr1g0f1 = dcIfIs
        Else
            Set mGr1g0f1 = dcIfNot
        End If
    End If
End Function

Public Function mGr1g0f2( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim pr As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        Set pr = dcOb(.Item(ky))
        If pr Is Nothing Then
            Stop
        Else
            'rt.Add ky, dumpLsKeyVal( _
                mGr1g0f1(pr) _
            )
        End If
    Next: End With
    Set mGr1g0f2 = rt
End Function

Public Function dcGeniusPropsPartPre20180530( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' REV[2022.08.26.1204]
    ''' moved to module modGPUpdateATrev
    ''' from modGPUpdateAT to get it out
    ''' of the way, while keeping it on
    ''' hand for reference, just in case.
    '''
    Dim rt As Scripting.Dictionary
    ''
    Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
    ''
    Dim prFamily    As Inventor.Property
    Dim prRawMatl   As Inventor.Property 'pnRawMaterial
    Dim prRmUnit    As Inventor.Property 'pnRmUnit
    Dim prRmQty     As Inventor.Property 'pnRmQty
    ''
    Dim Family As String
    Dim pnStock As String
    Dim qtUnit As String
    Dim bomStruct As Inventor.BOMStructureEnum
    Dim ck As VbMsgBoxResult
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartPre20180530 = _
        dcGeniusPropsPartPre20180530( _
        invDoc, New Scripting.Dictionary)
    Else
        Set rt = dc
        
        With invDoc
            ' Get the custom property set.
            Set aiPropsUser = .PropertySets.Item(gnCustom)
            Set aiPropsDesign = .PropertySets.Item(gnDesign)
            Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                ''[2018-03-13:Add 1 to create RM property if not found]
                ''[2018-05-15:Add following to get props for RM Unit & Qty]
            Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
            Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
            ''
            Set prFamily = aiGetProp(aiPropsDesign, pnFamily)
            
            
            ''' We should check HERE for possibly misidentified purchased parts
            ''' UPDATE[2018.02.06]: Using new UserForm; see below
            With .ComponentDefinition
                ck = vbNo
                If InStr(1, invDoc.FullFileName, _
                    "\Doyle_Vault\Designs\purchased\" _
                ) > 0 Then
                    ''' UPDATE[2018.02.06]: Using new UserForm
                    '''     to show image and details
                    '''     of part to be verified.
                    ck = newFmTest2().AskAbout(invDoc, , _
                        "Is this a Purchased Part?" _
                    )
                    ''' Commented text below may be removed
                    ''' upon validation of new form process.
                    ''' See below for additional revs pending
                    'Stop
                    'If MsgBox("Is this a purchased part?", _
                        vbYesNo, invDoc.FullFileName _
                    ) = vbYes Then
                        'If .BOMStructure <> kPurchasedBOMStructure Then
                            '.BOMStructure = kPurchasedBOMStructure
                        'End If
                    'End If
                ElseIf InStr(1, _
                    "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                    "|" & prFamily.Value & "|" _
                ) > 0 Then
                ''' This ElseIf condition should be combinable
                ''' with the initial If condition above
                ''' to simplify this check process.
                ''' All/most text in this clause should be
                ''' redundant and removable, once updated
                ''' check process has been validated.
                    ''' UPDATE[2018.02.06]: Using same
                    '''     new UserForm as noted above.
                    ck = newFmTest2().AskAbout(invDoc, , _
                        "Is this a Purchased Part?" _
                    )
                    'Stop
                    'If MsgBox("Is this a purchased part?", _
                        vbYesNo, invDoc.FullFileName _
                    ) = vbYes Then
                        '.BOMStructure = kPurchasedBOMStructure
                    'End If
                End If
                
                ''' Check process below replaces duplicate check/responses above.
                ''' Should be able to merge back into main branch
                ''' once code above is validated and refactored.
                If ck = vbYes Then
                    If .BOMStructure <> kPurchasedBOMStructure Then
                        .BOMStructure = kPurchasedBOMStructure
                    End If
                End If
                
                bomStruct = .BOMStructure
            End With
            
            'Request #1: Get the Mass in Pounds and add to Custom Property GeniusMass
            With .ComponentDefinition.MassProperties
                Set rt = dcWithProp(aiPropsUser, pnMass, Round(.Mass * cvMassKg2LbM, 4), rt)
            End With
            
            '----------------------------------------------------'
            If .SubType = guidSheetMetal Then 'for SheetMetal ---'
            '----------------------------------------------------'
                'Request #4: Change Cost Center iProperty.
                'If BOMStructure = Normal, then Family = D-MTO,
                'else if BOMStructure = Purchased then Family = D-PTS.
                'With .ComponentDefinition
                
                If bomStruct = kNormalBOMStructure Then '.BOMStructure
'If prRawMatl.Value = "" Or cnGnsDoyle().Execute("select I.Family from vgMfiItems As I where I.Item='" & prRawMatl.Value & "';").Fields("Family").Value = "DSHEET" Then
                    'Request #3:
                    '   Get sheet metal extent area
                    '   and add to custom property "RMQTY"
                    Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                    
                    'Moved to start of block to check for NON sheet metal
                    
                    'NOTE: THIS call might best be combined somehow
                    '   with the flat pattern prop pickup above.
                    '   Note especially that if dcFlatPatProps
                    '   FINDS NO .FlatPattern, then there should
                    '   BE NO sheet metal part number!
                    If prRawMatl Is Nothing Then
                        If rt.Exists("~OFFTHK") Then
                            pnStock = ""
                        Else
                            Stop 'because we don't know IF this is sheet metal yet
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                        End If
                    Else
                        If prRawMatl.Value = "" Then
                            'Stop 'because we're not sure what we have.
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                        Else ''
                            Stop
                            'With cnGnsDoyle().Execute(sqlOf_simpleSelWhere( _
                                "vgMfiItems", "Family", "Item", prRawMatl.Value _
                            ))
                            With cnGnsDoyle().Execute( _
                                "select i.Family from vgMfiItems i " & _
                                "where i.Item='" & prRawMatl.Value & "';" _
                            )
                                With .Fields("Family")
                                    If .Value = "DSHEET" Then
                                        'We should be okay. This is sheet metal stock
                                    ElseIf .Value = "D-BAR" Then
                                        Stop 'because we might want a D-BAR handler
                                    Else
                                        Stop 'because we don't know WHAT do with it
                                    End If
                                End With
                            End With
                            pnStock = prRawMatl.Value
                        End If
                    End If
                    
                    If Len(pnStock) > 0 Then
                        'Stop
                        Family = "D-MTO"
                        ''' This needs to change.
                        ''' Got code above that might better determind
                        ''' whether this should be D-MTO or R-MTO
                        ''' and don't forget about D/R-PTS/O options
                        '''
                        ''' NEW IDEA[2018-05-04]
                        ''' We SHOULD be able to guess between D-MTO and R-MTO
                        ''' based on the family of the raw material,
                        ''' as determined in the prior section.
                        ''' Plan to adjust the code here and above
                        ''' to allow for that opportunity.
                        '''
                    Else
                        '''
                        ''' We MIGHT have an incorrectly marked PURCHASED part
                            'Stop
                        ''' We'll want to see about fixing that here, maybe?
                        '''
                        
                        With newFmTest1()
                            'aiSMdef.Document
                            If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                            If .AskAbout(invDoc, "No Stock Found! Please Review") = vbYes Then
                            'Join(Array( _
                                Join(Array("NO STOCK# for", _
                                    Format(invSheetMetalThickness, "0.000") & "in", _
                                    invSheetMetalMaterial), " " _
                                ), _
                                "in " & docName, " ", "Stop/pause here?" _
                            ), vbNewLine)
                                With .ItemData '.Synch
                                    If .Exists(pnFamily) Then
                                        Family = .Item(pnFamily)
                                        Debug.Print pnFamily & "=" & Family
                                    End If
                                    
                                    If .Exists(pnRawMaterial) Then
                                        pnStock = .Item(pnRawMaterial)
                                        Debug.Print pnRawMaterial & "=" & pnStock
                                    End If
                                End With
                                'Stop
                            End If
                        End With
                    End If
                    
'Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
'Set rt = dcAddProp(prRmUnit, rt)
                    qtUnit = "FT2"
                    prRawMatl.Value = pnStock
                    Set rt = dcAddProp(prRawMatl, rt)
                    'Set rt = dcWithProp(aiPropsUser, pnRawMaterial, pnStock, rt)
                    '''
                    'If aiGetProp(aiPropsUser, pnRmUnit) Is Nothing Then
                        'Stop
                    'Else
                        'If aiGetProp(aiPropsUser, pnRmUnit).Value <> "FT2" Then
'''         'Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                    If Len(prRmUnit.Value) > 0 Then
                        If prRmUnit.Value <> qtUnit Then
                            Stop 'so we DON'T automatically "fix" the RMUNIT value
                            '"FT2"
                        End If
                    Else 'we're setting a new quantity unit
                    End If
                    'End If
                    ''' When this Stop activates, skip the next line
                    Set rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                    ''' Want to change this part to allow for alternate RMUNIT values
                    ''' When prior Stop is activated, use Ctrl-F9
                    ''' to continue at the Stop line below.
                    If 0 Then Stop 'to give us a skipover point
                    
                    'Moved Flat Pattern data collection to beginning of block
                    'to facilitate detection of possible NON sheet metal part
'Else
    'Debug.Print prRawMatl.Value
    'Stop
'End If
                ElseIf bomStruct = kPurchasedBOMStructure Then '.BOMStructure
                    Family = "D-PTS"
                Else
                    Stop 'because we might need to do something else
                End If
                'End With
            '--------------------------'
            Else 'for standard Part ---'
            '--------------------------'
                'Request #2: Change Cost Center iProperty.
                'If BOMStructure = Purchased and not content center,
                'then Family = D-PTS, else Family = D-HDWR.
                With .ComponentDefinition
                    If bomStruct = kPurchasedBOMStructure _
                    And .IsContentMember = False Then '.BOMStructure
                        Family = "D-PTS"
                    Else
                        'Family = "D-HDWR"
                        Family = ""
                    End If
                End With
            End If 'Sheetmetal vs Part
            
            ' Get the design tracking property set,
            ' and update the Cost Center Property
            If Len(Family) > 0 Then
                Set rt = dcWithProp(aiPropsDesign, pnFamily, Family, rt)
            End If
        End With
        
        Call iSyncPartFactory(invDoc) 'Backport Properties to iPart Factory
        Set dcGeniusPropsPartPre20180530 = rt
    End If
End Function

