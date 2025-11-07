

Public Function dcGnsMatlSpecPairings() As Scripting.Dictionary
    '''
    ''' dcGnsMatlSpecPairings -- Genius Raw Material Spec Relations
    '''     Return a Dictionary of Dictionaries
    '''     keyed to each Specification value
    '''     found in ANY Spec field of any
    '''     Raw Material Item, each listing
    '''     all OTHER Spec values found in
    '''     conjunction with each value.
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    Dim dcVl As Scripting.Dictionary
    Dim dcAl As Scripting.Dictionary
    Dim kyVl As Variant
    Dim dxAl As Variant
    Dim kyAl As String
    
    Set rt = New Scripting.Dictionary
    
    Set wk = dcDxFromRecSetDc(dcFromAdoRS( _
        cnGnsDoyle().Execute(sqlOf_MatlSpecXref()) _
    ))
    If wk Is Nothing Then
    ElseIf wk.Exists("val") Then
        Set dcVl = wk.Item("val")
        With dcVl
            For Each kyVl In .Keys
                rt.Add kyVl, New Scripting.Dictionary
            Next
            
            For Each kyVl In .Keys
                Set dcAl = rt.Item(kyVl)
                With dcOb(.Item(kyVl))
                For Each dxAl In .Keys
                    kyAl = dcOb(.Item(dxAl)).Item("also")
                    With rt
                    If .Exists(kyAl) Then
                        dcAl.Add kyAl, .Item(kyAl)
                    Else
                        Stop 'because something went wrong
                    End If: End With
                Next: End With
            Next
        End With
    Else
    End If
    
    Set dcGnsMatlSpecPairings = rt
End Function

Public Function dcOfDcWithXrefsDep1st( _
    dc As Scripting.Dictionary, _
    Optional wk As Scripting.Dictionary = Nothing, _
    Optional pt As String = "#" _
) As Scripting.Dictionary
    '''
    ''' dcOfDcWithXrefsDep1st
    '''     Replace rudundant / recursive Dictionary
    '''     Objects in hierarchical Dictionary structure
    '''     '
    '''     This is a depth-first implementation, which
    '''     might locate an initial Dictionary reference
    '''     deep inside an early branch before finding a
    '''     shallower instance that might be preferable.
    '''     '
    '''     A breadth-first implementation might be preferred.
    '''
    Dim rt As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim ar As Variant
    Dim ky As Variant
    Dim sp As String
    
    If wk Is Nothing Then
        Set rt = dcOfDcWithXrefsDep1st(dc, _
        New Scripting.Dictionary)
    Else
        Set rt = New Scripting.Dictionary
        With dc
            For Each ky In .Keys
                ar = Array(.Item(ky))
                Set ck = dcOb(obOf(ar(0)))
                
                If ck Is Nothing Then
                Else
                    If wk.Exists(ck) Then
                        ar = Array(wk.Item(ck))
                    Else
                        ''' prep new $ref path
                        sp = pt & "/" & CStr(ky)
                        
                        ''' add new $ref to wk
                        With nuDcPopulator( _
                        ).Setting("$ref", sp)
                            wk.Add ck, .Dictionary
                        End With
                        
                        ''' go ahead and process
                        ''' subdictionary
                        ar = Array(dcOfDcWithXrefsDep1st(ck, wk, sp))
                        ''' with new $ref in wk BEFORE the call,
                        ''' it should be picked up for any new
                        ''' references to the same directory
                    End If
                End If
                
                rt.Add ky, ar(0)
            Next
        End With
        
        Set dcOfDcWithXrefsDep1st = rt
    End If
    
    Set dcOfDcWithXrefsDep1st = rt
'send2clipBdWin10 ConvertToJson(dcOfDcWithXrefsDep1st(dcGnsMatlSpecPairings()), vbTab)
End Function

Public Function dcOfDcWithXrefsBrd1st(dc As Scripting.Dictionary, _
    Optional wk As Scripting.Dictionary = Nothing, _
    Optional pt As String = "#" _
) As Scripting.Dictionary
    '''
    ''' dcOfDcWithXrefsBrd1st
    '''     Replace rudundant / recursive Dictionary
    '''     Objects in hierarchical Dictionary structure
    '''     '
    '''     This is a depth-first implementation, which
    '''     might locate an initial Dictionary reference
    '''     deep inside an early branch before finding a
    '''     shallower instance that might be preferable.
    '''     '
    '''     A breadth-first implementation might be preferred.
    '''
    Dim rt As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim ls As Scripting.Dictionary
    Dim ar As Variant
    Dim ky As Variant
    Dim sp As String
    
    If wk Is Nothing Then
        Set rt = dcOfDcWithXrefsBrd1st(dc, _
        New Scripting.Dictionary)
    Else
        ''' create returned Dictionary
        Set rt = New Scripting.Dictionary
        
        ''' create local working
        ''' Dictionary of Dictionaries
        Set ls = New Scripting.Dictionary
        
        ''' being processing
        ''' supplied Dictionary
        With dc
            ''' first pass: collect and process
            ''' all sub Dictionary Objects
            For Each ky In .Keys
                Set ck = dcOb(obOf(.Item(ky)))
                
                If Not ck Is Nothing Then
                If wk.Exists(ck) Then
                    ''' add existing $ref Dictionary
                    ''' to Dictionary list. thinking
                    ''' recursion should NOT be an issue
                    ls.Add ky, wk.Item(ck)
                Else
                    ''' add new Dictionary to list
                    ''' for subsequent recursion
                    ls.Add ky, .Item(ky)
                    
                    ''' prep new $ref path
                    sp = pt & "/" & CStr(ky)
                    
                    With wk
                        ''' add new $ref Dictionary
                        .Add ck, New Scripting.Dictionary
                        
                        ''' add path to Dictionary
                        dcOb(.Item(ck)).Add "$ref", sp
                    End With
                End If
                End If
            Next
            
            For Each ky In .Keys
                If ls.Exists(ky) Then
                    rt.Add ky, dcOfDcWithXrefsBrd1st(ls.Item(ky), _
                        wk, pt & "/" & CStr(ky) _
                    )
                Else
                    rt.Add ky, .Item(ky)
                End If
            Next
        End With
    End If
    
    Set dcOfDcWithXrefsBrd1st = rt
'send2clipBdWin10 ConvertToJson(dcOfDcWithXrefsBrd1st(dcGnsMatlSpecPairings()), vbTab)
End Function

Public Function dcGnsMatlSpecPairings4json() As Scripting.Dictionary
    '''
    ''' dcGnsMatlSpecPairings4json -- check on dcGnsMatlSpecPairings
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim k2 As Variant
    
    Set rt = dcGnsMatlSpecPairings()
    
    With rt: For Each ky In .Keys()
        '.Item(ky) = Join(dcOb(.Item(ky)).Keys)
        With dcOb(.Item(ky))
        For Each k2 In .Keys()
            .Item(k2) = Join(dcOb(.Item(k2)).Keys)
        Next: End With
    Next: End With
    
    Set dcGnsMatlSpecPairings4json = rt
'send2clipBdWin10 ConvertToJson(dcGnsMatlSpecPairings4json(), vbTab)
End Function

Public Function dcSpecSubsetWith( _
    txSpec As String, _
    inDc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If inDc.Exists(txSpec) Then
        Set dcSpecSubsetWith = dcKeysInCommon(inDc, _
            dcOb(inDc.Item(txSpec)), 1 _
        )
    Else
        Set dcSpecSubsetWith = New Scripting.Dictionary
    End If
'Debug.Print Join(dcSpecSubsetWith("ROUND", dcSpecSubsetWith("BAR", dcGnsMatlSpecPairings())).Keys)
End Function

Public Function dcSpecSubsetWithAll( _
    dcSpec As Scripting.Dictionary, _
    inDc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = inDc
    For Each ky In dcSpec.Keys
        Set rt = dcSpecSubsetWith(CStr(ky), rt)
    Next
    Set dcSpecSubsetWithAll = rt
End Function

Public Function dcSpecSetFromUser( _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim dc As Scripting.Dictionary
    Dim fm As fmSelectorList
    Dim nx As String
    
    Set rt = New Scripting.Dictionary
    Set dc = dcGnsMatlSpecPairings()
    'Debug.Print Join(dc.Keys)
    
    Do
        Set fm = nuSelFromDict(dc)
        nx = fm.GetReply(, "")
        
        If Len(nx) > 0 Then
            rt.Add nx, nx
            Set dc = dcSpecSubsetWith(nx, dc)
            If dc.Count = 0 Then nx = ""
        End If
    Loop While Len(nx) > 0
    
    'Stop
    Set dcSpecSetFromUser = rt
'Debug.Print Join(dcSpecSetFromUser().Keys)
End Function

Public Function d2g3f1( _
    Part As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' d2g3f1 -- Return Dictionary
    '''     of relevant Part Properties
    '''     and information for use in
    '''     Genius data extraction
    '''
    Dim rt As Scripting.Dictionary
    Dim txPartNum As String
    
    If dc Is Nothing Then
        Set rt = d2g3f1(Part, _
        New Scripting.Dictionary)
    Else
        Set rt = dc
        
        With Part
            With .PropertySets.Item(gnDesign)
                rt.Add pnPartNum, .Item(pnPartNum)
                rt.Add pnFamily, .Item(pnFamily)
            End With
            
            'rt.Add "subType", .SubType 'aiSubType
            
            'With .ComponentDefinition
            '    rt.Add "bomStr", .BOMStructure 'aiBomType
            '
            '    With nuAiBoxData().SortingDims(.RangeBox)
            '        With .UsingInches()
            '            rt.Add "Width", .SpanX
            '            rt.Add "Length", .SpanY
            '            rt.Add "Height", .SpanZ
            '        End With
            '    End With
            'End With
        End With
        
        Set rt = dcGnsInfoCompDef(aiCompDefOf(Part), rt)
        'aiCompDefOf replaces obAiCompDefAny
    End If
    
    Set d2g3f1 = rt
End Function

Public Function dcGnsInfoAiDocBase( _
    AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoAiDocBase (formerly d2g3f1a)
    '''     Return Dictionary of Document Properties
    '''     and information relevant to Genius
    '''     for data extraction
    '''
    Dim rt As Scripting.Dictionary
    Dim txPartNum As String
    
    If dc Is Nothing Then
        Set rt = dcGnsInfoAiDocBase(AiDoc, _
        New Scripting.Dictionary)
    Else
        Set rt = dc
        
        With AiDoc
            With .PropertySets.Item(gnDesign)
                rt.Add pnPartNum, .Item(pnPartNum)
                rt.Add pnFamily, .Item(pnFamily)
            End With
            
            If False Then
            rt.Add "subType", .SubType 'aiSubType
            rt.Add "docType", .DocumentType
            rt.Add "dsbType", .DocumentSubType.DocumentSubTypeID
            End If
        End With
        
        Set rt = dcGnsInfoCompDef(aiCompDefOf(AiDoc), rt)
        'aiCompDefOf replaces obAiCompDefAny
    End If
    
    Set dcGnsInfoAiDocBase = rt
End Function

Public Function dcGnsInfoCompDef( _
    CpDef As Inventor.ComponentDefinition, _
    Optional dcWkg As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoCompDef -- Generate and/or populate
    '''     Dictionary (new or supplied) with data for
    '''     Genius from supplied ComponentDefinition.
    ''' This is the "generic" variant, which dispatches
    '''     the supplied ComponentDefinition to a variant
    '''     more specific to its Class. (some Class
    '''     variants remain to be implemented)
    ''' Note that this function follows the convention
    '''     of a recursive call with a new Dictionary
    '''     object when none is supplied. Duplication
    '''     of the basic function structure should ensure
    '''     this pattern is followed by all specialized
    '''     variants. While this should not usually be
    '''     necessary under normal usage (dispatch to
    '''     specialized variants from here), it should
    '''     help accommodate the possibility of direct
    ''      calls from other client functions.
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcWkg
    If rt Is Nothing Then
        Set rt = dcGnsInfoCompDef(CpDef, _
        New Scripting.Dictionary)
    ElseIf CpDef Is Nothing Then 'Do Nothing
        'cuz we got Nothing to Do With Nothing
    Else
        With CpDef '.ComponentDefinition
            rt.Add "bomStr", .BOMStructure 'aiBomType
            If .BOMStructure = kNormalBOMStructure Then rt.Add "Type", "M"
            If .BOMStructure = kPurchasedBOMStructure Then rt.Add "Type", "R"
            
            With nuAiBoxData().UsingBox(.RangeBox) '.SortingDims
                With .UsingInches() 'WARNING[2021.12.15]
                ''' Forcing inch conversion MAY lead
                ''' to issues in future development.
                ''' It is absolutely ESSENTIAL that
                ''' unit measurement be tracked and
                ''' kept consistent throughout the
                ''' entire management process.
                    rt.Add pnLength, Round(.SpanX, 6)
                    rt.Add pnWidth, Round(.SpanY, 6)
                    rt.Add "Height", Round(.SpanZ, 6)
                End With
            End With
        End With
        
        If TypeOf CpDef Is Inventor.SheetMetalComponentDefinition Then
            Set rt = dcGnsInfoCompDefShtMtl(CpDef, rt)
        ElseIf TypeOf CpDef Is Inventor.WeldmentComponentDefinition Then
            Stop 'using general Assembly handler
            Set rt = dcGnsInfoCompDefAssy(CpDef, rt)
        ElseIf TypeOf CpDef Is Inventor.WeldsComponentDefinition Then
            Stop 'using general Assembly handler
            Set rt = dcGnsInfoCompDefAssy(CpDef, rt)
        ElseIf TypeOf CpDef Is Inventor.PartComponentDefinition Then
            Set rt = dcGnsInfoCompDefPart(CpDef, rt)
        ElseIf TypeOf CpDef Is Inventor.AssemblyComponentDefinition Then
            Set rt = dcGnsInfoCompDefAssy(CpDef, rt)
        Else
        End If
    End If
    
    Set dcGnsInfoCompDef = rt
End Function

Public Function dcGnsInfoCompDefShtMtl( _
    CpDef As Inventor.SheetMetalComponentDefinition, _
    Optional dcWkg As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoCompDefShtMtl -- Generate and/or populate Dictionary
    '''     (new or supplied) with data for Genius
    '''     from supplied ComponentDefinition.
    ''' This is the Assembly variant.
    '''     '
    '''
    Dim rt As Scripting.Dictionary
    Dim wd As Double 'width
    Dim lg As Double 'length
    Dim ht As Double 'height
    Dim tk As Double 'thickness
    Dim ar As Double 'area
    Dim ck As Double 'check height vs thickness
    
    Dim rm As Scripting.Dictionary
    Dim s6 As String
    Dim ky As Variant
    
    Set rt = dcWkg
    If rt Is Nothing Then
        Set rt = dcGnsInfoCompDefShtMtl( _
        CpDef, New Scripting.Dictionary)
    Else
        Set rt = dcGnsInfoCompDefPart(CpDef, rt)
        With rt
            If Not .Exists("SPEC06") Then
                Stop
                .Add "SPEC06", steelSpec6("", 1)
            End If
            s6 = .Item("SPEC06")
            
            If Not .Exists("RMLIST") Then
                .Add "RMLIST", New Scripting.Dictionary
            End If
            Set rm = dcOb(.Item("RMLIST"))
        End With
        
        With CpDef
            tk = .Thickness.Value / cvLenIn2cm
            ' NOTE conversion to Inches from Centimeters.
            ' keep in mind we're grabbing Thickness HERE
            ' and will use Height (below) in an effort
            ' to validate the Flat Pattern, and determine
            ' this Part is MEANT to be Sheet Metal.
            
            If .HasFlatPattern Then
                With .FlatPattern
                    With nuAiBoxData().UsingBox(.RangeBox)
                        With .UsingInches()
                            ht = Round(.SpanZ, 6)
                            ' remember, Height here is meant
                            ' to verify Sheet Metal Part
                            
                            lg = Round(.SpanX, 6)
                            wd = Round(.SpanY, 6)
                        End With
                    End With
                End With
                
                ck = Round(Abs(ht - tk), 6)
                If ck > 0.002 Then
                    With dcFlatPatSpansByVertices(.FlatPattern)
                        If ht > .Item("Z") Then
                            ht = .Item("Z")
                            
                            Debug.Print .Item("X") - lg
                            Debug.Print .Item("Y") - wd
                            Debug.Print ; 'Breakpoint Landing
                        End If
                    End With
                Else
                End If
                
                If Round(Abs(ht - tk), 6) > 0 Then
                End If
                
                ar = lg * wd '.SpanX * .SpanY
                'does this need to be divided by 144?
                'to get to ft^2? or do we stick to in^2?
            Else
                With rt
                    If .Exists(pnLength) Then
                        lg = .Item(pnLength)
                        '.Remove pnLength
                    End If
                    
                    If .Exists(pnWidth) Then
                        wd = .Item(pnWidth)
                        '.Remove pnWidth
                    End If
                    
                    If .Exists("Height") Then
                        ht = .Item("Height")
                        '.Remove "Height"
                    End If
                End With
                
                ar = 0 'STOPGAP[2021.12.08]
                ''' might want to consider using this
                ''' to store whatever material quantity
                ''' might be obtained, regardless
                ''' of stock type
            End If
            '''
            ''' At this point, we should have either
            ''' likely dimensions of the flat pattern, OR
            ''' the original dimensions of the part itself.
            '''
            ''' The next step is to determine whether they
            ''' are consistent with a valid sheet metal part.
            ''' If not, it's likely a structural one.
            '''
            ''' The key criterion is how closely the height
            ''' dimension matches the given thickness.
            '''
            ck = Round(Abs(ht - tk), 6)
            If ck > 0.002 Then
            Else
            End If
            
            ''' REV[2021.12.15]:
            '''     add material option collection
            '''     specific to sheet metal
            With dcGnsMatlOps(dcCtOfEach( _
                Array(tk, lg, wd, ht) _
            ), s6)
                For Each ky In .Keys
                If Not rm.Exists(ky) Then
                    rm.Add ky, .Item(ky)
                End If: Next
            End With
            
            '''
            '''
            '''
            With rt
                ' first, remove any previous
                ' dimensional values
                If .Exists(pnLength) Then .Remove pnLength
                If .Exists(pnWidth) Then .Remove pnWidth
                If .Exists("Height") Then .Remove "Height"
                '(not sure this is the best way
                ' but going to try it for now)
                
                .Add pnThickness, tk
                .Add pnLength, lg
                .Add pnWidth, wd
                .Add pnArea, ar
                .Add "Height", ht
            End With
        End With
    End If
    
    Set dcGnsInfoCompDefShtMtl = rt
End Function

Public Function dcFlatPatSpansByVertices( _
    smFlat As Inventor.FlatPattern _
) As Scripting.Dictionary
    '''
    ''' dcFlatPatSpansByVertices -- get extents of
    '''     Sheet Metal Flat Pattern
    '''     from a scan of its Vertices.
    '''     this is a last resort,
    '''     in case an erroneous Z span
    '''     reported from the Range Box
    '''     fails to match Thickness.
    '''
    Dim rt As Scripting.Dictionary
    Dim vx As Inventor.Vertex
    Dim xmn As Double
    Dim xmx As Double
    Dim ymn As Double
    Dim ymx As Double
    Dim zmn As Double
    Dim zmx As Double
    
    Set rt = New Scripting.Dictionary
    
    If Not smFlat.Body Is Nothing Then
        With smFlat.Body '.Vertices'.RangeBox
        For Each vx In .Vertices
            With vx.Point
                If .X < xmn Then xmn = .X
                If .X > xmx Then xmx = .X
                If .Y < ymn Then ymn = .Y
                If .Y > ymx Then ymx = .Y
                If .Z < zmn Then zmn = .Z
                If .Z > zmx Then zmx = .Z
            End With
        Next
        End With
        '''
        
        '''
    End If
    
    With rt
        .Add "X", xmx - xmn
        .Add "Y", ymx - ymn
        .Add "Z", zmx - zmn
    End With
    
    Set dcFlatPatSpansByVertices = rt
End Function

Public Function dcGnsInfoCompDefPart( _
    CpDef As Inventor.PartComponentDefinition, _
    Optional dcWkg As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoCompDefPart -- Generate and/or populate Dictionary
    '''     (new or supplied) with data for Genius
    '''     from supplied ComponentDefinition.
    ''' This is the general Part variant.
    '''     It's grown somewhat in complexity
    '''     since development's begun.
    ''' Here is rough flow map:
    '''     (inside component definition)
    '''     - stage supplied Dictionary for return
    '''       (start a new one, if none supplied
    '''        one USUALLY should be)
    '''     - get Mass -- don't add to Dictionary yet
    '''       (no data should be added to Dictionary
    '''        until all data are collected and verified)
    '''     - get Active Material, and its name
    '''     - use this to set target Spec 6
    '''     (inside returning Dictionary)
    '''     - collect length, width, and height
    '''       dimensions from Dictionary
    '''       (this is why one should be supplied)
    '''     - collect raw material candidate items
    '''       from Genius into a Recordset, using
    '''       an SQL query generated from collected
    '''       dimensions, and material Spec 6
    '''     - generate Dictionary of candidates
    '''       from the Recordset, keyed on item names
    '''     - add data to Dictionary:
    '''       - mass
    '''       - material name
    '''       - spec 6
    '''       - Dictionary of raw material
    '''         item candidates
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary 'may be temporary
    Dim d2 As Scripting.Dictionary 'may be temporary
    Dim mt As Inventor.MaterialAsset
    Dim mtName As String
    Dim s6 As String
    Dim ms As Double 'mass
    Dim ky As Variant
    Dim ck As Double
    
    Set rt = dcWkg
    If rt Is Nothing Then
        Set rt = dcGnsInfoCompDefPart(CpDef, _
        New Scripting.Dictionary)
    Else
        With CpDef
            With .MassProperties
                ms = Round(.Mass * cvMassKg2LbM, 4)
                'Round( _
                    ThisApplication( _
                    ).UnitsOfMeasure( _
                    ).ConvertUnits( _
                    .Mass, "kg", "lb" _
                ), 4)
                ''' Apparently, empty parentheses may be
                ''' placed after both ThisApplication
                ''' and UnitsOfMeasure without error.
                ''' This makes it rather easy to lay out
                ''' code in a nice, compact, and maybe
                ''' even more readable form.
            End With
            
            'ptNumShtMetal
            Set mt = aiDocPart(.Document).ActiveMaterial
            If mt Is Nothing Then
                mtName = ""
                'Stop
            Else
                mtName = mt.DisplayName
                ''' NOTE[2021.12.03]:
                ''' mt also has a .CategoryName that might
                ''' want included in full material designator.
                ''' also not sure if constant pnMaterial is
                ''' best choice for Dictionary Key, though
                ''' probably so. Keep this point in mind.
            End If
            s6 = steelSpec6(mtName)
        End With
        
        With rt
            Set wk = New Scripting.Dictionary
            For Each ky In Array(pnLength, pnWidth, "Height")
                If .Exists(ky) Then
                    wk.Add ky, Round(CDbl(.Item(ky)), 6)
                '''
                ''' REV[2021.12.15]:
                '''     disabling everything from here
                '''     to end of Then block. Instead,
                '''     will collect target dimension
                '''     values, then submit them to
                '''     function dcCtOfEach to generate
                '''     the "histogram".
                '''
                '''     that way, the same function may
                '''     be used by other callers, like
                '''     dcGnsInfoCompDefShtMtl
                '''
                '    'ck = Round(CDbl(.Item(ky)), 6)
                '    ''' NOTE[2021.12.08]:
                '    '''     The conversion kludge here
                '    '''     might NOT be reliable for
                '    '''     long term use. Be prepared
                '    '''     to deal with issues here.
                '    With wk
                '    If .Exists(ck) Then
                '        .Item(ck) = _
                '        .Item(ck) + 1
                '    Else
                '        .Add ck, 1 'ky
                '    End If: End With
                    ''' NOTE[2021.12.08]:
                    '''     Dictionary wk counts occurrences
                    '''     of each dimension value. While not
                    '''     presently used, the count might
                    '''     prove helpful in prioritizing raw
                    '''     material candidate items.
                    ''' REV[2021.12.15]:
                    '''     occurrence count has been moved
                    '''     below, outside the For-Next loop,
                    '''     and into a call to new function
                    '''     dcCtOfEach.
                End If
            Next
            Set wk = dcCtOfEach(wk.Items)
            If wk.Count = 0 Then wk.Add 0.075, 1
            ' another kludge to trap an error
            ' which should NOT occur as long as
            ' a prepared Dictionary is supplied.
            
            '''
            ''' Here is where we'll attempt to collect
            ''' raw material Item candidates from Genius
            '''
            
            'present setup -- plan to change
            '"select d.v "
            'Debug.Print "from (values (" & txDumpLs(wk.Keys, "), (") & "))"
            '" as d(v)"
            
            'future proposal -- counts occurrences
            '"select d.v, d.c "
            'Debug.Print "from (values (" & dumpLsKeyVal(wk, ", ", "), (") & "))"
            '" as d(v, c)"
            
            Set wk = dcGnsMatlOps(wk, s6)
            ''' REV[2021.12.15]:
            '''     preceding line replaces With block below,
            '''     moving Genius material options request
            '''     to function dcGnsMatlOps, so it can be
            '''     called from other functions, like,
            '''     again, dcGnsInfoCompDefShtMtl
            'With cnGnsDoyle()
            '    Dim rs As ADODB.Recordset
            '
            '    'wk.RemoveAll
            '    On Error Resume Next
            '
            '    Err.Clear
            '    Set rs = .Execute( _
            '    sqlOf_GnsMatlOptions( _
            '        s6, wk.Keys _
            '    ))
            '
            '    If Err.Number = 0 Then
            '        With dcFromAdoRS(rs, "") 'Set wk =
            '        For Each ky In .Keys
            '            Set d2 = dcOb(.Item(ky))
            '            If d2 Is Nothing Then
            '                Stop
            '            Else
            '                wk.Add d2.Item("Item"), d2
            '            End If
            '            'Stop
            '            ''' ENDOFDAY[2021.12.08]:
            '            '''     Need to setup process of remapping
            '            '''     raw material Items from Genius
            '            '''     to their Item names
            '        Next: End With
            '
            '        rs.Close
            '    Else
            '        Stop
            '        Err.Clear
            '    End If
            '    On Error GoTo 0
            '
            '    .Close
            'End With
            
            '.Add pnRawMaterial, wk
            
            .Add pnMass, ms
            .Add pnMaterial, mtName
            .Add "SPEC06", s6
            
            'If False Then
            'not quite ready for this one yet
            .Add "RMLIST", wk
            'End If
            Debug.Print ; 'Breakpoint Landing
        End With
    End If
    
    Set dcGnsInfoCompDefPart = rt
End Function

Public Function dcGnsInfoCompDefAssy( _
    CpDef As Inventor.AssemblyComponentDefinition, _
    Optional dcWkg As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoCompDefAssy -- Generate and/or populate Dictionary
    '''     (new or supplied) with data for Genius
    '''     from supplied ComponentDefinition.
    ''' This is the general Assembly variant.
    '''     '
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcWkg
    If rt Is Nothing Then
        Set rt = dcGnsInfoCompDefAssy(CpDef, _
        New Scripting.Dictionary)
    Else
        With CpDef
            With .MassProperties
                rt.Add pnMass, Round( _
                .Mass * cvMassKg2LbM, 4)
                '''
                ''' see dcGnsInfoCompDefPart
                ''' for alternate
                ''' implementation
                '''
            End With
        End With
    End If
    
    Set dcGnsInfoCompDefAssy = rt
End Function

Public Function dcGnsInfoCompDefTBD( _
    CpDef As Inventor.ComponentDefinition, _
    Optional dcWkg As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoCompDefTBD (formerly d2g4f2zz)
    '''     Generate and/or populate Dictionary
    '''     (new or supplied) with data for Genius
    '''     from supplied ComponentDefinition.
    ''' This is the <TBD> variant. (formerly <zz>)
    '''     Use it as a template for others.
    '''     (be sure to modify comments accordingly)
    '''
    Dim rt As Scripting.Dictionary
    
    Set rt = dcWkg
    If rt Is Nothing Then
        Set rt = dcGnsInfoCompDefTBD(CpDef, _
        New Scripting.Dictionary)
    Else
        With CpDef
        End With
    End If
    
    Set dcGnsInfoCompDefTBD = rt
End Function

Public Function dcGnsInfoSQLitem( _
    Item As String _
) As Scripting.Dictionary
    '''
    ''' dcGnsInfoSQLitem -- Return a Dictionary
    '''     of Part data from Genius
    '''     for the indicated Item
    '''
    Dim rt As Scripting.Dictionary
    Dim mt As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim ky As Variant
    
    With cnGnsDoyle()
        On Error Resume Next
        Err.Clear
        Set rs = .Execute(sqlOf_GnsPartInfo(Item)) 'sqlOf_ASDF
        If Err.Number = 0 Then
            Set rt = dcFromAdoRSrow(rs, "")
            'With rs
            '    If Not .EOF Then
            '        .MoveNext
            '        If Not .EOF Then
            '            Stop 'to handle multiple raw materials
            '            Debug.Print ; 'Breakpoint Landing
            '        End If
            '    End If
            '
            '    .Close
            'End With
        Else
            Debug.Print Err.Number
            Debug.Print Err.Description
            Stop
        End If
        
        Set rs = .Execute(sqlOf_GnsPartMatl(Item)) 'sqlOf_ASDF
        If Err.Number = 0 Then
            Set mt = dcFromAdoRS(rs, "")
            With mt
                If .Count > 0 Then
                    If .Count > 1 Then
                        Stop 'to handle multiple raw materials
                        Debug.Print ; 'Breakpoint Landing
                    Else
                        With dcOb(.Item(.Keys(0)))
                        For Each ky In .Keys
                            If rt.Exists(ky) Then
                                Stop 'to deal with collision
                                'which should NOT happen here
                                'because field sets returned
                                'by each query should have
                                'NO names in common.
                            Else
                                rt.Add ky, .Item(ky)
                            End If
                        Next: End With
                    End If
                End If
            End With
        Else
            Debug.Print Err.Number
            Debug.Print Err.Description
            Stop
        End If
        
        On Error GoTo 0
        .Close
    End With
    
    Set dcGnsInfoSQLitem = rt
'Debug.Print dumpLsKeyVal(dcGnsInfoSQLitem(aiProperty(d2g3f1(aiDocPart(userChoiceFromDc())).Item(pnPartNum)).Value), " = ")
End Function

Public Function d2g3f4(Part As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' d2g3f4 -- Return a Dictionary
    '''     of Properties and info from
    '''     Inventor Part Document for
    '''     Genius Interface.
    '''
    Dim rt As Scripting.Dictionary
    'Dim rs As ADODB.Recordset
    
    Set rt = dcProps4genius(Part, d2g3f1(Part, dc), 0)
    
    Set d2g3f4 = rt
'Debug.Print dumpLsKeyVal(d2g3f4(aiDocPart(userChoiceFromDc())), " = ")
End Function

Public Function d2g3f5( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    '''
    ''' d2g3f5 -- Gather Dictionaries of Inventor
    '''     Properties and Genius info from supplied
    '''     Document for correlation and potential
    '''     revision.
    ''' REV[2021.12.15]:
    '''     Parameter Part renamed to AiDoc, with Class
    '''     changed from PartDocument to the more general
    '''     Document, as it would appear all supporting
    '''     functions will accept and work with it.
    '''
    Dim dcPt As Scripting.Dictionary
    '   base Genius info and inherent Properties
    Dim dcPr As Scripting.Dictionary
    '   base + custom Genius Properties
    Dim dcVlAi As Scripting.Dictionary
    '   values of all collected Properties
    Dim dcVlPr As Scripting.Dictionary
    '   values of all collected Properties
    Dim dcGn As Scripting.Dictionary
    '   information from Genius database
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set dcPt = dcGnsInfoAiDocBase(AiDoc)
    Set dcVlAi = dcMapAiProps2vals(dcPt)
    ''' REV[2021.12.16]:
    '''     additional value Dictionary
    '''     collects values ONLY from
    '''     the inherent document data
    '''     and Properties, which gets
    '''     overridden at the next step:
    
    Set dcPr = dcProps4genius(AiDoc, dcCopy(dcPt), 2)
    ''' REV[2021.12.15] argument 2 replaces 0 in order
    ''' to generate references to missing Properties
    ''' (see dcGnsPropsListed) without trying to
    ''' create them. That way, client functions may
    ''' be made aware of Properties that need created.
    '''
    ''' Modifications to those functions might be needed
    ''' to be prepared for missing Properties, whose
    ''' names will map to Nothing (void references)
    ''' UPDATE[2021.12.16]:
    '''     just happened today. created new function
    '''     blankIfNoValElseSelf to address this issue.
    '''     see dcMapAiProps2vals for application
    '''
            '= dcProps4genius(AiDoc, d2g3f1(AiDoc, dc), 0)
            '= d2g3f4(AiDoc)
    
    Set dcVlPr = dcMapAiProps2vals(dcPr) 'dcPt
    With dcVlPr
        For Each ky In Array( _
            pnThickness, pnWidth, pnLength, pnArea, pnRmQty _
        ) ''' THIS IS A KLUDGE
            ' to temporarily "fix" an issue with Width,
            ' Length, and Area values that don't match
            ' up between Inventor Properties and Genius,
            ' even when their numeric values are equal.
            '
            ' A more thorough review and revision
            ' will probably be needed eventually.
            '
            ' REV[2021.12.07]:
            ' Thickness is also affected
            ' and has been added to the list.
            '
            ''' REV[2021.12.16]:
            '''     added pnRmQty to list as quick and
            '''     dirty method to force a blank value
            '''     to zero, and prevent an error in
            '''     the correction code after the loop.
            '''
            '''     This must be the sort of cruft Joel
            '''     Spolsky was talking about in that
            '''     essay of his. Still not a justified
            '''     defense for crud programming, which
            '''     this is, let's face it! Right there
            '''     with you, Brando!
            '''
            If .Exists(ky) Then
                .Item(ky) = Val(Split( _
                    "0" + CStr(.Item(ky)), " " _
                )(0))
                ''' REV[2021.12.16]:
                '''     minor switch in order of operations:
                '''     now prepending "0" to string BEFORE
                '''     splitting, to avoid the emtpy array
                '''     returned by splitting an empty string
            End If
        Next
        
        If .Exists(pnRmQty) Then
            .Item(pnRmQty) = Round(.Item(pnRmQty), 8)
            ' THIS is intended to fix a precision discrepancy
            ' between Inventor, which seems to store material
            ' quantity with at least twelve digits of precision,
            ' and Genius, which keeps only eight.
        End If
    End With
    
    Set dcGn = dcGnsInfoSQLitem( _
        dcVlPr.Item(pnPartNum) _
    )
    
    Set rt = New Scripting.Dictionary
    With rt
        .Add "aiVal", dcVlAi
        .Add "inv", dcVlPr
        .Add "gns", dcGn
        .Add "prp", dcPr
    End With
    
    Set d2g3f5 = rt
'send2clipBdWin10 ConvertToJson(d2g3f5(aiDocPart(userChoiceFromDc())), vbTab)
End Function

Public Function d2g3f5as( _
    Assy As Inventor.AssemblyDocument, _
    Optional ThisToo As Long = 0 _
) As Scripting.Dictionary
    '''
    ''' d2g3f5as -- Assembly counterpart to d2g3f5
    '''     not sure what's actually to be done with it yet.
    '''     probably just remove it; d2g3f5 can handle both.
    '''
    Dim dc As Scripting.Dictionary
    
    Set dc = dcRemapByPtNum( _
        dcAiDocComponents( _
            Assy, , ThisToo _
        ) _
    )
End Function

Public Function dcMapAiProps2vals( _
    dc As Scripting.Dictionary, _
    Optional Flags As Long = 0 _
) As Variant
    '''
    ''' dcMapAiProps2vals --
    '''     Return a Dictionary
    '''     containing the Values of
    '''     any Inventor Properties
    '''     in supplied Dictionary,
    '''     with all other members
    '''     returned as they are.
    '''
    ''' related functions:
    '''     dcOfDcAiPropVals
    '''     dcAiPropValsFromDc
    '''     dcOfPropsInAiDoc
    '''
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcNewIfNone(dc): For Each ky In .Keys
        rt.Add ky, blankIfNoValElseSelf( _
        valIfAiPropElseSelf(.Item(ky)))
        ''' REV[2021.12.16]:
        '''     add call to blankIfNoValElseSelf
        '''     against RESULT of valIfAiPropElseSelf
        '''     so blind checks against expected key
        '''     values don't fail when a null Object
        '''     (AKA Nothing) is encountered.
    Next: End With
    Set dcMapAiProps2vals = rt
End Function

Public Function valIfAiPropElseSelf( _
    vl As Variant _
) As Variant
    '''
    ''' valIfAiPropElseSelf --
    '''     Return the Value of any
    '''     supplied Inventor Property.
    '''     Any other type of argument
    '''     should be returned directly.
    '''
    Dim pr As Inventor.Property
    
    If IsObject(vl) Then
        Set pr = aiProperty(obOf(vl))
        If pr Is Nothing Then
            Set valIfAiPropElseSelf = vl
        Else
            valIfAiPropElseSelf = pr.Value
        End If
    Else
        valIfAiPropElseSelf = vl
    End If
End Function

Public Function blankIfNoValElseSelf( _
    vl As Variant _
) As Variant
    '''
    ''' blankIfNoValElseSelf --
    '''     Return the Value of any
    '''     supplied Inventor Property.
    '''     Any other type of argument
    '''     should be returned directly.
    '''
    Dim pr As Inventor.Property
    
    If IsObject(vl) Then
        If obOf(vl) Is Nothing Then
            blankIfNoValElseSelf = ""
        Else
            Set blankIfNoValElseSelf = vl
        End If
    ElseIf IsNull(vl) Then
        blankIfNoValElseSelf = ""
    ElseIf IsEmpty(vl) Then
        blankIfNoValElseSelf = ""
    Else
        blankIfNoValElseSelf = vl
    End If
End Function

Public Function d2g3f7( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    '''
    ''' d2g3f7 --
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    'Set rt = New Scripting.Dictionary
    With d2g3f5(AiDoc)
        Debug.Print ; 'Breakpoint Landing
        Set rt = _
            dcTreeReKeyedInPlc("src1", "gns", _
            dcTreeReKeyedInPlc("src0", "inv", _
            dcWBQbyCmpResult( _
            dcCmpTextOf2dc( _
                .Item("inv"), _
                .Item("gns") _
            ) _
        )))
        
        With rt
            'Stop
        End With
        rt.Add "prp", .Item("prp")
        rt.Add "doc", AiDoc
        ' might want this in order
        ' to grab custom PropertySet
        
        'With dcKeysInCommon(.Item("inv"), .Item("gns"))
        'End With
    End With
    Set d2g3f7 = rt
'send2clipBdWin10 ConvertToJson(d2g3f7(aiDocPart(userChoiceFromDc())), vbTab)
'send2clipBdWin10 ConvertToJson(nuDcPopulator().Setting(Format$(Now, "\[YYYY.MM.DD@HH.NN.SS\]"), d2g3f7(aiDocPart(userChoiceFromDc()))).Dictionary(), vbTab)
End Function

Public Function d2g3f8( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Scripting.Dictionary
    '''
    ''' d2g3f8 --
    '''
    Dim rt As Scripting.Dictionary
    'Dim ck As Inventor.Document
    Dim ky As Variant
    
    If AiDoc Is Nothing Then
        With ThisApplication
            If .ActiveDocument Is Nothing Then
                Stop
            Else
                Set rt = d2g3f8(.ActiveDocument)
            End If
        End With
    Else
        Set rt = New Scripting.Dictionary
        
        With nuPicker(New kyPickAiPartVsAssy _
        ).AfterScanning(dcAiDocComponents(AiDoc))
            '''
            With .dcIn() 'Parts
                For Each ky In .Keys
                    rt.Add ky, d2g3f7(aiDocPart(obOf(.Item(ky))))
                Next
            End With
            
            With .dcOut() 'Assemblies
            End With
        End With
    End If
    
    Set d2g3f8 = rt
'send2clipBdWin10 ConvertToJson(nuDcPopulator().Setting(Format$(Now, "\[YYYY.MM.DD@HH.NN.SS\]"), d2g3f8(aiDocument(obOf(userChoiceFromDc())))).Dictionary(), vbTab)
End Function

Public Function dcTreeMembersWithKey( _
    tg As Variant, dc As Scripting.Dictionary, _
    Optional wk As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' dcTreeMembersWithKey (formerly d2g5f1)
    '''     Given a Dictionary that might contain
    '''     other Dictionaries, check it and any
    '''     sub Dictionaries for target key (tg)
    '''     and return a Dictionary of those
    '''     Dictionaries containing it, each
    '''     keyed to the number already found.
    '''     This should ensure a unique key
    '''     for each match found, with no
    '''     need to track any other keys.
    '''
    ''' The ultimate goal of this function is to
    '''     support a Key Find/Replace operation
    '''     across a hierarchy of Dictionaries.
    '''
    ''' This is initially and specifically to map
    '''     comparison keys "src0" and "src1" to
    '''     the names of sources they represent.
    '''
    ''' This is of course the 'Find' component
    '''     of the ultimate product
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If wk Is Nothing Then
        Set rt = dcTreeMembersWithKey(tg, dc, _
        New Scripting.Dictionary)
    Else
        Set rt = wk
        If Not dc Is Nothing Then
            With dc
                If .Exists(tg) Then
                    With rt
                        .Add .Count, dc
                    End With
                End If
                
                For Each ky In .Keys
                    Set rt = dcTreeMembersWithKey(tg, _
                    dcOb(obOf(.Item(ky))) _
                    , rt)
                Next
            End With
        End If
    End If
    Set dcTreeMembersWithKey = rt
End Function

Public Function dcTreeMemWithReplcmt( _
    rp As Variant, dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcTreeMemWithReplcmt (formerly d2g5f2)
    '''     Given a Dictionary of Dictionaries,
    '''     check for any Dictionary containing target
    '''     replacement Key rp, and return a Dictionary
    '''     containing any results.
    '''
    ''' This is a check for potential Key collisions.
    '''     The Dictionary returned should be empty.
    '''
    ''' This is presently accomplished by first calling
    '''     dcTreeMembersWithKey against the supplied Dictionary,
    '''     which is normally expected to be the result
    '''     of a PRIOR call to dcTreeMembersWithKey using the target
    '''     key to be replaced.
    '''
    ''' It is therefore possible that the supplied
    '''     Dictionary might contain replacement key rp,
    '''     and thus be included in the local result.
    '''     That Dictionary should NOT be included
    '''     in the FINAL result returned.
    '''
    ''' It is therefore necessary to scan the result
    '''     of the local dcTreeMembersWithKey call, and remove it,
    '''     if found.
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = dcTreeMembersWithKey(rp, dc)
    With rt: For Each ky In .Keys
        If dcOb(obOf(.Item(ky))) Is dc Then
            Stop
            .Remove ky
        End If
    Next: End With
    Set dcTreeMemWithReplcmt = rt
End Function

Public Function dcTreeReKeyedInPlc( _
    tg As Variant, rp As Variant, _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcTreeReKeyedInPlc (formerly d2g5f3)
    '''     Given a target Key tg (to be replaced),
    '''     a replacement Key rp, and a Dictionary
    '''     that includes other Dictionaries,
    '''     attempt to replace all instances of the
    '''     target Key with the replacement Key in
    '''     all Dictionaries within the hierarchy.
    '''
    ''' Note that this is a DESTRUCTIVE replacement
    '''     operation. A preferable option might be
    '''     to generate a NEW hierarchical Dictionary
    '''     replicating the original, with the desired
    '''     key substitution. Will consider that for
    '''     a later implementation.
    '''
    ''' Note also that error checking/handling
    '''     in this implementation is presently minimal.
    '''     A more robust process should also be considered.
    '''
    Dim wk As Scripting.Dictionary
    Dim ck As Scripting.Dictionary
    Dim ky As Variant
    
    Set wk = dcTreeMembersWithKey(tg, dc)
    Set ck = dcTreeMemWithReplcmt(rp, wk)
    
    If ck.Count > 0 Then
        Stop
    Else
        With wk: For Each ky In .Keys
            With dcOb(obOf(.Item(ky)))
            ''' A Dictionary object is assumed, here.
            ''' Though typically risky in a With block,
            ''' it SHOULD be guaranteed here,
            ''' so no error should occur.
            ''' Don't be surprised if it does, though.
            If .Exists(rp) Then
                Stop 'because this
                'should NOT be happening!
                'dcTreeMemWithReplcmt
                'should have caught
                'any replacement key
                'collisions already.
            Else
                'a proper error handler might
                'be desired here in future
                'On Error Resume Next
                'for now, keep disabled
                
                'note order of operations here
                .Add rp, .Item(tg) 'FIRST
                
                .Remove tg 'ONLY AFTER
                'associated Item is added
                'under replacement Key
                
                'this ensures the associated Item
                'is retained under AT LEAST ONE Key,
                'and not lost in the event of some
                'fault or error, which really
                'shouldn't occur, BUT...
                
                'On Error GoTo 0
                'potential error handler
                'to end here, unless moved
            End If: End With
        Next: End With
    End If
    
    Set dcTreeReKeyedInPlc = dc
End Function

Public Function userChoiceFromDc(Optional dc _
    As Scripting.Dictionary = Nothing, _
    Optional ifNone As Variant = Nothing _
) As Variant
    '''
    ''' userChoiceFromDc (formerly d2g3f2)
    '''     Request User Selection from
    '''     a Dictionary of options.
    '''
    '''     A list of Dictionary Keys is
    '''     presented to the user. After
    '''     User selects a Key, matching
    '''     Item is returned for use.
    '''
    Dim ck As VbMsgBoxResult
    Dim msNoSel As String
    Dim rp As Variant
    Dim rt As Variant
    
    ''' REV[2023.05.17.1304]
    ''' add ifNone processing to present
    ''' User with information on default
    ''' option(s), if supplied
    On Error Resume Next
        Err.Clear
        msNoSel = CStr(ifNone)
        
        If Err.Number = 0 Then
            If Len(msNoSel) > 0 Then
            msNoSel = "Use default value (" & msNoSel & ")?"
            End If
        Else
            msNoSel = ""
            Err.Clear
            
            If IsObject(ifNone) Then
                If Not ifNone Is Nothing Then
                msNoSel = Join(Array( _
                    "Use default " & TypeName(ifNone) & " Object?", _
                    "(Object details not available)" _
                ), vbNewLine)
                End If
            Else
                Stop
            End If
        End If
    On Error GoTo 0
    
    If Len(msNoSel) > 0 Then
        msNoSel = vbNewLine & msNoSel
    End If
    'msNoSel = Join(Array( _
        "User selection was requested" _
        , "with no available options!" _
        , msNoSel _
    ), vbNewLine)
    
    If dc Is Nothing Then
        rt = Array(userChoiceFromDc(dcAiDocsVisible()))
    Else
        If dc.Count > 0 Then
            rp = nuSelFromDict(dc _
            ).GetReply()
                ' , , , , , _
                , Join(Array( _
                    "No option selected!", _
                    msNoSel _
                ), vbNewLine)
            If dc.Exists(rp) Then
                rt = Array(dc.Item(rp))
            Else
                rt = Array(ifNone)
            End If
        Else
            ck = MsgBox(Join(Array( _
                "User selection was requested", _
                "with no available options!" _
            ), vbNewLine _
            ), vbOKOnly, "No Options!" _
            )
            ''' IIf(Len(msNoSel) > 0, _
                vbYesNo, vbOKOnly _
            ), _
                msNoSel
            '''
            If ck = vbNo Then
                rt = Array(Nothing)
                'not the best option, but
                'not sure what else to do
            Else
                rt = Array(ifNone)
            End If
        End If
    End If
    
    If IsObject(rt(0)) Then
        Set userChoiceFromDc = rt(0)
    Else
        Let userChoiceFromDc = rt(0)
    End If
End Function

Public Function dcGnsPrpPtDvl_2021_1112( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim dc01 As Scripting.Dictionary
    Dim dcVlGn  As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    ''
    Dim aiPartNum   As String
    Dim aiFamily    As String
    Dim aiSubType   As String
    ''
    'Dim aiPropsUser As Inventor.PropertySet
    Dim aiPropsDesign As Inventor.PropertySet
    ''
    Dim prPartNum   As Inventor.Property
    Dim prFamily    As Inventor.Property
    ''
    'Dim aiPartNum   As String 'will be same as gnPartNum
    'Dim aiPartFam   As String
    'Dim aiMatlNum   As String
    'Dim aiMatlFam   As String
    'Dim aiMatlQty   As Double
    'Dim aiQtyUnit   As String
    Dim aiBomType   As Inventor.BOMStructureEnum
    ''
    ''
    Set rt = New Scripting.Dictionary
    Set dc01 = New Scripting.Dictionary
    
    With invDoc
        ' Get Property Sets
        With .PropertySets
            'Set aiPropsUser = .Item(gnCustom)
            Set aiPropsDesign = .Item(gnDesign)
        End With
        aiBomType = .ComponentDefinition.BOMStructure
        aiSubType = .SubType
    End With
    
    ' Get Part Number and Family
    ' Properties from Design set
    With aiPropsDesign
        Set prPartNum = .Item(pnPartNum)
        Set prFamily = .Item(pnFamily)
    End With
    
    ' Get Values of Part Number
    ' and Family Properties
    aiPartNum = prPartNum.Value
    aiFamily = prFamily.Value
    ''' NOTE[2021.11.12]
    '''     The preceding three sections
    '''     can PROBABLY be consolidated
    '''     into one, using fewer variables
    '''     and probably just one With block
    
    'dc01
    With cnGnsDoyle()
        On Error Resume Next
        Err.Clear
        Set rs = .Execute( _
            sqlOf_ASDF(aiPartNum) _
        ) '
        If Err.Number = 0 Then
            Set dcVlGn = dcFromAdoRSrow(rs, "")
            With dcVlGn
                'gnPartNum = .Item("Item")
                'gnPartFam = .Item("Family")
                'gnBomType = .Item("bomStr")
                ''Set fdOrder = .Item("Ord")
                'gnMatlNum = .Item("Material")
                'gnMatlFam = .Item("MtFamily")
                'gnMatlQty = .Item("Qty")
                'gnQtyUnit = .Item("Unit")
            End With
            
            With rs
                If .BOF And .EOF Then
                Else
                    With .Fields
                        'Set fdItem = .Item("Item")
                        'gnPartNum = .Item("Item")
                        'should ALWAYS match aiPartNum
                        'IF it's found in Genius
                        'otherwise, always BLANK
                        
                        'Set fdFamly = .Item("Family")
                        'gnPartFam = .Item("Family").Value
                        
                        'gnBomType = .Item("bomStr").Value
                        
                        'Set fdOrder = .Item("Ord")
                        
                        'Set fdMatrl = .Item("Material")
                        'gnMatlNum = .Item("Material").Value
                        
                        'Set fdMtFam = .Item("MtFamily")
                        'gnMatlFam = .Item("MtFamily").Value
                        
                        'Set fdQty = .Item("Qty")
                        'gnMatlQty = .Item("Qty").Value
                        
                        'Set fdUnit = .Item("Unit")
                        'gnQtyUnit = .Item("Unit").Value
                        
                        'Stop 'to check things out
                        'not doing anything else with this yet
                        'but want to start matching against model
                    End With
                    
                    .MoveNext
                    If Not .EOF Then
                        Stop 'to handle multiple raw materials
                        Debug.Print ; 'Breakpoint Landing
                    End If
                End If
                
                .Close
            End With
        Else
            Debug.Print Err.Number
            Debug.Print Err.Description
            Stop
        End If
        On Error GoTo 0
        .Close
    End With
    
    If aiBomType = kNormalBOMStructure Then
        If aiSubType = guidSheetMetal Then
            'try to get flat pattern data here
        End If
    ElseIf aiBomType = kPurchasedBOMStructure Then
        Stop
    ElseIf aiBomType = kPhantomBOMStructure Then
        Stop
    ElseIf aiBomType = kInseparableBOMStructure Then
        Stop
    ElseIf aiBomType = kReferenceBOMStructure Then
        Stop
    ElseIf aiBomType = kNormalBOMStructure Then
        Stop
    ElseIf aiBomType = kPhantomBOMStructure Then
        Stop
    ElseIf aiBomType = kDefaultBOMStructure Then
        Stop
    ElseIf aiBomType = kVariesBOMStructure Then
        Stop
    ElseIf aiBomType = kDefaultBOMStructure Then
        Stop
    Else
        Stop
    End If
    
    ''
    Set dcGnsPrpPtDvl_2021_1112 = rt
End Function

Public Function dcGeniusPropsPartRev20180530_broken2( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    'Dim dcPr As Scripting.Dictionary
    Dim dcVlGn  As Scripting.Dictionary
    Dim dcProp  As Scripting.Dictionary
    Dim dcVlPr  As Scripting.Dictionary
    Dim dcVlAi  As Scripting.Dictionary
    Dim dcVlFP  As Scripting.Dictionary
    Dim pr      As Inventor.Property
    Dim ky      As Variant
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
    ''' UPDATE[2021.11.08] MAJOR CHANGE
    '''     Overhaul variable names to better
    '''     reflect TWO distinct value sets
    '''         one from Genius
    '''         another from Inventor
    '''     in order to better compare
    '''     and synchronize them.
    '''
    ''' First set are the Genius variables:
    ''' the original set, renamed en masse:
    '''
    Dim gnPartNum   As String 'was pnModel
    Dim gnPartFam   As String 'was ptFamily
    Dim gnMatlNum   As String 'was pnStock
    Dim gnMatlFam   As String 'was mtFamily
    Dim gnMatlQty   As Double 'was qtRawMatl
    Dim gnQtyUnit   As String 'was qtUnit
    Dim gnBomType   As Inventor.BOMStructureEnum
    '''
    ''' Second set are the new Inventor variables.
    ''' These should replace the Genius instances
    ''' anywhere their values are taken
    ''' from the model.
    '''
    Dim aiPartNum   As String 'will be same as gnPartNum
    Dim aiPartFam   As String
    Dim aiMatlNum   As String
    Dim aiMatlFam   As String
    Dim aiMatlQty   As Double
    Dim aiQtyUnit   As String
    Dim aiBomType   As Inventor.BOMStructureEnum
    '''
    '''
    ''
    Dim ck          As VbMsgBoxResult
    Dim bd          As aiBoxData
    ''' UPDATE[2021.11.03]:
    '''
    '''
    Dim rs      As ADODB.Recordset
    Dim fdItem  As ADODB.Field
    Dim fdFamly As ADODB.Field
    Dim fdOrder As ADODB.Field
    Dim fdMatrl As ADODB.Field
    Dim fdMtFam As ADODB.Field
    Dim fdQty   As ADODB.Field
    Dim fdUnit  As ADODB.Field
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartRev20180530_broken2 = _
        dcGeniusPropsPartRev20180530_broken2( _
            invDoc, New Scripting.Dictionary _
        )
    Else
        aiBomType = invDoc.ComponentDefinition.BOMStructure
        '''UPDATE[2021.11.11]
        '''     Moved Property Set collection
        '''     to top of program to permit
        '''     collection of Design Properties
        '''     in second step. Also pulled up
        '''     BOM Structure capture (above)
        '''     along with the Values of
        '''     of Design Properties.
            
        With invDoc
            With .PropertySets
                Set aiPropsDesign = .Item(gnDesign)
                Set aiPropsUser = .Item(gnCustom)
            End With
            
            aiBomType = .ComponentDefinition.BOMStructure
            
            If aiBomType = kNormalBOMStructure Then
                If .SubType = guidSheetMetal Then
                End If
            End If
        End With
        
        ' Part Number and Family properties
        ' are from Design, NOT Custom set
        With aiPropsDesign 'we know they're present
            'so we can grab them directly
            Set prPartNum = .Item(pnPartNum)
            Set prFamily = .Item(pnFamily)
        End With
        aiPartNum = prPartNum.Value
        aiPartFam = prFamily.Value
        
        Set dcProp = dcGnsPropsPart(invDoc, , 0) 'dcAiPropsInSet
        Set dcVlPr = New Scripting.Dictionary
        With dcProp
            .Add pnPartNum, prPartNum
            .Add pnFamily, prFamily
            
            For Each ky In .Keys
                Set pr = aiProperty(.Item(ky))
                If pr Is Nothing Then
                    Stop
                Else
                    dcVlPr.Add ky, pr.Value
                End If
            Next
            
            If .Exists(pnRawMaterial) Then
                Set prRawMatl = .Item(pnRawMaterial)
                aiMatlNum = prRawMatl.Value
            Else
                aiMatlNum = ""
            End If
            
            If .Exists(pnRmUnit) Then
                Set prRmUnit = .Item(pnRmUnit)
                aiQtyUnit = prRmUnit.Value
            Else
                aiQtyUnit = ""
            End If
            
            If .Exists(pnRmUnit) Then
                Set prRmQty = .Item(pnRmQty)
                aiMatlQty = prRmQty.Value
            Else
                aiMatlQty = 0
            End If
        End With
        Debug.Print "=== Check Existing Model Genius Properties ==="
        Debug.Print dumpLsKeyVal(dcVlPr, "=")
        Debug.Print
        Stop
        
        ''' NOTE[2021.11.11]
        '''     Assignment of initial rt Dictionary
        '''     now essentially duplicates the new
        '''     process now preceding this section.
        '''     The only difference is, that version
        '''     does NOT apply Genius Property col-
        '''     lection to the supplied Dictionary dc.
        Set rt = dcGnsPropsPart(invDoc, dc, 0) 'dcAiPropsInSet
        Set dcVlAi = New Scripting.Dictionary
        With rt
            .Add pnPartNum, prPartNum
            .Add pnFamily, prFamily
            
            For Each ky In .Keys
                Set pr = aiProperty(.Item(ky))
                If pr Is Nothing Then
                    Stop
                Else
                    dcVlAi.Add ky, pr.Value
                End If
            Next
            Set pr = Nothing
        End With
        '''     Ultimately, processes which populate
        '''     returned Dictionary rt, and set the
        '''     Properties it should receive, should
        '''     be moved toward the end of the function.
        
        With cnGnsDoyle()
            'Pre-clear all relevant variables
            'to be set from query results,
            'if available.
            
            'gnPartNum = aiPartFam
            'gnPartFam = ""
            'Set fdOrder = .Item("Ord")
            'gnMatlNum = ""
            'gnMatlFam = ""
            'gnMatlQty = 0
            'gnQtyUnit = ""
            'gnBomType = kDefaultBOMStructure
                'use this to indicate no BOM type
                'or structure returned from Genius
            
            On Error Resume Next
            Err.Clear
            Set rs = .Execute( _
                sqlOf_ASDF(aiPartNum) _
            ) '
            If Err.Number = 0 Then
                Set dcVlGn = dcFromAdoRSrow(rs, "")
                With dcVlGn
                    gnPartNum = .Item("Item")
                    gnPartFam = .Item("Family")
                    gnBomType = .Item("bomStr")
                    'Set fdOrder = .Item("Ord")
                    gnMatlNum = .Item("Material")
                    gnMatlFam = .Item("MtFamily")
                    gnMatlQty = .Item("Qty")
                    gnQtyUnit = .Item("Unit")
                End With
                
                With rs
                    If .BOF And .EOF Then
                    Else
                        With .Fields
                            'Set fdItem = .Item("Item")
                            'gnPartNum = .Item("Item")
                            'should ALWAYS match aiPartNum
                            'IF it's found in Genius
                            'otherwise, always BLANK
                            
                            'Set fdFamly = .Item("Family")
                            'gnPartFam = .Item("Family").Value
                            
                            'gnBomType = .Item("bomStr").Value
                            
                            'Set fdOrder = .Item("Ord")
                            
                            'Set fdMatrl = .Item("Material")
                            'gnMatlNum = .Item("Material").Value
                            
                            'Set fdMtFam = .Item("MtFamily")
                            'gnMatlFam = .Item("MtFamily").Value
                            
                            'Set fdQty = .Item("Qty")
                            'gnMatlQty = .Item("Qty").Value
                            
                            'Set fdUnit = .Item("Unit")
                            'gnQtyUnit = .Item("Unit").Value
                            
                            'Stop 'to check things out
                            'not doing anything else with this yet
                            'but want to start matching against model
                        End With
                        
                        .MoveNext
                        If Not .EOF Then
                            Stop 'to handle multiple raw materials
                            Debug.Print ; 'Breakpoint Landing
                        End If
                    End If
                    
                    .Close
                End With
            Else
                Debug.Print Err.Number
                Debug.Print Err.Description
                Stop
            End If
            On Error GoTo 0
            .Close
        End With
        
        Debug.Print "== Prop Check =="
        Debug.Print "---- Genius ----"
        Debug.Print dumpLsKeyVal(dcVlGn, "=")
        Debug.Print "--- Inventor ---"
        Debug.Print dumpLsKeyVal(dcVlPr, "=") 'dcVlAi
        Debug.Print "================"
        Stop
        
        With invDoc
            '''UPDATE[2021.11.11]
            '''     Moved Property Set collection
            '''     to top of program, along with
            '''     collection of Design Properties
            '''     and their values. BOM Structure
            '''     as well.
            
            ''' We should check HERE for possibly misidentified purchased parts
            ''' UPDATE[2021.11.08]
            '''     Another MAJOR overhaul, here:
            '''     change Purchased Parts identification
            '''     to defer to Genius. Only attempt to guess
            '''     when no value comes back from Genius.
            'Stop 'BKPT-2021-1108-1608
            ''' CHANGE NEEDED[2021.11.08]:
            '''     indeterminate -- stopping work @endOfDay
            '''     effort here is to separate collection
            '''     and potential reassignment of
            '''     based on Part's family, file location,
            '''     and whatever other criteria, if any.
            '''
            '''     Likely need a counterpart variable
            '''     which takes its value from the Model.
            '''     The most likely Genius equivalent is
            '''     probably the ItemType field in view
            '''     table vgMfiItems, which will need
            '''     translation.
            '''
            If gnBomType = kDefaultBOMStructure Then
                'Genius didn't return an Item type
                'or BOM structure. We need to get it here.
                
                ''' BKPT-2021-1109-1042
                '''     Checkpoint here. Verify desired
                '''     behavior here prior to removal.
                Stop
                
                    ''' Get BOM Structure type, correcting if appropriate,
                    ''' and prepare Family value for part, if purchased.
                    '''
                    ''' UPDATE[2018.02.06]
                    '''     Using new UserForm; see below
                    ''' UPDATE[2018.05.31]
                    '''     Combined both InStr checks by addition
                    '''     to generate a single test for > 0
                    '''     If EITHER string match succeeds, the total
                    '''     SHOULD exceed zero, so this SHOULD work.
                    ''' UPDATE[2021.11.08]
                    '''     Removed extraneoous code previously
                    '''     disabled under preceding update[2018.05.31]
                    '''     Also reseparated InStr checks previously combined
                If aiBomType = kPurchasedBOMStructure Then 'it's purchased.
                    'Just assume that's what it's supposed to be.
                    gnBomType = aiBomType
                ElseIf InStr(1, _
                    "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                    "|" & aiPartFam & "|" _
                ) > 0 Then
                    'it needs to be SET purchased.
                    gnBomType = kPurchasedBOMStructure
                Else
                    'might need to ask User
                    If InStr(1, invDoc.FullFileName, _
                        "\Doyle_Vault\Designs\purchased\" _
                    ) > 0 Then 'it'a LIKELY purchased.
                        'Double check with User.
                        ck = newFmTest2().AskAbout(invDoc, , _
                            "Is this a Purchased Part?" _
                        )
                    Else
                        ck = vbNo
                    End If
                    
                    'Stop 'BKPT-2021-1105-0942
                    ''' CHANGE NEEDED[2021.11.05]:
                    '''     ONLY COLLECT desired BOMStructure here
                    '''     while keeping track of current value.
                    '''     Reassignment should take place along
                    '''     with collective property changes
                    ''' UPDATE[2021.11.09]
                    '''     This section now reduced to setting
                    '''     gnBomType from User response, if any.
                    '''     Code to assign Model BOM structure
                    '''     moved toward bottom for further work.
                    '''
                    ''' Check process below replaces duplicate check/responses above.
                    If ck = vbYes Then 'User said it IS purchased
                        gnBomType = kPurchasedBOMStructure
                    Else
                        gnBomType = aiBomType
                        dcVlGn.Item("bomStr") = gnBomType
                    End If
                    
                    'Request #2: Change Cost Center iProperty.
                    'If BOMStructure = Purchased and not content center,
                    'then Family = D-PTS, else Family = D-HDWR.
                    '
                    ''' UPDATE[2018.05.30]: Value produced here
                    '''     will now be held for later processing,
                    '''     more toward the end of this function.
                    ''' UPDATE[2021.11.09]
                    '''     Changed to set target (Genius)
                    '''     Family, and ONLY if not set already.
                    '''     '
                    '''     MIGHT want to set up a more robust check
                    '''     system, but see how this holds up, first.
                    '''     '
                    If Len(gnPartFam) = 0 Then
                        If gnBomType = kPurchasedBOMStructure Then
                            If .IsContentMember Then
                                Stop 'BKPT-2021-1105-0946
                                gnPartFam = "D-HDWR"
                            Else
                                Stop 'BKPT-2021-1105-0947
                                gnPartFam = "D-PTS"
                                'NOTE: NON Content Center members
                                '       might still be D-HDWR
                                '       Additional checks might
                                '       be recommended
                            End If
                        Else
                        End If
                    End If
                End If
            'Else 'disabled unless/until needed, or removed
                'aiPartFam = gnPartFam 'no, DON'T!
                ''' UPDATE[2021.11.09]: Disabled this
                '''     Keep the Model Part Family value AS IS
                '''     so it may be used to check for equality
                '''     and the need to update at the end.
                '''     '
            End If
            
            With .ComponentDefinition
                ''' Request #1: Get the Mass in Pounds
                ''' and add to Custom Property GeniusMass
                With .MassProperties
                    'Stop 'BKPT-2021-1110-1551
                    ''' CHANGE NEEDED[2021.11.10]
                    '''     '
                    '''     '
                    '''     '
                    If dcVlPr.Exists(pnMass) Then
                        If Round(cvMassKg2LbM * .Mass, 4) - CDbl(dcVlPr.Item(pnMass)) = 0 Then
                        Else
                            'Stop
                            dcVlPr.Item(pnMass) = Round(cvMassKg2LbM * .Mass, 4)
                        End If
                    Else
                        dcVlPr.Add pnMass, Round( _
                            cvMassKg2LbM * .Mass, 4 _
                        )
                    End If
                    ''' UPDATE[2021.11.09]
                    '''     Part Mass Value now assigned
                    '''     to new Values Dictionary, instead
                    '''     of directly to Genius Mass Property.
                    '''     That assignment now moved toward the
                    '''     end of this function, where it can
                    '''     be set alongside other Properties
                    '''     in one straight process.
                    '''     '
                    '''     Note that value in Genius is not
                    '''     yet collected for comparison. This
                    '''     will likely require modification
                    '''     of an SQL query, and a new variable
                    '''     or Dictionary to keep track of it.
                End With
            End With
            ''' At this point, gnPartFam SHOULD be set
            ''' to a non-blank value if Item is purchased.
            ''' We should be able to check this later on,
            ''' if Item BOMStructure is NOT Normal
            
            'Stop 'BKPT-2021-1109-1053
            ''' HERE is where it starts to get interesting
            ''' Actually, just a little further down, where
            ''' Part SubType is checked for Sheet Metal.
            ''' At that point, the function divides into two
            ''' LONG, and possibly nearly identical branches.
            ''' Ideally, these should be refactored, with as
            ''' much of their processes as possible combined
            ''' into a single path.
            
            'Request #4: Change Cost Center iProperty.
            'If BOMStructure = Normal, then Family = D-MTO,
            'else if BOMStructure = Purchased then Family = D-PTS.
            If aiBomType = kNormalBOMStructure Then
                ' Get Custom Properties
                'Stop 'BKPT-2021-1105-1144
                ''' CHANGE NEEDED[2021.11.05]:
                '''     these properties should NOT
                '''     be added immediately, but only
                '''     when it's time to set them,
                '''     towards the END of this function.
                ''' UPDATE[2021.11.09]
                '''     Custom Property collection/generation
                '''     moved into Normal BOM Part handling, as
                '''     no earlier usage appears to take place.
                '''     '
                '''     If possible, may wish to move even further.
                '''     Plan to review later, as time permits.
                ''' UPDATE[2021.11.10]
                '''     Disabled Genius Property collection here
                '''     since a Dictionary of ALL Genius Properties
                '''     is generated towards the beginning.
                '''     '
                '''     '
                'With rt
                '    If .Exists(pnRawMaterial) Then Set prRawMatl = .Item(pnRawMaterial)
                '    Set prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                '    If .Exists(pnRmUnit) Then Set prRmUnit = .Item(pnRmUnit)
                '    Set prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                '    If .Exists(pnRmQty) Then Set prRmQty = .Item(pnRmQty)
                '    Set prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                'End With
                '''     Collecting them at this point
                '''     might still be appropriate,
                '''     or it may be more desirable
                '''     to hold off until later.
                '''     '
                ''' UPDATE[2021.11.08]
                '''     Design Properties have been moved
                '''     toward the top, as proposed.
                '''     Commentary recommending this
                '''     has been removed as extraneous.
                '''     '
                ''' UPDATE[2021.11.09]
                '''     BOM Structure collection has also been
                '''     moved up, alongside Design Properties,
                '''     using renamed variable aiBomType
                '''     (formerly bomStruct)
                '''     '
                
                '----------------------------------------------------'
                If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                '----------------------------------------------------'
                    'Request #3:
                    '   Get sheet metal extent area
                    '   and add to custom property "RMQTY"
                    ''' UPDATE[2021.11.10]
                    '''     Now collecting Flat Pattern Values
                    '''     instead of the Properties for them.
                    '''     If necessary, Properties should be
                    '''     assigned in a separate function
                    ''' CHANGE NEEDED[2021.11.05]:
                    '''     not quite sure on this one yet,
                    '''     but dcFlatPatProps might need its
                    '''     own set of revisions to generate
                    '''     assignment recommendations WITHOUT
                    '''     performing them itself
                    ''' UPDATE[2021.11.10]
                    '''     Embedded Flat Pattern Property collection
                    '''     in bypassed If branch. Preceding Stop,
                    '''     when enabled, offers user/developer
                    '''     an opportunity to run it, if desired.
                    Stop 'BKPT-2021-1105-1105
                    If True Then
                        Set dcVlFP = dcFlatPatVals( _
                            .ComponentDefinition _
                        ) 'dcVlAi
                        With dcVlFP
                            aiMatlFam = .Item("mtFamily")
                            .Remove "mtFamily"
                            For Each ky In .Keys
                                If dcVlPr.Exists(ky) Then
                                    If CStr(dcVlPr.Item(ky)) = CStr(.Item(ky)) Then
                                    Else
                                        Stop
                                        ''' need to add value to a new 'change' Dictionary
                                    End If
                                Else
                                    Stop
                                    ''' need to add value to a new 'change' Dictionary
                                End If
                            Next
                        End With
                        Stop
                    Else
                        Set rt = dcFlatPatProps(.ComponentDefinition, rt)
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
                            ''' UPDATE[2018.05.30]:
                            '''     Restoring original key check
                            '''     and adding code for debug
                            '''     Previously changed to "~OFFTHK"
                            '''     to avoid this block and its issues.
                            '''     (Might re-revert if not prepped to fix now)
                            Debug.Print aiProperty(rt.Item("OFFTHK")).Value
                            Stop 'because we're going to need to do something with this.
                            
                            gnMatlNum = "" 'Originally the ONLY line in this block.
                            ' A more substantial response is required here.
                            
                            If 0 Then Stop '(just a skipover)
                        Else
                            If Len(gnMatlNum) = 0 Then
                                Stop 'because we don't know IF this is sheet metal yet
                                gnMatlNum = ptNumShtMetal(.ComponentDefinition)
                            End If
                        End If
                    Else
                        ''  ACTION ADVISED[2018.09.14]:
                        ''  gnMatlNum can probably be set
                        ''  to prRawMatl.Value and THEN
                        ''  checked for length to see
                        ''  if lookup needed.
                        ''  This might also allow us to check
                        ''  for machined or other non-sheet
                        ''  metal parts.
                        
                        Stop
                        ''' !!!WARNING!!![2021.11.04]:
                        ''' Following section has been shuffled
                        ''' and should be considered HIGHLY
                        ''' UNSTABLE until verified functional
                        ''' and SAFE! TWO Stop commands are
                        ''' placed to emphasize the need for
                        ''' EXTREME CAUTION at this point
                        Stop
                        ''' UPDATE[2021.11.04]:
                        '''     This section is being adjusted
                        '''     in an attempt to improve the raw
                        '''     material determination process.
                        '''
                        '''     This particular segment should
                        '''     ONLY be invoked if gnMatlNum is not
                        '''     successfully retrieved from Genius
                        '''
                        If Len(gnMatlNum) = 0 Then
                            'no stock retrieved from Genius
                            'attempt to retrieve from Model
                            'gnMatlNum = aiMatlNum
                            
                            If Len(aiMatlNum) > 0 Then 'gnMatlNum
                            'need to verify it against Genius
                            'by retrieving its Family there
                            '''
                            ''' This With block copied and modified [2021.03.11]
                            ''' from elsewhere in this function as a temporary measure
                            ''' to address a stopping situation later in the function.
                            ''' See comment below for details.
                            '''
                            ''' UPDATE[2021.11.04]:
                            '''     This section MIGHT be removed in future,
                            '''
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & gnMatlNum & "';" _
                            )
                                If .BOF Or .EOF Then
                                    'Stop 'because Material value likely invalid
                                    Stop 'because we do NOT want to set gnMatlNum!
                                    ''' want to assign it to a separate RETURN variable
                                    ''' or most likely, the return Dictionary.
                                    gnMatlNum = ptNumShtMetal(invDoc.ComponentDefinition)
                                    Debug.Print ; 'Breakpoint Landing
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
                                Else
                                    ''  This section retained from source,
                                    ''  but disabled to avoid potential issues
                                    ''  with subsequent operations, just in case
                                    ''  anything depends on gnMatlFam remaining
                                    ''  uninitialized up to that point.
                                    ''' UPDATE[2021.11.09]
                                    '''     Re-enabling Genius Material
                                    '''     Family assignment, as it SHOULD
                                    '''     be set to match what Genius
                                    '''     returns from this query.
                                    '''     '
                                    '''     Might not be the best place
                                    '''     to do this, though. If ptNumShtMetal
                                    '''     returns a valid Material Item above,
                                    '''     a Family is still needed.
                                    '''     '
                                    '''     NOTE: Fix disabled With block between runs
                                    ''  With .Fields
                                          Stop 'because we do not want to set gnMatlFam
                                          '''   for same reasons as above
                                          gnMatlFam = .Fields.Item("Family").Value
                                    ''  End With
                                End If
                            End With
                            End If
                        End If
                        
                        If Len(gnMatlNum) = 0 Then
                            ''' UPDATE[2018.05.30]:
                            '''     Pulling ALL code/text from this section
                            '''     to get rid of excessive cruft.
                            '''
                            '''     In fact, reversing logic to go directly
                            '''     to User Prompt if no stock identified
                            '''
                            '''     IN DOUBLE FACT, hauling this WHOLE MESS
                            '''     RIGHT UP after initial gnMatlNum assignment
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
                                            gnPartFam = .Item(pnFamily)
                                            Debug.Print pnFamily & "=" & gnPartFam
                                        End If
                                        
                                        If .Exists(pnRawMaterial) Then
                                            gnMatlNum = .Item(pnRawMaterial)
                                            Debug.Print pnRawMaterial & "=" & gnMatlNum
                                        End If
                                    End With
                                    If 0 Then Stop 'Use this for a debugging shim
                                End If
                            End With
                        ElseIf Left$(gnMatlNum, 2) = "LG" Then 'it's probably lagging
                            Debug.Print aiPartNum & ": PROBABLE LAGGING"
                            Debug.Print "  TRY TO IDENTIFY, AND FILL IN BELOW."
                            Debug.Print "  PRESS ENTER ON gnMatlNum LINE WHEN"
                            Debug.Print "  COMPLETED, THEN F5 TO CONTINUE."
                            Debug.Print "  gnMatlNum = """ & gnMatlNum & """"
                            Stop
                        End If
                        
                        If Len(gnMatlNum) > 0 Then 'and ONLY then
                        'do we look for a Raw Material Family!
                            
                            ''' NOTE[2021.11.10]
                            '''     This query is probably WAY more than needed here.
                            '''     Spec fields are probably not needed at all,
                            '''     and it's not clear which of the others might be.
                            '''     '
                            '''     It might also be possible to REMOVE this query
                            '''     based on the earlier one, which would return
                            '''     Material Family along with Part Family,
                            '''     providing the Part/Item were found in Genius.
                            '''     '
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & gnMatlNum & "';" _
                            ) '", Description1, Unit, " & _
                                "Specification1, Specification2, Specification3, " & _
                                "Specification4, Specification5, Specification6, " & _
                                "Specification7, Specification8, Specification9, " & _
                                "Specification15, Specification16 " & _
                            '''
                            ''' UPDATE[2021.11.10]
                            '''     Removed (likely) unneeded fields from query text.
                            '''     Will keep a lookout for any resulting errors.
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]:
                                    ''  Will need to address this situation
                                    ''  in a more robust manner.
                                    ''  A more thorough query above
                                    ''  might also be called for.
                                Else
                                    With .Fields
                                        If Len(gnMatlFam) > 0 Then
                                            If gnMatlFam = .Item("Family").Value Then
                                            Else
                                                Stop
                                                'as with gnMatlNum, want to be careful
                                                'about changing gnMatlFam, although
                                                'in this case, it SHOULD be okay
                                                'since it should just align with
                                                'the selected material's Family
                                            End If
                                        Else
                                            gnMatlFam = .Item("Family").Value
                                        End If
                                    End With
                            ''' NOTE[2021.11.10]
                            '''     Else branch should PROBABLY end here
                            '''     to permit Recordset to be closed,
                            '''     and probably a new If/Then block
                            '''     proceed based on results.
                            '''     '
                                    
                                    ''' UPDATE[2021.06.18]:
                                    '''     New pre-check for Material Item
                                    '''     in Purchased Parts Family.
                                    '''     VERY basic handler simply
                                    '''     maps Material Family to D-BAR
                                    '''     to force extra processing below.
                                    '''     Further refinement VERY much needed!
                                    If gnMatlFam Like "?-MT*" Then
                                        'Debug.Print aiPartNum & " [" & aiMatlNum & "]: " & aiPropsDesign(pnDesc).Value
                                        Debug.Print aiPartNum & "[" & prRmQty.Value & gnQtyUnit & "*" & gnMatlNum & ": " & aiPropsDesign(pnDesc).Value & "]" ' aiMatlNum
                                        Stop 'FULL Stop!
                                    ElseIf gnMatlFam = "D-PTS" Then
                                        gnPartFam = "D-RMT"
                                        Stop 'NOT SO FAST!
                                        gnMatlFam = "D-BAR"
                                    ElseIf gnMatlFam = "R-PTS" Then
                                        gnPartFam = "R-RMT"
                                        Stop 'NOT SO FAST!
                                        gnMatlFam = "D-BAR"
                                    End If
                                    
                                    If gnMatlFam = "DSHEET" Then
                                        'We should be okay. This is sheet metal stock
                                        
                                        ''' UPDATE[2021.11.04]:
                                        '''     Expanding gnPartFam and gnQtyUnit
                                        '''     assignments to check for pre-
                                        '''     existing values, and validate
                                        '''     them if found.
                                        If Len(gnPartFam) = 0 Then
                                            gnPartFam = "D-RMT"
                                        Else
                                            If gnPartFam = "D-RMT" Then
                                            Else
                                                Stop 'because we have
                                                'an unexpected situation
                                            End If
                                        End If
                                        
                                        If Len(gnQtyUnit) = 0 Then
                                            gnQtyUnit = "FT2"
                                        Else
                                            If gnQtyUnit = "FT2" Then
                                            Else
                                                Stop 'because we have
                                                'an unexpected situation
                                            End If
                                        End If
                                        
                                        ''' UPDATE[2018.05.30]:
                                        '''     Moving part family assignment
                                        '''     to this section for better mapping
                                        '''     and updating to new Family names
                                        '''     as well as pulling up gnQtyUnit assignment
                                    Stop 'BKPT-2021-1105-1120
                                    ''' CHANGE NEEDED[2021.11.05]:
                                    '''     probably want to demote this ElseIf
                                    '''     along with the subsequent Else into
                                    '''     a replacement Else clause, and thus
                                    '''     allow for on post-interactive check
                                    '''     following whichever branch is taken.
                                    '''     '
                                    '''     As is, a separate check is required
                                    '''     within but this ElseIf and the Else
                                    '''     '
                                    ''' UPDAGE[SAME_DAY]:
                                    '''     Change completed. All that remains
                                    '''     is to indent the embedded With block
                                    '''     when safe to do so.
                                    Else
                                        If gnMatlFam = "D-BAR" Then
                                            ''' UPDATE[2021.06.18]:
                                            '''     Added check for Part Family already set
                                            '''     to more properly handle new situation (above)
                                            If Len(gnPartFam) = 0 Then
                                                gnPartFam = "R-RMT"
                                                'mignt not want to use
                                                'fixed constant like this
                                                'see gnQtyUnit below
                                            Else
                                                If gnPartFam = "R-RMT" Then
                                                Else
                                                    Stop
                                                End If
                                                Debug.Print ; 'Breakpoint Landing
                                                'Stop
                                            End If
                                            
                                            If Len(gnQtyUnit) = 0 Then
                                                gnQtyUnit = "IN" 'prRmUnit.Value '
                                            Else
                                                If gnQtyUnit = "IN" Then 'prRmUnit.Value '
                                                Else
                                                    Stop
                                                End If
                                                Debug.Print ; 'Breakpoint Landing
                                                'Stop
                                            End If
                                            ''may want function here
                                            ''' UPDATE[2018.05.30]: As noted above
                                            '''     Will keep Stop for now
                                            '''     pending further review,
                                            '''     hopefully soon
                                            Debug.Print aiPartNum & " [" & gnMatlNum & "]: " & aiPropsDesign(pnDesc).Value 'aiMatlNum
                                            ''' UPDATE[2021.03.11]: Replaced
                                            ''' aiPropsDesign.Item(pnPartNum)
                                            ''' with prPartNum (and now aiPartNum)
                                            ''' since it's used in several places
                                            Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); gnQtyUnit; ". IF CHANGE NEEDED,"
                                            Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                            Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
                                            Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                                        'Stop 'BKPT-2021-1105-1137
                                        ''' CHANGE NEEDED[2021.11.05]:
                                        '''     indent the following With,
                                        '''     when possible to do so
                                        '''     without resetting project
                                        With invDoc.ComponentDefinition.RangeBox
                                            Debug.Print _
                                            (.MaxPoint.X - .MinPoint.X) / 2.54, _
                                            (.MaxPoint.Y - .MinPoint.Y) / 2.54, _
                                            (.MaxPoint.Z - .MinPoint.Z) / 2.54
                                        End With
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
                                            Debug.Print ""
                                            Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                                            'Debug.Print "gnQtyUnit = """; gnQtyUnit; """"
                                            Debug.Print "gnQtyUnit = ""IN"""
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
                                        Else
                                            Debug.Print "NON-STANDARD MATERIAL FAMILY (" & gnMatlFam & ")"
                                            Debug.Print "PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW"
                                            Debug.Print "PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED"
                                            Debug.Print "PRESS [F5] WHEN READY TO CONTINUE"
                                            Debug.Print ""
                                            Debug.Print "gnPartFam = """ & gnPartFam & """ 'PART FAMILY"
                                            Debug.Print "gnQtyUnit = """ & gnQtyUnit & """ 'UNIT OF MEASURE"
                                            Stop 'because we don't know WHAT to do with it
                                                 'but might NOT want to clear variables
                                            'gnPartFam = ""
                                            'gnQtyUnit = "" 'may want function here
                                            ''' UPDATE[2018.05.30]: As noted above
                                            '''     However, might need more handling here.
                                        End If
                                        
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); gnQtyUnit; ". IF OKAY, CONTINUE."
                                        Stop
                                        
                                        Stop 'BKPT-2021-1105-1117
                                        ''' CHANGE NEEDED[2021.11.05]:
                                        '''     Property assignment needs moved
                                        '''     to collective assignment sequence
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    End If
                                End If
                            End With
                        Else
                            If 0 Then Stop 'and regroup
                            ''' Things are looking a right royal mess
                            ''' at the moment I'm writing this comment.
                        End If
                    End If
                    
                    ''' UPDATE[2021.11.10]
                    '''     Transported prRawMatl and prRmUnit assignments
                    '''     to point following this If/Else block
                    '''     to consolidate duplicate processes.
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
                        If invDoc.ComponentDefinition.Document Is invDoc Then
                        'following needs indented if not already
                        
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
                        Stop 'BKPT-2021-1105-0955
                        ''' CHANGE NEEDED[2021.11.05]:
                        '''     Probably want to move this
                        '''     outside of this With block,
                        '''     and closer to the beginning
                        '''     of this function, as it could
                        '''     prove helpful at other points.
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
                                    If Len(gnPartFam) = 0 Then
                                        gnPartFam = .Item(pnFamily)
                                    Else
                                        If gnPartFam = .Item(pnFamily) Then
                                        Else
                                            Debug.Print "====="
                                            Debug.Print "Model Family differs from Genius"
                                            Debug.Print "Genius: " & pnFamily
                                            Debug.Print "Model:  " & .Item(pnFamily)
                                            Debug.Print "gnPartFam = .Item(pnFamily) 'Press [ENTER] on this line to fix, and/or [F5] to continue'"
                                            Stop
                                        End If
                                    End If
                                    Debug.Print pnFamily & "=" & gnPartFam
                                End If
                                
                                If .Exists(pnRawMaterial) Then
                                    If Len(gnMatlNum) = 0 Then
                                        gnMatlNum = .Item(pnRawMaterial)
                                    Else
                                        If gnMatlNum = .Item(pnRawMaterial) Then
                                        Else
                                            Debug.Print "====="
                                            Debug.Print "Model Raw Material differs from Genius"
                                            Debug.Print "Genius: " & gnMatlNum
                                            Debug.Print "Model:  " & .Item(pnRawMaterial)
                                            Debug.Print "gnMatlNum = .Item(pnRawMaterial) 'Press [ENTER] on this line to fix, and/or [F5] to continue'"
                                            Stop
                                        End If
                                    End If
                                    Debug.Print pnRawMaterial & "=" & gnMatlNum
                                End If
                            End With
                            If 0 Then Stop 'Use this for a debugging shim
                            ''  We're going to need something here
                            ''  to make sure raw material gets added
                            ''  for non sheet metal parts, as well
                            ''  What we're going to need to do
                            ''  is refactor this whole bloody thing.
                        Else
                            Stop 'shouldn't actually hit this line
                            'as the condition checked should always
                            'be true, at least for now.
                        End If
                        Else
                            Stop 'because we've got a serious mismatch
                        End If
                    End With
                    '''
                    '''
                    '''
                    
                    If Len(gnMatlNum) > 0 Then 'and ONLY then
                    'do we look for a Raw Material Family!
                        
                        ''' This enclosing With block should NOT be necessary
                        ''' since the newFmTest1 above takes care of collecting
                        ''' the Stock Family along with the Stock itself
                        With cnGnsDoyle().Execute( _
                            "select Family " & _
                            "from vgMfiItems " & _
                            "where Item='" & gnMatlNum & "';" _
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
                                    If Len(gnMatlFam) = 0 Then
                                        gnMatlFam = .Item("Family").Value
                                    ElseIf gnMatlFam = .Item("Family").Value Then
                                    Else
                                        Stop
                                    End If
                                End With
                            End If
                        End With
                        
                        If gnMatlFam = "DSHEET" Then
                            Stop 'because we should NOT be doing Sheet Metal in this section.
                                 ' This might require further investigation and/or development, if encountered.
                            'We should be okay. This is sheet metal stock
                            ''' UPDATE[2021.11.04]:
                            '''     Expanding gnPartFam and gnQtyUnit
                            '''     assignments to check for pre-
                            '''     existing values, and validate
                            '''     them if found.
                            If Len(gnPartFam) = 0 Then
                                gnPartFam = "D-RMT"
                            Else
                                If gnPartFam = "D-RMT" Then
                                Else
                                    Stop 'because we have
                                    'an unexpected situation
                                End If
                            End If
                            
                            If Len(gnQtyUnit) = 0 Then
                                gnQtyUnit = "FT2"
                            Else
                                If gnQtyUnit = "FT2" Then
                                Else
                                    Stop 'because we have
                                    'an unexpected situation
                                End If
                            End If
                            ''' UPDATE[2018.05.30]:
                            '''     Moving part family assignment
                            '''     to this section for better mapping
                            '''     and updating to new Family names
                            '''     as well as pulling up gnQtyUnit assignment
                        ElseIf gnMatlFam = "D-BAR" Then
                            gnPartFam = "R-RMT"
                            If Len(gnQtyUnit) = 0 Then
                                'this might have to change
                                'to better handle case
                                'of missing prRmUnit
                                gnQtyUnit = prRmUnit.Value '"IN"
                            End If
                            ''may want function here
                            ''' UPDATE[2018.05.30]: As noted above
                            '''     Will keep Stop for now
                            '''     pending further review,
                            '''     hopefully soon
                            Debug.Print aiPartNum & " [" & gnMatlNum & "]: " & CStr(aiPropsDesign(pnDesc).Value) 'prRawMatl.Value
                            ''' UPDATE[2021.03.11]: Replaced
                            ''' aiPropsDesign.Item(pnPartNum)
                            ''' as noted above
                            Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(gnMatlQty); gnQtyUnit; ". IF CHANGE NEEDED," 'prRmQty.Value
                            Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                            Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
                            Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
                            Debug.Print (invDoc.ComponentDefinition.RangeBox.MaxPoint.X - invDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (invDoc.ComponentDefinition.RangeBox.MaxPoint.Y - invDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (invDoc.ComponentDefinition.RangeBox.MaxPoint.Z - invDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54
                            Debug.Print ""
                            Debug.Print "PLACE CURSOR ON gnQtyUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED."
                            Debug.Print "PRESS ENTER/RETURN TWICE. THEN CONTINUE."
                            Debug.Print ""
                            Debug.Print "gnMatlQty = "; CStr(gnMatlQty) 'prRmQty.Value
                            Debug.Print "gnQtyUnit = ""IN"""
                            Debug.Print ""
                            Stop 'because we might want a D-BAR handler
                            ''' Actually, we might NOT need to stop here
                            ''' if bar stock is already selected,
                            ''' because quantities would presumably
                            ''' have been established already.
                            ''' Any D-BAR handler probably needs
                            ''' to be implemented in prior section(s)
                            Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(gnMatlQty); gnQtyUnit; ". IF OKAY, CONTINUE." 'prRmQty.Value
                            Stop
                            Stop 'BKPT-2021-1110-1647
                            ''' CHANGE NEEDED[2021.11.10]
                            '''     This Dictionary Property assignment
                            '''     MUST be moved to the END of the function!
                            Set rt = dcWithProp(aiPropsUser, pnRmQty, gnMatlQty, rt) 'dcAddProp(prRmQty, rt)
                            Debug.Print ; 'Landing line for debugging. Do not disable.
                        Else
                            Debug.Print "NON-STANDARD MATERIAL FAMILY (" & gnMatlFam & ")"
                            Debug.Print "PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW"
                            Debug.Print "PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED"
                            Debug.Print "PRESS [F5] WHEN READY TO CONTINUE"
                            Debug.Print ""
                            Debug.Print "gnPartFam = """ & gnPartFam & """ 'PART FAMILY"
                            Debug.Print "gnQtyUnit = """ & gnQtyUnit & """ 'UNIT OF MEASURE"
                            'gnPartFam = ""
                            'gnQtyUnit = "" 'may want function here
                            ''' UPDATE[2018.05.30]: As noted above
                            '''     However, might need more handling here.
                            Stop 'because we don't know WHAT to do with it
                        End If
                    Else
                        If 0 Then Stop 'and regroup
                        ''' Things are looking a right royal mess
                        ''' at the moment I'm writing this comment.
                    End If
                End If 'Sheetmetal vs Part
                    
                    'Stop 'BKPT-2021-1105-1011
                    ''' UPDATE[2021.11.10]
                    '''     Disabled this prRawMatl assignment pending removal.
                    '''     Counterpart moved below from sheet metal branch
                    '''     should serve in place of both branch instances.
                    '''     '
                    '''     Extraneous commentary removed.
                    '''     '
                    'With prRawMatl
                    '    If Len(Trim$(.Value)) > 0 Then
                    '        If gnMatlNum <> .Value Then
                    '            'Debug.Print "Raw Stock Selection"
                    '            'Debug.Print "  Current : " & prRawMatl.Value
                    '            'Debug.Print "  Proposed: " & gnMatlNum
                    '            'Stop 'because we might not want to change existing stock setting
                    '            'if
                    '            ck = MsgBox( _
                    '                Join(Array( _
                    '                    "Raw Stock Change Suggested", _
                    '                    "  for Item " & aiPartNum, _
                    '                    "", _
                    '                    "  Current : " & prRawMatl.Value, _
                    '                    "  Proposed: " & gnMatlNum, _
                    '                    "", "Change It?", "" _
                    '                ), vbNewLine), _
                    '                vbYesNo, aiPartNum & " Stock" _
                    '            )
                    '            '"Change Raw Material?"
                    '            '"Suggested Sheet Metal"
                    '            If ck = vbYes Then .Value = gnMatlNum
                    '        End If
                    '    Else
                    '        .Value = gnMatlNum
                    '    End If
                    'End With
                    'Set rt = dcAddProp(prRawMatl, rt)
                    
                    'Stop 'BKPT-2021-1110-1130
                    ''' UPDATE[2021.11.10]
                    '''     Disabled this prRmUnit assignment pending removal.
                    '''     Duplicate moved below from sheet metal branch
                    '''     should serve in place of both branch instances.
                    '''     '
                    '''     Also moved End If AHEAD of this block to minimize
                    '''     comment clutter WITHIN branches.
                    'With prRmUnit
                    '    If Len(.Value) > 0 Then
                    '        If Len(gnQtyUnit) > 0 Then
                    '            If .Value <> gnQtyUnit Then
                    '                Stop 'and check both so we DON'T
                    '                'automatically "fix" the RMUNIT value
                    '
                    '                .Value = gnQtyUnit
                    '
                    '                If 0 Then Stop 'Ctrl-9 here to skip changing
                    '            End If
                    '        End If
                    '    Else 'we're setting a new quantity unit
                    '        .Value = gnQtyUnit
                    '    End If
                    'End With
                    'Set rt = dcAddProp(prRmUnit, rt)
                
                Stop 'BKPT-2021-1109-1610
                ''' UPDATE[2021.11.10]
                '''     Transported this prRawMatl assignment
                '''     from sheet metal branch to consolidate
                '''     both instances of duplicated process
                '''     into one following both branches.
                '''     '
                '''     Extraneous commentary also removed.
                '''     '
                If prRawMatl Is Nothing Then
                    Set rt = dcWithProp(aiPropsUser, pnRawMaterial, gnMatlNum, rt)
                    Debug.Print ; 'Breakpoint Landing
                Else
                With prRawMatl
                    If Len(Trim$(.Value)) > 0 Then
                        If gnMatlNum <> .Value Then
                            'Debug.Print "Raw Stock Selection"
                            'Debug.Print "  Current : " & prRawMatl.Value
                            'Debug.Print "  Proposed: " & gnMatlNum
                            'Stop 'because we might not want to change existing stock setting
                            'if
                            ck = MsgBox( _
                                Join(Array( _
                                    "Raw Stock Change Suggested", _
                                    "  for Item " & aiPartNum, _
                                    "", _
                                    "  Current : " & prRawMatl.Value, _
                                    "  Proposed: " & gnMatlNum, _
                                    "", "Change It?", "" _
                                ), vbNewLine), _
                                vbYesNo, aiPartNum & " Stock" _
                            )
                            '"Change Raw Material?"
                            '"Suggested Sheet Metal"
                            If ck = vbYes Then .Value = gnMatlNum
                        End If
                    Else
                        .Value = gnMatlNum
                    End If
                End With
                Set rt = dcAddProp(prRawMatl, rt)
                End If
                
                'Stop 'BKPT-2021-1110-1133
                ''' UPDATE[2021.11.10]
                '''     Transported this prRmUnit assignment
                '''     from sheet metal branch to consolidate
                '''     both instances of duplicated process
                '''     into one following both sheet metal
                '''     and structural branches
                If prRmUnit Is Nothing Then
                    Set rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt)
                    Debug.Print ; 'Breakpoint Landing
                Else
                With prRmUnit
                    If Len(.Value) > 0 Then
                        If Len(gnQtyUnit) > 0 Then
                            If .Value <> gnQtyUnit Then
                                Stop 'and check both so we DON'T
                                'automatically "fix" the RMUNIT value
                                
                                .Value = gnQtyUnit
                                
                                If 0 Then Stop 'Ctrl-9 here to skip changing
                            End If
                        End If
                    Else 'we're setting a new quantity unit
                        .Value = gnQtyUnit
                    End If
                End With
                Set rt = dcAddProp(prRmUnit, rt)
                End If
                'Set rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt) 'gnQtyUnit WAS "FT2"
                ''' Plan to remove commented line above,
                ''' superceded by the one above that
                Debug.Print ; 'Breakpoint Landing
                    
                
                'Stop 'BKPT-2021-1110-1133
                ''' UPDATE[2021.11.09]
                '''     This is a VERY crude implementation
                '''     of the closing BOM Structure assignment.
                '''     Plan on revision and cleanup in future.
                If gnBomType = aiBomType Then
                    With .ComponentDefinition
                        If .BOMStructure <> gnBomType Then
                            On Error Resume Next
                            .BOMStructure = gnBomType
                            If Err.Number = 0 Then
                        '        aiBomType = .BOMStructure
                            Else
                                Stop
                                Debug.Print ; 'Breakpoint Landing
                        '        aiBomType = kPurchasedBOMStructure
                        '
                        '        ''' WARNING: NOT a good way to go about this
                        '        '''     but will go with it for now
                            End If
                            On Error GoTo 0
                            Stop
                            Debug.Print ; 'Breakpoint Landing
                        End If
                        Debug.Print ; 'Breakpoint Landing
                    End With
                Else
                    Stop
                    Debug.Print ; 'Breakpoint Landing
                End If
                'this With block was pulled down
                'from the BOMStructure section above.
                'It MIGHT want to be moved ahead
                'of the two preceding With blocks.
                'With .ComponentDefinition 'appears to be extraneous now.
                ''' Disabling it does not lead to compilation errors, indicating
                ''' nothing (active) within the block depends on it anymore.
                ''' It does look like this disabled If block
                ''' would still need it, though.
                    ''' This If block is meant to set BOMStructure
                    ''' according to information gathered from the
                    ''' Model, its Vault location, Genius, and if
                    ''' all else fails, the Users themselves.
                    '''
                    ''' Since it SETS a Model attribute, it belongs
                    ''' toward the bottom, with the blocks that set
                    ''' Model Properties. Plan to move it there.
                    '''
                    'If .BOMStructure <> kPurchasedBOMStructure Then
                    '    On Error Resume Next
                    '    .BOMStructure = kPurchasedBOMStructure
                    '    If Err.Number = 0 Then
                    '        aiBomType = .BOMStructure
                    '    Else
                    '        aiBomType = kPurchasedBOMStructure
                    '
                    '        ''' WARNING: NOT a good way to go about this
                    '        '''     but will go with it for now
                    '    End If
                    '    On Error GoTo 0
                    'Else
                    '    aiBomType = .BOMStructure 'to make sure this is captured
                    'End If
                'End With
            ElseIf aiBomType = kPurchasedBOMStructure Then
                ''' As mentioned above, gnPartFam
                ''' SHOULD be set at this point
                If Len(gnPartFam) = 0 Then
                    If 1 Then Stop 'because we might
                    'need to check out the situation
                    gnPartFam = "D-PTS" 'by default
                End If
            Else
                Stop 'because we might need
                    'to do something else
                    'based on an unexpected
                    'BOM Structure
            End If
            
            'Stop 'BKPT-2021-1105-1020
            ''' CHANGE NEEDED[2021.11.05]:
            '''     Family assignment should be
            '''     ported up into collective
            '''     Property assignment, although
            '''     its position here assures
            '''     one instance of the sequence
            '''     catches ALL divergent cases
            '''     leading up to this point.
            '''     '
            '''     Ultimately, those cases probably
            '''     need to be consolidated HERE
            '''     if, or WHEN possible.
            '''     '
            ' Get the design tracking property set,
            ' and update the Cost Center Property
            If invDoc.ComponentDefinition.IsContentMember Then
                ' Don't muck around with the Family!
            Else
                If Len(gnPartFam) > 0 Then
                    dcVlGn.Item("Family") = gnPartFam
                    If aiPartFam = gnPartFam Then
                    Else
                        On Error Resume Next
                        prFamily.Value = gnPartFam
                        If Err.Number Then
                            Debug.Print "CHGFAIL[FAMILY]{'" _
                                & prFamily.Value & "' -> '" & gnPartFam & "'}: " _
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
                    Debug.Print ; 'Breakpoint Landing
                    'Set rt = dcWithProp(aiPropsDesign, pnFamily, gnPartFam, rt)
                End If
            End If
        End With
        ''' UPDATE[2021.11.09]
        '''     Moved Part Mass Property assignment
        '''     out of the main With block, modified
        '''     to take its value from the new Values
        '''     Dictionary.
        Set rt = dcWithProp( _
            aiPropsUser, pnMass, _
            dcVlPr.Item(pnMass), rt _
        ) 'Round(cvMassKg2LbM * .Mass, 4)
        
        Call iSyncPartFactory(invDoc) 'Backport Properties to iPart Factory
        Set dcGeniusPropsPartRev20180530_broken2 = rt
    End If
End Function

Public Function dcGeniusPropsPartRev20180530_broken( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim dcChg As Scripting.Dictionary
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
    
    Dim txFilePath As String
    
    If dc Is Nothing Then
        Set dcGeniusPropsPartRev20180530_broken = _
        dcGeniusPropsPartRev20180530_broken( _
            invDoc, New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        Set dcChg = New Scripting.Dictionary
        
        With invDoc
            txFilePath = .FullFileName
            
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
            Set prPartNum = aiGetProp(aiPropsDesign, pnPartNum) 'ADDED 2021.03.11
            pnModel = prPartNum.Value
            Set prFamily = aiGetProp(aiPropsDesign, pnFamily)
            
            ''' UPDATE[2018.02.06]: Using new UserForm; see below
            With .ComponentDefinition
                ''' Request #1: Get the Mass in Pounds
                With .MassProperties
                    Set rt = dcWithProp( _
                        aiPropsUser, pnMass, _
                        Round(cvMassKg2LbM * .Mass, 4), rt _
                    )
                End With
                
                bomStruct = .BOMStructure ' kDefaultBOMStructure '''''''''''
                Set dcChg = d2g1f1(prFamily, dcChg)
            End With
            ''' At this point, nmFamily SHOULD be set
            
            'Request #4: Change Cost Center iProperty.
            If bomStruct = kNormalBOMStructure Then
                '----------------------------------------------------'
                If .SubType = guidSheetMetal Then 'for SheetMetal ---'
                '----------------------------------------------------'
                ''' NOTE[2018-05-31]: At this point, we MAY wish
                    'Request #3: Get sheet metal extent area
                    Set rt = dcFlatPatProps(.ComponentDefinition, rt)
                    ''' NOTE[2018-05-30]: Raw Material Quantity value
                    
                    'NOTE: THIS call might best be combined somehow
                    If prRawMatl Is Nothing Then
                        If rt.Exists("OFFTHK") Then
                            ''' UPDATE[2018.05.30]: Restoring original key check
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
                        ''  ACTION ADVISED[2018.09.14]: pnStock can probably be set
                        If Len(prRawMatl.Value) > 0 Then
                            pnStock = prRawMatl.Value
''' This With block copied and modified [2021.03.11]
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    'Stop 'because Material value likely invalid
                                    pnStock = ptNumShtMetal(invDoc.ComponentDefinition)
                                    Debug.Print ; 'Breakpoint Landing
                                    ''  ACTION TAKEN[2021.03.11]: temporary measure to try to ensure
                                Else
                                    ''  This section retained from source,
                                    ''  With .Fields
                                    ''      mtFamily = .Item("Family").Value
                                    ''  End With
                                End If
                            End With
'''
''' This section likely should be removed when primary issue resolved.
'''
                        Else
                            pnStock = ptNumShtMetal(.ComponentDefinition)
                        End If
                        
                        If Len(pnStock) = 0 Then
                            ''' UPDATE[2018.05.30]: Pulling ALL code/text from this section
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
                                ''' UPDATE[2018.05.30]: Pulling some extraneous commented code
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
                            Debug.Print pnModel & ": PROBABLE LAGGING"
                            Debug.Print "  TRY TO IDENTIFY, AND FILL IN BELOW."
                            Debug.Print "  PRESS ENTER ON pnStock LINE WHEN"
                            Debug.Print "  COMPLETED, THEN F5 TO CONTINUE."
                            Debug.Print "  pnStock = """ & pnStock & """"
                            Stop
                        End If
                        
                        If Len(pnStock) > 0 Then 'and ONLY then do we look for a Raw Material Family!
                            
                            With cnGnsDoyle().Execute( _
                                "select Family, Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16 " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]: Will need to address this situation
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
                                    
                                    ''' UPDATE[2021.06.18]: New pre-check for Material Item
                                    If mtFamily Like "?-MT*" Then
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
                                        ''' UPDATE[2018.05.30]: Moving part family assignment
                                    ElseIf mtFamily = "D-BAR" Then
                                        ''' UPDATE[2021.06.18]: Added check for Part Family already set
                                        If Len(nmFamily) = 0 Then
                                            nmFamily = "R-RMT"
                                        Else
                                            Debug.Print ; 'Breakpoint Landing
                                            'Stop
                                        End If
                                        
                                        qtUnit = prRmUnit.Value '"IN"
                                        ''may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
                                        Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        ''' UPDATE[2021.03.11]: Replaced aiPropsDesign.Item(pnPartNum)
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
                                        'Debug.Print "CURRENT RAW MATERIAL QUANTITY (";
                                        'Debug.Print CStr(prRmQty.Value); ") IS SHOWN BELOW."
                                        Debug.Print ""
                                        Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                                        'Debug.Print "qtUnit = """; qtUnit; """"
                                        Debug.Print "qtUnit = ""IN"""
                                        'Debug.Print ""
                                        Stop 'because we might want a D-BAR handler
                                        ''' Actually, we might NOT need to stop here
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                                        Stop
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
                                    Stop 'and check both so we DON'T automatically "fix" the RMUNIT value
                                    
                                    .Value = qtUnit
                                    
                                    If 0 Then Stop 'Ctrl-9 here to skip changing
                                End If
                            End If
                        Else 'we're setting a new quantity unit
                            .Value = qtUnit
                        End If
                    End With
                    Set rt = dcAddProp(prRmUnit, rt)
                    Debug.Print ; 'Another landing line
                    
                '--------------------------------------------'
                Else 'for standard Part (NOT Sheet Metal) ---'
                '--------------------------------------------'
                            ''' [2018.07.31 by AT] Duped following block from above
                            With newFmTest1()
                                If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                                
                                ''' [2018.07.31 by AT] Added the following to try to
                                Set bd = nuAiBoxData().UsingInches.SortingDims( _
                                    invDoc.ComponentDefinition.RangeBox _
                                )
                                
                                ck = .AskAbout(invDoc, _
                                    "Please Select Stock for Machined Part" _
                                    & vbNewLine & vbNewLine & bd.Dump(0) _
                                )
                                
                                If ck = vbYes Then
                                ''' UPDATE[2018.05.30]: Pulling some extraneous commented code
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
                                End If
                            End With
                            '''
                            '''
                            '''
'''
''' The following If block is copied wholesale from sheet metal section above.
                        If Len(pnStock) > 0 Then 'and ONLY then do we look for a Raw Material Family!
                            
''' This enclosing With block should NOT be necessary
                            With cnGnsDoyle().Execute( _
                                "select Family " & _
                                "from vgMfiItems " & _
                                "where Item='" & pnStock & "';" _
                            )
                                If .BOF Or .EOF Then
                                    Stop 'because Material value likely invalid
                                    ''  ACTION ADVISED[2018.09.14]: Will need to address this situation
                                Else
                                    With .Fields
                                        mtFamily = .Item("Family").Value
                                    End With
'''
''' Content formerly here moved BELOW and OUT of this section
                                End If
                            End With
''' These closing statements moved up from below following If block
'''

                                    If mtFamily = "DSHEET" Then
                                        Stop
'because we should NOT be doing Sheet Metal in this section.
                                        nmFamily = "D-RMT"
                                        qtUnit = "FT2"
                                        ''' UPDATE[2018.05.30]: Moving part family assignment
                                    ElseIf mtFamily = "D-BAR" Then
                                        nmFamily = "R-RMT"
                                        qtUnit = prRmUnit.Value '"IN"
                                        ''may want function here
                                        ''' UPDATE[2018.05.30]: As noted above Will keep Stop for now
                                        Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                                        ''' UPDATE[2021.03.11]: Replaced aiPropsDesign.Item(pnPartNum)
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
                                        ''' Actually, we might NOT need to stop here
                                        Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                                        Stop
                                        Set rt = dcAddProp(prRmQty, rt)
                                        Debug.Print ; 'Landing line for debugging. Do not disable.
                                    Else
                                        nmFamily = ""
                                        qtUnit = "" 'may want function here
                                        ''' UPDATE[2018.05.30]: As noted above
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
                                    Stop 'and check both so we DON'T automatically "fix" the RMUNIT value
                                    
                                    .Value = qtUnit
                                    
                                    If 0 Then Stop 'Ctrl-9 here to skip changing
                                End If
                            End If
                        Else 'we're setting a new quantity unit
                            .Value = qtUnit
                        End If
                    End With
                    Set rt = dcAddProp(prRmUnit, rt)
                End If 'Sheetmetal vs Part
            ElseIf bomStruct = kPurchasedBOMStructure Then
                ''' As mentioned above, nmFamily SHOULD be set at this point
                If Len(nmFamily) = 0 Then
                    If 1 Then Stop 'because we might need to check out the situation
                    nmFamily = "D-PTS" 'by default
                End If
            Else
                Stop 'because we might need to do something else
            End If
            
            ' Get the design tracking property set,
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
        Set dcGeniusPropsPartRev20180530_broken = rt
    End If
End Function

Public Function dcGeniusPropsPartDvl20210929( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
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
        Set dcGeniusPropsPartDvl20210929 = _
        dcGeniusPropsPartDvl20210929( _
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
            Set prPartNum = aiGetProp(aiPropsDesign, pnPartNum) 'ADDED 2021.03.11
            pnModel = prPartNum.Value
            Set prFamily = aiGetProp(aiPropsDesign, pnFamily)
            
            ''' Request #1: Get the Mass in Pounds
            With .ComponentDefinition.MassProperties
                Set rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * .Mass, 4), rt)
            End With
            
            ''' NOTE[2021.10.01]: This block is for Purchased Part Determination! (see below)
            ''' UPDATE[2018.02.06]: Using new UserForm; see below
            With .ComponentDefinition
                ''' Get BOM Structure type, correcting if appropriate,
                ck = vbNo
                ''' UPDATE[2018.05.31]: Combined both InStr checks
''' look at N++ tab "new 7" for content here
            End With
            ''' NOTE[2021.10.01]: END OF BLOCK for Purchased Part Determination!
        
            'Request #4: Change Cost Center iProperty.
            If bomStruct = kNormalBOMStructure Then
''' look at N++ tab "new 7" for content here
            ElseIf bomStruct = kPurchasedBOMStructure Then
''' look at N++ tab "new 7" for content here
            Else
''' look at N++ tab "new 7" for content here
            End If
            
            ' Get the design tracking property set,
            If invDoc.ComponentDefinition.IsContentMember Then
''' look at N++ tab "new 7" for content here
            Else
''' look at N++ tab "new 7" for content here
            End If
        End With
    End If
End Function

Public Function d2g2f1( _
    invDoc As Inventor.PartDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary 'Inventor.BOMStructureEnum
    '''
    ''' d2g2f1 -- (to be determined)
    '''
    ''' code here extracted for development
    ''' from Function dcGeniusPropsPartRev20180530
    '''     in module modGPUpdateAT (start line 559)
    '''     lines 1056 (497 down from start)
    '''        to 1201 (146 lines copied)
    ''' along with necessary declarations:
    Dim pnModel     As String
    Dim nmFamily    As String
    Dim mtFamily    As String
    Dim pnStock     As String
    Dim qtUnit      As String
    Dim bd          As aiBoxData
    Dim ck          As VbMsgBoxResult
    '''
    ''' followed by new declarations:
    Dim rt As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    
    With invDoc.ComponentDefinition
        If .Document Is invDoc Then
            Set bd = nuAiBoxData( _
            ).UsingInches.SortingDims( _
                .RangeBox _
            )
            With newFmTest1() '''== Original Line 1056 ==
                'If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                'moved this check outside this form block (see above)
                
                ''' [2018.07.31 by AT]
                ''' Added the following to try to
                ''' preselect non-sheet metal stock
                '.dbFamily.Value = "D-BAR"
                '.lbxFamily.Value = "D-BAR"
                ''' Doesn't quite do it.
                
                ck = .AskAbout(invDoc, _
                    "Please Select Stock for Machined Part" _
                    & vbNewLine & vbNewLine & bd.Dump(0) _
                )
                
                If ck = vbYes Then
                ''' UPDATE[2018.05.30]:
                '''     Pulling some extraneous commented code
                '''     from here and beginning of block
                    With .ItemData()
                        If .Exists(pnFamily) Then
                            nmFamily = .Item(pnFamily)
                            rt.Add pnFamily, nmFamily
                            Debug.Print pnFamily & "=" & nmFamily
                        End If
                        
                        If .Exists(pnRawMaterial) Then
                            pnStock = .Item(pnRawMaterial)
                            rt.Add pnRawMaterial, pnStock
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
        Else
            Stop
        End If
    End With
    
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
            End If
        End With
        
        If mtFamily = "DSHEET" Then
            Stop 'because we should NOT be doing Sheet Metal in this section.
                 'might require further investigation and/or development, if encountered.
            'We should be okay. This is sheet metal stock
            nmFamily = "D-RMT"
            qtUnit = "FT2"
            ''' UPDATE[2018.05.30]:
            '''     Moving part family assignment
            '''     to this section for better mapping
            '''     and updating to new Family names
            '''     as well as pulling up qtUnit assignment
        ElseIf mtFamily = "D-BAR" Then
            nmFamily = "R-RMT"
            Stop 'and note disabled qtUnit -- needs work here
            'qtUnit = prRmUnit.Value '"IN"
            ''may want function here
            ''' UPDATE[2018.05.30]: As noted above
            '''     Will keep Stop for now
            '''     pending further review,
            '''     hopefully soon
            Stop 'and note disabled prRawMatl too
            'Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
            ''' UPDATE[2021.03.11]: Replaced
            ''' aiPropsDesign.Item(pnPartNum)
            ''' as noted above
            Stop 'and note disabled prRmQty
            'Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
            Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
            Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."
            Debug.Print "X SPAN", "Y SPAN", "Z SPAN"
            Debug.Print (invDoc.ComponentDefinition.RangeBox.MaxPoint.X - invDoc.ComponentDefinition.RangeBox.MinPoint.X) / 2.54, (invDoc.ComponentDefinition.RangeBox.MaxPoint.Y - invDoc.ComponentDefinition.RangeBox.MinPoint.Y) / 2.54, (invDoc.ComponentDefinition.RangeBox.MaxPoint.Z - invDoc.ComponentDefinition.RangeBox.MinPoint.Z) / 2.54
            Debug.Print ""
            Debug.Print "PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED."
            Debug.Print "PRESS ENTER/RETURN TWICE. THEN CONTINUE."
            Debug.Print ""
            Stop 'and note disabled prRmQty again
            'Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
            Debug.Print "qtUnit = ""IN"""
            Debug.Print ""
            Stop 'because we might want a D-BAR handler
            ''' Actually, we might NOT need to stop here
            ''' if bar stock is already selected,
            ''' because quantities would presumably
            ''' have been established already.
            ''' Any D-BAR handler probably needs
            ''' to be implemented in prior section(s)
            Stop 'and note one moredisabled prRmQty
            'Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
            Stop
            Stop 'and note prRmQty once more disabled
            'this one really DOES need removed from this function
            'Set rt = dcAddProp(prRmQty, rt)
            Debug.Print ; 'Landing line for debugging. Do not disable.
        Else
            Stop 'because we don't know WHAT to do with it
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
    
    Set d2g2f1 = rt
'Debug.Print ConvertToJson(d2g2f1(aiDocPart(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences.ItemByName("04-18-102-1006:1").Definition.Document)), "  ")
End Function

Public Function d2g1f1( _
    prFamily As Inventor.Property, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary 'Inventor.BOMStructureEnum
    Dim rt As Scripting.Dictionary
    'Dim invDoc As Inventor.PartDocument
    Dim ck As VbMsgBoxResult
    Dim txFilePath As String
    Dim isNow As Inventor.BOMStructureEnum
    Dim shdBe As Inventor.BOMStructureEnum
    Dim gnFam As String
    Dim aiFam As String
    Dim ptNum As String
    Dim fm As New fmTest2
    
    If dc Is Nothing Then
        Set rt = d2g1f1(prFamily, New Scripting.Dictionary)
    Else
        Set rt = dc
        
        With prFamily
            ''' Get Family from Model
            aiFam = .Value
            
            With .Parent 'Property Set
                ''' Get Part Number from Model
                ptNum = .Item(pnPartNum).Value
                
                ''' Then try to get Family from Genius
                With cnGnsDoyle()
                    With .Execute(Join(Array( _
                        "select ISNULL(i.Family, '') Family", _
                        "from vgMfiItems i right join", _
                        "(values ('" & ptNum & "')) ls(Item)", _
                        "on i.Item = ls.Item", _
                        ";" _
                    ), vbNewLine))
                        If .BOF Or .EOF Then
                            Stop 'because something went wrong
                            gnFam = ""
                        Else
                            gnFam = .GetRows()(0, 0)
                        End If
                        .Close
                    End With
                End With
                
                With .Parent 'Set OF Property Sets
                    ''' Get File Path to check for Purchased Part
                    txFilePath = aiDocument(.Parent).FullFileName
                    
                    'Request #2: Change Cost Center iProperty.
                    If ck4ContentMember(.Parent) Then
                        If Len(gnFam) = 0 Then
                            gnFam = "D-HDWR"
                        ElseIf gnFam = "D-HDWR" Then
                        ElseIf gnFam = "D-PTS" Then
                        ElseIf gnFam = "R-PTS" Then
                        Else
                            Stop
                        End If
                    End If
                
                    isNow = bomStructOf(.Parent)
                    
                    'Set fm = newFmTest2()
                    
                    ''' Check Model Family against Genius Family,
                    ''' if defined, and if different, ask whether
                    ''' it should be changed.
                    If Len(gnFam) > 0 Then
                        If gnFam <> aiFam Then
                            ck = fm.AskAbout(.Parent, _
                                Join(Array( _
                                    "Model Family " & aiFam & " does not", _
                                    "match Genius Part Family " & gnFam _
                                ), vbNewLine), _
                                Join(Array( _
                                    "Should Model be updated", _
                                    "to match Genius?" _
                                ), vbNewLine) _
                            )
                            If ck = vbCancel Then
                                Stop
                            End If
                            
                            If ck = vbYes Then
                                rt.Add prFamily.Name, gnFam
                            Else
                                gnFam = aiFam
                                'going to need final family below
                                'and Genius Family will be changed
                                'anyway, if Model Family isn't
                            End If
                        End If
                    End If
                    
                    ''' Get BOM Structure type,
                    ''' correcting if appropriate,
                    ''' UPDATE[2018.05.31]: Combined both InStr checks
                    If InStr(1, txFilePath, _
                        "\Doyle_Vault\Designs\purchased\" _
                    ) + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", _
                        "|" & prFamily.Value & "|" _
                    ) > 0 Then
                        shdBe = kPurchasedBOMStructure
                    Else
                        shdBe = isNow
                    End If
                    
                    If shdBe <> isNow Then
                        ck = fm.AskAbout(.Parent, _
                            Join(Array( _
                                "Model Family " & gnFam & " or File Path", _
                                "(" & txFilePath & ")", _
                                "indicates a Purchased Part, but BOM", _
                                "Structure is NOT set to match" _
                            ), vbNewLine), _
                            Join(Array( _
                                "Should BOM Structure", _
                                "be set to Purchased?" _
                            ), vbNewLine) _
                        )
                        If ck = vbCancel Then
                            Stop
                        End If
                        
                        If ck = vbYes Then
                            'On Error Resume Next
                            '.BOMStructure = kPurchasedBOMStructure
                            'If Err.Number = 0 Then
                            '    bomStruct = .BOMStructure
                            'Else
                            '    bomStruct = kPurchasedBOMStructure
                            '    ''' WARNING: NOT a good way to go about this
                            '    '''     but will go with it for now
                            'End If
                            'On Error GoTo 0
                            rt.Add "BOMstructure", shdBe
                        Else
                            shdBe = isNow
                            'want to know what the USER says
                            'it should be, and BOM structure
                            'MIGHT affect Genius Part status
                            'in subsequent import operation.
                            
                            'bomStruct = .BOMStructure 'to make sure this is captured
                        End If
                    End If
                End With
            End With
            Debug.Print 'Breakpoint Landing
        End With
    End If
    
    Set d2g1f1 = rt
'Debug.Print txDumpLs(d2g1f1(aiDocActive().PropertySets(gnDesign).Item(pnFamily)).Keys)
End Function

Public Function dcCtOfEach(ls As Variant) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ck As Variant
    Dim mx As Long
    Dim dx As Long
    
    Set rt = New Scripting.Dictionary
    If IsArray(ls) Then
        mx = UBound(ls)
        dx = LBound(ls)
        With rt
            Do Until dx > mx
                ck = ls(dx)
                If .Exists(ck) Then
                    .Item(ck) = _
                    .Item(ck) + 1
                Else
                    .Add ck, 1
                End If
                
                dx = 1 + dx
            Loop
        End With
        
        Set dcCtOfEach = rt
    Else
        Set dcCtOfEach = _
            dcCtOfEach(Array(ls))
    End If
End Function

Public Function dcGnsMatlOps( _
    DimCt As Scripting.Dictionary, _
    Optional MtSpec As String = "" _
) As Scripting.Dictionary 'defaulted to SS, but maybe not such a great idea
    '''
    '''
    '''
    Dim rt As Scripting.Dictionary
    Dim rw As Scripting.Dictionary
    Dim rs As ADODB.Recordset
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With cnGnsDoyle()
        On Error Resume Next
        
        Err.Clear
        Set rs = .Execute( _
        sqlOf_GnsMatlOptions( _
            MtSpec, DimCt.Keys _
        ))
        
        If Err.Number = 0 Then
            With dcFromAdoRS(rs, "")
            For Each ky In .Keys
                Set rw = dcOb(.Item(ky))
                If rw Is Nothing Then
                    Stop
                Else
                    rt.Add rw.Item("Item"), rw
                End If
            Next: End With
            
            rs.Close
        Else
            Stop
            Err.Clear
        End If
        
        On Error GoTo 0
        .Close
    End With
    Set dcGnsMatlOps = rt
End Function

Public Function ck4ContentMember( _
    AiDoc As Inventor.Document _
) As Boolean
    ck4ContentMember = _
    ptIsContentMember( _
    aiDocPart(AiDoc))
End Function

Public Function ptIsContentMember( _
    AiDoc As Inventor.PartDocument _
) As Boolean
    If AiDoc Is Nothing Then
        ptIsContentMember = 0
    Else
        ptIsContentMember = AiDoc.ComponentDefinition.IsContentMember
    End If
End Function

Public Function bomStructOfPart( _
    AiDoc As Inventor.PartDocument _
) As Inventor.BOMStructureEnum
    If AiDoc Is Nothing Then
        bomStructOfPart = 0
    Else
        bomStructOfPart = AiDoc.ComponentDefinition.BOMStructure
    End If
End Function

Public Function bomStructOfAssy( _
    AiDoc As Inventor.AssemblyDocument _
) As Inventor.BOMStructureEnum
    If AiDoc Is Nothing Then
        bomStructOfAssy = 0
    Else
        bomStructOfAssy = AiDoc.ComponentDefinition.BOMStructure
    End If
End Function

Public Function bomStructOf( _
    AiDoc As Inventor.Document _
) As Inventor.BOMStructureEnum
    If AiDoc Is Nothing Then
        bomStructOf = 0
    ElseIf TypeOf AiDoc Is Inventor.PartDocument Then
        bomStructOf = bomStructOfPart(AiDoc)
    ElseIf TypeOf AiDoc Is Inventor.AssemblyDocument Then
        bomStructOf = bomStructOfAssy(AiDoc)
    Else
        bomStructOf = 0
    End If
End Function
