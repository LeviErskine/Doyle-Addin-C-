

Public Function dcCutTimePerimeter(ad As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional incTop As Long = 0 _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ActiveDoc As Inventor.Document
    Dim pt As Inventor.Document
    Dim ky As Variant
    
    'If dc Is Nothing Then
        'Set dcCutTimePerimeter = dcCutTimePerimeter( _
            ad, New Scripting.Dictionary, incTop _
        )
    'Else
        Set rt = New Scripting.Dictionary
        
        With dcAiDocComponents(ad, dc, incTop)
            For Each ky In .Keys
                Set pt = aiDocument(.Item(ky))
                rt.Add _
                    pt.PropertySets.Item( _
                        gnDesign).Item( _
                        pnPartNum).Value, _
                    fpPerimeterInch(pt) _
                    ', aiPropVal(aiPropShtMetalThickness(pt), -1)
                ''
            Next
        End With
        
        Set dcCutTimePerimeter = rt
    'End If
End Function
'Debug.Print dumpLsKeyVal(dcCutTimePerimeter(ThisApplication.ActiveDocument))

Public Function mdl1g0f0() As Long
    Dim dc As Scripting.Dictionary
    Dim ky As Variant
    Dim ad As Inventor.Document
    'Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    
    Set dc = dcAssyDocComponents(ThisApplication.Documents.ItemByName( _
        "C:\Doyle_Vault\Designs\Misc\andrewT\02\02-weldmentStd-01.iam" _
    ))
    With dc
        For Each ky In .Keys
            Set ad = aiDocument(.Item(ky))
            With dcGeniusProps(ad)
                If .Exists(pnRawMaterial) Then
                    Set pr = ad.PropertySets(gnCustom).Item(pnRawMaterial)
                    With pr
                        Debug.Print .Value
                        If .Value Like "FM-*" Then
                            With New fmTest1
                                If .AskAbout(ad) = vbYes Then
                                    With .ItemData
                                        '
                                        Stop
                                        pr.Value = .Item(pnRawMaterial)
                                    End With
                                End If
                            End With
                        End If
                    End With
                    'Stop
                Else
                    'Stop
                End If
            End With
        Next
    End With
End Function
'Debug.Print cnGnsDoyle.Execute("select I.ItemID, I.Thickness, I.Item, I.Description1 from vgMfiItems as I where I.Family='DSHEET'").GetString

Public Function mdl1g1f0() As Long
    With New fmTest1
        .AskAbout ThisApplication.ActiveDocument
    End With
End Function

Public Function mdl1g1f2(lb As MSForms.Label) As Single ', txt As String
    Dim x0 As Long, x1 As Single
    Dim y0 As Long, y1 As Single
    Dim ct As MSForms.Control
    
    Set ct = lb
    With ct
        'x0 = .Left
        'y0 = .Top
        x1 = .Width
        y1 = .Height
        With lb
            '.Caption = txt
            .AutoSize = True
            .AutoSize = False
        End With
        .Width = x1
        mdl1g1f2 = .Height - y1
    End With
End Function

Public Function mdl1g1f3( _
    ct As MSForms.Control, _
    byX As Single, byY As Single _
) As Single
    With ct
        .Left = .Left + byX
        .Top = .Top + byY
    End With
    mdl1g1f3 = Sqr(byX * byX + byY * byY)
End Function

''' For lack of a better place to put it, creating this node
''' The following is a basic example of accessing Parameters,
''' such as dimensions, from an Inventor Part Document.
Public Function mdl1g1f1() As Long
    With aiDocPart(ThisApplication.ActiveDocument)
        'aiDocPart casts an Inventor Part Document
        'from its general Document reference, if valid.
        With .ComponentDefinition.Parameters.Item("Thickness")
            Debug.Print .ExposedAsProperty
            Debug.Print .Value
        End With
    End With
End Function
''' This example was written as a quick one-off to see how
''' an Inventor Parameter like the Thickness setting for
''' Sheet Metal Parts might have its "Export" status
''' modified programmatically.

Public Function aiPropVal( _
    pr As Inventor.Property, _
    Optional ifNot As Variant = "" _
) As Variant
    If pr Is Nothing Then
        aiPropVal = ifNot
    Else
        aiPropVal = aiPropValAux( _
            pr.Value, ifNot _
        )
    End If
End Function
''' This example was written as a quick one-off to see how
''' an Inventor Parameter like the Thickness setting for
''' Sheet Metal Parts might have its "Export" status
''' modified programmatically.

Public Function aiPropValAux( _
    vl As Variant, Optional ifNot As Variant = "" _
) As Variant
    If IsObject(vl) Then
        If vl Is Nothing Then
            aiPropValAux = ifNot
        Else
            If TypeOf vl Is stdole.StdPicture Then 'IPictureDisp
                aiPropValAux = "<stdole.StdPicture>"
                Debug.Print ; 'Breakpoint Landing
            Else
                Stop 'and see what we need to do
                aiPropValAux = "<Object:" & TypeName(vl) & ">"
            End If
        End If
    Else
        aiPropValAux = vl
    End If
End Function

Public Function aiPropGnsItmFamily( _
    AiDoc As Inventor.Document _
) As Inventor.Property
    If AiDoc Is Nothing Then
        Set aiPropGnsItmFamily = Nothing
    Else
        Set aiPropGnsItmFamily = AiDoc.PropertySets( _
            gnDesign).Item(pnFamily _
        )
    End If
End Function

Public Function aiPropShtMetalThickness( _
    adPart As Inventor.PartDocument _
) As Inventor.Property
    If adPart Is Nothing Then
        Set aiPropShtMetalThickness = Nothing
    Else
        With adPart
            If .SubType = guidSheetMetal Then
                If smThicknessExposed(.ComponentDefinition) Then
                    Set aiPropShtMetalThickness = .PropertySets(gnCustom).Item(pnThickness)
                Else
                    Set aiPropShtMetalThickness = Nothing
                End If
            Else
                Set aiPropShtMetalThickness = Nothing
            End If
        End With
    End If
End Function

Public Function smThicknessExposed( _
    smDef As Inventor.SheetMetalComponentDefinition _
) As Long
    If smDef.Parameters.IsExpressionValid(pnThickness, "in") Then
    smThicknessExposed = parExposed( _
        smDef.Parameters(pnThickness), 1 _
    )
    'With smDef.Parameters(pnThickness)
        'If Not .ExposedAsProperty Then
            '.ExposedAsProperty = True
        'End If
        'smThicknessExposed = IIf(.ExposedAsProperty, -1, 0)
    'End With
    Else
        Stop
    End If
End Function

Public Function parExposed( _
    par As Inventor.Parameter, _
    Optional tryTo As Long = 0 _
) As Long
    ''  Check Inventor Parameter for exposure as Property.
    ''  Return 0 if not, unless caller requests exposure
    ''  (tryTo <> 0). Nonzero return indicates exposed
    ''  Parameter, with sign indicating initial status.
    ''  -1 indicates Parameter already exposed
    ''  1 indicates status change to expose it.
    ''  No provision is made for failure to expose,
    ''  nor to reverse exposure status.
    With par
        If .ExposedAsProperty Then
            parExposed = -1
        ElseIf tryTo Then
            .ExposedAsProperty = True
            parExposed = 1 And parExposed(par)
        Else
            parExposed = 0
        End If
    End With
End Function

Public Function dcGnsPropsListed( _
    ad As Inventor.Document, ls As Variant, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional ifNone As Long = 1 _
) As Scripting.Dictionary
    '''
    ''' dcGnsPropsListed --
    '''     Return a Dictionary of any
    '''     Properties in the supplied list
    '''     from the "custom" PropertySet.
    '''
    '''     Missing Property names are addressed
    '''     in one of three (present) ways,
    '''     based on optional argument ifNone:
    '''         0 - do not add to Dictionary
    '''             missing name is missing
    '''         1 - attempt to create. failure
    '''             returns Nothing, which is
    '''             still not added
    '''         2 - add Nothing to Dictionary
    '''             under missing name
    '''         3 - attempt to create, adding
    '''             Nothing for any failures
    '''             (combines options 1 and 2)
    '''
    Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    Dim ky As Variant
    Dim mkNf As Long 'try to make any not found
    Dim rtNf As Long 'return Nothing for not found
    Dim wk As Variant
    
    'Set rt = New Scripting.Dictionary
    Set ps = ad.PropertySets.Item(gnCustom)
    
    If dc Is Nothing Then
        Set dcGnsPropsListed = dcGnsPropsListed( _
            ad, ls, New Scripting.Dictionary, ifNone _
        )
    Else
        If IsArray(ls) Then
            mkNf = ifNone And 1 'IIf(ifNone = 1, 1, 0)
            rtNf = ifNone And 2 'IIf(ifNone = 2, 1, 0)
            '''
            ''' originally used IIf construct
            ''' to force mapping of exact values
            ''' of ifNone to corresponding behaviors.
            '''
            ''' changed to bitcode matching once clear
            ''' that each bit would map exclusively
            ''' to a particular behavior, and could
            ''' be combined with the other, if desired.
            '''
            
            For Each ky In ls 'Array(pnMass _
                , pnArea, pnWidth, pnLength _
                , pnRawMaterial, pnRmQty, pnRmUnit _
            ) ' _
                , "SPEC01", "SPEC02", "SPEC03" _
                , "SPEC04", "SPEC05", "SPEC06" _
                , "SPEC07", "SPEC08", "SPEC16" _
            '
                Set pr = aiGetProp(ps, CStr(ky), mkNf)
                wk = Array(pr)
                
                If pr Is Nothing Then 'check
                    'if supposed to return Nothings
                    
                    If rtNf = 0 Then wk = Array()
                    'and empty the array if not
                End If
                
                If UBound(wk) < LBound(wk) Then
                    'we have an empty array,
                    'and nothing to add,
                    'not even Nothing!
                Else
                    If dc.Exists(ky) Then
                        dc.Remove ky
                        ''' WARNING[2021.11.19]
                        '''     This was added to permit
                        '''     replacement of elements
                        '''     already present under a
                        '''     supplied key. It might
                        '''     NOT be the best way to
                        '''     address this situation.
                        '''     Be prepared to correct
                        '''     this with a more robust
                        '''     solution in future.
                        ''' Meanwhile, have a
                        Debug.Print ; 'Breakpoint Landing
                        '''
                    End If
                    
                    dc.Add ky, pr 'rt
                End If
            Next
            Set dcGnsPropsListed = dc 'rt
        ElseIf VarType(ls) = vbString Then
            Set dcGnsPropsListed _
            = dcGnsPropsListed( _
                ad, Array(ls), _
                dc, ifNone _
            )
        Else
            Stop
        End If
    End If
End Function

Public Function dcGnsPropsPart(ad As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional ifNone As Long = 1 _
) As Scripting.Dictionary
    '''
    ''' dcGnsPropsPart
    '''
    ''' REV[2021.11.18]:
    '''     Added pnThickness to list
    '''     of Properties to return.
    '''
    Set dcGnsPropsPart = dcGnsPropsListed( _
    ad, Array(pnMass _
        , pnArea, pnWidth, pnLength, pnThickness _
        , pnRawMaterial, pnRmQty, pnRmUnit _
    ), dc, ifNone)
End Function

Public Function dcGnsPropsAssy(ad As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional ifNone As Long = 1 _
) As Scripting.Dictionary
    'Dim rt As Scripting.Dictionary
    'Dim ps As Inventor.PropertySet
    'Dim pr As Inventor.Property
    'Dim ky As Variant
    
    Set dcGnsPropsAssy = dcGnsPropsListed(ad, _
        Array(pnMass _
            , "SPEC01", "SPEC02", "SPEC03" _
            , "SPEC04", "SPEC05", "SPEC06" _
            , "SPEC07", "SPEC08", "SPEC16" _
        ), dc, ifNone _
    )
    
    'Set rt = New Scripting.Dictionary
    'Set ps = ad.PropertySets.Item(gnCustom)
    
    'If dc Is Nothing Then
        'Set dcGnsPropsAssy = dcGnsPropsAssy( _
            ad, New Scripting.Dictionary _
        )
    'Else
        'For Each ky In Array(pnMass _
            , "SPEC01", "SPEC02", "SPEC03" _
            , "SPEC04", "SPEC05", "SPEC06" _
            , "SPEC07", "SPEC08", "SPEC16" _
        ) ' _
            , pnArea, pnWidth, pnLength _
            , pnRawMaterial, pnRmQty, pnRmUnit _
        '
            'Set pr = aiGetProp(ps, CStr(ky), 1)
            'If pr Is Nothing Then
                'nothing we can do (as yet?)
            'Else
                'dc.Add ky, pr 'rt
            'End If
        'Next
        'Set dcGnsPropsAssy = dc 'rt
    'End If
End Function

Public Function dcProps4genius(ad As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing, _
    Optional Create As Long = 1 _
) As Scripting.Dictionary
    'Dim rt As Scripting.Dictionary
    Dim ps As Inventor.PropertySet
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    If dc Is Nothing Then
        Set dcProps4genius = dcProps4genius( _
            ad, New Scripting.Dictionary, Create _
        )
    Else
        With ad
        If .DocumentType = kAssemblyDocumentObject Then
            Set dcProps4genius = dcGnsPropsAssy(ad, dc, Create)
        ElseIf .DocumentType = kPartDocumentObject Then
            Set dcProps4genius = dcGnsPropsPart(ad, dc, Create)
        Else
            Set dcProps4genius = dc
        End If
        End With
    End If
    
    'Stop 'With Exit Function above, should never get here
    '
    ''Set rt = New Scripting.Dictionary
    'Set ps = ad.PropertySets.Item(gnCustom)
    '
    'If dc Is Nothing Then
    '    Set dcProps4genius = dcProps4genius( _
    '        ad, New Scripting.Dictionary _
    '    )
    'Else
    '    For Each ky In Array(pnMass _
    '        , pnArea, pnWidth, pnLength _
    '        , pnRawMaterial, pnRmQty, pnRmUnit _
    '        , "SPEC01", "SPEC02", "SPEC03" _
    '        , "SPEC04", "SPEC05", "SPEC06" _
    '        , "SPEC07", "SPEC08", "SPEC16" _
    '    ) ' _
    '    '
    '        Set pr = aiGetProp(ps, CStr(ky), 1)
    '        If pr Is Nothing Then
    '            'nothing we can do (as yet?)
    '        Else
    '            dc.Add ky, pr 'rt
    '        End If
    '    Next
    '    Set dcProps4genius = dc 'rt
    'End If
End Function

Public Function mdl1g2f1(ad As Inventor.Document) As Inventor.WorkPlanes
    Set mdl1g2f1 = aiDocPart(ad).ComponentDefinition.WorkPlanes
End Function

Public Function mdl1g3f0(ad As Inventor.Document) As Double
    Dim rt As Double
    
    Select Case ad.DocumentType
    Case kPartDocumentObject
        rt = aiDocPart(ad).ComponentDefinition.MassProperties.Mass
    Case kAssemblyDocumentObject
        rt = aiDocAssy(ad).ComponentDefinition.MassProperties.Mass
    Case Else
        rt = 0#
    End Select
    
    With ad.UnitsOfMeasure
        mdl1g3f0 = .ConvertUnits(rt, _
            kKilogramMassUnits, _
            kLbMassMassUnits _
        ) '.MassUnits)
    End With
End Function

Public Function mdl1g4f0() As Long
    Dim mx As Long
    Dim dx As Long
    
    With ThisApplication.CommandManager.ControlDefinitions
        mx = .Count
        For dx = 1 To mx
            With .Item(dx)
                If .InternalName Like "*ault*" Then
                    Debug.Print CStr(dx) & ": " & .InternalName & "/" & .DisplayName
                End If
            End With
        Next
    End With
End Function

Public Function mdl1g5f0(ad As Inventor.Document) As Scripting.Dictionary
    ''' The purpose of this function is to return a Dictionary
    ''' of Genius Family Inventor Properties
    ''' for each component Document of an assembly
    ''' or a single part Document.
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dcAiDocComponents(ad, New Scripting.Dictionary, 1) 'sc
        For Each ky In .Keys
            With aiDocument(.Item(ky))
                rt.Add .FullFileName, aiPropGnsItmFamily(.PropertySets.Parent)
            End With
        Next
    End With
    Set mdl1g5f0 = rt
End Function

Public Function mdl1g5f1(ad As Inventor.Document) As Scripting.Dictionary
    ''' This function calls mdl1g5f0 to retrieve a Dictionary
    ''' of Genius Family Inventor Properties, and then
    ''' transforms it into a Dictionary of Dictionaries
    ''' grouped by Family Property Value
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim ky As Variant
    Dim fm As String
    Dim pr As Inventor.Property
    
    Set rt = New Scripting.Dictionary
    With mdl1g5f0(ad)
        For Each ky In .Keys
            Set pr = aiProperty(.Item(ky))
            fm = pr.Value
            With rt
                If Not .Exists(fm) Then
                    .Add fm, New Scripting.Dictionary
                End If
                dcOb(.Item(fm)).Add ky, pr
            End With
        Next
    End With
    Set mdl1g5f1 = rt
End Function
'Debug.Print txDumpLs(mdl1g5f1(ThisApplication.ActiveDocument).Keys)
'Debug.Print txDumpLs(dcOb(mdl1g5f1(ThisApplication.ActiveDocument).Item("")).Keys)

Public Function mdl1g5f2(ad As Inventor.Document) As Scripting.Dictionary
    ''' The purpose of this function is to return a Dictionary
    ''' of Genius Family Inventor Properties
    ''' for each component Document of an assembly
    ''' or a single part Document.
    Dim rt As Scripting.Dictionary
    Dim fm As fmTest1
    Dim ky As Variant
    
    Set fm = New fmTest1
    Set rt = New Scripting.Dictionary
    With mdl1g5f1(ad)
        If .Exists("") Then
            With dcOb(.Item(""))
                For Each ky In .Keys
                    With aiProperty(.Item(ky))
                        If fm.AskAbout(.Parent) = vbOK Then
                            Stop
                        Else
                            Stop
                        End If
                        Stop
                        'rt.Add .FullFileName, aiPropGnsItmFamily(.PropertySets.Parent)
                    End With
                Next
            End With
        End If
    End With
    Set mdl1g5f2 = rt
End Function

Public Function mdl1g5f3(ad As Inventor.AssemblyDocument) As Scripting.Dictionary
    ''' Scan immediate members of Assembly document
    ''' and group in Dictionary by declared Part Number
    ''' and sub-grouped by Full Document Name.
    '''
    ''' (I wonder if an ADO Recordset wouldn't be a better choice?)
    '''
    Dim oc As Inventor.ComponentOccurrence
    Dim sd As Inventor.Document
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    'Dim fm As fmTest1
    'Dim ky As Variant
    Dim pn As String
    
    'Set fm = New fmTest1
    
    Set rt = New Scripting.Dictionary
    For Each oc In ad.ComponentDefinition.Occurrences
        Set sd = oc.Definition.Document 'aiDocument()
        pn = sd.PropertySets.Item(gnDesign).Item(pnPartNum).Value
        With rt
            If .Exists(pn) Then
                Set gp = dcAiDocsByFullDocName(sd, .Item(pn))
            Else
                .Add pn, dcAiDocsByFullDocName(sd, _
                    New Scripting.Dictionary _
                )
            End If
        End With
    Next
    Set mdl1g5f3 = rt
End Function
'Debug.Print txDumpLs(mdl1g5f3(ThisApplication.ActiveDocument).Keys)

Public Function mdl1g5f4(dc As Scripting.Dictionary) As Scripting.Dictionary
    ''' Transform keys from supplied Dictionary
    ''' (expected from mdl1g5f3)
    ''' into header/member indented form.
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim dl As String
    
    dl = vbNewLine & vbTab
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            rt.Add ky & vbTab & Join( _
                dcOb(.Item(ky)).Keys, _
                vbNewLine & ky & vbTab _
            ), .Item(ky)
        Next
    End With
    Set mdl1g5f4 = rt
End Function
'Debug.Print txDumpLs(mdl1g5f4(mdl1g5f3(ThisApplication.ActiveDocument)).Keys)

Public Function dcAiDocsByFullDocName( _
    ad As Inventor.Document, _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    ''' Add supplied Inventor Document
    ''' to supplied Dictionary
    ''' under its Full Document Name
    ''' (supports mdl1g5f3)
    Dim ky As String
    
    ky = ad.FullDocumentName
    If dc Is Nothing Then
        Set dcAiDocsByFullDocName = dcAiDocsByFullDocName(ad, New Scripting.Dictionary)
    Else
        With dc
            If .Exists(ky) Then
                .Item(ky) = 1 + .Item(ky)
            Else
                .Add ky, 1
            End If
        End With
        Set dcAiDocsByFullDocName = dc
    End If
End Function

Public Function dcAssyDocsByPtNum(ad As Inventor.AssemblyDocument) As Scripting.Dictionary
    ''' Derived from mdl1g5f3
    '''
    ''' Scan immediate members of Assembly Document
    ''' and collect source Documents in Dictionary,
    ''' grouped by declared Part Number.
    '''
    Dim oc As Inventor.ComponentOccurrence
    Dim sd As Inventor.Document
    Dim rt As Scripting.Dictionary
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    For Each oc In ad.ComponentDefinition.Occurrences
        Set sd = oc.Definition.Document
        pn = sd.PropertySets.Item(gnDesign).Item(pnPartNum).Value
        With rt
            If .Exists(pn) Then
                If sd Is .Item(pn) Then 'we're okay
                    'so carry on
                Else 'we've got a problem, so
                    Stop 'and check it out
                End If
            Else
                .Add pn, sd
            End If
        End With
    Next
    Set dcAssyDocsByPtNum = rt
End Function
'Debug.Print txDumpLs(dcAssyDocsByPtNum(ThisApplication.ActiveDocument).Keys)

Public Function dcAiDocsByCompList(dc As Scripting.Dictionary) As Scripting.Dictionary
    ''' Derived from mdl1g5f4
    ''' Transform keys from supplied Dictionary
    ''' (expected from dcAssyDocsByPtNum)
    ''' into tab-delimited list form.
    Dim rt As Scripting.Dictionary
    Dim sd As Inventor.Document
    Dim ky As Variant
    Dim dl As String
    
    dl = vbNewLine & vbTab
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set sd = aiDocument(.Item(ky))
            With sd
                If .DocumentType = kAssemblyDocumentObject Then
                    rt.Add ky & vbTab & Join( _
                        Split(txDumpLs( _
                            mdl1g5f4(mdl1g5f3(sd)).Keys _
                        ), vbNewLine), _
                        vbNewLine & ky & vbTab _
                    ), sd
                ElseIf .DocumentType = kPartDocumentObject Then
                    'Stop
                    With dcAiPropsInSet(sd.PropertySets.Item(gnCustom))
                        If .Exists(pnRawMaterial) Then
                            dl = Trim$(aiProperty(.Item(pnRawMaterial)).Value)
                            If Len(dl) = 0 Then
                                dl = "NO_RAW_STOCK" & vbTab & "<No Raw Stock Declared>"
                            Else
                                Stop
                                With cnGnsDoyle.Execute(Join(Array( _
                                    "select Description1", _
                                    "from vgMfiItems", _
                                    "where Item = '" & dl & "';" _
                                ), vbNewLine))
                                    If .BOF Or .EOF Then
                                        dl = dl & vbTab & "<Stock Number Not Found>"
                                    Else
                                        dl = dl & vbTab & .Fields(0).Value
                                    End If
                                End With
                            End If
                        Else
                            dl = "NO_RAW_STOCK" & vbTab & "<No Raw Stock Declared>"
                        End If
                    End With
                    rt.Add ky & vbTab & dl, sd
                    'rt.Add ky & vbTab & "(RAW STOCK NOT YET IMPLMENTED)", sd
                Else
                    rt.Add ky & vbTab & "(UNSUPPORTED DOCUMENT TYPE)", sd
                End If
            End With
        Next
    End With
    Set dcAiDocsByCompList = rt
End Function
'Debug.Print txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)
'send2clipBd txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)

Public Function rsWinUpdHist() As ADODB.Recordset
    ''' Windows Update History
    Dim it As WUApiLib.IUpdateHistoryEntry
    Dim rt As ADODB.Recordset
    Dim ls As Variant
    
    Set rt = rsNewWinUpdHist
    ls = Array("ResultCode", "Operation", "Title", "Description", "Date")
    With New WUApiLib.UpdateSession '.CreateUpdateSearcher
        With .CreateUpdateSearcher
            For Each it In .QueryHistory(0, .GetTotalHistoryCount)
                With it
                    'Debug.Print .ResultCode, .Operation, .Title, .Description, .Date
                    rt.AddNew ls, Array( _
                        .ResultCode, .Operation, _
                        .Title, .Description, .Date _
                    )
                End With
            Next
            rt.Filter = ""
        End With
    End With
    Set rsWinUpdHist = rt
End Function

Public Function rsNewWinUpdHist() As ADODB.Recordset
    Dim rt As ADODB.Recordset
    Set rt = New ADODB.Recordset
    With rt
        With .Fields
            '.Append "", adBigInt
            '.Append "", adVarChar, 1024
            .Append "ResultCode", adBigInt
            .Append "Operation", adBigInt
            .Append "Title", adVarChar, 256
            .Append "Description", adVarChar, 1024
            .Append "Date", adDBDate
        End With
        .Open
    End With
    Set rsNewWinUpdHist = rt
End Function

Public Function rsShtMtlCutPars( _
    ad As Inventor.Document, _
    Optional incTop As Long = 0 _
) As ADODB.Recordset
    ''' Windows Update History
    Dim rt As ADODB.Recordset
    Dim ActiveDoc As Inventor.Document
    Dim pt As Inventor.Document
    Dim ls As Variant
    Dim ky As Variant
    
    Set rt = rsNewShtMtlCutPars
    ls = Array("Item", "Description", "Thickness", "Perimeter")
    
    With dcAiDocComponents(ad, , incTop)
        For Each ky In .Keys
            Set pt = aiDocument(.Item(ky))
            With pt.PropertySets.Item(gnDesign)
                rt.AddNew ls, Array( _
                    .Item(pnPartNum).Value, _
                    .Item(pnDesc).Value, _
                    aiPropVal(aiPropShtMetalThickness(aiDocPart(pt)), -1), _
                    fpPerimeterInch(pt) _
                )
            End With
        Next
        rt.Filter = ""
    End With
    
    Set rsShtMtlCutPars = rt
End Function
'send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , "|")
'send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , vbTab)

Public Function rsNewShtMtlCutPars() As ADODB.Recordset
    Dim rt As ADODB.Recordset
    
    Set rt = New ADODB.Recordset
    With rt
        With .Fields
            '.Append "", adBigInt
            '.Append "", adVarChar, 1024
            '.Append "Date", adDBDate
            '
            .Append "Item", adVarChar, 32, adFldKeyColumn
            .Append "Description", adVarChar, 128
            .Append "Thickness", adDouble
            .Append "Perimeter", adDouble
        End With
        .Open
    End With
    
    Set rsNewShtMtlCutPars = rt
End Function
