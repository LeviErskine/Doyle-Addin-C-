

Public Function d0g0f0( _
    cd As Inventor.SheetMetalComponentDefinition, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    '''
    ''' New Sheet Metal Part processing function
    ''' Reference function d0g0f0
    '''
    Dim rt As Scripting.Dictionary
    ''
    Dim pt As Inventor.PartDocument
    Dim ps As Inventor.PropertySet
    Dim prThk As Inventor.Parameter
    ''
    Dim ck As VbMsgBoxResult
    Dim ec As Long
    Dim ed As String
    'Dim op As Boolean
    Dim v1 As Inventor.View
    Dim v2 As Inventor.View
    ''
    Dim dLength As Double
    Dim dWidth As Double
    Dim dArea As Double
    Dim strWidth As String
    Dim strLength As String
    Dim strArea As String
    ''
    Dim dHeight As Double
    Dim dfHtThk As Double
    Dim strDVNs As String
    ''
    
    If dc Is Nothing Then
        Set d0g0f0 = d0g0f0(cd, _
        New Scripting.Dictionary)
    Else
        Set rt = dc
        If cd Is Nothing Then
            'Stop
        Else
            Set pt = cd.Document
            Set ps = pt.PropertySets.Item(gnCustom)
            Set v1 = d0g1f0(pt)
            'op = pt.Open
            
            On Error Resume Next
            Err.Clear
            Set prThk = cd.Thickness
            ec = Err.Number
            ed = Err.Description
            On Error GoTo 0
            
            If ec = 0 Then 'we're good so far
                If Not cd.HasFlatPattern Then
                    ck = newFmTest2().AskAbout(pt, , _
                        "NO FLAT PATTERN!" & vbNewLine & _
                        "Try to generate one?" _
                    )
                    If ck = vbYes Then
                        'If Not op Then Stop
                        'Want to see if forcing an Unfold
                        'causes an unopened document to open.
                        'Also want to see how Open Property
                        'relates to a referenced Document
                        'not (yet) separately opened.
                        
                        Err.Clear
                        cd.Unfold
                        If Err.Number = 0 Then
                            If cd.HasFlatPattern Then
                                cd.FlatPattern.ExitEdit
                            End If
                        Else
                            Stop 'Couldn't make Flat Pattern
                        End If
                        Err.Clear
                        
                        Set v2 = d0g1f0(pt)
                        If Not v2 Is Nothing Then
                            If v1 Is Nothing Then v2.Close
                        End If
                    Else
                    End If
                End If
                
                If cd.HasFlatPattern Then
                    With cd.FlatPattern
                        'First, make sure it's VALID
                        With .Body.RangeBox
                            ' Check height against thickness
                            ' Valid flat pattern should return
                            ' zero or VERY minimal difference
                            dHeight = (.MaxPoint.Z - .MinPoint.Z)
                            dfHtThk = Abs(dHeight - prThk.Value)
                            
                            ' Get the extent of the face.
                            ' Extract the width, length and area from the range.
                            dLength = (.MaxPoint.X - .MinPoint.X)
                            dWidth = (.MaxPoint.Y - .MinPoint.Y)
                            dArea = dLength * dWidth
                        End With
                        'Stop
                        '''
                        ''' At this point, we should have enough
                        ''' to check at least a few things,
                        ''' and possibly pick out stock.
                        '''
                        If dfHtThk > 0.01 Then
                            'Stop 'and prep for machined (non sheet metal) specs
                            ''' Pretty sure dimension values
                            ''' come through in centimeters
                            ''' so try converting them here
                            'sort3dimsUp
                            d0g1f4 cd
                            
                            Set rt = d0g1f3(pt, _
                                sort3dimsUp( _
                                    dHeight / cvLenIn2cm, _
                                    dWidth / cvLenIn2cm, _
                                    dLength / cvLenIn2cm _
                                ), _
                            rt)
                        Else
                            'Stop 'and prep to verify sheet metal processing
                        End If
                        
                        If dArea > 0 Then 'this one's a longshot, BUT!
                            ''' an invalid flat pattern SHOULD have no geometry,
                            ''' which means it SHOULD have no area to speak of.
                            ''' '
                            ''' One would think this obvious, in retrospect,
                            ''' but one would not be surprised to be proven wrong.
                            ''' Again.
                            
                            With pt
                                
                                ' Convert values into document units.
                                ' This will result in strings that are identical
                                ' to the strings shown in the Extent dialog.
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
                    
                    ' Add area to custom property set
'                        Set rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt)
                    
                    ' Add Width to custom property set
'                        Set rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt)
                    
                    ' Add Length to custom property set
'                        Set rt = dcWithProp(aiPropSet, pnLength, strLength, rt)
                    
                    ' Add AreaDescription to custom property set
'                        Set rt = dcWithProp(aiPropSet, pnArea, strArea, rt)
                    
                    If Len(strDVNs) > 0 Then
'                            Set rt = dcWithProp(aiPropSet, "OFFTHK", strWidth, rt)
                    End If
                Else
                End If
            Else
                Stop
            End If
            On Error GoTo 0
        End If
        Set d0g0f0 = rt
    End If
    
    '''
    '''
    '''
End Function
'For Each dc In aiDocAssy(aiDocActive).ComponentDefinition.Occurrences: Debug.Print aiDocument(aiCompOcc(obOf(dc)).Definition.Document).Open, aiDocument(aiCompOcc(obOf(dc)).Definition.Document).FullDocumentName: Next
'Looks like Open property will NOT distinguish documents in tab list from those not
'All entries came up True

Public Function d0g1f0(rf As Inventor.Document) As Inventor.View
    Dim rt As Inventor.View
    Dim vw As Inventor.View
    
    Set rt = Nothing
    For Each vw In ThisApplication.Views
        If vw.Document Is rf Then Set rt = vw
    Next
    Set d0g1f0 = rt
End Function

Public Function d0g1f1( _
    rf As Inventor.PartDocument _
) As Inventor.SheetMetalComponentDefinition
    If rf Is Nothing Then
        Set d0g1f1 = Nothing
    Else
        Set d0g1f1 = aiCompDefShtMetal( _
            rf.ComponentDefinition _
        )
    End If
End Function

Public Function noVal(Optional vt As VbVarType) As Variant
    If vt And vbArray Then
        noVal = Array()
    Else
    Select Case vt
        Case vbString:          noVal = ""
        Case vbLong:            noVal = CLng(0)
        Case vbVariant:         noVal = Empty
        
        Case vbInteger:         noVal = CInt(0)
        Case vbSingle:          noVal = CSng(0)
        Case vbDouble:          noVal = CDbl(0)
        Case vbDecimal:         noVal = CDec(0)
        Case vbCurrency:        noVal = CCur(0)
        Case vbBoolean:         noVal = CBool(0)
        Case vbByte:            noVal = CByte(0)
        
        Case vbEmpty:           noVal = Empty
        Case vbNull:            noVal = Null
        
        Case vbObject:          Set noVal = Nothing
        
        Case vbDate:            Stop 'noVal = Empty
        Case vbError:           Stop 'noVal = Empty
        Case vbDataObject:      Stop 'noVal = Empty
        Case vbUserDefinedType: Stop 'noVal = Empty
    End Select
    End If
End Function

Public Function pt3d( _
    Optional d0 As Double = 0#, _
    Optional d1 As Double = 0#, _
    Optional d2 As Double = 0# _
) As Double()
    Dim rt(2) As Double
    
    rt(0) = d0
    rt(1) = d1
    rt(2) = d2
    
    pt3d = rt
End Function

Public Function sort3dimsUp( _
    d0 As Double, d1 As Double, d2 As Double _
) As Double()
    Dim rt() As Double
    
    If d1 < d0 Then
        rt = sort3dimsUp(d1, d0, d2)
    ElseIf d2 < d1 Then
        rt = sort3dimsUp(d2, d0, d1)
    Else
        ReDim rt(2)
        rt(0) = d0
        rt(1) = d1
        rt(2) = d2
    End If
    sort3dimsUp = rt
End Function

Public Function sort3dimsDn( _
    d0 As Double, d1 As Double, d2 As Double _
) As Double()
    Dim rt() As Double
    
    If d1 > d0 Then
        rt = sort3dimsDn(d1, d0, d2)
    ElseIf d2 > d1 Then
        rt = sort3dimsDn(d2, d0, d1)
    Else
        ReDim rt(2)
        rt(0) = d0
        rt(1) = d1
        rt(2) = d2
    End If
    sort3dimsDn = rt
End Function

Public Function aiBoxDims( _
    RefBox As Inventor.Box _
) As Double()
    Dim rt() As Double
    Dim mx As Inventor.Point
    Dim mn As Inventor.Point
    
    With RefBox
        Set mx = .MaxPoint
        Set mn = .MinPoint
    End With
    
    rt(0) = mx.X - mn.X
    rt(1) = mx.Y - mn.Y
    rt(2) = mx.Z - mn.Z
    
    aiBoxDims = rt
End Function

Public Function aiBoxSortDown( _
    RefBox As Inventor.Box _
) As Inventor.Box
    Dim rt As Inventor.Box
    Dim mx As Inventor.Point
    Dim mn As Inventor.Point
    
    With ThisApplication.TransientGeometry
        Set rt = .CreateBox()
        With RefBox
            Set mx = .MaxPoint
            Set mn = .MinPoint
            rt.PutBoxData pt3d(), _
            sort3dimsDn( _
                mx.X - mn.X, _
                mx.Y - mn.Y, _
                mx.Z - mn.Z _
            )
        End With
    End With
    
    Set aiBoxSortDown = rt
End Function

Public Function dcSteelType2Spec6() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    With rt
        .Add "Steel, Mild", "MS"
        .Add "Stainless Steel", "SS"
        .Add "Stainless Steel, Austenitic", "SS"
    End With
    
    Set dcSteelType2Spec6 = rt
End Function

Public Function steelSpec6( _
    stl As String, Optional Ask As Long = 0 _
) As String
    Dim rt As String
    'With dcSteelType2Spec6()
    '    If .Exists(stl) Then
    '        steelSpec6 = .Item(stl)
    '    Else
    '        steelSpec6 = ""
    '    End If
    'End With
    
    Select Case stl
    Case "Stainless Steel"
        rt = "SS"
    Case "Stainless Steel, Austenitic"
        rt = "SS"
    Case "Stainless Steel 304"
        rt = "SS"
    Case "Steel, Mild"
        rt = "MS"
    Case "Rubber"
        rt = ""  'LG
    Case "Rubber, Silicone"
        rt = ""  'LG
    Case "UHMW, White"
        rt = ""  'LG
    Case Else
        If Ask Then
            Debug.Print "=== UNKNOWN MATERIAL ==="
            Debug.Print "   (" & stl & ")"
            Debug.Print "Please supply a code for Specification 6,"
            Debug.Print "if applicable, on the line below, and"
            Debug.Print "press [ENTER] or [RETURN] to modify."
            Debug.Print "Press [F5] when ready to continue."
            Debug.Print "rt  = """" '<-( place code between double quotes )"
            Stop
        Else
            rt = ""
        End If
    End Select
    
    steelSpec6 = rt
End Function

Public Function d0g1f3( _
    rf As Inventor.PartDocument, dm() As Double, _
    Optional dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    ''
    Dim aspect As Double
    Dim offSqr As Double
    Dim length As Double
    ''
    Dim rmType As String
    Dim rmSpc6 As String
    Dim rmItem As String
    Dim rmUnit As String
    Dim rmQty As String
    
    If dc Is Nothing Then
        Set rt = d0g1f3(rf, dm, _
            New Scripting.Dictionary _
        )
    Else
        Set rt = dc
        
        aspect = dm(1) / dm(0)
        offSqr = aspect - 1#
        length = dm(2)
        rmType = rf.PropertySets(gnDesign).Item(pnMaterial).Value
        rmSpc6 = steelSpec6(rmType)
        
        'Debug.Print ".Add """ & rmType & """, """""
        Debug.Print "Material: " _
            & rmType & " (" _
            & rmSpc6 & ")"
        '''
        Debug.Print "Cross Section: " _
            & Format$(dm(1), "0.000") _
            & " X " _
            & Format$(dm(0), "0.000") _
        '''
        Debug.Print "Length: " _
            & Format$(dm(2), "0.000") _
        '''
        
        If offSqr < 0.01 Then
            Debug.Print "Likely Square or Round"
            'Stop 'probably square or round
        ElseIf offSqr > 20 Then
            Debug.Print "Likely Sheet or Plate"
            'Stop 'might be SOME sort of sheet/plate
        Else
            Debug.Print "Likely Rectangular, Uneven?"
            'Stop 'likely rectangular
        End If
        
        rmItem = rmType & "???" 'User has to help from info given
        rmQty = Format$(dm(2), "0.000")
        rmUnit = "IN"
        
        dc.Add "RM", rmItem
        dc.Add "RMQTY", rmQty
        dc.Add "RMUNIT", rmUnit
        
        Stop
    End If
    Set d0g1f3 = rt
End Function

Public Function d0g1f4( _
    cd As Inventor.SheetMetalComponentDefinition _
) As Scripting.Dictionary
    Dim sb As Inventor.SurfaceBody
    Dim fc As Inventor.Face
    Dim pt As Inventor.Point
    
    Dim rkMgr As Inventor.ReferenceKeyManager
    Dim kyContx As Long
    Dim kyBytes() As Byte
    Dim kyLabel As String
    
    Dim d0 As Scripting.Dictionary
    Dim d1 As Scripting.Dictionary
    Dim ky As Variant
    Dim k1 As Variant
    
    Set d0 = New Scripting.Dictionary
    Set d1 = New Scripting.Dictionary
    
    Set rkMgr = aiDocument( _
        cd.Document _
    ).ReferenceKeyManager
    
    kyContx = rkMgr.CreateKeyContext
    
    For Each sb In cd.SurfaceBodies
        For Each fc In sb.Faces
            fc.GetReferenceKey kyBytes, kyContx
            kyLabel = rkMgr.KeyToString(kyBytes)
            d1.Add kyLabel, fc '.InternalName
            
            With fc.Evaluator
                Debug.Print TypeName(fc.Geometry) & _
                "(" & CStr(0 _
                    + IIf(.IsExtrudedShape, 1, 0) _
                    + IIf(.IsRevolvedShape, 2, 0) _
                ) & ")"
                With .RangeBox
                    d0.Add d0.Count, .MaxPoint
                    d0.Add d0.Count, .MinPoint
                End With
            End With
            'If 0 Then
            'ElseIf fc.SurfaceType = kPlaneSurface Then
            '''
            'ElseIf fc.SurfaceType = kCylinderSurface Then
            'ElseIf fc.SurfaceType = kConeSurface Then
            '''
            'ElseIf fc.SurfaceType = kSphereSurface Then
            'ElseIf fc.SurfaceType = kTorusSurface Then
            '''
            'ElseIf fc.SurfaceType = kBSplineSurface Then
            'ElseIf fc.SurfaceType = kEllipticalCylinderSurface Then
            'ElseIf fc.SurfaceType = kEllipticalConeSurface Then
            '''
            'ElseIf fc.SurfaceType = kUnknownSurface Then
            'Else
            'End If
        Next
        
        For Each ky In d0.Keys
            Set pt = d0.Item(ky)
            For Each k1 In d1.Keys 'fc In sb.Faces
                Set fc = d1.Item(k1)
                'If d1.Exists(k1) Then
                If Not fc.Evaluator.RangeBox.Contains(pt) Then
                    d1.Remove k1 'fc.InternalName
                End If
                'End If
            Next
            Stop
        Next
    Next
End Function

Public Function d0g1f5( _
    sb As Inventor.SurfaceBody _
) As Scripting.Dictionary
    Dim fc As Inventor.Face
    Dim pt As Inventor.Point
    
    Dim dFc As Scripting.Dictionary
    Dim dPt As Scripting.Dictionary
    Dim kPt As Variant
    
    Set dFc = New Scripting.Dictionary
    Set dPt = New Scripting.Dictionary
    
    For Each fc In sb.Faces
        dFc.Add fc.InternalName, fc
        With fc.Evaluator
            Debug.Print TypeName(fc.Geometry) & _
            "(" & CStr(0 _
                + IIf(.IsExtrudedShape, 1, 0) _
                + IIf(.IsRevolvedShape, 2, 0) _
            ) & ")"
            With .RangeBox
                dPt.Add dPt.Count, .MaxPoint
                dPt.Add dPt.Count, .MinPoint
            End With
        End With
    Next
    
    For Each kPt In dPt.Keys
        Set pt = dPt.Item(kPt)
        For Each fc In sb.Faces
            If dFc.Exists(fc.InternalName) Then
                If Not fc.Evaluator.RangeBox.Contains(pt) Then
                    dFc.Remove fc.InternalName
                    If dFc.Count = 0 Then Stop
                End If
            End If
        Next
        Stop
    Next
    
    Set d0g1f5 = dFc
End Function

Public Function d0g1f6( _
    fc As Inventor.Face _
) As String
    Dim kyBytes() As Byte
    Dim kyContx As Long
    Dim rt As String
    
    With aiDocument( _
        fc.SurfaceBody.ComponentDefinition.Document _
    ).ReferenceKeyManager '.CreateKeyContext
        kyContx = .CreateKeyContext
        fc.GetReferenceKey kyBytes, kyContx
        rt = .KeyToString(kyBytes)
    End With
    d0g1f6 = rt
End Function

Public Function aiPoint(ob As Object) As Inventor.Point
    If TypeOf ob Is Inventor.Point Then
        Set aiPoint = ob
    Else
        Set aiPoint = Nothing
    End If
End Function

''' d0g2: Testing
'''
'''

Public Function d0g2f1()
    ''' Verify 3-way sorting function sort3dimsUp
    Dim ck() As Double
    ck = sort3dimsUp(2, 3, 5): Stop
    ck = sort3dimsUp(2, 5, 3): Stop
    ck = sort3dimsUp(3, 2, 5): Stop
    ck = sort3dimsUp(3, 5, 2): Stop
    ck = sort3dimsUp(5, 2, 3): Stop
    ck = sort3dimsUp(5, 3, 2): Stop
End Function

Public Function d0g2f2()
    ''' Testing new spec pickup system
    Dim ky As Variant
    
    With dcAiDocComponents( _
        ThisApplication.ActiveDocument, , 0 _
    )
        For Each ky In .Keys
            Debug.Print ky
            Set .Item(ky) = d0g0f0(aiCompDefShtMetal( _
                aiCompDefOf(aiDocPart(.Item(ky))) _
            ))
            If .Item(ky) Is Nothing Then
                .Remove ky
            Else
                'Stop
            End If
        Next
    End With
End Function

Public Function d0g2f3()
    ''' Checking some behaviors
    ''' on string arrays
    ''' vs variants
    Dim ky As Variant
    With New aiPropSetter
        Debug.Print Join(.PropList(), "|")
        For Each ky In .PropList()
            Debug.Print ky
        Next
    End With
End Function

Public Function d0g2f4(dc As Scripting.Dictionary) As Scripting.Dictionary
    '''
    ''' Return Dictionary of ALLOCATED Property
    ''' Values (True/False) attached to all components
    ''' and subcomponents of the active Document.
    '''
    ''' Where the ALLOCATED Property is not present,
    ''' represent it as "<default>"
    '''
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            'Debug.Print ky
            With aiDocument(.Item(ky)).PropertySets.Item(gnCustom)
                On Error Resume Next
                Err.Clear
                Set pr = .Item("ALLOCATED")
                If Err.Number = 0 Then
                    rt.Add ky, CStr(pr.Value) & "|" & ky
                Else
                    rt.Add ky, "<default>" & "|" & ky
                End If
                On Error GoTo 0
            End With
        Next
    End With
    Set d0g2f4 = rt
End Function
'Debug.Print Join(d0g2f4(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbNewLine)

Public Function d0g2f6( _
    dc As Scripting.Dictionary, pn As String, _
    Optional gn As String = gnCustom, _
    Optional df As String = "<NOPROP>" _
) As Scripting.Dictionary
    '''
    ''' Return Dictionary of named Property Values
    ''' attached to all Inventor Documents
    ''' in supplied Dictionary.
    '''
    '''
    ''' Where the ALLOCATED Property is not present,
    ''' represent it as "<default>"
    '''
    Dim ad As Inventor.Document
    Dim rt As Scripting.Dictionary
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set ad = aiDocument(.Item(ky))
            If ad Is Nothing Then
            Else
            End If
            With ad.PropertySets.Item(gn)
                On Error Resume Next
                Err.Clear
                Set pr = .Item("ALLOCATED")
                If Err.Number = 0 Then
                    rt.Add ky, CStr(pr.Value) & "|" & ky
                Else
                    rt.Add ky, "<default>" & "|" & ky
                End If
                On Error GoTo 0
            End With
        Next
    End With
    Set d0g2f6 = rt
End Function
'Debug.Print Join(d0g2f6(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbNewLine)

Public Function d0g2f5(dc As Scripting.Dictionary) As Scripting.Dictionary
    '''
    ''' Attempt to "transpose" contents of Dictionary
    ''' and return a dictionary of Items mapped
    ''' to sub-Dictionaries containing all keys
    ''' which mapped to each value
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim kv As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            kv = d0g2f5a(.Item(ky))
            With rt
                If Not .Exists(kv) Then
                    .Add kv, New Scripting.Dictionary
                End If
                
                With dcOb(.Item(kv))
                    .Add .Count, ky
                End With
            End With
        Next
    End With
End Function

Public Function d0g2f5a(vr As Variant) As Variant
    '''
    ''' Return any Variant that is NOT an Object
    ''' Object handling MAY be addressed later.
    '''
    If IsObject(vr) Then
        Stop
    Else
        d0g2f5a = vr
    End If
End Function

Public Function dcMaterialUsage( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim pt As Inventor.PartDocument
    Dim ky As Variant
    Dim pn As String
    Dim mt As String
    
    Set rt = New Scripting.Dictionary
    For Each ky In dc
        Set pt = aiDocPart(obOf(dc.Item(ky)))
        If pt Is Nothing Then
            'do Nothing
        Else
            pn = pt.PropertySets(gnDesign).Item(pnPartNum).Value
            mt = pt.PropertySets(gnDesign).Item(pnMaterial).Value
            With rt
                If .Exists(mt) Then
                    .Item(mt) = .Item(mt) & vbNewLine & vbTab & pn
                Else
                    .Add mt, mt & vbNewLine & vbTab & pn
                End If
            End With
            Set pt = Nothing
        End If
    Next
    Set dcMaterialUsage = rt
End Function
'lsDump dcMaterialUsage(dcAiDocsOfType(kPartDocumentObject, dcAiDocComponents(aiDocActive()))).Items

Public Function d0g3f0() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim mt As Inventor.Asset
    
    Set rt = New Scripting.Dictionary
    With ThisApplication.ActiveMaterialLibrary
        For Each mt In .MaterialAssets
            rt.Add mt.DisplayName, mt
        Next
    End With
    Set d0g3f0 = rt
End Function
'lsDump d0g3f0().Keys

Public Function dcGrpByPtNum( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    ''
    ''  Returns Dictionary of Dictionaries
    ''  grouping Inventor Documents in
    ''  supplied Dictionary by their Part
    ''  Numbers.
    ''
    ''  Ideally, each Document's Part Number
    ''  should be unique, and each sub Dictionary
    ''  should contain only one Document, however,
    ''  it is possible for more than one Document
    ''  to have the same Part Number.
    ''
    ''  By returning a Dictionary of Dictionaries,
    ''  this function provides a way for the client
    ''  to detect and respond to any conflicts.
    ''
    Dim rt As Scripting.Dictionary
    Dim pt As Inventor.Document
    Dim ky As Variant
    Dim pn As String
    Dim dn As String
    
    Set rt = New Scripting.Dictionary
    With dc: For Each ky In .Keys
        Set pt = aiDocument(.Item(ky))
        dn = pt.FullDocumentName
        pn = CStr(aiDocPropVal( _
            pt, pnPartNum, gnDesign _
        ))
        With rt
            If Not .Exists(pn) Then
                .Add pn, New Scripting.Dictionary
            End If
            
            With dcOb(.Item(pn))
                If .Exists(dn) Then
                    Stop 'because something went wrong
                    'should NOT get same Document twice!
                Else
                    .Add dn, pt
                End If
            End With
        End With
        Debug.Print ;
    Next: End With
    
    Set dcGrpByPtNum = rt
'lsDump               dcGrpByPtNum(dcAiDocComponents(aiDocActive())).Keys
'Debug.Print txDumpLs(dcGrpByPtNum(dcAiDocComponents(aiDocActive())).Keys)
End Function

Public Function dcRemapByPtNum( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    ''  Returns Dictionary of Inventor
    ''  Documents keyed on Part Number
    Dim rt As Scripting.Dictionary
    Dim xt As Scripting.Dictionary
    Dim pt As Inventor.Document
    Dim ky As Variant
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    Set xt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set pt = aiDocument(.Item(ky))
            pn = CStr(aiDocPropVal( _
                pt, pnPartNum, gnDesign _
            ))
'''
'''
'''
            'If pt.DocumentType = kPartDocumentObject Then
            '    ''' UPDATE[2021.06.22]
            '    ''' moving hardware component check outside of
            '    ''' and before collision check in order to skip
            '    ''' known hardware items entirely.
            '    ''' '
            '    ''' UPDATE[2021.06.21]
            '    ''' implementing a set of checks for hardware components
            '    ''' probably want to move outside Dictionary check
            '    ''' to catch hardware elements before they're added.
            '    ''' Could lead to trouble if this prevents new hardware
            '    ''' Items from being added to Genius, but don't believe
            '    ''' this is a high risk.
            '    If aiDocPart(pt).ComponentDefinition.IsContentMember Then
            '        'it's commodity hardware
            '        With cnGnsDoyle().Execute( _
            '            "select Family from vgMfiItems where Item = '" _
            '            & pn & "';" _
            '        )
            '            If .BOF And .EOF Then
            '                'probably not in Genius
            '                'keep it -- may need added
            '            ElseIf Split(.GetString( _
            '                adClipString, , "", vbVerticalTab _
            '            ), vbVerticalTab)(0) = "D-HDWR" Then
            '                    Set pt = Nothing 'and move on
            '                Else
            '                    Debug.Print ; 'Breakpoint Landing
            '                    'Stop
            '                End If
            '            End If
            '        End With
            '    ElseIf pt.PropertySets.Item(gnDesign).Item(pnFamily).Value = "D-HDWR" Then
            '        'it's in commodity hardware family
            '        Set pt = Nothing 'and move on
            '    ElseIf InStr(1, "|D-HDWR|D-PTS|R-PTS|", "|" & pt.PropertySets.Item(gnDesign).Item(pnFamily).Value & "|") > 0 Then
            '        'it's PROBABLY hardware
            '        'but keep it, just in case
            '        Debug.Print ; 'Breakpoint Landing
            '    Else 'nothing special to worry about, probably
            '        Debug.Print ; 'Breakpoint Landing
            '        'Stop
            '    End If
            'Else 'we've got an Assembly
            '    Debug.Print ; 'Breakpoint Landing
            '    'Stop
            'End If
'''
'''
'''
            If Len(pn) > 0 Then
                With rt
                    If .Exists(pn) Then 'we have Key collsion
                        With xt 'report' it here
                            If Not .Exists(pn) Then
                                .Add pn, New Scripting.Dictionary
                            End If
                            
                            With dcOb(.Item(pn))
                                .Add pt.FullDocumentName, pt
                            End With
                        End With
                    Else
                        .Add pn, pt
                    End If
                End With
                Debug.Print ;
            Else
                Debug.Print InputBox("This component has no part number:" & vbNewLine & pt.DisplayName & vbNewLine & CStr(aiDocPropVal(pt, pnDesc, gnDesign)) & vbNewLine & vbNewLine & "Copy file path from text box for later review.", pt.DisplayName, pt.FullDocumentName)
                If getFromClipBdWin10() = pt.FullDocumentName Then
                    'Stop
                Else
                    If MsgBox("Are you sure you want to continue" & vbNewLine & "without recording this file path?", vbExclamation + vbYesNo, "File Path not copied!") = vbNo Then
                        Stop
                    End If
                End If
            End If
        Next
    End With
    
    If xt.Count > 0 Then
        Debug.Print MsgBox( _
            Join(Array( _
                "The following Part Numbers are", _
                "assigned to more than one Model:", _
                "", _
                vbTab & Join(xt.Keys, vbNewLine & vbTab), _
                "" _
            ), vbNewLine), _
            vbOKOnly Or vbInformation, _
            "Duplicate Part Numbers!" _
        )
    End If
    
    Set dcRemapByPtNum = rt
'lsDump               dcRemapByPtNum(dcAiDocComponents(aiDocActive())).Keys
'Debug.Print txDumpLs(dcRemapByPtNum(dcAiDocComponents(aiDocActive())).Keys)
End Function

Public Function dcRemapByFilePath( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    ''  Returns Dictionary of Inventor
    ''  Documents re-keyed to File Path.
    ''  Typically for a Dictionary
    ''  previously remapped to another
    ''  key (most likely Part Number)
    Dim rt As Scripting.Dictionary
    Dim pt As Inventor.Document
    Dim ky As Variant
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set pt = aiDocument(.Item(ky))
            pn = pt.FullDocumentName
            rt.Add pn, pt
        Next
    End With
    Set dcRemapByFilePath = rt
'lsDump dcRemapByFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys
'send2clipBd txDumpLs(dcRemapByFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys)
End Function

Public Function dcRemapByPtNumFilePath( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    ''  Returns Dictionary of Inventor
    ''  Documents keyed on Part Number
    ''  combined with original key
    ''  (which SHOULD be full doc path)
    Dim rt As Scripting.Dictionary
    Dim pt As Inventor.Document
    Dim ky As Variant
    Dim pn As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set pt = aiDocument(.Item(ky))
            pn = CStr(aiDocPropVal( _
                pt, pnPartNum, gnDesign _
            )) & vbTab & pt.FullDocumentName 'ky
            rt.Add pn, pt
        Next
    End With
    Set dcRemapByPtNumFilePath = rt
'lsDump dcRemapByPtNumFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys
'send2clipBd txDumpLs(dcRemapByPtNumFilePath(dcAiDocComponents(aiDocActive(), , 1)).Keys)
End Function

Public Function dcGeniusItems() As Scripting.Dictionary
    ''  Generates Dictionary of Items in Genius
    ''  Formerly d0g3f2
    Dim rt As Scripting.Dictionary
    Dim ky As ADODB.Field
    Dim vl As ADODB.Field
    
    Set rt = New Scripting.Dictionary
    With cnGnsDoyle()
        With .Execute("select Item, ItemID from vgMfiItems")
            Set ky = .Fields("Item")
            Set vl = .Fields("ItemID")
            Do Until .EOF Or .BOF
                rt.Add ky.Value, vl.Value
                .MoveNext
            Loop
        End With
    End With
    Set dcGeniusItems = rt
End Function

Public Function d0g3f3( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary()
    ''  Appears intended to separate
    ''  parts in Genius from those
    ''  not there yet. I don't think
    ''  this one was working quite
    ''  right yet. New kyPick system
    ''  should handle this properly,
    ''  now, in any case.
    Dim ky As Variant
    Dim rt(1) As Scripting.Dictionary
    
    Set rt(0) = New Scripting.Dictionary
    Set rt(1) = New Scripting.Dictionary
    With dcGeniusItems()
        For Each ky In dc
            If .Exists(ky) Then
                rt(1).Add ky, .Item(ky)
            Else
                rt(0).Add ky, ""
            End If
        Next
    End With
    d0g3f3 = rt
End Function

Public Function deConstrainAssyComponent(co As Inventor.ComponentOccurrence) As Long
    ''  Deletes all constraints on an occurrence
    ''  !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
    Dim cs As Inventor.AssemblyConstraint
    Dim ct As Long
    
    ct = 0
    For Each cs In co.Constraints
        cs.Delete
        ct = ct + 1
    Next
    
    deConstrainAssyComponent = ct
End Function

Public Function deConstrainAssyDocument(ad As Inventor.AssemblyDocument) As Long
    ''  Calls deConstrainAssyComponent over all occurrences
    ''  in an assembly to remove all their constraints
    ''  !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
    ''  !!!That goes DOUBLE for THIS function!!!
    Dim co As Inventor.ComponentOccurrence
    Dim ct As Long
    
    ct = 0
    For Each co In ad.ComponentDefinition.Occurrences
        ct = ct + deConstrainAssyComponent(co)
    Next
    
    deConstrainAssyDocument = ct
End Function

Public Function dcPartsInGeniusOrNot() As Scripting.Dictionary
    Dim dcInGns As Scripting.Dictionary
    Dim dcNotIn As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    Set dcInGns = New Scripting.Dictionary
    Set dcNotIn = dcRemapByPtNum(dcAiDocComponents(aiDocActive()))
    'Set dcNotIn = dcAiDocComponents(aiDocActive())
    'Set dcNotIn = dcAssyDocsByPtNum(aiDocActive())
    With dcFrom2Fields(cnGnsDoyle().Execute( _
        "select Item from vgMfiItems" _
    ), "Item", "Item")
        For Each ky In .Keys
            With dcNotIn
                If .Exists(ky) Then
                    dcInGns.Add ky, .Item(ky)
                    .Remove ky
                End If
            End With
        Next
    End With
    'Stop
    
    rt.Add "INGNS", dcInGns
    rt.Add "NOTIN", dcNotIn
    Set dcPartsInGeniusOrNot = rt
End Function

Public Function d0g4f1() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim cn As ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Set rt = New Scripting.Dictionary
    Set cn = cnGnsDoyle()
    Set rs = cn.Execute("select Item from vgMfiItems")
    Set d0g4f1 = rt
End Function

Public Function compOccFromProxy( _
    oc As Inventor.ComponentOccurrence _
) As Inventor.ComponentOccurrence
    Dim px As Inventor.ComponentOccurrenceProxy
    
    'If TypeOf oc Is Inventor.ComponentOccurrenceProxy Then
        'Set px = oc
        'Stop
        'Set compOccFromProxy = compOccFromProxy(px.ContainingOccurrence)
    'Else
        Set compOccFromProxy = oc
    'End If
End Function

Public Function nuPicker( _
    Optional ob As kyPick = Nothing _
) As kyPick
    If ob Is Nothing Then
        Set nuPicker = New kyPick
    Else
        Set nuPicker = ob
    End If
End Function

Public Function nuSplitter( _
    Optional ob As dcSplitter = Nothing _
) As dcSplitter
    If ob Is Nothing Then
        Set nuSplitter = New dcSplitter
    Else
        Set nuSplitter = ob
    End If
End Function

Public Function dcAiPartDocsWithRMv0( _
    dcIn As Scripting.Dictionary, _
    Optional WantOut As Long = 0 _
) As Scripting.Dictionary
    Dim ky As Variant
    Dim rt(1) As Scripting.Dictionary
    
    If WantOut < 0 Or WantOut > 1 Then
        Set dcAiPartDocsWithRMv0 = dcAiPartDocsWithRMv0(dcIn, 1)
    Else
        'With nuSplitter().WithSel(New kyPickAiDocWithRM)
        With New kyPickAiDocWithRM
            Set rt(0) = .dcIn()
            Set rt(1) = .dcOut()
            For Each ky In dcIn.Keys
                With .dcFor(dcIn.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add ky, dcIn.Item(ky)
                    End If
                End With
            Next
        End With
        Set dcAiPartDocsWithRMv0 = rt(WantOut)
    End If
End Function
'Debug.Print txDumpLs(dcAiPartDocsWithRMv0(dcAiDocComponents(aiDocActive()), 1).Keys)

Public Function kyScanned( _
    dcIn As Scripting.Dictionary, _
    Optional pkr As kyPick = Nothing _
) As kyPick
    If pkr Is Nothing Then
        Set kyScanned = kyScanned(dcIn, New kyPick)
    Else
        Set kyScanned = pkr.AfterScanning(dcIn)
    End If
End Function
'Debug.Print txDumpLs(kyScanned(dcAiDocComponents(aiDocActive()), New kyPickAiPartVsAssy).dcIn().Keys)

Public Function dcAiDocsPicked( _
    dcIn As Scripting.Dictionary, _
    Optional pkr As kyPick = Nothing, _
    Optional WantOut As Long = 0 _
) As Scripting.Dictionary
    Dim ky As Variant
    Dim rt(1) As Scripting.Dictionary
    
    If pkr Is Nothing Then
        Set dcAiDocsPicked = dcAiDocsPicked(dcIn, New kyPick, WantOut)
    ElseIf WantOut < 0 Or WantOut > 1 Then
        Set dcAiDocsPicked = dcAiDocsPicked(dcIn, pkr, 1)
    Else
        With pkr
            Set rt(0) = .dcIn()
            Set rt(1) = .dcOut()
            For Each ky In dcIn.Keys
                With .dcFor(dcIn.Item(ky))
                    If .Exists(ky) Then
                        Stop
                    Else
                        .Add ky, dcIn.Item(ky)
                    End If
                End With
            Next
        End With
        Set dcAiDocsPicked = rt(WantOut)
    End If
End Function
'Debug.Print txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), 1).Keys)
'Debug.Print txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), New kyPickAiDocContentCtr, 0).Keys)

Public Function dcAiPartDocsWithRM( _
    dcIn As Scripting.Dictionary, _
    Optional WantOut As Long = 0 _
) As Scripting.Dictionary
    Set dcAiPartDocsWithRM = dcAiDocsPicked(dcAiDocsPicked(dcIn, _
        New kyPickAiPartVsAssy, 0), _
        New kyPickAiDocWithRM, WantOut)
End Function
'Debug.Print txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)

Public Function d0g5f0()
    '
End Function

Public Function dcAiDocGrpsByForm( _
    dcIn As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcAiDocGrpsByForm -- Separate a Dictionary
    '''     of Inventor Documents into
    '''     categorical sub-Dictionaries
    '''     according to various criteria:
    '''     -   PRCH    Purchased Items
    '''     -   ASSY    Assemblies
    '''     -   HDWR    Hardware (Content Center)
    '''     -   DBAR    Structural Parts (was BSTK)
    '''                 Subtype NOT Sheet Metal
    '''                 (some "Sheet Metal" Parts
    '''                  might also technically
    '''                  belong here. see below)
    '''     -   MAYB    Likely Structural Parts
    '''                 Sheet Metal subtype, but
    '''                 has either no flat pattern,
    '''                 or an invalid one.
    '''     -   SHTM    Sheet Metal Parts
    '''                 Indicated both by Subtype
    '''                 and presence of a valid
    '''                 flat pattern.
    '''
    '''     Presence in Genius, a distinction
    '''     originally intended to be made here,
    '''     is now planned to be made to a separate
    '''     Dictionary, possibly also subcategorized,
    '''     to be processed in conjunction with
    '''     the results of this function.
    '''
    '''     The notion of passing different subgroups
    '''     of this Dictionary to separate handlers
    '''     for more specialized processing, while
    '''     still an option, is no longer considered
    '''     its primary role. Instead, the set is
    '''     expected to be used in a form application
    '''     which will present the various groups to
    '''     the user for review, and modification as
    '''     desired or necessary.
    '''
    '''     REV[2022.03.08.1212] All new text
    '''     in function description above.
    '''     see notes_2022-0308_general-01.txt
    '''     for prior description
    '''
    Dim rt As Scripting.Dictionary
    'Dim pkGns As kyPick
    Dim pkBuy As kyPick
    Dim pkPrt As kyPick
    Dim pkCtC As kyPick
    Dim pkSht As kyPick
    Dim pkMbe As kyPick
    
    Set rt = New Scripting.Dictionary
    
    ''' REV[2022.03.08.1112]
    '''     Disabled split on presence
    '''     in Genius. Believe better
    '''     addressed separately
    ''  separate items already in Genius
    ''  from those not yet in
    'Set pkGns = nuPicker( _
    '    New kyPickInGenius _
    ').AfterScanning(dcIn)
    ''  NOTE: no further processing
    ''  implemented on this yet
    ''  MIGHT be better applied
    ''  at a different stage?
    
    ''' REV[2022.03.08.1115]
    '''     Add division on Purchased Parts
    '''     with "out" Dictionary replacing
    '''     main for Part/Assy separation.
    Set pkBuy = nuPicker( _
        New kyPickAiDocPurchased _
    ).AfterScanning(dcIn)
    
    With pkBuy
        rt.Add "PRCH", .dcIn
        
        ''  separate parts from assemblies
        Set pkPrt = nuPicker( _
            New kyPickAiPartVsAssy _
        ).AfterScanning(.dcOut)
    End With
    
    With pkPrt
        rt.Add "ASSY", .dcOut
    
        ''  isolate Content Center
        ''  parts from the rest
        Set pkCtC = nuPicker( _
            New kyPickAiDocContentCtr _
        ).AfterScanning(.dcIn)
    End With
    
    With pkCtC
        rt.Add "HDWR", .dcIn
        
        ''  separate (potential) sheet
        ''  metal parts from non-sheet
        Set pkSht = nuPicker( _
            New kyPickAiSheetMetal _
        ).AfterScanning(.dcOut)
    End With
    
    With pkSht
        rt.Add "DBAR", .dcOut
        
        Set pkMbe = nuPicker( _
            New kyPickAiShMtl4sure _
        ).AfterScanning(.dcIn)
    End With
    
    With pkMbe
        rt.Add "SHTM", .dcIn
        rt.Add "MAYB", .dcOut
    End With
    
    Debug.Print ; 'Breakpoint Landing
    
    Set dcAiDocGrpsByForm = rt
'send2clipBd Join(Array("#If False Then", ConvertToJson(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive())), vbTab), "#End If"), vbNewLine)
'send2clipBd ConvertToJson(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive())), vbTab)
'Debug.Print dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive()))
'Debug.Print txDumpLs(pkCtC.dcIn.Keys)
'Debug.Print dcAiDocGrpsByForm(dcAssyDocsByPtNum(aiDocActive()))
'Debug.Print txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)
End Function

Public Function dcAiDocGrpsByFormAndIfac( _
    dcIn As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcAiDocGrpsByFormAndIfac -- Separate a Dictionary
    '''     of Inventor Documents into
    '''     categorical sub-Dictionaries
    '''     according to various criteria:
    '''     -   PRCH    Purchased Items
    '''     -   ASSY    Assemblies
    '''     -   IASM    iAssembly Factories
    '''     -   IPRT    iPart Factories
    '''     -   HDWR    Hardware (Content Center)
    '''     -   DBAR    Structural Parts (was BSTK)
    '''                 Subtype NOT Sheet Metal
    '''                 (some "Sheet Metal" Parts
    '''                  might also technically
    '''                  belong here. see below)
    '''     -   MAYB    Likely Structural Parts
    '''                 Sheet Metal subtype, but
    '''                 has either no flat pattern,
    '''                 or an invalid one.
    '''     -   SHTM    Sheet Metal Parts
    '''                 Indicated both by Subtype
    '''                 and presence of a valid
    '''                 flat pattern.
    '''
    ''' REV[2023.01.24.0921]
    ''' copied from dcAiDocGrpsByForm to produce
    ''' new variant with additional groupings for
    ''' iPart (IPRT) and iAssembly (IASM) members.
    ''' additional groups for their corresponding
    ''' Factories will likely also be added.
    '''
    ''' text of prior REV[2022.03.08.1212] removed.
    ''' see dcAiDocGrpsByForm for that.
    '''
    Dim rt As Scripting.Dictionary
    Dim wk As Scripting.Dictionary
    ''' REV[2023.01.24.1156]
    ''' add working Dictionary
    ''' to collect iAssembly
    ''' and iPart Factories
    Dim ky As Variant
    Dim mb As Inventor.Document
    Dim md As Inventor.Document
    Dim fp As String
    
    'Dim pkGns As kyPick
    Dim pkBuy As kyPick
    Dim pkPrt As kyPick
    Dim pkCtC As kyPick
    Dim pkSht As kyPick
    Dim pkMbe As kyPick
    
    ''' REV[2023.01.24.1009]
    ''' add new pickers for
    ''' iAssemblies and iParts
    Dim pkIas As kyPick
    Dim pkIpt As kyPick
    
    Set rt = New Scripting.Dictionary
    
    ''' REV[2022.03.08.1112]
    '''     Disabled split on presence
    '''     in Genius. Believe better
    '''     addressed separately
    ''  separate items already in Genius
    ''  from those not yet in
    'Set pkGns = nuPicker( _
    '    New kyPickInGenius _
    ').AfterScanning(dcIn)
    ''  NOTE: no further processing
    ''  implemented on this yet
    ''  MIGHT be better applied
    ''  at a different stage?
    
    ''' REV[2022.03.08.1115]
    '''     Add division on Purchased Parts
    '''     with "out" Dictionary replacing
    '''     main for Part/Assy separation.
    Set pkBuy = nuPicker( _
        New kyPickAiDocPurchased _
    ).AfterScanning(dcIn)
    
    With pkBuy
        If dcIn.Count > 0 Then rt.Add "PRCH", .dcIn
        
        ''  separate parts from assemblies
        Set pkPrt = nuPicker( _
            New kyPickAiPartVsAssy _
        ).AfterScanning(.dcOut)
    End With
    
    With pkPrt
        ''  separate iAssembly members
        ''  from stand-alone assemblies
        Set pkIas = nuPicker( _
            New kyPickAiAssyMember _
        ).AfterScanning(.dcOut)
    
        ''  isolate Content Center
        ''  parts from the rest
        Set pkCtC = nuPicker( _
            New kyPickAiDocContentCtr _
        ).AfterScanning(.dcIn)
    End With
    
    With pkIas
        If .dcOut.Count > 0 Then rt.Add "ASSY", .dcOut
        
        With .dcIn
            Set wk = New Scripting.Dictionary
            
            For Each ky In .Keys
                With aiDocAssy(.Item(ky)).ComponentDefinition
                    Set mb = .Parent
                    With .iAssemblyMember.ParentFactory.Parent
                        Set md = .Parent
                        fp = md.FullDocumentName
                        
                        With wk
                            If Not .Exists(fp) Then
                                .Add fp, New Scripting.Dictionary
                                dcOb(.Item(fp)).Add "", md
                            End If
                            
                            With dcOb(.Item(fp))
                            If .Exists(ky) Then
                                Stop
                            Else
                                .Add ky, mb
                            End If
                            End With
                            
                        End With
                    End With
                End With
                'Debug.Print aiDocAssy(.Item(ky)).ComponentDefinition.iAssemblyMember.ParentFactory.Parent.Parent.FullDocumentName
                'Stop
            Next
            
            If wk.Count > 0 Then rt.Add "IASM", wk
        End With
    End With
    
    With pkCtC
        If .dcIn.Count > 0 Then rt.Add "HDWR", .dcIn
        
        ''  separate iPart members
        ''  from stand-alone parts
        Set pkIpt = nuPicker( _
            New kyPickAiPartMember _
        ).AfterScanning(.dcOut)
    End With
    
    With pkIpt
        With .dcIn
            Set wk = New Scripting.Dictionary
            
            For Each ky In .Keys
            With aiDocPart(.Item(ky)).ComponentDefinition
                Set mb = .Document '.Parent
                Set md = aiDocPart(.iPartMember.ParentFactory.Parent) ' .PropertySets.Parent ' .Parent
                With md
                    fp = md.FullDocumentName
                    
                    With wk
                        If Not .Exists(fp) Then
                            .Add fp, New Scripting.Dictionary
                            dcOb(.Item(fp)).Add "", md
                        End If
                        
                        With dcOb(.Item(fp))
                        If .Exists(ky) Then
                            Stop
                        Else
                            .Add ky, mb
                        End If
                        End With
                    End With
                End With
            End With: Next
            
            If wk.Count > 0 Then rt.Add "IPRT", wk
        End With
        
        ''  add iPart Factories to
        ''  Dictionary of non-Members
        With .dcOut: For Each ky In wk.Keys
            If .Exists(ky) Then
                Stop
            Else
                .Add ky, dcOb(wk.Item(ky)).Item("")
            End If
        Next: End With
        
        ''  separate (potential) sheet
        ''  metal parts from non-sheet
        Set pkSht = nuPicker( _
            New kyPickAiSheetMetal _
        ).AfterScanning(.dcOut)
    End With
    
    With pkSht
        If .dcOut.Count > 0 Then rt.Add "DBAR", .dcOut
        
        Set pkMbe = nuPicker( _
            New kyPickAiShMtl4sure _
        ).AfterScanning(.dcIn)
    End With
    
    With pkMbe
        If .dcIn.Count > 0 Then rt.Add "SHTM", .dcIn
        If .dcOut.Count > 0 Then rt.Add "MAYB", .dcOut
    End With
    
    Debug.Print ; 'Breakpoint Landing
    
    Set dcAiDocGrpsByFormAndIfac = rt
'send2clipBd Join(Array("#If False Then", ConvertToJson(dcAiDocGrpsByFormAndIfac(dcAiDocComponents(aiDocActive())), vbTab), "#End If"), vbNewLine)
'send2clipBd ConvertToJson(dcAiDocGrpsByFormAndIfac(dcAiDocComponents(aiDocActive())), vbTab)
'Debug.Print dcAiDocGrpsByFormAndIfac(dcAiDocComponents(aiDocActive()))
'Debug.Print txDumpLs(pkCtC.dcIn.Keys)
'Debug.Print dcAiDocGrpsByFormAndIfac(dcAssyDocsByPtNum(aiDocActive()))
'Debug.Print txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)
End Function

Public Function d0g5f2( _
    dcIn As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' function d0g5f2
    '''
    ''' INITIATED[2021.03.23]
    ''' this variant on dcAiDocGrpsByForm
    ''' is intended to separate items
    ''' in Genius from those not yet in
    ''' and purchased items from those
    ''' to be made, cross-referencing
    ''' the two to determine individual
    ''' needs for processing.
    '''
    ''' presently in a nonfunctional state
    ''' as End of Day approaches. will hope
    ''' to continue development tomorrow
    '''
    Dim rt As Scripting.Dictionary
    Dim pkGns As kyPick
    Dim pkPvA As kyPick
    Dim pkCtC As kyPick
    Dim pkSht As kyPick
    Dim pkBuy As kyPick
    
    Set rt = New Scripting.Dictionary
    
    ''  separate items already in Genius
    ''  from those not yet in
    Set pkGns = nuPicker( _
        New kyPickInGenius _
    ).AfterScanning(dcIn)
    
    ''  separate purchased items
    ''  from those to be made
    Set pkBuy = nuPicker( _
        New kyPickAiDocPurchased _
    ).AfterScanning(dcIn)
    ''  NOTE: no further processing
    ''  implemented on this yet
    ''  MIGHT be better applied
    ''  at a different stage?
    
    
    
    ''  separate parts from assemblies
    'Set pkPvA = nuPicker( _
    '    New kyPickAiPartVsAssy _
    ').AfterScanning(dcIn)
    ''rt.Add "ASSY", dck pkPvA.dcOut
    
    ''  isolate Content Center
    ''  parts from the rest
    'Set pkCtC = nuPicker( _
    '    New kyPickAiDocContentCtr _
    ').AfterScanning(pkPvA.dcIn)
    ''rt.Add "HDWR", pkCtC.dcIn
    
    ''  separate (potential)
    ''  sheet metal parts
    ''  from non-sheet
    'Set pkSht = nuPicker( _
    '    New kyPickAiSheetMetal _
    ').AfterScanning(pkCtC.dcOut)
    'rt.Add "SHTM", pkSht.dcIn
    'rt.Add "BSTK", pkSht.dcOut
    
    Debug.Print ; 'Breakpoint Landing
    
    Set d0g5f2 = rt
End Function

Public Function d0g5f3( _
    dcIn As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' function d0g5f3 -- essentially a recreation of dcAiDocGrpsByForm
    '''
    Dim rt As Scripting.Dictionary
    Dim pkBuy As kyPick
    Dim pkPrt As kyPick
    Dim pkCtC As kyPick
    Dim pkSht As kyPick
    Dim pkMbe As kyPick
    'Dim pkGns As kyPick
    'Dim pk___ As kyPick
    
    Set rt = New Scripting.Dictionary
    'Set pkGns = nuPicker( _
        New kyPickInGenius _
    ).AfterScanning(dcIn)
    
    Set pkBuy = nuPicker( _
        New kyPickAiDocPurchased _
    ).AfterScanning(dcIn)
    
    With pkBuy
        rt.Add "PRCH", .dcIn
        Set pkPrt = nuPicker( _
            New kyPickAiPartVsAssy _
        ).AfterScanning(.dcOut)
    End With
    
    With pkPrt
        rt.Add "ASSY", .dcOut
        Set pkCtC = nuPicker( _
            New kyPickAiDocContentCtr _
        ).AfterScanning(.dcIn)
    End With
    
    With pkCtC
        rt.Add "HDWR", .dcIn
        Set pkSht = nuPicker( _
            New kyPickAiSheetMetal _
        ).AfterScanning(.dcOut)
    End With
    
    With pkSht
        rt.Add "DBAR", .dcOut
        
        Set pkMbe = nuPicker( _
            New kyPickAiShMtl4sure _
        ).AfterScanning(.dcIn)
    End With
    
    With pkMbe
        rt.Add "MAYB", .dcOut
        rt.Add "SHTM", .dcIn
    End With
    
    Set d0g5f3 = rt
End Function

Public Function dcDepthAiDocGrp( _
    dc As Scripting.Dictionary _
) As Long
    Dim vr As Variant
    Dim ob As Object
    Dim mx As Long
    Dim dx As Long
    Dim ck As Long
    Dim rt As Long
    
    With dc
        mx = .Count
        
        If mx = 0 Then
            dcDepthAiDocGrp = 0 'indeterminate
        Else
            dx = 0
            
            Do
            'vr = Array(.Item(.Keys(dx)))
            Set ob = obOf(.Item(.Keys(dx))) 'vr(0)
            If ob Is Nothing Then 'IsObject(vr(0))
                ck = -1 'invalid
            Else
                If TypeOf ob Is Scripting.Dictionary Then
                    ck = dcDepthAiDocGrp(ob)
                    If ck > 0 Then ck = 1 + ck
                ElseIf TypeOf ob Is Inventor.Document Then
                    ck = 1
                Else
                    ck = -1 'invalid
                End If
            End If
            
            dx = dx + 1
            If dx > mx Then ck = -1 'invalid
            
            Loop While ck = 0 'indeterminate
            
            dcDepthAiDocGrp = ck
        End If
    End With
End Function

Public Function nu_fmIfcTest05A( _
    Optional dcIn As Scripting.Dictionary = Nothing _
) As fmIfcTest05A
    With New fmIfcTest05A
        Set nu_fmIfcTest05A = .Using(dcIn)
    End With
'nu_fmIfcTest05A(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive()))).Show 1
'nu_fmIfcTest05A(dcAiDocGrpsByForm(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive())))).Show 1
'nu_fmIfcTest05A(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive()))).Show 1
End Function

Public Function nu_fmTest05A( _
    Optional dcIn As Scripting.Dictionary = Nothing _
) As fmTest05
    Stop 'DO NOT USE THIS FUNCTION!
    'instead, use the Interface
    'generator nu_fmIfcTest05A
    
    With New fmTest05
        Set nu_fmTest05A = .Holding(dcIn) ' .Using(dcIn)
    End With
'nu_fmTest05A(dcAiDocGrpsByForm(dcAiDocComponents(aiDocActive()))).Show 1
'nu_fmTest05A(dcAiDocGrpsByForm(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive())))).Show 1
'nu_fmTest05A(dcAiDocsByPtNum(dcAiDocComponents(aiDocActive()))).Show 1
End Function

Public Function lsAssyMembers( _
    aiAssy As Inventor.AssemblyDocument _
) As String
    Dim dc As Scripting.Dictionary
    Dim pn As String
    Dim rt As String
    Dim ky As Variant
    
    Set dc = dcAiDocsByPtNum(dcAssyComponentsImmediate(aiAssy)) 'dcAiDocPartNumbers
    pn = vbNewLine & aiAssy.PropertySets.Item( _
        gnDesign).Item(pnPartNum).Value & vbTab
    rt = pn & Join(dc.Keys, pn)
    
    With nuPicker( _
        New kyPickAiPartVsAssy _
    ).AfterScanning(dc)
        With .dcOut
            For Each ky In .Keys
                rt = rt & lsAssyMembers(aiDocument(obOf(.Item(ky))))
            Next
        End With
    End With
    
    lsAssyMembers = rt
End Function

Public Function d0g6f0(AiDoc As Inventor.Document) As String
    ''  Try to pick a distinct listing name
    ''  for a supplied Inventor Document
    Dim rt As String
    Dim ds As String
    
    With AiDoc
        With .PropertySets(gnDesign)
            rt = Trim$(.Item(pnPartNum).Value)
            ds = Trim$(.Item(pnDesc).Value)
        End With
        
        If Len(rt) > 0 Then
            If Len(ds) > 0 Then rt = rt & ": " & ds
        ElseIf Len(ds) > 0 Then
            rt = ds
        End If
        
        If Len(rt) = 0 Then
            ds = .FullFileName
            If Len(ds) > 0 Then
                With nuFso().GetFile(ds)
                    rt = .Name & " (" & .ParentFolder.Path & ")"
                End With
            Else
                rt = .DisplayName
            End If
        End If
        
        d0g6f0 = rt
    End With
End Function

Public Function d0g6f1()
    ''
    ''  testing form class fmTest0
    ''
    With New fmTest0
        .imTNail.Visible = False
        Debug.Print .Controls.Count
        Stop
    End With
End Function

Public Function d0g6f2(dc As Scripting.Dictionary) As Scripting.Dictionary
    ''' Call this one from inside dcAiDocGrpsByForm (above)
    ''' Try: debug.Print txDumpLs(d0g6f2(pkPvA.dcIn).Keys)
    '''
    Dim rt As Scripting.Dictionary
    Dim ad As Inventor.Document
    Dim pr As Inventor.Property
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    
    With dc
        For Each ky In .Keys
            Set ad = aiDocument(.Item(ky))
            If ad Is Nothing Then
                'nothing we can do with it
            Else
                Set pr = ad.PropertySets(gnDesign).Item(pnFamily)
                rt.Add ky, pr
                With pr
                    .Value = "R-PTS"
                    'Stop
                End With
            End If
        Next
    End With
    
    Set d0g6f2 = rt
    '''
    '''
End Function

Public Function d0g6f3()
    ''
    ''  testing new empty form class fmEmpty
    ''
    With New fmEmpty
        '.imTNail.Visible = False
        With .Controls.Add("Forms.ComboBox.1", "test", True)
            Debug.Print .Name
            .Left = 10
            .Top = 10
        End With
        Debug.Print .Controls.Count
        .Show 1
        Stop
    End With
End Function

Public Function d0g7f0() As String
    '''
    ''' This function used to transfer Property Values
    ''' from blank model files GR12 ~ GR20
    ''' to new versions generated from Intraflo's
    ''' supplied STEP files. Save for reference,
    ''' but this version should not likely be used
    ''' as is for other tasks without review.
    '''
    Dim rt As Scripting.Dictionary
    Dim dcPr As Scripting.Dictionary
    Dim sd As Inventor.Document
    Dim td As Inventor.Document
    Dim psSc As Inventor.PropertySet
    Dim psTg As Inventor.PropertySet
    Dim prSc As Inventor.Property
    Dim prTg As Inventor.Property
    Dim ky As Variant
    Dim pn As Variant
    Dim sn As String
    
    Set rt = dcAiDocsByPtNum(dcAssyComponentsImmediate(aiDocActive()))
    With rt
        For Each ky In .Keys
            Debug.Print ky 'This might want to be disabled
            Set sd = aiDocument(obOf(.Item(ky)))
            If UCase$(Left$(ky, 2)) = "GR" Then
                sn = sd.PropertySets(gnDesign).Item(pnStockNum).Value
                If .Exists(sn) Then
                    Set td = aiDocument(obOf(.Item(sn)))
                    Debug.Print ;: 'Stop
                    Set psTg = td.PropertySets(gnCustom)
                    Set dcPr = dcAiPropsInSet(psTg)
                    Set psSc = sd.PropertySets(gnCustom)
                    For Each prSc In psSc
                        If dcPr.Exists(prSc.Name) Then
                            With psTg.Item(prSc.Name)
                                .Value = prSc.Value
                                Debug.Print ; 'Landing Point -- Ctrl-F8 to here
                            End With
                        Else
                            With prSc
                                psTg.Add .Value, .Name ', .PropId
                                Debug.Print ; 'Landing Point -- Ctrl-F8 to here
                            End With
                        End If
                    Next
                    
                    Set psSc = sd.PropertySets(gnDesign)
                    With td.PropertySets(gnDesign)
                        For Each pn In Array( _
                            pnPartNum, pnStockNum, _
                            pnFamily, pnDesc, _
                            pnCatWebLink _
                        )
                            '.Item(pnStockNum).Value = psSc.Item(pnStockNum).Value
                            '.Item(pnFamily).Value = psSc.Item(pnFamily).Value
                            '.Item(pnCatWebLink).Value = psSc.Item(pnCatWebLink).Value
                            '.Item(pnDesc).Value = psSc.Item(pnDesc).Value
                            '.Item(pnPartNum).Value = psSc.Item(pnPartNum).Value
                            '.Item(pn).Value = psSc.Item(pn).Value
                            .Item(CStr(pn)).Value = psSc.Item(CStr(pn)).Value
                            Debug.Print ; 'Landing Point -- Ctrl-F8 to here
                        Next
                    End With
                Else
                End If
            Else
            End If
        Next
    End With
End Function

Public Function d0g8f0() As String
    Dim dx As Long
    Dim fn As String
    
    For dx = 1 To 16
        fn = "Specification" & CStr(dx)
        With cnGnsDoyle().Execute(Join(Array( _
            "select distinct", fn, _
            "from vgMfiItems", _
            "where Family = 'D-BAR'", _
            "and", fn, "is not null", _
            "and", fn, "<> ''", _
            "order by", fn, _
            ";" _
        ), " "))
            If .BOF Or .EOF Then
                'Debug.Print "<EMPTY>"
                'Debug.Print
            Else
                Debug.Print "[" & fn & "]"
                Debug.Print .GetString
            End If
        End With
    Next
End Function

Public Function d0g9f0( _
    Optional ad As Inventor.Document = Nothing, _
    Optional pn As String = "" _
) As String
    
    Dim vw As Inventor.View
    Dim cm As Inventor.Camera
    Dim bp As String
    
    If ad Is Nothing Then
        d0g9f0 = d0g9f0(ThisApplication.ActiveDocument, pn)
    ElseIf Len(pn) < 1 Then
        d0g9f0 = d0g9f0(ad, d0g9f3(ad))
    Else
        bp = "C:\Doyle_Vault\Designs\Misc\andrewT\"
        Set vw = ad.Views.Item(1) 'ThisApplication.ActiveView
        Set cm = vw.Camera
        
        With vw
            'Debug.Print .Left, .Top
            'Debug.Print .Width, .Height
            
            Debug.Print ; 'breakpoint anchor
            
            With .Camera
                .ViewOrientationType = kIsoTopRightViewOrientation
                .Fit
                .Apply
            End With
            .Fit
            .Update
            Debug.Print ; 'breakpoint anchor
            '.SaveAsBitmapWithOptions pn & "-I.png", 0, 0
            .SaveAsBitmap bp & pn & "-I.png", .Width, .Height '0, 0
            
            With .Camera
                .ViewOrientationType = kFrontViewOrientation
                .Fit
                .Apply
            End With
            .Fit
            .Update
            Debug.Print ; 'breakpoint anchor
            '.SaveAsBitmapWithOptions pn & "-I.png", 0, 0
            .SaveAsBitmap bp & pn & "-F.png", 0, 0
            
            With .Camera
                .ViewOrientationType = kTopViewOrientation
                .Fit
                .Apply
            End With
            .Fit
            .Update
            Debug.Print ; 'breakpoint anchor
            '.SaveAsBitmapWithOptions pn & "-I.png", 0, 0
            .SaveAsBitmap bp & pn & "-T.png", 0, 0
            
            .GoHome
            .Update
        End With
    End If
    
    Debug.Print ; 'breakpoint anchor
    
    d0g9f0 = ""
End Function

Public Function d0g9f1(ad As Inventor.AssemblyDocument) As String
    Dim rw As Inventor.iAssemblyTableRow
    
    With ad.ComponentDefinition
        If .IsiAssemblyFactory Then
            With .iAssemblyFactory
                For Each rw In .TableRows
                    With rw
                        Debug.Print .MemberName
                    End With
                    'Set .DefaultRow = rw
                    'Stop
                Next
            End With
        Else
        End If
    End With
    
    d0g9f1 = ""
End Function

Public Function d0g9f2(oc As Inventor.ComponentOccurrence) As String
    With oc
        If .IsiAssemblyMember Then
            Stop
            '.Definition
            '.Replace
            
        ElseIf .IsiPartMember Then
            Stop
        Else
            Stop
        End If
    End With
End Function

Public Function d0g9f2as(cd As Inventor.AssemblyComponentDefinition) As String
    With cd
        '.IsiAssemblyMember
        '.iAssemblyMember
        With .iAssemblyMember
            '.ParentFactory
            '.Row
        End With
    End With
End Function

Public Function d0g9f3(ad As Inventor.AssemblyDocument) As String
    Dim cp As Inventor.AssemblyDocument
    
    With ad.ComponentDefinition.Occurrences.Item(1)
        Set cp = aiDocAssy(.Definition.Document)
        If cp Is Nothing Then
            d0g9f3 = "NO-NUM-ASSY"
        Else
            d0g9f3 = _
                cp.PropertySets( _
                gnDesign).Item( _
                pnPartNum).Value
        End If
    End With
End Function

Public Sub PlaceInAssembly()
    Dim dc As Scripting.Dictionary
    Dim cm As Inventor.CommandManager
    Dim cd As Inventor.Document
    Dim ad As Inventor.Document
    Dim rp As VbMsgBoxResult
    Dim nm As String
    
    With ThisApplication
        If .ActiveDocumentType = kPartDocumentObject _
        Or .ActiveDocumentType = kAssemblyDocumentObject _
        Then
            Set cd = .ActiveDocument
            Set dc = dcAiAssyDocs( _
                dcAiDocsVisible())
            dc.Remove cd.FullDocumentName
            
            With nuSelAiDoc().WithList(dc.Keys)
                Do
                    nm = .GetReply()
                    If dc.Exists(nm) Then
                        Set ad = dc.Item(nm)
                        rp = vbOK
                    Else
                        Set ad = Nothing
                        rp = MsgBox( _
                            "No Valid Assembly Selected.", _
                            vbRetryCancel, "No Assembly" _
                        ) ' Try Again?
                    End If
                Loop While rp = vbRetry
            End With
            
            If ad Is Nothing Then
                Debug.Print ;
            Else
                ad.Activate
                Set cm = .CommandManager
                With cm
                    .PostPrivateEvent kFileNameEvent, cd.FullDocumentName
                    .ControlDefinitions.Item("AssemblyPlaceComponentCmd").Execute
                End With
            End If
        End If
    End With
End Sub
