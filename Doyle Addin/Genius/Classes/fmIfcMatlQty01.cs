

Private WithEvents fm As fmMatlQty
Attribute fm.VB_VarHelpID = -1
Private WithEvents lbxMatlQty   As MSForms.ListBox
Attribute lbxMatlQty.VB_VarHelpID = -1
Private WithEvents txbMatlQty   As MSForms.TextBox
Attribute txbMatlQty.VB_VarHelpID = -1
Private WithEvents cbxUnitQty   As MSForms.ComboBox
Attribute cbxUnitQty.VB_VarHelpID = -1
Private WithEvents imgThmNail   As MSForms.Image
Attribute imgThmNail.VB_VarHelpID = -1

'Private WithEvents cmdOK        As MSForms.CommandButton
'Private WithEvents cmdCancel    As MSForms.CommandButton
'Private WithEvents cmdOK        As MSForms.CommandButton

Private lblPartNumber           As MSForms.Label
Private lblPartInfo             As MSForms.Label
Private lblMatlNumber           As MSForms.Label
Private lblMatlInfo             As MSForms.Label
'lblMatlQty
'lblNoImg

'imThmNail

Private dcResult As Scripting.Dictionary

'Private dcGiven As Scripting.Dictionary
'Private dcWorkg As Scripting.Dictionary
'Private fmStatus As VbMsgBoxResult

Private Const fmVersion As String = "Material Quantity Form Interface 0.0.0.0 [2022.03.04]"
'''
'''
'''

Private Sub Class_Initialize()
    'Dim ctl As MSForms.Control
    
    Set dcResult = New Scripting.Dictionary
    With dcResult
        .Add pnRmQty, 0
        .Add pnRmUnit, ""
    End With
    
    Set fm = New fmMatlQty
    With fm
        'For Each ctl In .Controls
        '    Debug.Print ctl.Name
        'Next
        
        Set cbxUnitQty = .cbxUnitQty
        cbxUnitQty.List = Split("IN FT FT2 IN2 EA")
        
        Set lbxMatlQty = .lbxMatlQty
        Set txbMatlQty = .txbMatlQty
        
        Set imgThmNail = .imThmNail
        Set lblPartNumber = .lblPartNumber
        Set lblPartInfo = .lblPartInfo
        Set lblMatlNumber = .lblMatlNumber
        Set lblMatlInfo = .lblMatlInfo
    End With
End Sub

Public Function Result() As Scripting.Dictionary
    Set Result = dcCopy(dcResult)
End Function

Private Function Changes( _
    wkg As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With wkg: For Each ky In dcResult.Keys
        If .Exists(ky) Then
            If .Item(ky) = dcResult.Item(ky) Then
                'no difference; skip it
            Else
                rt.Add ky, .Item(ky)
            End If
        Else 'not set; not sure what to do here
            'rt.Add ky, Empty
            'don't really like this option
            'so leaving it disabled for now
        End If
    Next: End With
    
    Set Changes = dcCopy(dcResult)
End Function

Private Function Commit( _
    src As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim ky As Variant
    
    With dcResult
    For Each ky In .Keys
        If src.Exists(ky) Then
            .Item(ky) = src.Item(ky)
        End If
    Next: End With
    
    Set Commit = dcCopy(dcResult)
End Function

Public Function SeeUser( _
    Optional About As Object = Nothing _
) As Scripting.Dictionary 'fmIfcMatlQty01
    Dim ky As String
    Dim ck As String
    
    If About Is Nothing Then
        Set SeeUser = SeeUser(nuDcPopulator( _
            ).Setting(pnRmQty & "()", _
                nuDcPopulator( _
                    ).Setting(4, 1 _
                    ).Setting(2, 1 _
                    ).Setting(24, 1 _
                ).Dictionary() _
            ).Setting(pnRmQty, 24 _
            ).Setting(pnRmUnit, "IN" _
            ).Setting(pnPartNum, "NO-ITM-GIVEN" _
            ).Setting(pnRawMaterial, "NO-MTL-GIVEN" _
        ).Dictionary())
    ElseIf TypeOf About Is Scripting.Dictionary Then
        Set SeeUser = SeeUserWithDict(About)
    ElseIf TypeOf About Is Inventor.PartDocument Then
        Set SeeUser = SeeUserWithPart(About)
    ElseIf TypeOf About Is Inventor.Property Then
        Set SeeUser = SeeUserWithQtyProp(About)
    Else
        Set SeeUser = SeeUser()
    End If
End Function

''' make this one Public later
''' once Part version is working
Private Function SeeUserWithModel( _
    About As Inventor.Property _
) As fmIfcMatlQty01
End Function

Public Function SeeUserWithPart( _
    About As Inventor.PartDocument _
) As Scripting.Dictionary 'fmIfcMatlQty01
    If About Is Nothing Then
        Set SeeUserWithPart = SeeUser(About)
    Else
        Dim dcPr As Scripting.Dictionary
        Dim obPr As Inventor.Property
        Dim kyPr As Variant
        Dim op As Long
        
        Set dcPr = New Scripting.Dictionary
        With About.PropertySets
            With .Item(gnDesign)
                dcPr.Add pnPartNum, .Item(pnPartNum).Value
                dcPr.Add pnDesc, .Item(pnDesc).Value
            End With
            
            On Error Resume Next
            With .Item(gnCustom)
                For Each kyPr In Array( _
                    pnRawMaterial, pnRmQty, pnRmUnit _
                )
                    Err.Clear
                    Set obPr = .Item(CStr(kyPr))
                    If Err.Number = 0 Then
                        dcPr.Add kyPr, obPr.Value
                    Else
                        Debug.Print Err.Description
                        Stop
                        Err.Clear
                    End If
                Next
            End With
            On Error GoTo 0
        End With
        
        '''
        ''' prepare Dictionary of Dimensions
        ''' with Count of Each
        '''
        Dim dcDm As Scripting.Dictionary
        Dim vlDm As Variant
        Dim ctDm As Long
        
        Set dcDm = New Scripting.Dictionary
        
        With nuAiBoxData().UsingInches(1)
            For op = 0 To 1
            With .UsingModel(About, op)
                For Each vlDm In Array( _
                    Round(.SpanX, 4), _
                    Round(.SpanY, 4), _
                    Round(.SpanZ, 4), _
                0)
                With dcDm
                If vlDm > 0 Then
                    If .Exists(vlDm) Then
                        ctDm = .Item(vlDm) + 1
                        .Item(vlDm) = ctDm
                    Else
                        .Add vlDm, ctDm
                    End If
                End If: End With: Next
            End With: Next
        End With
        
        With dcPr
            .Add pnRmQty & "()", dcDm
            .Add "img", About.Thumbnail
        End With
        
        Set SeeUserWithPart _
        = SeeUserWithDict(dcPr)
    End If
End Function

Private Function SeeUserWithQtyProp( _
    About As Inventor.Property _
) As Scripting.Dictionary 'fmIfcMatlQty01
    '''
    ''' this one will have to be heavily modified
    ''' likely dumping a bunch of code now implemented
    ''' in SeeUserWithPart, which can simply be
    ''' called with the Document containing
    ''' the supplied Property
    '''
    
    If About Is Nothing Then
        Stop
    Else
        ''' these variables are for use
        ''' in separating quantity from
        ''' unit of measure in Value of
        ''' supplied Property
        Dim vlIn As String
        Dim arIn As Variant
        Dim qtIn As Double
        Dim unIn As String
        ''' split incoming Property Value into
        ''' Quantity and Unit of Measurement
        vlIn = CStr(About.Value) & " "
        ''' note: concatenated space at end
        ''' of Value text should ensure two
        ''' members of arIn, as follows
        arIn = Split(vlIn, " ", 2)
        
        qtIn = Round(Val(arIn(0)), 4)
        If UBound(arIn) > 0 Then
            unIn = Trim$(arIn(1))
        End If
        ''' this section and its associated variables
        ''' will likely be exported to a separate function
        
        ''' force blank Unit of
        ''' Measure to default inches
        If Len(unIn) = 0 Then unIn = "IN"
        
        ''' the following section SHOULD be
        ''' implemented now in SeeUserWithPart
        ''' it should be possible to simply
        ''' call that function, completely
        ''' ignoring the supplied Property
        '''
        ''' prepare Dictionary of Dimensions
        ''' with Count of Each
        '''
        Dim dcDm As Scripting.Dictionary
        Dim vlDm As Variant
        Dim ctDm As Long
        
        Set dcDm = New Scripting.Dictionary
        If qtIn > 0 Then dcDm.Add qtIn, 1
        
        '''
        ''' get all necessary information
        ''' from Inventor Model
        '''
        Dim md As Inventor.Document
        Dim mdPt As Inventor.Property
        Dim mdMt As Inventor.Property
        
        Set md = aiDocument(About.Parent.Parent.Parent)
        With md.PropertySets
            Set mdPt = .Item(gnDesign).Item(pnPartNum)
            On Error Resume Next
            Err.Clear
            Set mdMt = .Item(gnCustom).Item(pnRawMaterial)
            If Err.Number = 0 Then
            Else
                Stop
            End If
            On Error GoTo 0
        End With
        
        With nuAiBoxData( _
        ).UsingInches(1 _
        ).UsingModel(About)
            For Each vlDm In Array( _
                Round(.SpanX, 4), _
                Round(.SpanY, 4), _
                Round(.SpanZ, 4), _
            0)
            With dcDm
            If vlDm > 0 Then
                If .Exists(vlDm) Then
                    ctDm = .Item(vlDm) + 1
                    .Item(vlDm) = ctDm
                Else
                    .Add vlDm, ctDm
                End If
            End If: End With: Next
        End With
        
        With nuDcPopulator( _
            ).Setting(pnRmQty & "()", dcDm _
            ).Setting(pnRmQty, qtIn _
            ).Setting(pnRmUnit, unIn _
            ).Setting(pnPartNum, mdPt.Value _
            ).Setting(pnRawMaterial, mdMt.Value _
        )
        Set SeeUserWithQtyProp _
            = SeeUserWithDict( _
            .Dictionary() _
        )
        End With
    End If
End Function

Public Function SeeUserWithDict( _
    About As Scripting.Dictionary _
) As Scripting.Dictionary 'fmIfcMatlQty01
    Dim ky As String
    Dim ck As String
    
    If About Is Nothing Then
        Set SeeUserWithDict = SeeUserWithDict(nuDcPopulator( _
            ).Setting(pnRmQty & "()", _
                nuDcPopulator( _
                    ).Setting(4, 1 _
                    ).Setting(2, 1 _
                    ).Setting(24, 1 _
                ).Dictionary() _
            ).Setting(pnRmQty, 24 _
            ).Setting(pnRmUnit, "IN" _
            ).Setting(pnPartNum, "NO-ITM-GIVEN" _
            ).Setting(pnRawMaterial, "NO-MTL-GIVEN" _
        ).Dictionary())
    Else
        With About
            '.Add "img", About.Thumbnail
            If .Exists("img") Then
                Set imgThmNail.Picture _
                = .Item("img")
            End If
            
            If .Exists(pnDesc) Then
                txbMatlQty.Value = Val( _
                CStr(.Item(pnDesc)))
            End If
            
            ky = pnRmQty & "()"
            If .Exists(ky) Then
                lbxMatlQty.List = _
                dcOb(.Item(ky)).Keys
            End If
            
            If .Exists(pnRmQty) Then
                txbMatlQty.Value = Val( _
                CStr(.Item(pnRmQty)))
            End If
            
            If .Exists(pnRmUnit) Then
                On Error Resume Next
                
                Err.Clear
                cbxUnitQty.Value = .Item(pnRmUnit)
                If Err.Number Then
                    Debug.Print ; 'Breakpoint Landing
                    cbxUnitQty.Value = "IN"
                End If
                On Error GoTo 0
            End If
            
            '''
            ''' Following are "boilerplate" elements
            ''' for Part/Item and Raw Material numbers,
            ''' along with their descriptions.
            '''
            ''' A thumbnail image of the Part is also
            ''' expected to be supplied at some point,
            ''' but will be held off for now, pending
            ''' successful testing of the form's main
            ''' functions.
            '''
            ''' Part/Item Number
            If .Exists(pnPartNum) Then
                lblPartNumber.Caption _
                = CStr(.Item(pnPartNum))
            End If
            
            ''' Material Number
            If .Exists(pnRawMaterial) Then
                lblMatlNumber.Caption _
                = CStr(.Item(pnRawMaterial))
            End If
            
            ''' Item Description
            If .Exists(pnDesc) Then
                lblPartInfo.Caption _
                = CStr(.Item(pnDesc))
            End If
            
            ''' Material Description
            ''' (not expected at this time)
            ky = pnRawMaterial & ":"
            If .Exists(ky) Then
                lblMatlInfo.Caption _
                = CStr(.Item(ky))
            End If
            
            'imThmNail
        End With
        
        With Commit(About)
        End With
        
        fm.Show 1
        'Stop
        
        With nuDcPopulator( _
        ).Setting(pnRmQty, _
            Round(Val( _
            txbMatlQty.Value _
            ), 4) _
        ).Setting(pnRmUnit, _
            cbxUnitQty.Value _
        ) 'Mapping...
        ' txbMatlQty -> pnRmQty
        ' cbxUnitQty -> pnRmUnit
        
            Set SeeUserWithDict = Commit(.Dictionary)
        End With
    End If
End Function

Public Function Version()
    Version = fmVersion
    'fmStatus=vbRetry
End Function

Private Sub Class_Terminate()
    'Set dcWorkg = Nothing
    'Set dcGiven = Nothing
    
    Set imgThmNail.Picture = Nothing
    Set imgThmNail = Nothing
    
    Set cbxUnitQty = Nothing
    
    Set lbxMatlQty = Nothing
    Set txbMatlQty = Nothing
    
    Set lblPartNumber = Nothing
    Set lblPartInfo = Nothing
    Set lblMatlNumber = Nothing
    Set lblMatlInfo = Nothing
    
    Set fm = Nothing
End Sub

Private Sub fm_Sent(Signal As VbMsgBoxResult)
    Dim ck As VbMsgBoxResult
    
    If Signal = vbCancel Then
        ck = MsgBox(Join(Array( _
            "Material Quantity", _
            "and Units will", _
            "remain unchanged." _
        ), vbNewLine), vbYesNo, _
            "Cancel Update?" _
        )
        'Stop
        
        If ck = vbYes Then
            With dcResult
                txbMatlQty.Value = CStr(.Item(pnRmQty))
                cbxUnitQty.Value = .Item(pnRmUnit)
            End With
            fm.Hide
        ElseIf ck = vbCancel Then
            Stop 'drop and debug
            'NOTE: Without Cancel Button
            'available on MsgBox, this
            'option won't be accessible.
            
            'a proposed "debug" mode that
            'would add the Cancel Button to
            'the MsgBox has not yet been
            'implemented, but might in future.
        End If
        Debug.Print ; 'Breakpoint Landing
    ElseIf Signal = vbOK Then
        ck = MsgBox(Join(Array( _
            "Update Material", _
            "Quantity to " _
            & CStr(Round(Val( _
                txbMatlQty.Value _
            ), 4)) _
            & cbxUnitQty.Value & "?" _
        ), vbNewLine), vbYesNo, _
            "Update Quantity?" _
        )
        'Stop
        
        If ck = vbYes Then
            fm.Hide
            Debug.Print ; 'Breakpoint Landing
        ElseIf ck = vbCancel Then
            Stop 'drop and debug
            'NOTE: Without Cancel Button
            'available on MsgBox, this
            'option won't be accessible.
            
            'a proposed "debug" mode that
            'would add the Cancel Button to
            'the MsgBox has not yet been
            'implemented, but might in future.
        End If
        Debug.Print ; 'Breakpoint Landing
    Else
        Stop
    End If
End Sub

Private Sub lbxMatlQty_DblClick( _
    ByVal Cancel As MSForms.ReturnBoolean _
)
    txbMatlQty.Value = lbxMatlQty.Value
End Sub

Private Sub lbxMatlQty_MouseMove( _
    ByVal Button As Integer, ByVal Shift As Integer, _
    ByVal X As Single, ByVal Y As Single _
)
    Dim dt As MSForms.DataObject
    Dim ef As Integer
    
    If Button = 1 Then
        Set dt = New MSForms.DataObject
        dt.SetText lbxMatlQty.Value
        ef = dt.StartDrag()
    End If
End Sub

Private Sub txbMatlQty_Change()
    Dim ck As Double
    Dim tx As String
    Dim gp As Variant
    Dim mx As Long
    Dim dx As Long
    
    With txbMatlQty
        tx = .Value
        gp = Split(tx, ".")
        mx = UBound(gp)
        
        For dx = LBound(gp) To mx
            ck = Val(gp(dx))
            
            If ck > 0 Then
                gp(dx) = CStr(ck)
            ElseIf dx > 0 Then
                gp(dx) = ""
            Else
                gp(dx) = "0"
            End If
        Next
        tx = Join(gp, ".")
        
        If tx <> .Value Then
            DoEvents
            .Value = tx
        End If
    End With
End Sub
