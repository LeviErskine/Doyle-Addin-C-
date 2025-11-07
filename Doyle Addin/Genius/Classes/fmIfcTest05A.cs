

Public Event Sent(Signal As VbMsgBoxResult)
Public Event GroupIs(Now As String)
Public Event ItemIs(Now As String)

Private obClient As Object
''' this is meant to hold the "client" Object
''' expected to make calls to this interface.
''' not sure if we want to do anything with this yet

Private WithEvents fm As fmTest05
Attribute fm.VB_VarHelpID = -1
'Private WithEvents fm As MSForms.UserForm 'fmTest04

Private WithEvents lbxItems         As MSForms.ListBox
Attribute lbxItems.VB_VarHelpID = -1
Private WithEvents tbsItemGrps      As MSForms.TabStrip
Attribute tbsItemGrps.VB_VarHelpID = -1
Private WithEvents lblPartNum       As MSForms.Label
Attribute lblPartNum.VB_VarHelpID = -1
Private WithEvents lblDesc          As MSForms.Label
Attribute lblDesc.VB_VarHelpID = -1
Private WithEvents imgOfItem        As MSForms.Image
Attribute imgOfItem.VB_VarHelpID = -1
Private WithEvents cmdOpenItem      As MSForms.CommandButton
Attribute cmdOpenItem.VB_VarHelpID = -1
'Private WithEvents cmdEndCancel     As MSForms.CommandButton
'Private WithEvents cmdEndSave       As MSForms.CommandButton

'Private dcActvSet   As Scripting.Dictionary
Private itemsFlat As Scripting.Dictionary
Private allGroups As Scripting.Dictionary
Private itmPicked As Scripting.Dictionary
Private gdcActive As Scripting.Dictionary

Private docActive As Inventor.Document

Private Const txVersion As String = ""

Private Const kyInitList As String = "initSpecs"    'key name identifying initial spec list
Private Const vsnString As String = "Form Test04 Interface A v0.1.0.0 [2022.03.17]"
''' prior values                    "Form Test04 Interface A v0.0.0.0 [2022.03.03]"
'''                                 ""
'''                                 ""
'''
'''
'''

Public Function Itself() As fmIfcTest05A
    ''' returns this fmIfcTest04A class instance "Itself"
    ''' should be HIGHLY useful inside a With context
    Set Itself = Me
End Function

Public Function Using( _
    Optional Dict As Scripting.Dictionary = Nothing _
) As fmIfcTest05A 'fmTest05
    Dim ky As Variant
    Dim dp As Long
    
    If Dict Is Nothing Then
        Set Using = Me.Using( _
        New Scripting.Dictionary)
    Else
        Set itemsFlat = Nothing
        Set allGroups = Nothing
        dp = dcDepthAiDocGrp(Dict)
        
        Select Case dp
        Case 1:
            Set Using = withDcFlat(Dict)
        Case 2:
            Set Using = withDcGrpd(Dict)
            
            Set itemsFlat = New Scripting.Dictionary
            With allGroups
            For Each ky In .Keys
                Set itemsFlat = dcKeysCombined( _
                    dcOb(.Item(ky)), itemsFlat, 1 _
                )
            Next: End With
        Case Else:
            Set Using = Me
        End Select
    End If
End Function

Public Function GroupNow() As String
    GroupNow = fm.GroupNow()
End Function

Public Function InGroup( _
    GrpId As String _
) As fmIfcTest05A
    If fm.InGroup(GrpId _
    ).GroupNow() = GrpId Then
        ''' change succeeded!
    Else
        ''' couldn't change!
    End If
    
    Set InGroup = Me
End Function

Public Function ItemNow() As String
    ItemNow = fm.ItemNow()
End Function

Public Function OnItem( _
    ItemId As String _
) As fmIfcTest05A
    If fm.OnItem(ItemId _
    ).ItemNow() = ItemId Then
        ''' change succeeded!
    Else
        ''' couldn't change!
        Stop
    End If
    
    Set OnItem = Me
End Function

Public Function Show( _
    Modal As Variant _
) As fmIfcTest05A
    fm.Show Modal
    Set Show = Me
End Function

Public Function Hide() As fmIfcTest05A
    fm.Hide
    Set Hide = Me
End Function

Public Function SaveAll() As Scripting.Dictionary
    'debug.Print aiDocument(itemsFlat.Item(itemsFlat.Keys(0))).Dirty
    ''' NOTE[2022.04.13.1224] (copied from ...)
    ''' want to initiate 'save all' operation here
    ''' or somewhere nearby. note immediate mode
    ''' command in comment above
    Dim rtGd As Scripting.Dictionary
    'Dim rtBd As Scripting.Dictionary
    Dim wk As Inventor.Document
    Dim ky As Variant
    'Dim mx As Long
    'Dim dx As Long
    
    Set rtGd = New Scripting.Dictionary
    'Set rtBd = New Scripting.Dictionary
    
    With itemsFlat
        On Error Resume Next
        For Each ky In .Keys
            Set wk = .Item(ky)
            With wk
            If .Dirty Then
                Err.Clear
                .Save2
                If Err.Number = 0 Then
                    rtGd.Add ky, wk
                Else
                End If
            Else
                rtGd.Add ky, wk
            End If: End With
        Next
        On Error GoTo 0
    End With
    
    Set SaveAll = rtGd
End Function

Private Function withDcFlat( _
    Dict As Scripting.Dictionary _
) As fmIfcTest05A 'fmTest05
    Set itemsFlat = dcCopy(Dict)
    Set withDcFlat = withDcGrpd( _
    dcAiDocGrpsByForm(itemsFlat))
End Function

Private Function withDcGrpd( _
    Dict As Scripting.Dictionary _
) As fmIfcTest05A 'fmTest05
    Dim ky As Variant
    Dim ls As MSForms.Tabs
    
    Set itmPicked = New Scripting.Dictionary
    Set docActive = Nothing
    
    Set ls = tbsItemGrps.Tabs
    ls.Clear
    
    Set allGroups = Dict
    
    With allGroups: For Each ky In Split( _
        "MAYB DBAR SHTM ASSY PRCH HDWR" _
    ) 'instead of .Keys, to
    '  ensure preferred order
        If .Exists(ky) Then
        With dcOb(.Item(ky))
        ''  check for group members
        If .Count > 0 Then 'select the first
            itmPicked.Add ky, .Keys(0)
            ls.Add ky 'this will want more development later
        Else 'paint it blank
            'itmPicked.Add ky, ""
            ''  actually, don't do anything
            ''  like, don't even add the tab
            ''  if nothing's going to be
            ''  there anyway
        End If: End With: End If
    Next: End With
    Set withDcGrpd = Me
End Function

Private Function gpActive() As String
    Dim tb As MSForms.Tab
    
    With tbsItemGrps
        Set tb = .Tabs.Item(.Value)
    End With
    gpActive = tb.Name
End Function

Private Sub cmdOpenItem_Click()
    Dim ck As VbMsgBoxResult
    Dim pn As String
    
    If docActive Is Nothing Then
    '
    Else
        With docActive
            pn = .PropertySets.Item(gnDesign).Item(pnPartNum).Value
            
            ''' REV[2022.05.06.1142]
            ''' added check for Part Document to avoid error
            ''' trying to edit Material for Assembly Documents.
            If TypeOf docActive Is Inventor.PartDocument Then
                ck = MsgBox(Join(Array( _
                    "Would you rather just edit", _
                    "material for " & pn & "?", "", _
                    "(No to go ahead and open)" _
                ), vbNewLine), _
                    vbYesNoCancel + vbQuestion, _
                    "Edit Material?" _
                )
            Else
                ck = vbNo
            End If
            
            If ck = vbCancel Then
                Stop
            ElseIf ck = vbYes Then
                ''' NOTE[2022.05.06.1143]
                ''' this section throws an error
                ''' if the Document is an assembly.
                ''' REV[2022.05.06.1142] above adds
                ''' a check to prevent this branch
                ''' from being taken in that case.
                Debug.Print ConvertToJson( _
                askUserForPartMatlUpdate( _
                itemsFlat.Item(pn)), vbTab)
            Else
                If .Open Then
                    ck = vbYes
                Else
                    ck = MsgBox(Join(Array( _
                            "Document " & pn, _
                            "is not presently open.", _
                            "Go ahead and open it?" _
                        ), vbNewLine), _
                        vbYesNo, "Open " & pn & "?" _
                    )
                End If
                
                If ck = vbYes Then
                    On Error Resume Next
                    
                    Err.Clear
                    .Activate
                    
                    If Err.Number = 0 Then
                    Else
                        If ThisApplication.Documents.Open( _
                            .FullDocumentName, True _
                        ) Is docActive Then
                            Debug.Print ; 'Breakpoint Landing
                        Else
                            Stop
                            Debug.Print ; 'Breakpoint Landing
                        End If
                        
                        Err.Clear
                        .Activate
                        If Err.Number Then Stop
                        
                        Debug.Print ; 'Breakpoint Landing
                    End If
                    
                    On Error GoTo 0
                End If
            End If
        End With
    End If
End Sub

Private Sub fm_GroupIs(Now As String)
    '''
    RaiseEvent GroupIs(Now)
End Sub

Private Sub fm_ItemIs(Now As String)
    '''
    RaiseEvent ItemIs(Now)
End Sub

Private Sub fm_Sent(Signal As VbMsgBoxResult)
    'Public Event Sent(Signal As VbMsgBoxResult)
    '''
    Dim ck As VbMsgBoxResult
    
    If obClient Is Nothing Then
        ck = vbRetry
        
        Select Case Signal
        Case vbOK
            'ck = MsgBox(Join(Array( _
            '    "Save and Close", _
            '    "Operation Selected" _
            '), vbNewLine), _
            '    vbYesNoCancel, _
            '    "Save Documents?" _
            ')
            
            'debug.Print aiDocument(itemsFlat.Item(itemsFlat.Keys(0))).Dirty
            ''' NOTE[2022.04.13.1224]
            ''' want to initiate 'save all' operation here
            ''' or somewhere nearby. note immediate mode
            ''' command in comment above
            With dcKeysMissing( _
                itemsFlat, _
                SaveAll() _
            )
                If .Count > 0 Then
                    ck = MsgBox(Join(Array( _
                        "Errors encountered trying to", _
                        "save the following Documents:", _
                        vbTab & txDumpLs(.Keys, _
                            vbNewLine & vbTab _
                        ), _
                        "", _
                        "Close anyway?" _
                    ), vbNewLine), _
                        vbYesNoCancel, _
                        "Errors on Save!" _
                    )
                Else
                    ck = vbYes
                End If
            End With
        Case vbAbort, vbCancel
            ck = MsgBox(Join(Array( _
                "Cancel", _
                "Operation", _
                "Selected" _
            ), vbNewLine), _
                vbYesNoCancel, _
                "Finished?" _
            )
        'Case Else
        Case Else
        End Select
        
        If ck = vbCancel Then
            Stop
        ElseIf ck = vbYes Then
            Hide
        End If
    Else
        Stop
        RaiseEvent Sent(Signal)
    End If
End Sub

Private Sub lbxItems_Change()
    Dim pn As String
    Dim pc As stdole.StdPicture
    
    'Stop
    pn = lbxItems.Value
    With gdcActive
        If .Exists(pn) Then
            Set docActive = aiDocument(.Item(pn))
            With docActive
                With .PropertySets.Item(gnDesign)
                    lblPartNum.Caption = .Item(pnPartNum).Value
                    lblDesc.Caption = .Item(pnDesc).Value
                End With
                
                On Error Resume Next
                    Set pc = .Thumbnail
                    
                    If Err.Number = 0 Then
                    Else
                        Set pc = Nothing
                    End If
                    
                    Set imgOfItem.Picture = pc
                On Error GoTo 0
            End With
            'docActive.Thumbnail
        Else
            Set docActive = Nothing
            lblPartNum.Caption = "(select part)"
            lblDesc.Caption = ""
        End If
    End With
    
    ''' REV[2022.03.17.1348]
    '''     add DoEvents for rapid visual feedback
    '''     (see tbsItemGrps_Change for details)
    DoEvents
End Sub

Private Sub tbsItemGrps_Change()
    Dim tb As MSForms.Tab
    Dim nm As String
    
    'With tbsItemGrps
    '    Set tb = .Tabs.Item(.Value)
    'End With
    nm = gpActive() 'tb.Name
    
    Set gdcActive = dcOb( _
        allGroups.Item(nm) _
    )
    With lbxItems
        .List = gdcActive.Keys ' dcOb( _
            allGroups.Item(nm) _
        ).Keys
        'If gdcActive.Count > 0 Then
        .Value = itmPicked.Item(nm) 'gdcActive.Item()
    End With
    
    ''' REV[2022.03.17.1348]
    '''     adding DoEvents steps to various
    '''     Change Event handlers to try to
    '''     ensure timely visual feedback
    '''     to the User in-process
    DoEvents
End Sub

Private Sub Class_Initialize()
    'Set fm =
    With New fmTest05
    'End With
    'With fm
        Set lbxItems = .lbxItems
        Set tbsItemGrps = .tbsItemGrps
        Set lblPartNum = .lblPartNum
        Set lblDesc = .lblDesc
        Set imgOfItem = .imgOfItem
        Set cmdOpenItem = .cmdOpenItem
        'Set cmdEndCancel = .cmdEndCancel
        'Set cmdEndSave = .cmdEndSave
        
        Set fm = .Holding(Me)
    End With
    
    'Set dcActvSet = New Scripting.Dictionary
    Set allGroups = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    Set allGroups = Nothing
    
    Set fm = fm.Dropping(Me)
    
    'Set cmdEndSave = Nothing
    'Set cmdEndCancel = Nothing
    Set cmdOpenItem = Nothing
    Set imgOfItem = Nothing
    Set lblDesc = Nothing
    Set lblPartNum = Nothing
    Set tbsItemGrps = Nothing
    
    Set lbxItems = Nothing
    Set lbxItems = Nothing
    
    Set fm = Nothing
End Sub

