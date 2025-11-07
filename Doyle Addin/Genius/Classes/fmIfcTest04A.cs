

Private fm As fmTest04
'Private fmSpec As fmTest04
'Private WithEvents fm As MSForms.UserForm 'fmTest04
Private WithEvents lbxSpecOps As MSForms.ListBox
Attribute lbxSpecOps.VB_VarHelpID = -1
Private WithEvents lbxSpecSet As MSForms.ListBox
Attribute lbxSpecSet.VB_VarHelpID = -1

Private dcSpecPairs As Scripting.Dictionary
Private dcInitList  As Scripting.Dictionary
Private dcActvOps   As Scripting.Dictionary
Private dcActvSet   As Scripting.Dictionary

Private Const kyInitList As String = "initSpecs"    'key name identifying initial spec list
Private Const vsnString As String = "Form Test04 Interface A v0.1.0.0 [2022.03.17]"
''' prior values                    "Form Test04 Interface A v0.0.0.0 [2022.03.03]"
'''                                 ""
'''                                 ""
'''
'''
'''

Private Sub Class_Initialize()
    Set fm = New fmTest04
    Set lbxSpecOps = fm.lbxSpecOps
    Set lbxSpecSet = fm.lbxSpecSel
    
    Set dcSpecPairs = dcGnsMatlSpecPairings()
    ''' probably want to make this controllable
    ''' from client processes to support more
    ''' general usage. this will do for now.

    Set dcInitList = dcSpecPairs
    ''' like dcSpecPairs, probably want this
    ''' to be controllable from client to
    ''' facilitate flexible usage, but again,
    ''' this will serve for the moment.
    
    Set dcActvOps = dcInitList
    Set dcActvSet = New Scripting.Dictionary
    ''' since the user will not have been met
    ''' at this point, no options would be
    ''' selected at this moment, but ALL options
    ''' in dcInitList should be presently active.
End Sub

Private Sub Class_Terminate()
    Set dcActvSet = Nothing
    Set dcActvOps = Nothing
    Set dcInitList = Nothing
    Set dcSpecPairs = Nothing
    
    Set lbxSpecSet = Nothing
    Set lbxSpecOps = Nothing
    Set fm = Nothing
End Sub

Public Function Itself() As fmIfcTest04A
    ''' returns this fmIfcTest04A class instance "Itself"
    ''' should be HIGHLY useful inside a With context
    Set Itself = Me
End Function

Public Function Using( _
    Optional About As Scripting.Dictionary = Nothing _
) As fmIfcTest04A
    If About Is Nothing Then
        Set Using = Using(nuDcPopulator( _
            ).Setting(kyInitList, dcInitList _
        ).Dictionary)
    ElseIf About.Exists(kyInitList) Then
        Set dcInitList = dcOb(About.Item(kyInitList))
        
        Set dcActvOps = dcInitList
        Set dcActvSet = New Scripting.Dictionary
        ''' noting these steps are also taken
        ''' in Class_Initialize, it's tempting
        ''' to wonder why they should appear in
        ''' both places, however, the purpose
        ''' THERE is to ensure a valid setup
        ''' at the earliest possible moment.
        '''
        ''' it would likely be appropriate
        ''' to consolidate the two into a single
        ''' procedure to be called from either place.
        
        lbxSpecOps.List = dcActvOps.Keys
        lbxSpecSet.List = dcActvSet.Keys
    End If
    
    Set Using = Me
End Function

Public Function SeeUser( _
    Optional About As Scripting.Dictionary = Nothing _
) As fmIfcTest04A
    ''' REV[2022.03.17.1308]
    '''     disabling If-Then-Else blocking,
    '''     including content of If block.
    '''     '
    '''     Only active sections of Else
    '''     block to remain active.
    '''     '
    '''     Since majority of process formerly
    '''     performed here is now addressed by
    '''     new Function Using, it should now
    '''     be sufficient to call that Function
    '''     for preparation, and then present
    '''     the UserForm for user's response.
    '''     '
    'If About Is Nothing Then
    '    Set SeeUser = SeeUser(nuDcPopulator( _
    '        ).Setting(kyInitList, dcInitList _
    '    ).Dictionary)
    'Else
        ''' REV[2022.03.17.1258] -- IMPORTANT!
        '''     the following section, copied
        '''     to new Method Function Using
        '''     (see above) has been disabled
        '''     here pending removal
        'Set dcInitList = dcOb(About.Item(kyInitList))
        '
        'Set dcActvOps = dcInitList
        'Set dcActvSet = New Scripting.Dictionary
        ''' noting these steps are also taken
        ''' in Class_Initialize, it's tempting
        ''' to wonder why they should appear in
        ''' both places, however, the purpose
        ''' THERE is to ensure a valid setup
        ''' at the earliest possible moment.
        '''
        ''' it would likely be appropriate
        ''' to consolidate the two into a single
        ''' procedure to be called from either place.
        '
        'lbxSpecOps.List = dcActvOps.Keys
        'lbxSpecSet.List = dcActvSet.Keys
        '''
        ''' REV[2022.03.17.1258] ENDS HERE
        
        ''' REV[2022.03.17.1301]
        '''     implementation of Method Function
        '''     Using, having taken over the steps
        '''     disabled immediately above, is now
        '''     called in their stead. Separation
        '''     of that sequence into its own
        '''     Function enables the preparation
        '''     of this Class instance without
        '''     immediately invoking the UserForm.
        With Me.Using(About)
            fm.Show 1
            
            '''
            
            Set SeeUser = .Itself
        End With
    'End If
End Function

Public Function Version() As String
    Version = vsnString
End Function

Private Function clsAddSpec(sp As String) As Long
    Dim rt As Long
    
    rt = 0
    If dcActvOps.Exists(sp) Then
        If dcActvSet.Exists(sp) Then
            rt = 1 'spec already set
            Stop '''
        Else
            dcActvSet.Add sp, sp
        End If
        lbxSpecSet.List = dcActvSet.Keys
        
        Set dcActvOps = dcSpecSubsetWith(sp, dcActvOps)
        lbxSpecOps.List = dcActvOps.Keys
        Debug.Print ; 'Breakpoint Landing
    Else
        rt = 2
        Stop '''
    End If
    
    clsAddSpec = rt
End Function

Private Function clsDropSpec(sp As String) As Long
    Dim rt As Long
    Dim ky As Variant
    
    rt = 0
    
    dcActvSet.Remove sp
    
    ' first attempt to reinitialize
    ' dcActvOps from dcInitList
    Set dcActvOps = dcSpecSubsetWithAll( _
        dcActvSet, dcInitList _
    )
    
    ' that SHOULD have left sp back
    ' in dcActvOps. if not there,
    ' try the FULL set dcSpecPairs
    If Not dcActvOps.Exists(sp) Then
    Set dcActvOps = dcSpecSubsetWithAll( _
        dcActvSet, dcSpecPairs _
    )
    End If
    
    ' check once more to maks sure it's in
    ' if not, we've got a REAL problem.
    If Not dcActvOps.Exists(sp) Then
        rt = 1
        Stop
    End If
    
    lbxSpecSet.List = dcActvSet.Keys
    lbxSpecOps.List = dcActvOps.Keys
    ''' this might be more flexibly implemented
    ''' in a separate function or procedure
    
    clsDropSpec = rt
End Function

Private Sub lbxSpecOps_DblClick( _
    ByVal Cancel As MSForms.ReturnBoolean _
)
    'Dim sp As String
    Dim ck As Long
    
    ck = clsAddSpec(lbxSpecOps.Value)
    If ck Then 'we got an error
        Stop
    End If
    
    'sp = lbxSpecOps.Value
    'If dcActvOps.Exists(sp) Then
    '    If dcActvSet.Exists(sp) Then
    '        Stop '''
    '    Else
    '        dcActvSet.Add sp, sp
    '    End If
    '    lbxSpecSet.List = dcActvSet.Keys
    '
    '    Set dcActvOps = dcSpecSubsetWith(sp, dcActvOps)
    '    lbxSpecOps.List = dcActvOps.Keys
    '    Debug.Print ; 'Breakpoint Landing
    'Else
    '    Stop '''
    'End If
End Sub

Private Sub lbxSpecSet_DblClick( _
    ByVal Cancel As MSForms.ReturnBoolean _
)
    Dim sp As String
    Dim ky As Variant
    
    sp = lbxSpecSet.Value
    dcActvSet.Remove sp
    lbxSpecSet.List = dcActvSet.Keys
    
    Debug.Print ; 'Breakpoint Landing
    ''' NOTE: this section resets dcActvOps to
    ''' the original dcInitList (NOT dcSpecPairs)
    ''' and then sequentially re-applies the
    ''' active terms remaining in dcActvSet.
    Set dcActvOps = dcInitList
    For Each ky In dcActvSet.Keys
        Set dcActvOps = dcSpecSubsetWith( _
            CStr(ky), dcActvOps _
        )
    Next
    lbxSpecOps.List = dcActvOps.Keys
    ''' this might be more flexibly implemented
    ''' in a separate function or procedure
    
    Debug.Print ; 'Breakpoint Landing
End Sub

Private Sub lbxSpecOps_Change()
    'Stop '''
End Sub

Private Sub lbxSpecOps_Click()
    ''' Stop '''
End Sub

Private Sub lbxSpecOps_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Stop '''
End Sub

Private Sub lbxSpecSet_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Stop '''
End Sub

Private Sub lbxSpecSet_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal Action As MSForms.fmAction, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Stop '''
End Sub

Private Sub lbxSpecSet_Change()
    'Stop '''
End Sub

Private Sub lbxSpecSet_Click()
    ''' Stop '''
End Sub

Private Sub lbxSpecSet_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Stop '''
End Sub
