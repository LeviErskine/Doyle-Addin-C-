

Private md As Inventor.Document
Private ap As Inventor.Application
'Private pj As Inventor.DesignProject
Private ad As Inventor.ApplicationAddIn
Private il As Object 'Inventor.ApplicationAddIn
Private dc As Scripting.Dictionary

'Private mp As Inventor.NameValueMap 'for WithArgs
''' NOTE: this NameValueMap is not presently used
''' as method WithArgs is not yet implemented,
''' or even defined.
'''

Public Function RuleSource() As Inventor.Document
    Set RuleSource = md
End Function

Public Function WithRulesIn( _
    Optional AiDoc As Inventor.Document = Nothing _
) As iLogicIfc
    '''
    ''' DO NOT CALL this Method with ThisDocument
    ''' For some unknown reason, this will trigger
    ''' a Type Mismatch Error (Err 13) on exit.
    '''
    Dim rl As Variant
    
    If AiDoc Is Nothing Then
        Set md = ThisDocument
    Else
        Set md = AiDoc
    End If
    
    Set ap = md.Parent
    With ap
        With .ApplicationAddIns
            Set ad = .ItemById(guidILogicAdIn)
            With ad
                If Not .Activated Then .Activate
                dc.RemoveAll
                
                If .Activated Then
                    Set il = .Automation
                    For Each rl In il.rules(md)
                        dc.Add rl.Name, rl
                    Next
                Else
                    Set il = Nothing
                End If
            End With
        End With
        
        '''
        ''' generic iLogic interface does
        ''' not implement Vault access.
        ''' this section is therefore disabled
        '''
        'With .DesignProjectManager
        '    If .ActiveDesignProject Is Nothing Then
        '        Set pj = Nothing
        '    Else
        '        Set pj = .ActiveDesignProject
        '        If pj.ProjectType = kVaultMode Then
        '            vBase = pj.VaultVirtualPath
        '            fBase = fileIfPresent( _
        '                pj.FullFileName _
        '            ).ParentFolder.Path
        '        Else
        '            Set pj = Nothing
        '        End If
        '    End If
        '
        '    If pj Is Nothing Then
        '        vBase = ""
        '        fBase = ""
        '    End If
        'End With
    End With
    
    Set WithRulesIn = Me
End Function

Public Function RuleNames( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Variant
    If AiDoc Is Nothing Then
        RuleNames = dc.Keys
    Else
        RuleNames = WithRulesIn( _
        AiDoc).RuleNames()
    End If
End Function

Public Function RuleDefs( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim nm As Variant
    
    Set rt = New Scripting.Dictionary
    For Each nm In RuleNames(AiDoc)
        rt.Add nm, TextOf(CStr(nm))
    Next
    Set RuleDefs = rt
End Function

Public Function TextOf(ruleName As String) As String
    '''
    ''' TextOf - retrieve text of rule indicated
    '''     NOTE[2023.03.06.1424] will need error traps
    '''
    Dim rl As Object
    
    If il Is Nothing Then 'Or pj Is Nothing
        Set rl = Nothing
    Else
        On Error Resume Next
        Err.Clear
        Set rl = il.GetRule(md, ruleName)
        ''' NOTE[2023.03.06.1511] might want
        ''' to use Dictionary instead, and thus
        ''' reduce anticipated need for error traps.
        ''' NOTE, however, that a rule's name
        ''' MIGHT be changed during run time,
        ''' so the Dictionary approach
        ''' might still fail.
        If Err.Number <> 0 Then
            Stop
            Set rl = Nothing
        End If
        On Error GoTo 0
    End If
    
    If rl Is Nothing Then
        TextOf = "''' NORULE '''"
    Else
        TextOf = rl.Text()
    End If
End Function

Public Function Apply(ruleName As String, _
    Optional Args As Object = Nothing _
) As Scripting.Dictionary
    'Optional dcArgs As Scripting.Dictionary = Nothing, _
    'Optional NVMap As Inventor.NameValueMap = Nothing, _
    '
    ''' NOTE: Some iLogic Rules might behave
    '''     differently when supplied with
    '''     arguments (in a NameValueMap)
    '''     than without. For example, a
    '''     Rule which would normally add
    '''     its results to the supplied
    '''     NameValueMap might instead
    '''     present them in a message box.
    '''
    ''' If such behavior is not desired
    '''     in a call with no arguments,
    '''     an empty NameValueMap or
    '''     Dictionary may be supplied
    '''     to avoid it, provided the
    '''     Rule supports this.
    '''
    Dim mp As Inventor.NameValueMap
    Dim rt As Scripting.Dictionary
    
    If il Is Nothing Then 'Or pj Is Nothing
        Set rt = New Scripting.Dictionary
        Stop 'and debug
    ElseIf Args Is Nothing Then 'dcArgs
        il.RunRule md, ruleName 'WithArguments
        
        Set rt = New Scripting.Dictionary
        'Apply(ruleName , )
    ElseIf TypeOf Args Is Inventor.NameValueMap Then
        il.RunRuleWithArguments md, ruleName, Args 'mp
        ''' !!!WARNING!!! This call MIGHT result
        ''' in changes to supplied NameValueMap,
        ''' which might or might not be a problem
        ''' for the client process.
        '''
        ''' for now, will keep this way,
        ''' but might reconsider in future
        ''' for safety or other reasons.
        '''
        ''' UPDATE[2023.03.06.1401]
        ''' actually, this serves vital role
        ''' in receiving results FROM iLogic rule,
        ''' so this is almost certainly going
        ''' to remain as is
        
        Set rt = dcFromAiNameValMap(Args) 'mp
    ElseIf TypeOf Args Is Scripting.Dictionary Then
        Set rt = Apply(ruleName _
        , dc2aiNameValMap(Args)) 'dcArgs, NVMap
        'New Scripting.Dictionary)
        'Set mp =
    Else
        
    End If
    
    Set Apply = rt
'Debug.Print ConvertToJson(nuIfcVault().Apply("iLog01", New Scripting.Dictionary), vbTab)
End Function

'''
'''
'''
Public Function Itself() As iLogicIfc
    Set Itself = Me
End Function

Private Sub Class_Initialize()
    ''' REV[2023.03.06.1504] added Dictionary
    ''' to collect set of iLogic Rules
    Set dc = New Scripting.Dictionary
    '''
    ''' Initialization calls local method WithRulesIn
    ''' to set up private objects and variables.
    '''
    ''' Originally passed ThisDocument to initialize
    ''' from Inventor Document containing this Class.
    '''
    ''' However, for some unknown reason, a Type Mismatch
    ''' Error is triggered by the value returned by
    ''' WithRulesIn, even though it SHOULD be compatible.
    '''
    ''' After numerous failed attempts to correct the issue,
    ''' it was decided to pass Nothing to the function, which
    ''' interprets the null value as an indication to use
    ''' ThisDocument internally. THIS seems to work
    '''
    ''' One prior "solution" was to enclose the call in an
    ''' Error Trap that clears the Type Mismatch, however,
    ''' this felt too much a kludge, so this alternative
    ''' was chosen instead.
    '''
    'WithRulesIn ThisDocument
    'WithRulesIn Nothing
    WithRulesIn ThisApplication.Documents.ItemByName( _
        ThisDocument.FullDocumentName _
    )
    ''' REV[2023.03.06.1500] FIXED IT!
    ''' Instead of passing ThisDocument directly in,
    ''' a new reference to it is pulled from ThisApplication,
    ''' retrieving it from the Documents collection
    ''' by its full name. THAT actually WORKS!!!
End Sub

