

Private dcWkg As Scripting.Dictionary
Private dcFiled As Scripting.Dictionary
Private dcIndex As Scripting.Dictionary

Private fm As fmIfcTest05A

Private Sub Class_Initialize()
    '''
    Set dcWkg = New Scripting.Dictionary
    Set dcFiled = New Scripting.Dictionary
    Set dcIndex = New Scripting.Dictionary
    
    Set fm = New fmIfcTest05A
End Sub

Private Sub Class_Terminate()
    '''
    Set fm = Nothing
    
    dcWkg.RemoveAll
    dcFiled.RemoveAll
    dcIndex.RemoveAll
    
    Set dcWkg = Nothing
    Set dcFiled = Nothing
    Set dcIndex = Nothing
End Sub

Public Function Itself() As wkgCls0
    Set Itself = Me
End Function

Public Function Using( _
    Optional AiDoc As Inventor.Document = Nothing _
) As wkgCls0
    If AiDoc Is Nothing Then
        Set Using = Me
    Else
        Set Using = Collect(AiDoc)
    End If
End Function

Public Function Process( _
    Optional AiDoc As Inventor.Document = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    
    If AiDoc Is Nothing Then
        ''' THIS is where we start processing
        ''' Inventor Documents collected
        ''' in the internal Dictionary
        
        ''' at the moment, we simply pull up
        ''' the standard form and present it
        ''' with the current collection
        ''' of Inventor Documents
        With fm.Using(dcWkg)
            .Show 1
            Stop
        End With
        
        Set rt = dcCopy(dcWkg)
        ''' Right now, all this does
        ''' is copy out the internal
        ''' Dictionary as is.
    Else
        Set rt = Collect(AiDoc).Process()
    End If
    
    Set Process = rt
'Debug.Print nu_wkgCls0().Collect().Process().Count
End Function

Public Function Collect( _
    Optional AiDoc As Inventor.Document = Nothing _
) As wkgCls0 'Scripting.Dictionary
',optional dcWkg as Scripting.Dictionary=nothing
    '''
    ''' Method Function Collect
    '''
    '''     given a valid Inventor Document
    '''     (usually an assembly), gather
    '''     any and all Parts in it into
    '''     the internal Dictionary dcWkg,
    '''     and return a copy.
    '''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If AiDoc Is Nothing Then
        Collect ThisApplication.ActiveDocument
    ElseIf AiDoc Is ThisDocument Then
    'ElseIf AiDoc.DocumentType = kAssemblyDocumentObject Then
    'ElseIf AiDoc.DocumentType = kPartDocumentObject Then
    Else
        Set dcWkg = _
            dcAiDocGrpsByForm( _
            dcRemapByPtNum( _
            dcAiDocComponents( _
            AiDoc, , 0 _
        )))
        'With dcWkg: For Each ky In .Keys
        '    If dcFiled.Exists(ky) Then
        '        Stop 'going to have to deal
        '        'with merging two Dictionaries
        '    Else
        '        dcFiled.Add ky, .Item(ky)
        '    End If
        '
            'If dcWkg.Exists(ky) Then
            '    If .Item(ky) Is dcWkg.Item(ky) Then
            '        Stop 'should be okay
            '        'but want to be sure
            '        'this can even happen
            '    Else
            '        Stop 'to deal with part number conflict
            '        'there may be several possibilities to deal with
            '        'which will likely require a separate function.
            '    End If
            'Else
            '    dcWkg.Add ky, .Item(ky)
            'End If
        'Next: End With
    End If
    
    Set Collect = Me 'dcCopy(dcWkg)
End Function

