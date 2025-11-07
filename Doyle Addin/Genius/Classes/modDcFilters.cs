

Public Function dcAiDocsVisible() As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim AiDoc As Inventor.Document

    Set rt = New Scripting.Dictionary
    For Each AiDoc In ThisApplication.Documents.VisibleDocuments
        'rt.Add aiDoc.FullDocumentName, aiDoc
        rt.Add d0g6f0(AiDoc), AiDoc
    Next
    Set dcAiDocsVisible = rt
End Function

Public Sub lsAiDocsVisible()
    Debug.Print txDumpLs(dcAiDocsVisible().Keys)
End Sub

Public Function dcAiDocsByType(dc As Scripting.Dictionary) As Scripting.Dictionary
    '''
    ''' Split Dictionary of Inventor Documents
    ''' into separate "sub" Dictionaries,
    ''' keyed by Document Type
    '''
    Dim rt As Scripting.Dictionary
    Dim gp As Scripting.Dictionary
    Dim AiDoc As Inventor.Document
    Dim tp As Inventor.DocumentTypeEnum
    Dim fn As String
    Dim ky As Variant
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set AiDoc = aiDocument(.Item(ky))
            With AiDoc
                tp = .DocumentType
                fn = .FullFileName
            End With
            
            With rt
                If .Exists(tp) Then
                    Set gp = .Item(tp)
                Else
                    Set gp = New Scripting.Dictionary
                    .Add tp, gp
                End If
            End With
            
            With gp
                If .Exists(fn) Then
                    Stop
                Else
                    .Add fn, AiDoc
                End If
            End With
        Next
    End With
    Set dcAiDocsByType = rt
End Function
'Debug.Print Join(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, ", ")

Public Function dcAiDocsOfType( _
    tp As Inventor.DocumentTypeEnum, _
    Optional dc As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' Retrieve subDictionary for
    ''' given Inventor Document type
    '''
    With dcAiDocsByType(dc)
        If .Exists(tp) Then
            Set dcAiDocsOfType = .Item(tp)
        Else
            Set dcAiDocsOfType = New Scripting.Dictionary
        End If
    End With
End Function

Public Function dcAiPartDocs( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Set dcAiPartDocs = dcAiDocsOfType(kPartDocumentObject, dc)
End Function
'Debug.Print Join(dcAiPartDocs(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbNewLine)

Public Function dcAiAssyDocs( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Set dcAiAssyDocs = dcAiDocsOfType(kAssemblyDocumentObject, dc)
End Function

Public Function dcOf_iPartFactories( _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim pt As Inventor.PartDocument
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If dc Is Nothing Then
        Set rt = dcOf_iPartFactories( _
            dcAiDocsVisible() _
        )
    Else
        Set rt = New Scripting.Dictionary
        
        With dcAiPartDocs(dc)
        For Each ky In .Keys
            Set pt = aiDocPart(.Item(ky))
            With pt
            If .ComponentDefinition.IsiPartFactory Then
                rt.Add .FullFileName, pt
            End If
            End With
        Next
        End With
    End If
    
    Set dcOf_iPartFactories = rt
End Function

Public Function dcOf_iAssyFactories( _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim sm As Inventor.AssemblyDocument
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If dc Is Nothing Then
        Set rt = dcOf_iAssyFactories( _
            dcAiDocsVisible() _
        )
    Else
        Set rt = New Scripting.Dictionary
        
        With dcAiAssyDocs(dc)
        For Each ky In .Keys
            Set sm = aiDocAssy(.Item(ky))
            With sm
            If .ComponentDefinition.IsiAssemblyFactory Then
                rt.Add .FullFileName, sm
            End If
            End With
        Next
        End With
    End If
    
    Set dcOf_iAssyFactories = rt
End Function

Public Function dcOf_iAll_Factories( _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    
    If dc Is Nothing Then
        Set rt = dcOf_iAll_Factories( _
            dcAiDocsVisible() _
        )
    Else
        Set rt = dcOf_iPartFactories(dc)
        
        With dcOf_iAssyFactories(dc)
        For Each ky In .Keys
            rt.Add ky, .Item(ky)
        Next
        End With
    End If
    
    Set dcOf_iAll_Factories = rt
End Function

Public Function dcAiSheetMetal( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim ky As Variant
    Dim pt As Inventor.PartDocument
    Dim rt As Scripting.Dictionary
    
    Set rt = New Scripting.Dictionary
    With dcAiPartDocs(dc)
        For Each ky In .Keys
            With aiDocPart(.Item(ky))
                If .DocumentSubType.DocumentSubTypeID = guidSheetMetal Then
                    rt.Add .FullFileName, .ComponentDefinition.Document
                End If
            End With
        Next
    End With
    Set dcAiSheetMetal = rt
End Function
'Debug.Print Join(dcAiSheetMetal(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbNewLine)
'
'Debug.Print dcAiPartDocs(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count
'Debug.Print dcAiSheetMetal(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count

Public Function dcAssyPartsPrimary( _
    aiAssy As Inventor.AssemblyDocument _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim oc As Inventor.ComponentOccurrence
    
    Set rt = New Scripting.Dictionary
    With aiAssy.ComponentDefinition
        For Each oc In .Occurrences
            With aiDocument(oc.Definition.Document)
                If Not rt.Exists(.FullDocumentName) Then
                    rt.Add .FullDocumentName, _
                    .PropertySets.Parent
                End If
            End With
        Next
    End With
    Set dcAssyPartsPrimary = rt
End Function

Public Function dcAiDocsByPtNum( _
    dcIn As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim pn As String
    'Dim oc As Inventor.ComponentOccurrence
    
    Set rt = New Scripting.Dictionary
    With dcIn 'aiAssy.ComponentDefinition
        For Each ky In .Keys
            With aiDocument(.Item(ky)).PropertySets.Item(gnDesign)
                pn = .Item(pnPartNum).Value
                If rt.Exists(pn) Then
                    Stop
                Else
                    rt.Add pn, .Parent.Parent
                End If
            End With
        Next
    End With
    Set dcAiDocsByPtNum = rt
End Function
'Set dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each ky In dc: Debug.Print txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.Item(ky)))).Keys, vbNewLine & vbTab): Next
'tx = "": Set dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each ky In dc: tx = tx & vbNewLine & ky & vbNewLine & vbTab & txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.Item(ky)))).Keys, vbNewLine & vbTab): Next: send2clipBd tx: Set dc = Nothing

Public Function dcItemsNotInGenius( _
    dcPts As Scripting.Dictionary _
) As Scripting.Dictionary
    '''
    ''' dcItemsNotInGenius --
    '''     takes a Dictionary of Items
    '''     (keyed by Item/Part Number)
    '''     and returns a Dictionary of
    '''     Items not yet found in Genius
    '''
    ''' NOTE: originally designed to take
    '''     a Dictionary of Inventor
    '''     Documents, it SHOULD be able
    '''     to process a Dictionary of
    '''     ANY sort of Items keyed
    '''     to Item/Part Number
    '''
    'Set dcPts = dcRemapByPtNum( _
        dcAiDocComponents(aiDoc) _
    )
    
    Set dcItemsNotInGenius _
    = dcKeysMissing(dcPts, dcOb( _
        dcDxFromRecSetDc(dcFromAdoRS( _
            cnGnsDoyle().Execute( _
            q1g1x2v2(dcPts) _
        )) _
    ).Item("Item")))
'Debug.Print txDumpLs(dcItemsNotInGenius(aiDocActive()).Keys)
End Function

Public Function dcAiPartsNotInGenius( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    '''
    ''' dcAiPartsNotInGenius --
    '''     calls dcItemsNotInGenius
    '''     against a Dictionary of Items
    '''     from supplied Inventor Document
    '''     to return a subset of Items
    '''     not yet added to Genius
    '''
    
    Set dcAiPartsNotInGenius _
    = dcItemsNotInGenius( _
        dcRemapByPtNum( _
        dcAiDocComponents( _
        AiDoc _
    )))
'Debug.Print txDumpLs(dcAiPartsNotInGenius(aiDocActive()).Keys)
End Function

Public Function mdf0g0f0( _
    AiDoc As Inventor.Document _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim bk As String
    
    bk = vbNewLine & vbTab
    Set rt = New Scripting.Dictionary
    With dcAiDocsByPtNum(dcAssyPartsPrimary(AiDoc))
        For Each ky In .Keys
            rt.Add ky & bk & txDumpLs( _
                dcAiDocsByPtNum(dcAssyPartsPrimary( _
                    aiDocument(.Item(ky)) _
                )).Keys, bk _
            ), .Item(ky)
        Next
    End With
    Set mdf0g0f0 = rt
End Function
'send2clipBd txDumpLs(mdf0g0f0(ThisApplication.ActiveDocument))

