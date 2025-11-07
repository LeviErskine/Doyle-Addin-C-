

Public Function obOf(vr As Variant) As Object
    If IsObject(vr) Then
        Set obOf = vr
    Else
        Set obOf = Nothing
    End If
End Function

Public Function dcOb(vr As Variant) As Scripting.Dictionary
    If IsObject(vr) Then
        If vr Is Nothing Then
            Set dcOb = Nothing
        ElseIf TypeOf vr Is Scripting.Dictionary Then
            Set dcOb = vr
        Else
            Set dcOb = Nothing
        End If
    Else
        Set dcOb = Nothing
    End If
End Function

Public Function fdOb(vr As Variant) As ADODB.Field
    If IsObject(vr) Then
        If TypeOf vr Is ADODB.Field Then
            Set fdOb = vr
        Else
            Set fdOb = Nothing
        End If
    Else
        Set fdOb = Nothing
    End If
End Function

Public Function aiDocument(doc As Object) As Inventor.Document
    If doc Is Nothing Then
        Set aiDocument = doc
    ElseIf TypeOf doc Is Inventor.Document Then
        Set aiDocument = doc
    Else
        Set aiDocument = Nothing
    End If
End Function
'For Each itm In ActiveDocsComponents(ThisApplication): Debug.Print aiDocument(obOf(itm)).FullFileName: Next

Public Function aiDocActive() As Inventor.Document
    Set aiDocActive = ThisApplication.ActiveDocument
End Function

Public Function aiDocPart( _
    doc As Inventor.Document _
) As Inventor.PartDocument
    If doc Is Nothing Then
        Set aiDocPart = doc
    ElseIf TypeOf doc Is Inventor.PartDocument Then
        Set aiDocPart = doc
    Else
        Set aiDocPart = Nothing
    End If
End Function

Public Function aiDocPartFromCCtr( _
    doc As Inventor.Document _
) As Inventor.PartDocument
    Dim rt As Inventor.PartDocument
    
    If doc Is Nothing Then
        Set aiDocPartFromCCtr = doc
    Else
        Set rt = aiDocPart(doc)
        If rt Is Nothing Then
            Set aiDocPartFromCCtr = rt
        ElseIf rt.ComponentDefinition.IsContentMember Then
            Set aiDocPartFromCCtr = rt
        Else
            Set aiDocPartFromCCtr = Nothing
        End If
    End If
End Function

Public Function aiDocAssy( _
    doc As Inventor.Document _
) As Inventor.AssemblyDocument
    If doc Is Nothing Then
        Set aiDocAssy = doc
    ElseIf TypeOf doc Is Inventor.AssemblyDocument Then
        Set aiDocAssy = doc
    Else
        Set aiDocAssy = Nothing
    End If
End Function

Public Function aiDocDwg( _
    doc As Inventor.Document _
) As Inventor.DrawingDocument
    If doc Is Nothing Then
        Set aiDocDwg = doc
    ElseIf TypeOf doc Is Inventor.DrawingDocument Then
        Set aiDocDwg = doc
    Else
        Set aiDocDwg = Nothing
    End If
End Function

Private Function aiCompDefinition( _
    doc As Object _
) As Inventor.ComponentDefinition
    '''
    ''' REV[2022.08.31.1313] OBSOLETED
    ''' -   no calls found to this function
    ''' -   aiCompDefOf serves same purpose
    '''     in (slightly?) more robust manner
    ''' -   changed scope to Private
    '''     to prevent future usage
    '''     outside local scope
    '''
    If TypeOf doc Is Inventor.ComponentDefinition Then
        Set aiCompDefinition = doc
    Else
        Set aiCompDefinition = Nothing
    End If
End Function

Public Function aiCompDefOf(doc As Object _
) As Inventor.ComponentDefinition 'Inventor.Document
    '''
    ''' aiCompDefOf -- Return the ComponentDefinition
    '''     of ANY Inventor Document which has one.
    ''' NOTE: currently returns ComponentDefinition objects
    '''     only from Part and Assembly Documents.
    ''' NOTE[2022.08.31.1202]: copied comments from redundant
    '''     function obAiCompDefAny prior to its deprecation
    '''
    If doc Is Nothing Then
        Set aiCompDefOf = Nothing
    ElseIf TypeOf doc Is Inventor.ComponentDefinition Then
        Set aiCompDefOf = doc
    ElseIf TypeOf doc Is Inventor.Document Then
        With aiDocument(doc)
            If .DocumentType = kAssemblyDocumentObject Then
                Set aiCompDefOf = aiDocAssy(doc).ComponentDefinition
            ElseIf .DocumentType = kPartDocumentObject Then
                Set aiCompDefOf = aiDocPart(doc).ComponentDefinition
            Else
                Set aiCompDefOf = Nothing
            End If
        End With
    Else
        Set aiCompDefOf = Nothing
    End If
End Function

Public Function obAiCompDefAny( _
    AiDoc As Inventor.Document _
) As Inventor.ComponentDefinition
    '''
    ''' obAiCompDefAny -- Return the ComponentDefinition
    '''     of ANY Inventor Document which has one.
    ''' NOTE: currently returns ComponentDefinition objects
    '''     only from Part and Assembly Documents.
    ''' NOTE[2022.08.31.1203]: rediscovered original
    '''     implementation aiCompDefOf; copied comments
    '''     there prior to deprecation of this implementation
    '''
    If AiDoc Is Nothing Then
        Set obAiCompDefAny = Nothing
    ElseIf AiDoc.DocumentType = kAssemblyDocumentObject Then
        Set obAiCompDefAny = aiDocAssy(AiDoc).ComponentDefinition
    ElseIf AiDoc.DocumentType = kPartDocumentObject Then
        Set obAiCompDefAny = aiDocPart(AiDoc).ComponentDefinition
    Else
        Set obAiCompDefAny = Nothing
    End If
'Debug.Print TypeName(obAiCompDefAny(aiDocument(userChoiceFromDc())))
End Function

Public Function aiCompDefPart(doc As Object _
) As Inventor.PartComponentDefinition
    '''
    ''' REV[2022.08.31.1247]
    ''' added ElseIf check for PartDocument
    ''' to accept Inventor Document as well
    ''' as ComponentDefinition
    ''' applied same to functions {
    '''     aiCompDefPart
    ''' }
    '''
    '''
    If doc Is Nothing Then
        Set aiCompDefPart = Nothing
    ElseIf TypeOf doc Is Inventor.PartComponentDefinition Then
        Set aiCompDefPart = doc
    ElseIf TypeOf doc Is Inventor.PartDocument Then
        Set aiCompDefPart = aiDocPart(doc).ComponentDefinition
    Else
        Set aiCompDefPart = Nothing
    End If
End Function

Public Function aiCompDefShtMetal(ob As Object _
) As Inventor.SheetMetalComponentDefinition
    If ob Is Nothing Then
        Set aiCompDefShtMetal = Nothing
    ElseIf TypeOf ob Is Inventor.SheetMetalComponentDefinition Then
        Set aiCompDefShtMetal = ob
    ElseIf TypeOf ob Is Inventor.PartDocument Then
        Set aiCompDefShtMetal = aiCompDefShtMetal( _
            aiDocPart(ob).ComponentDefinition _
        )
    Else
        Set aiCompDefShtMetal = Nothing
    End If
End Function

Public Function aiCompDefAssy(ob As Object _
) As Inventor.AssemblyComponentDefinition
    If ob Is Nothing Then
        Set aiCompDefAssy = Nothing
    ElseIf TypeOf ob Is Inventor.AssemblyComponentDefinition Then
        Set aiCompDefAssy = ob
    ElseIf TypeOf ob Is Inventor.AssemblyDocument Then
        Set aiCompDefAssy = aiDocAssy(ob).ComponentDefinition
    Else
        Set aiCompDefAssy = Nothing
    End If
End Function

Public Function aiProperty(ob As Object) As Inventor.Property
    If ob Is Nothing Then
        'Stop
        Set aiProperty = Nothing
    ElseIf TypeOf ob Is Inventor.Property Then
        Set aiProperty = ob
    Else
        'Stop 'because this is NOT a Property!
        Set aiProperty = Nothing
    End If
End Function

Public Function aiPlane(ob As Object) As Inventor.Plane
    If TypeOf ob Is Inventor.Plane Then
        Set aiPlane = ob
    Else
        Set aiPlane = Nothing
    End If
End Function

Public Function aiCompOcc( _
    ob As Object _
) As Inventor.ComponentOccurrence
    If TypeOf ob Is Inventor.ComponentOccurrence Then
        Set aiCompOcc = ob
    Else
        Set aiCompOcc = Nothing
    End If
End Function

Public Function obAiProp(ob As Object) As Inventor.Property
    If TypeOf ob Is Inventor.Property Then
        Set obAiProp = ob
    Else
        Set obAiProp = Nothing
    End If
End Function

Public Function obAiParam(ob As Object) As Inventor.Parameter
    If TypeOf ob Is Inventor.Parameter Then
        Set obAiParam = ob
    Else
        Set obAiParam = Nothing
    End If
End Function

Public Function obVbProject(ob As Object) As VBIDE.VBProject
    If TypeOf ob Is VBIDE.VBProject Then
        Set obVbProject = ob
    Else
        Set obVbProject = Nothing
    End If
End Function

Public Function obVbCodeMod(ob As Object) As VBIDE.CodeModule
    If TypeOf ob Is VBIDE.CodeModule Then
        Set obVbCodeMod = ob
    Else
        Set obVbCodeMod = Nothing
    End If
End Function

