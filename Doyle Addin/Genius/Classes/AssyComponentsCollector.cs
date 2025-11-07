
'''
''' Purpose of this module is to provide an alternate
''' method of collecting assembly components
''' using the native VBA Collection instead of
''' the Scripting Runtime's Dictionary.
''' Though less powerful/convenient, it does avoid
''' the need for a reference to the Scripting Runtime.
'''

Public Function CollectItem(Item As Variant, _
    Optional Key As Variant, _
    Optional coll As Collection = Nothing _
) As Collection
    Dim rt As Collection
    
    If coll Is Nothing Then
        Set rt = New Collection
    Else
        Set rt = coll
    End If
    
    On Error Resume Next
    With Err
        rt.Add Item, Key
        If .Number Then
            If .Number = 457 Then
                If IsObject(Item) Then
                    If IsObject(rt.Item(Key)) Then
                        If Item Is rt.Item(Key) Then
                            ''' OK! Same Object!
                        Else
                            Stop 'Different Objects!
                        End If
                    Else
                        Stop 'Object vs non-Object
                    End If
                ElseIf IsObject(rt.Item(Key)) Then
                    Stop 'Object vs non-Object
                Else
                    If Item = rt.Item(Key) Then
                        ''' OK! Equal Values!
                    Else
                        Stop 'Different Values!
                    End If
                End If
            Else
                Stop
            End If
        End If
    End With
    On Error GoTo 0
    
    Set CollectItem = rt
End Function

Public Function CollectComponents( _
    AiDoc As Inventor.Document, _
    Optional coll As Collection = Nothing _
) As Collection
    Dim aiDType As Inventor.DocumentTypeEnum
    Dim aiOcc As Inventor.ComponentOccurrence
    Dim rt As Collection
    
    If coll Is Nothing Then
        Set rt = CollectComponents(AiDoc, New Collection)
    Else
        Set rt = coll
        aiDType = AiDoc.DocumentType
        If aiDType = kAssemblyDocumentObject Then
            With aiDocAssy(AiDoc).ComponentDefinition
                For Each aiOcc In .Occurrences
                    If aiOcc.Definition.Document Is AiDoc Then 'skip it
                        'Otherwise, we'll dive into
                        'a bottomless pit of recursion.
                    Else
                        Set rt = CollectComponents( _
                            aiOcc.Definition.Document, rt _
                        )
                    End If
                Next
            End With
        ElseIf aiDType = kPartDocumentObject Then
            Set rt = CollectItem(AiDoc, AiDoc.FullFileName, rt)
        Else
            Stop 'cuz we dunno what to do with this one.
        End If
    End If
    
    Set CollectComponents = rt
End Function

Public Function ActiveDocsComponents(aiApp As Inventor.Application) As Collection
    Set ActiveDocsComponents = CollectComponents(aiApp.ActiveDocument)
End Function

Public Function strActiveDocsComponents(aiApp As Inventor.Application) As String
    Dim AiDoc As Inventor.Document
    Dim rt As String
    
    rt = ""
    For Each AiDoc In ActiveDocsComponents(aiApp)
        rt = rt & vbNewLine & AiDoc.FullFileName
    Next
    strActiveDocsComponents = rt
End Function
'Debug.Print strActiveDocsComponents(ThisApplication)

