

Public Function dgiG0t0() As Scripting.Dictionary
    Dim dcTree As Scripting.Dictionary
    Dim dcFlat As Scripting.Dictionary
    Dim nm As String
    Dim dt As String
    
    nm = nuSelAiDoc().GetReply()
    If Len(Trim(nm)) > 0 Then
        With ThisApplication.Documents
            Set dcTree = dgiAiDocClassified(.ItemByName(nm))
            dt = dgiG2f2(dgiG2f1(dcTree))
            If MsgBox( _
                "Send this text to the clipoard?" & vbNewLine & vbNewLine & dt, _
                vbYesNo + vbQuestion, _
                "Send to Clipboard?" _
            ) = vbYes Then
                On Error Resume Next
                Err.Clear
                send2clipBdWin10 dt
                If Err.Number = 0 Then
                    'MsgBox "PROMPT", vbOKOnly, "TITLE"
                    MsgBox CStr(Len(dt)) & " characters" & vbNewLine _
                    & "were copied to the clipboard.", _
                    vbOKOnly, "COPY SUCCESSFUL!"
                Else
                    If MsgBox( _
                        "Error Code " & Hex$(Err.Number) & ":" _
                        & vbNewLine & Err.Description & vbNewLine _
                        & vbNewLine & "Stop to attempt Debug?", _
                        vbYesNo, "COPY FAILED!" _
                    ) = vbYes Then
                        Stop
                    End If
                End If
                On Error GoTo 0
            Else
                MsgBox "No data sent to clipboard", _
                    vbOKOnly, "COPY CANCELED"
            End If
        End With
    End If
End Function

Public Function dgiAiDocClassified(AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    ''
    ''  Classify supplied Inventor Document
    ''  by basic Document Type. Retrieve or
    ''  generate sub Dictionary associated
    ''  with Document Type, and reference
    ''  Document there by its Full Name/Path
    ''
    Dim dt As Inventor.DocumentTypeEnum
    Dim fp As String
    'Dim st As String
    
    If dc Is Nothing Then
        Set dgiAiDocClassified = _
        dgiAiDocClassified(AiDoc, _
        New Scripting.Dictionary)
    Else
        With AiDoc
            fp = .FullDocumentName
            dt = .DocumentType
            'st = .SubType
        End With
        
        If Len(fp) > 0 Then
            With dc
                If Not .Exists(dt) Then
                    .Add dt, New Scripting.Dictionary
                End If
                With dcOb(.Item(dt))
                    If Not .Exists(fp) Then .Add fp, AiDoc
                End With
            End With
        End If
        
        If dt = kAssemblyDocumentObject Then
            Set dgiAiDocClassified = dgiMembersClassified(AiDoc, dc)
        Else
            Set dgiAiDocClassified = dc
        End If
    End If
End Function

Public Function dgiMembersClassified( _
    AiDoc As Inventor.AssemblyDocument, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    ''
    ''  Given an Assembly Document,
    ''  categorize its Components.
    ''
    Dim oc As Inventor.ComponentOccurrence
    Dim rt As Scripting.Dictionary
    
    Set rt = dc
    With AiDoc.ComponentDefinition
        For Each oc In .Occurrences
            With oc.Definition
                Set rt = dgiAiDocClassified(.Document, rt)
            End With
        Next
    End With
    Set dgiMembersClassified = rt
End Function

Public Function dgiFlatListed( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    ''
    ''  Flatten Dictionary
    ''  of Dictionaries of
    ''  Inventor Documents
    ''  into one singular
    ''  Dictionary for rescan.
    ''
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim fp As Variant
    Dim ct As Long
    
    Set rt = New Scripting.Dictionary
    
    If dc Is Nothing Then
    Else
        With dc
            For Each ky In .Keys
                With dcOb(.Item(ky))
                    For Each fp In .Keys
                        rt.Add fp, .Item(fp)
                    Next
                End With
            Next
        End With
    End If
    
    Set dgiFlatListed = rt
End Function

Public Function nuSelAiDoc( _
    Optional Default As String = "%$#@*&!" _
) As fmSelectorList
    Set nuSelAiDoc = nuSelector( _
    ).SetHdrCancel( _
        "Cancel Operation?" _
    ).SetHdrNoSelection( _
        "No Document Selected!" _
    ).SetHdrOK( _
        "Proceed With Operation?" _
    ).SetMsgCancel( _
        "No changes will be applied to any open Document." _
    ).SetMsgNoSelection(Join(Array( _
        "Do you wish to cancel the Operation?", _
        "(Click NO to return to list)" _
        ), vbNewLine) _
    ).SetMsgOK(Join(Array( _
        "The following Document(s) will be affected: ", _
        "%%%", _
        "(Click CANCEL to quit with no changes)" _
        ), vbNewLine) _
    ).WithList( _
        dcAiDocsVisible().Keys _
    ).SelectIfIn(Default)
    'lsAiDocsVisible()
    'lsWorkbooks()
End Function

Public Function AskUser4aiDoc( _
    Optional Default As Inventor.Document = Nothing, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Inventor.Document
    Dim nm As String
    
    If dc Is Nothing Then
        Set AskUser4aiDoc = AskUser4aiDoc( _
            Default, dcAiDocsVisible() _
        )
    Else
        If Default Is Nothing Then
            Set AskUser4aiDoc = AskUser4aiDoc( _
                ThisApplication.ActiveDocument, dc _
            )
        Else
            nm = d0g6f0(Default)
            With dc
                'If .Exists(nm) Then 'No, DON'T check this!
                    nm = nuSelAiDoc( _
                    ).WithList(dc.Keys _
                    ).SelectIfIn(nm _
                    ).GetReply()
                    
                    If .Exists(nm) Then
                        Set AskUser4aiDoc = .Item(nm)
                    Else
                        Set AskUser4aiDoc = Nothing
                    End If
                'Else 'we've probably got a real problem
                '     'No, the 'default' just might not be there
                '    Stop
                '    Set AskUser4aiDoc = Nothing
                'End If
            End With
        End If
    End If
End Function

'''
'''
'''
Public Function dgiG0f0() As Variant
    dgiG0f0 = Empty
End Function

Public Function dgiG0f1(AiDoc As Inventor.Document, _
    Optional dc As Scripting.Dictionary = Nothing _
) As Scripting.Dictionary
    ''
    ''  "Junk" function originally intended to
    ''  collect and categorize Inventor Documents.
    ''  See following functions for preferred approach.
    ''
    If dc Is Nothing Then
        Set dgiG0f1 = dgiG0f1(AiDoc, _
        New Scripting.Dictionary)
    Else
        With AiDoc
            If .DocumentType = kAssemblyDocumentObject Then
                If Not dc.Exists(kAssemblyDocumentObject) Then
                    dc.Add kAssemblyDocumentObject, _
                    New Scripting.Dictionary
                End If
                With dcOb(dc.Item(kAssemblyDocumentObject))
                    If Not .Exists(AiDoc.FullDocumentName) Then
                        .Add AiDoc.FullDocumentName, AiDoc
                    End If
                End With
            ElseIf .DocumentType = kPartDocumentObject Then
            Else
                Stop
                If .DocumentType = kDesignElementDocumentObject Then
                    Stop
                ElseIf .DocumentType = kDrawingDocumentObject Then
                    Stop
                ElseIf .DocumentType = kForeignModelDocumentObject Then
                    Stop
                ElseIf .DocumentType = kNoDocument Then
                    Stop
                ElseIf .DocumentType = kPresentationDocumentObject Then
                    Stop
                ElseIf .DocumentType = kSATFileDocumentObject Then
                    Stop
                ElseIf .DocumentType = kUnknownDocumentObject Then
                    Stop
                End If
            End If
        End With
        Set dgiG0f1 = dc
    End If
End Function

Public Function dgiG1f0( _
    dc As Scripting.Dictionary _
) As Long
    ''
    ''  Return the grand total count
    ''  of entries in all Dictionaries
    ''  within supplied Dictionary.
    ''
    ''  This is meant to check for
    ''  any additions to the collection
    ''  after each processing pass
    ''
    Dim ky As Variant
    Dim ct As Long
    
    ct = 0
    
    If dc Is Nothing Then
    Else
        With dc
            For Each ky In .Keys
                With dcOb(.Item(ky))
                    ct = ct + .Count
                End With
            Next
        End With
    End If
    
    dgiG1f0 = ct
End Function

Public Function dgiG1f1( _
    dc As Scripting.Dictionary, _
    Optional ck As Long = -1 _
) As Scripting.Dictionary
    ''
    ''  Build up Dictionary of Inventor
    ''  Part and Assembly Documents
    ''
    Dim AiDoc As Inventor.Document
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim fp As Variant
    Dim ct As Long
    
    If dc Is Nothing Then
        Set dgiG1f1 = New Scripting.Dictionary
    ElseIf ck < 0 Then
        Set dgiG1f1 = dgiG1f1(dc, dgiG1f0(dc))
    Else
        Set rt = dc
        With dgiFlatListed(rt)
            For Each ky In .Keys
                Set AiDoc = aiDocument(obOf(.Item(ky)))
                Set rt = dgiAiDocClassified(AiDoc, rt)
                
                If AiDoc.DocumentType _
                = kAssemblyDocumentObject Then
                    Set rt = dgiMembersClassified(AiDoc, rt)
                End If
            Next
        End With
        ct = dgiG1f0(dc)
        
        If ct > ck Then
            Set dgiG1f1 = dgiG1f1(rt, ct)
        ElseIf ct = ck Then
            Set dgiG1f1 = rt
        Else
            Stop 'cuz something went wrong
        End If
    End If
End Function

Public Function dgiG2f0( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim id As Inventor.Document
    Dim ky As Variant
    Dim sb As String
    Dim fp As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            Set id = aiDocument(.Item(ky))
            
            With id
                fp = .FullDocumentName
                sb = .SubType
            End With
            
            With rt
                If Not .Exists(sb) Then
                    .Add sb, New Scripting.Dictionary
                End If
                With dcOb(.Item(sb))
                    If Not .Exists(fp) Then .Add fp, id
                End With
            End With
        Next
    End With
    Set dgiG2f0 = rt
End Function

Public Function dgiG2f1( _
    dc As Scripting.Dictionary _
) As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim id As Inventor.Document
    Dim ky As Variant
    Dim sb As String
    Dim fp As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            rt.Add ky, dgiG2f0(dcOb(.Item(ky)))
        Next
    End With
    Set dgiG2f1 = rt
End Function

Public Function dgiG2f2( _
    dc As Scripting.Dictionary, _
    Optional pfx As String = "", _
    Optional dlm As String = "|", _
    Optional brk As String = vbNewLine _
) As String
    Dim rt As Scripting.Dictionary
    Dim ky As Variant
    Dim it As Variant
    Dim tx As String
    
    Set rt = New Scripting.Dictionary
    With dc
        For Each ky In .Keys
            tx = CStr(ky)
            If Len(pfx) > 0 Then tx = pfx & dlm & tx
            If TypeOf .Item(ky) Is Scripting.Dictionary Then
                rt.Add dgiG2f2( _
                    dcOb(obOf(.Item(ky))), _
                    tx, dlm, brk _
                ), 0
            Else 'If TypeOf .Item(ky) Is Inventor.Document Then
                rt.Add tx, 0
            'Else
                'Stop
            End If
        Next
    End With
    dgiG2f2 = Join(rt.Keys, brk)
End Function
'Debug.Print dgiG2f2(dgiG2f1(dgiAiDocClassified(ThisApplication.Documents.VisibleDocuments(3))))
'send2clipBd dgiG2f2(dgiG2f1(dgiAiDocClassified(ThisApplication.Documents.VisibleDocuments(3))))
'send2clipBd dgiG2f2(dgiG2f1(dgiAiDocClassified(ThisApplication.Documents.ItemByName(nuSelAiDoc().getReply()))))

