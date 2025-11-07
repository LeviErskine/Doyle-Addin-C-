

Public Sub Update_Genius_Properties()
    '''
    ''' use Find (Ctrl-F or H) to jump to:
    '''     WAYPOINT:UPDATE
    '''
    ''' NOTE[2023.06.20.1115]
    ''' This Sub moved to module 'app' from 'modMacros'
    ''' for better placement in Macro selection dialog.
    ''' Also renamed to add underscores, thus spacing
    ''' out the title for easier recognition.
    '''
    'Dim invProgressBar As Inventor.ProgressBar
    
    Dim fc As gnsIfcAiDoc
    Dim dc As Scripting.Dictionary
    Dim rt As Scripting.Dictionary
    Dim goAhead As VbMsgBoxResult
    Dim ActiveDoc As Document
    Dim txOut As String
    Dim ky As Variant
    Dim kyPt As Variant
    Dim ct As Long
    
    'Dim dx As Long
    
    Dim fm As fmIfcTest05A
    
    ''' Enable the following Message Box
    ''' when procedure in developmental transition.
    'goAhead = MsgBox( _
        Join(Array( _
            "Procedure Update_Genius_Properties is under MAJOR revision!", _
            "Significant changes are in effect which may", _
            "present issues in routine processing.", _
            "Watch for any problems, and be prepared", _
            "to respond appropriately." _
        ), " "), _
        vbOKOnly + vbCritical, _
        "!!!!! WARNING !!!!!" _
    )
    
    ''' Confirm User Request
    ''' to process active Document
    goAhead = vbYes ''''MsgBox( _
''''        Join(Array( _
''''            "Are you sure you want to process this document?", _
''''            "The process may require a few minutes depending on assembly size.", _
''''            "Suppressed and excluded parts will not be processed." _
''''        ), " "), _
''''        vbYesNo + vbQuestion, _
''''        "Process Document Custom iProperties" _
''''    )
    If goAhead = vbYes Then
        Set ActiveDoc = ThisApplication.ActiveDocument
        If ActiveDoc.DocumentType = kAssemblyDocumentObject Then
            ''' Check whether User wants to process main document.
            ''' Simple part/assembly collections generally
            ''' should not be processed.
            goAhead = vbYes ''''MsgBox( _
''''                Join(Array( _
''''                    "Do you want to process the primary assembly?", _
''''                    "If the main assembly document is just a collection", _
''''                    "of separate parts and assemblies to be processed,", _
''''                    "it's generally best not to include it in processing." _
''''                ), " "), _
''''                vbYesNo + vbQuestion, _
''''                "Include Main Assembly?" _
''''            )
            If goAhead = vbYes Then ct = 1
            ''' This result will be passed
            ''' to the component retrieval
            ''' process immediately following.
        ElseIf ActiveDoc.DocumentType = kPartDocumentObject Then
            goAhead = vbYes
            ct = 1
        End If
        
        ''' Collect Components for Processing
        ''' NOTE[2019.08.22]: Added call to dcRemapByPtNum
        '''     against original result to remap Keys
        '''     from file names to Part Numbers
        ''' !!!WARNING!!! This presents a significant risk
        '''     of Key collision, since different models
        '''     MIGHT be assigned the same Part Number.
        '''     This may be especially true if/when
        '''     Bolted Connections become involved.
        '''
        'Set dc = dcRemapByPtNum( _
            dcAiDocComponents( _
                ActiveDoc, , ct _
            ) _
        )
        ''' REV[2022.05.24.0956]
        ''' replacing the preceding call with the following
        ''' section to more effectively manage the collection
        ''' process, with potentially more flexibility
        With dcAiDocCompSetsByPtNum(ActiveDoc, ct) 'AiDoc
            If .Exists("") Then
                Stop 'for now
                'don't expect this situation
                'to occur frequently, so won't
                'worry about a handler just yet
            End If
            
            If .Exists(2) Then
                'THIS situation IS known to occur,
                'if not TERRIBLY frequently, so a
                'handler here is a good idea.
                '
                With dcOb(.Item(2))
                    'fortunately, we have one ready made
                    'in the dcRemapByPtNum function this
                    'section is replacing (see above).
                    Debug.Print MsgBox( _
                        Join(Array( _
                            "The following Part Numbers are", _
                            "assigned to more than one Model:", _
                            "", _
                            vbTab & Join(.Keys, vbNewLine & vbTab), _
                            "" _
                        ), vbNewLine), _
                        vbOKOnly Or vbInformation, _
                        "Duplicate Part Numbers!" _
                    ) 'with just a slight modification,
                    'this serves to notify the user
                    'just as dcRemapByPtNum did before.
                    'a more sophisticated response may
                    'eventually be called for, but for
                    'now, this will do.
                End With
            End If
            
            'and HERE is the step which ACTUALLY
            'replaces the prior version above.
            'Key 1 is guaranteed to be present
            'in the Dictionary returned, so no
            'need to check for it here.
            Set dc = dcOb(.Item(1))
        End With
        
        ''' REV[2022.03.23.1344]
        '''     disabling weedout filter of known
        '''     purchased parts. now that a display
        '''     and review form is used to present
        '''     the set of processed parts, exclusion
        '''     of parts known to be present may
        '''     prove a source of confusion
        ''' Filter out Purchased Parts Before Proceeding
        ''' (implementation pending/in-progress)
        'If ActiveDoc.DocumentType = kAssemblyDocumentObject Then
        '    goAhead = MsgBox(Join(Array( _
        '        "Some Items may already be recognized", _
        '        "as Purchased Parts in Genius. Models", _
        '        "for these components do not normally", _
        '        "require further updates.", _
        '        "", _
        '        "Would you like to skip these parts?", _
        '        "" _
        '    ), vbNewLine), vbYesNo, _
        '        "Skip Purchased Parts?" _
        '    )
        '
        '    If goAhead = vbYes Then
        '        MsgBox Join(Array( _
        '            "Purchased Parts in Genius", _
        '            "will not be processed.", _
        '            "" _
        '        ), vbNewLine), vbOKOnly, _
        '            "Skipping Known Purchased"
        '        Set dc = dcKeysMissing(dc, _
        '            dcAiPurch01fromDict( _
        '            dcRemapByPtNum(dc) _
        '        ))
        '    Else
        '        MsgBox Join(Array( _
        '            "You will be prompted to verify", _
        '            "all Purchased Parts, including", _
        '            "any already in Genius.", _
        '            "" _
        '        ), vbNewLine), vbOKOnly, _
        '        "Including ALL Purchased"
        '    End If
        'Else 'nevermind
        ''   don't need to check single parts
        ''   for purchased components
        'End If
        
        ''' REV[2022.03.14.1135]
        '''     Adding subdivision of gathered
        '''     Items into subgroups by form:
        '''     MAYB - probable R-RTM #parts : D-BAR #rawStock
        '''         #subtype #shtMetal indicates SHTM
        '''         but #invalid #flatPattern
        '''             suggests otherwise
        '''     DBAR - definite R-RTM #parts : D-BAR #rawStock
        '''         #subtype NOT #shtMetal
        '''     SHTM - D-RTM #parts : DSHEET #rawStock
        '''         #subtype #shtMetal
        '''         with #valid #flatPattern
        '''     ASSY - D/R-MTO #assemblies
        '''     PRCH - D/R-PTS/O #purchased #items
        '''     HDWR - D-HDWR #hardware #items
        'Set dc = dcAiDocGrpsByForm(dc)
        
        ''' Create a new ProgressBar object.
        ''' REV[2022.03.14.1137]
        '''     Disabling Progress Bar to avoid
        '''     complications arising from new
        '''     subdivision. MIGHT restore later.
        'ct = .Count
        'dx = 0
        'Set invProgressBar = ThisApplication.CreateProgressBar( _
            True, ct, "Progressing: " _
        )
        
        
        Set fc = New gnsIfcAiDoc
        ''' REV[2022.03.16.1318]
        
        Set fm = nu_fmIfcTest05A(dc) 'nu_fmTest05A
        ''' REV[2022.03.22.1448]
        ''' REV[2022.03.17.1324]
        ''' REV[2022.02.09.0829]
        Set rt = New Scripting.Dictionary
        
        fm.Show vbModeless
        ''' REV[2022.03.17.1354]
        ''' REV[2022.03.14.1526]
        With dcAiDocGrpsByForm(dc)
            ''' Process the full Component Collection
            For Each ky In Array( _
                "ASSY", "SHTM", "MAYB", "DBAR", "PRCH" _
            )
            ''' note how we're also skipping "HDWR" entirely
            ''' also plan on handling "PRCH" items separately
                
                ''' REV[2022.02.09.1432]
                If .Exists(ky) Then
                If dcOb(.Item(ky)).Count > 0 Then
                If fm.InGroup(CStr(ky)).GroupNow = ky Then
                    ''' REV[2022.03.22.1225]
                    ''' REV[2022.03.17.1339]
                    With dcOb(.Item(ky))
                    For Each kyPt In .Keys
                        ''' REV[2022.03.22.1246]
                        
                        ''' Update message for the progress bar
                        ''' REV[2022.03.14.1140]
                        '''     Disabling Progress Bar updates
                        '''     per REV[2022.03.14.1137] (above)
                        'dx = 1 + dx
                        'With invProgressBar
                        '    .Message _
                        '        = "Processing - " & ky _
                        '        & " - " & dx _
                        '        & "/" & ct
                        '    .UpdateProgress
                        'End With
                        
                        ''' WAYPOINT:UPDATE
                        ''' Process Genius Properties for next Item
                        ''' THIS is where ALL the magic happens!
                        
                        ''' REV[2022.03.22.1246]
                        If fm.OnItem(CStr(kyPt)).ItemNow = kyPt Then
                        
                        On Error Resume Next
                        
                        Err.Clear
                        If False Then 'stick to old method
                            rt.Add kyPt, dcGeniusProps(.Item(kyPt))
                        Else 'try the new one.
                            rt.Add kyPt, fc.Props( _
                            aiDocument(.Item(kyPt))) 'dcGeniusProps
                        End If
                        
                        If Err.Number = 0 Then
                        With dcOb(rt.Item(kyPt))
                            .Add "FORM", ky
                        End With
                        End If
                        
                        ''' REV[2022.03.22.1246]
                        Else
                            Stop
                        End If
                        
                        ''' REV[2022.03.16.1330]
                    Next: End With
                    Debug.Print ; 'Breakpoint Landing
                Else
                    Stop
                End If
                End If
                End If
                ''' REV[2022.03.22.1225]
            Next
            
            Debug.Print ; 'Breakpoint Landing
        End With
        
        ''' REV[2022.02.09.0844]
        With rt
            ''' Dump all Processing Results
            ''' REVISION[2021.08.09]
            For Each ky In .Keys
                Set .Item(ky) = dcPropVals( _
                    dcOb(.Item(ky)) _
                )
            Next
        End With
        
        ''' NOTE!![2022.03.14.1522]
        ''' ''' following section to matching NOTE!! below
        ''' ''' will require review to address changes
        ''' ''' resulting from addition of FORM
        ''' ''' subleveling
        ''' REV[2022.02.09.0847]
        Set dc = dcKeysMissing(dc, rt)
        
        Set rt = dcRecSetDcDx4json( _
            dcDxFromRecSetDc(rt) _
        )
        
        ''' REV[2022.02.09.0847]
        If dc.Count > 0 Then
            Stop 'so we can check how this is going to work
            rt.Add "NOTPROCESSED", dc.Keys
        End If
        ''' NOTE!![2022.03.14.1522]
        ''' section above ENDS here
        
        txOut = ConvertToJson(Array( _
            "[[ DELETE THIS PLACEHOLDER (KEEP COMMA IF NEEDED) ]]" _
        , rt), vbTab) 'dc
        ''' REV[2022.02.09.1035]
        'Debug.Print txOut
        
''''        goAhead = MsgBox( _
''''            Join(Array( _
''''                "Assembly Name:", _
''''                ActiveDoc.DisplayName, _
''''                "Process Completed", _
''''                "", _
''''                "Copy report text", _
''''                "(JSON format)", _
''''                "to Clipboard?", _
''''                "", _
''''                "(Cancel for Debug)" _
''''            ), vbNewLine _
''''            ), vbYesNoCancel, "Update Complete" _
''''        )
''''        If goAhead = vbCancel Then
''''            Stop
''''        ElseIf goAhead = vbYes Then
''''            send2clipBdWin10 txOut 'send2clipBd
''''        End If
''''    Else 'do nothing
    End If
''' '
''' REV[2022.03.23.1142]
''' removed descriptive text of revisions listed below
''' full text archived and available in archived source:
''' notes_2022-0323_vbaSrc#UpdateGeniusProperties_01.vb
''' '
''' REV entries affected
''' REV[2022.03.16.1318]
''' REV[2022.03.22.1448]
''' REV[2022.03.17.1324]
''' REV[2022.02.09.0829]
''' REV[2022.03.17.1354]
''' REV[2022.03.14.1526]
''' REV[2022.02.09.1421] (completely gone)
''' REV[2022.02.09.1432]
''' REV[2022.03.22.1225]
''' REV[2022.03.17.1339]
''' REV[2022.03.22.1246]
''' REV[2022.03.14.1140]
''' REV[2022.03.22.1246]
''' REV[2022.03.16.1330]
''' REV[2022.03.22.1225]
''' REV[2022.02.09.0844]
''' REVISION[2021.08.09]
''' REV[2022.02.09.0847]
''' REV[2022.02.09.0847]
''' REV[2022.02.09.1035]
''' '
End Sub

Public Sub Update_iPtAssy_Genius_Props()
    Dim md As Inventor.Document
    Dim dc As Scripting.Dictionary
    Dim ck As VbMsgBoxResult
    
    Set md = aiDocActive()
    Set dc = gnsUpdtAll_iFact(compDefOf(md))
    If dc.Count > 0 Then
    Else
        ck = MsgBox( _
            Join(Array( _
                "", _
                "" _
            ), vbNewLine), _
            vbOKOnly, _
            "" _
        )
    End If
    
    'If TypeOf md Is Inventor.PartDocument Then
    '    Set dc = gnsUpdtAll_iPart(aiDocPart(md).ComponentDefinition)
    'ElseIf TypeOf md Is Inventor.AssemblyDocument Then
    '    Set dc = gnsUpdtAll_iAssy(aiDocAssy(md).ComponentDefinition)
    'Else
    '    Set dc = New Scripting.Dictionary
    'End If
End Sub

'''
'''
'''
Private Function app() As String
    app = Array( _
        "module app version date 2023.06.20", _
        "" _
    )(0)
End Function
