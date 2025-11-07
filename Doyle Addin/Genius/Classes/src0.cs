

Public Sub GeniusPropertiesUpdater()

Dim answer As Integer

answer = MsgBox("Are you sure you want to process this document? The process may require a few minutes depending on assembly size. Suppressed and excluded parts will not be processed", vbYesNo + vbQuestion, "Process Document Custom iProperties")

If answer = vbYes Then


        Dim ActiveDoc As Document
        Set ActiveDoc = ThisApplication.ActiveDocument
        
        If ActiveDoc.DocumentType = kAssemblyDocumentObject Then
            
            ' Get the active assembly.
            Dim invAsmDoc As AssemblyDocument
            Set invAsmDoc = ThisApplication.ActiveDocument
            'MsgBox ("Assembly Name: " & invAsmDoc.DisplayName)
                       
                'Call IterateAssy(invAsmDoc.ComponentDefinition.Occurrences, 1)
                       
                MsgBox ("Process completed")
            
        Else
        
            Dim invPartDoc As PartDocument
            Set invPartDoc = ThisApplication.ActiveDocument
            
                'Call IteratePart(invPartDoc)
                
                    MsgBox ("Process completed: Part " & invPartDoc.DisplayName & " processed")
            
        
        End If
        
Else

    'do nothing
    
End If

End Sub
'
'''

Public Function IterateAssyRevA0(Occurences As ComponentOccurrences, Level As Integer) As Long
    'Iterate through the assembly
    Dim invOcc As ComponentOccurrence
    Dim TotalStepCount As Long
    Dim CurrentStepCount As Long
    Dim invProgressBar As ProgressBar
    Dim invDoc As PartDocument
    Dim invSheetMetalComp As SheetMetalComponentDefinition
    Dim invCustomPropertySet As PropertySet
    Dim invSheetMetalMass As Double
    Dim invGeniusMassProperty As Property
    Dim invGeniusMaterial As String
    Dim invSheetMetalName As String
    Dim invSheetMetalMaterial  As String
    Dim invRMProperty As Property
    Dim invRMUOMProperty As Property
    Dim invFlatPattern As FlatPattern
    Dim oExtent As Box
    Dim dLength As Double
    Dim dWidth As Double
    Dim dArea As Double
    Dim oUOM As UnitsOfMeasure
    Dim strWidth As String
    Dim strLength As String
    Dim strArea As String
    Dim invRMQTYProperty As Property
    Dim invWidthProperty As Property
    Dim invLengthProperty As Property
    Dim invAreaProperty As Property
    Dim Family As String
    Dim invDesignInfo As PropertySet
    Dim invCostCenterProperty As Property
    Dim invPartDocComp As PartComponentDefinition
    Dim invPartMass As Double
    Dim invCustomPartPropertySet As PropertySet
    Dim invGeniusPartMassProperty As Property
    'Dim Family as String
    'Dim invDesignInfo As PropertySet
    'Dim invCostCenterProperty As Property
    
    For Each invOcc In Occurences
        'MsgBox ("TYPE: " & invOcc.Definition.Type & vbNewLine & "VISIBLE: " & invOcc.Visible & vbNewLine & "NAME: " & invOcc.Name & vbNewLine & "Suboccurence: " & invOcc.SubOccurrences.Count & vbNewLine & "Occurence Type: " & invOcc.Definition.Occurrences.Type & vbNewLine & "BOMStructure: " & invOcc.BOMStructure)
        'Remove suppressed and excluded parts from the process
        If invOcc.Definition.Type <> kAssemblyComponentDefinitionObject And invOcc.Definition.Type <> kWeldmentComponentDefinitionObject Then
            If False And invOcc.Visible And Not invOcc.Suppressed And Not invOcc.Excluded And invOcc.Definition.Type <> kWeldsComponentDefinitionObject Then
                '-------------------------------'
                'Create Progress Bar Information
                '-------------------------------'
                
                'Define Total Steps
                TotalStepCount = Occurences.Count
                
                'Define Current Step
                CurrentStepCount = CurrentStepCount + 1
                
                'Create a new ProgressBar object.
                Set invProgressBar = ThisApplication.CreateProgressBar(True, TotalStepCount, "Progressing: ")
                
                ' Set the message for the progress bar
                invProgressBar.Message = "Processing - " & invOcc.Name & " - " & CurrentStepCount & "/" & TotalStepCount
                invProgressBar.UpdateProgress
                
                ' Get the active part document.
                Set invDoc = invOcc.Definition.Document
                
                '-------------------'
                'Check if SheetMetal'
                '-------------------'
                If False And invDoc.SubType = guidSheetMetal Then
                    Set invSheetMetalComp = invDoc.ComponentDefinition
                    
                    ' Get the custom property set.
                    Set invCustomPropertySet = _
                        invDoc.PropertySets.Item("Inventor User Defined Properties")
                    
                    'Request #1: Get the Mass in Pounds and add to Custom Property GeniusMass
                    invSheetMetalMass = Round(invSheetMetalComp.MassProperties.Mass * cvMassKg2LbM, 4)
                    
                    ' Attempt to get an existing custom property named "GeniusMass".
                    On Error Resume Next
                    Set invGeniusMassProperty = invCustomPropertySet.Item(pnMass)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add(invSheetMetalMass, pnMass)
                    Else
                        ' Got the property so update the value.
                        invGeniusMassProperty.Value = invSheetMetalMass
                    End If
                    'Request #2: Get Genius SheetMetal by matching Style Name and Material. Add to Custom Property RM
                    
                    
                    invSheetMetalName = invSheetMetalComp.ActiveSheetMetalStyle.Name
                    
                    invSheetMetalMaterial = invSheetMetalComp.ActiveSheetMetalStyle.Material.Name
                    
                    'Map combination to corresponding Genius Part Number
                    If invSheetMetalMaterial = "Stainless Steel" Then
                        If invSheetMetalName = "18 GA" Then
                            invGeniusMaterial = "FS-48x96x0.048"
                        ElseIf invSheetMetalName = "14 GA" Then
                            invGeniusMaterial = "FS-60x120x0.075"
                        ElseIf invSheetMetalName = "13 GA" Then
                            invGeniusMaterial = "FS-60x97x0.09"
                        ElseIf invSheetMetalName = "12 GA" Then
                            invGeniusMaterial = "FS-60x120x0.105"
                        ElseIf invSheetMetalName = "10 GA" Then
                            invGeniusMaterial = "FS-60x144x0.135"
                        ElseIf invSheetMetalName = "3/16""" Then
                            invGeniusMaterial = "FS-60x144x0.188"
                        ElseIf invSheetMetalName = "1/4""" Then
                            invGeniusMaterial = "FS-60x144x0.25"
                        ElseIf invSheetMetalName = "5/16""" Then
                            invGeniusMaterial = "FS-60x144x0.313"
                        ElseIf invSheetMetalName = "3/8""" Then
                            invGeniusMaterial = "FS-60x144x0.375"
                        ElseIf invSheetMetalName = "1/2""" Then
                            invGeniusMaterial = "FS-60x144x0.5"
                        Else
                            invGeniusMaterial = ""
                        End If
                    ElseIf invSheetMetalMaterial = "Steel, Mild" Then
                        If invSheetMetalName = "14 GA" Then
                            invGeniusMaterial = "FM-60x144x0.075"
                        ElseIf invSheetMetalName = "12 GA" Then
                            invGeniusMaterial = "FM-60x144x0.105"
                        ElseIf invSheetMetalName = "10 GA" Then
                            invGeniusMaterial = "FM-60x144x0.135"
                        ElseIf invSheetMetalName = "3/16""" Then
                            invGeniusMaterial = "FM-60x144x0.188"
                        ElseIf invSheetMetalName = "1/4""" Then
                            invGeniusMaterial = "FM-60x144x0.25"
                        ElseIf invSheetMetalName = "5/16""" Then
                            invGeniusMaterial = "FM-60x144x0.313"
                        ElseIf invSheetMetalName = "3/8""" Then
                            invGeniusMaterial = "FM-60x144x0.375"
                        ElseIf invSheetMetalName = "1/2""" Then
                            invGeniusMaterial = "FM-60x144x0.5"
                        ElseIf invSheetMetalName = "5/8""" Then
                            invGeniusMaterial = "FM-60x144x0.625"
                        ElseIf invSheetMetalName = "3/4""" Then
                            invGeniusMaterial = "FM-60x120x0.75"
                        ElseIf invSheetMetalName = "1""" Then
                            invGeniusMaterial = "FM-48x120x1"
                        Else
                            invGeniusMaterial = ""
                        End If
                    Else
                        invGeniusMaterial = ""
                    End If 'Mapping of material
                    
                    ' Attempt to get an existing custom property named "RM".
                    On Error Resume Next
                    Set invRMProperty = invCustomPropertySet.Item(pnRawMaterial)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add(invGeniusMaterial, pnRawMaterial)
                    Else
                        ' Got the property so update the value.
                        invRMProperty.Value = invGeniusMaterial
                    End If
                    
                    ' Attempt to get an existing custom property named "RMUOM".
                    On Error Resume Next
                    Set invRMUOMProperty = invCustomPropertySet.Item(pnRmUnit)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add("FT2", pnRmUnit)
                    Else
                        ' Got the property so update the value.
                        invRMUOMProperty.Value = "FT2"
                    End If
                    
                    'Request #3: Get sheet metal extent area and add to custom property "RMQTY"
                    Set invFlatPattern = invSheetMetalComp.FlatPattern
                    
                    'Check to see if flat exists
                    If Not invFlatPattern Is Nothing Then
                    
                    ' Get the extent of the face.
                    Set oExtent = invFlatPattern.Body.RangeBox
                    
                    ' Extract the width, length and area from the range.
                    dLength = (oExtent.MaxPoint.X - oExtent.MinPoint.X)
                    dWidth = (oExtent.MaxPoint.Y - oExtent.MinPoint.Y)
                    dArea = dLength * dWidth
                    
                    ' Convert these values into the document units.
                    ' This will result in strings that are identical
                    ' to the strings shown in the Extent dialog.
                    Set oUOM = invDoc.UnitsOfMeasure
                    strWidth = oUOM.GetStringFromValue(dWidth, oUOM.GetStringFromType(oUOM.LengthUnits))
                    strLength = oUOM.GetStringFromValue(dLength, oUOM.GetStringFromType(oUOM.LengthUnits))
                    strArea = oUOM.GetStringFromValue(dArea, oUOM.GetStringFromType(oUOM.LengthUnits) & "^2")
                    
                    ' Add area to custom property set
                    ' Attempt to get an existing custom property named "RMQTY".
                    On Error Resume Next
                    Set invRMQTYProperty = invCustomPropertySet.Item(pnRmQty)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add((dArea * cvArSqCm2SqFt), pnRmQty)
                    Else
                        ' Got the property so update the value.
                        invRMQTYProperty.Value = (dArea * cvArSqCm2SqFt)
                    End If
                    
                    ' Add Width to custom property set
                    ' Attempt to get an existing custom property named "Extent_Width".
                    On Error Resume Next
                    Set invWidthProperty = invCustomPropertySet.Item(pnWidth)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add(strWidth, pnWidth)
                    Else
                        ' Got the property so update the value.
                        invWidthProperty.Value = strWidth
                    End If
                    
                    ' Add Length to custom property set
                    ' Attempt to get an existing custom property named "Extent_Length".
                    On Error Resume Next
                    Set invLengthProperty = invCustomPropertySet.Item(pnLength)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add(strLength, pnLength)
                    Else
                        ' Got the property so update the value.
                        invLengthProperty.Value = strLength
                    End If
                    
                    ' Add AreaDescription to custom property set
                    ' Attempt to get an existing custom property named "Extent_Area".
                    On Error Resume Next
                    Set invAreaProperty = invCustomPropertySet.Item(pnArea)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPropertySet.Add(strArea, pnArea)
                    Else
                        ' Got the property so update the value.
                        invAreaProperty.Value = strArea
                    End If
                End If
                
                'Request #4: Change Cost Center iProperty. If BOMStructure = Normal, then Family = D-MTO, else if BOMStructure = Purchased then Family = D-PTS.
                
                If invSheetMetalComp.BOMStructure = kNormalBOMStructure Then
                    Family = "D-MTO"
                ElseIf invSheetMetalComp.BOMStructure = kPurchasedBOMStructure Then
                    Family = "D-PTS"
                End If
                
                ' Get the design tracking property set.
                Set invDesignInfo = _
                    invDoc.PropertySets.Item("Design Tracking Properties")
                
                ' Update the Cost Center Property
                Set invCostCenterProperty = invDesignInfo.Item(pnFamily)
                invCostCenterProperty.Value = Family
                Family = Family 'Just put this in for a next line to run to (Ctrl-F8).
                'Otherwise, stepping in or through previous line would run to end with no break
                
                '----------------------'
                'Else, if standard Part'
                '----------------------'
                
                Else
                    'Get the Parts Component Definition
                    Set invPartDocComp = invDoc.ComponentDefinition
                    
                    'Request #1: Get the Mass in Pounds and add to Custom Property GeniusMass
                    invPartMass = Round(invPartDocComp.MassProperties.Mass * cvMassKg2LbM, 4)
                    
                    ' Get the custom property set.
                    Set invCustomPartPropertySet = _
                        invDoc.PropertySets.Item("Inventor User Defined Properties")
                    
                    ' Attempt to get an existing custom property named "GeniusMass".
                    On Error Resume Next
                    Set invGeniusPartMassProperty = invCustomPartPropertySet.Item(pnMass)
                    If Err.Number <> 0 Then
                        ' Failed to get the property, which means it doesn't exist so we'll create it.
                        Call invCustomPartPropertySet.Add(invPartMass, pnMass)
                    Else
                        ' Got the property so update the value.
                        invGeniusPartMassProperty.Value = invPartMass
                    End If
                    
                    'Request #2: Change Cost Center iProperty. If BOMStructure = Purchased and not content center, then Family = D-PTS, else Family = D-HDWR.
                    'Dim Family as String
                    
                    If invPartDocComp.BOMStructure = kPurchasedBOMStructure And invPartDocComp.IsContentMember = False Then
                        Family = "D-PTS"
                    Else
                        Family = "D-HDWR"
                    End If
                    
                    ' Get the design tracking property set.
                    'Dim invDesignInfo As PropertySet
                    Set invDesignInfo = _
                        invDoc.PropertySets.Item("Design Tracking Properties")
                    
                    ' Update the Cost Center Property
                    'Dim invCostCenterProperty As Property
                    Set invCostCenterProperty = invDesignInfo.Item(pnFamily)
                    invCostCenterProperty.Value = Family
                End If 'Sheetmetal vs Part
                
                ' Terminate the progress bar.
                invProgressBar.Close
                
            Else
            End If 'Visible, suppressed, excluded or Welds
        Else 'assembly, iterate through next level
            Debug.Print IterateAssyRevA0(invOcc.SubOccurrences, Level + 1)
        End If 'part or assembly
    Next
    
               
End Function
'Debug.Print IterateAssyRevA0(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences, 1)

