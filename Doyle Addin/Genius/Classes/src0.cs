using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class src0
{
    public void GeniusPropertiesUpdater()
    {
        var answer = (int)MessageBox.Show(
            "Are you sure you want to process this document? The process may require a few minutes depending on assembly size. Suppressed and excluded parts will not be processed",
            Constants.vbYesNo & Constants.vbQuestion, "Process Document Custom iProperties");

        if (answer != (int)Constants.vbYes) return;
        Document ActiveDoc = ThisApplication.ActiveDocument;

        if (ActiveDoc.DocumentType == kAssemblyDocumentObject)
        {
            // the active assembly.
            var invAsmDoc = ThisApplication.ActiveDocument;

            // MessageBox.Show ("Assembly Name: " & invAsmDoc.DisplayName)
            // Call IterateAssy(invAsmDoc.ComponentDefinition.Occurrences, 1)
            MessageBox.Show(@"Process completed");
        }
        else
        {
            var invPartDoc = ThisApplication.ActiveDocument;

            // Call IteratePart(invPartDoc)
            MessageBox.Show(@"Process completed: Part " + invPartDoc.DisplayName + @" processed");
        }
    }
    // 
    // 

    public long IterateAssyRevA0(ComponentOccurrences Occurences, int Level)
    {
        // Iterate through the assembly

        invGeniusMassProperty;

        // Dim Family as String
        // Dim invDesignInfo As PropertySet
        // Dim invCostCenterProperty As Property

        foreach (ComponentOccurrence invOcc in Occurences)
        {
            // MessageBox.Show ("TYPE: " & invOcc.Definition.Type & vbCrLf & "VISIBLE: " & invOcc.Visible & vbCrLf & "NAME: " & invOcc.Name & vbCrLf & "Suboccurence: " & invOcc.SubOccurrences.Count & vbCrLf & "Occurence Type: " & invOcc.Definition.Occurrences.Type & vbCrLf & "BOMStructure: " & invOcc.BOMStructure)
            // Remove suppressed and excluded parts from the process
            if (invOcc.Definition.Type != kAssemblyComponentDefinitionObject &
                invOcc.Definition.Type != kWeldmentComponentDefinitionObject)
            {
                continue;
                // -------------------------------'
                // Create Progress Bar Information
                // -------------------------------'

                // Define Total Steps
                long TotalStepCount = Occurences.Count;

                // Define Current Step
                long CurrentStepCount = CurrentStepCount + 1;

                // Create a new ProgressBar dynamic.
                var invProgressBar = ThisApplication.CreateProgressBar(true, TotalStepCount, "Progressing: ");

                // the message for the progress bar
                invProgressBar.Message =
                    "Processing - " + invOcc.Name + " - " + CurrentStepCount + "/" + TotalStepCount;
                invProgressBar.UpdateProgress();

                // the active part document.
                PartDocument invDoc = invOcc.Definition.Document;

                // -------------------'
                // Check if SheetMetal'
                // -------------------'
                string Family;
                Property invCostCenterProperty;
                PropertySet invDesignInfo;
                if (false & invDoc.SubType == guidSheetMetal)
                {
                    SheetMetalComponentDefinition invSheetMetalComp = invDoc.ComponentDefinition;

                    // the custom property set.
                    var invCustomPropertySet = invDoc.PropertySets.get_Item("Inventor User Defined Properties");

                    // Request #1: the Mass in Pounds and add to Custom Property GeniusMass
                    var invSheetMetalMass = double.Round(invSheetMetalComp.MassProperties.Mass * cvMassKg2LbM, 4);

                    // Attempt to get an existing custom property named "GeniusMass".

                    invGeniusMassProperty = invCustomPropertySet.get_Item(pnMass);
                    if (Information.Err().Number != 0)
                        // Failed to get the property, which means it doesn't exist so we'll create it.
                        invCustomPropertySet.Add(invSheetMetalMass, pnMass);
                    else
                        // Got the property so update the value.
                        invGeniusMassProperty.Value = invSheetMetalMass;
                    // Request #2: Genius SheetMetal by matching Style Name and Material. Add to Custom Property RM

                    var invSheetMetalName = invSheetMetalComp.ActiveSheetMetalStyle.Name;

                    var invSheetMetalMaterial = invSheetMetalComp.ActiveSheetMetalStyle.Material.Name;

                    var invGeniusMaterial = invSheetMetalMaterial switch
                    {
                        // Map combination to corresponding Genius Part Number
                        "Stainless Steel" => invSheetMetalName switch
                        {
                            "18 GA" => "FS-48x96x0.048",
                            "14 GA" => "FS-60x120x0.075",
                            "13 GA" => "FS-60x97x0.09",
                            "12 GA" => "FS-60x120x0.105",
                            "10 GA" => "FS-60x144x0.135",
                            "3/16\"" => "FS-60x144x0.188",
                            "1/4\"" => "FS-60x144x0.25",
                            "5/16\"" => "FS-60x144x0.313",
                            "3/8\"" => "FS-60x144x0.375",
                            "1/2\"" => "FS-60x144x0.5",
                            _ => ""
                        },
                        "Steel, Mild" => invSheetMetalName switch
                        {
                            "14 GA" => "FM-60x144x0.075",
                            "12 GA" => "FM-60x144x0.105",
                            "10 GA" => "FM-60x144x0.135",
                            "3/16\"" => "FM-60x144x0.188",
                            "1/4\"" => "FM-60x144x0.25",
                            "5/16\"" => "FM-60x144x0.313",
                            "3/8\"" => "FM-60x144x0.375",
                            "1/2\"" => "FM-60x144x0.5",
                            "5/8\"" => "FM-60x144x0.625",
                            "3/4\"" => "FM-60x120x0.75",
                            "1\"" => "FM-48x120x1",
                            _ => ""
                        },
                        _ => ""
                    };

                    // Attempt to get an existing custom property named "RM".
                    var invRMProperty = invCustomPropertySet.get_Item(pnRawMaterial)
                    if (Information.Err().Number != 0)
                        // Failed to get the property, which means it doesn't exist so we'll create it.
                        invCustomPropertySet.Add(invGeniusMaterial, pnRawMaterial);
                    else
                        // Got the property so update the value.
                        invRMProperty.Value = invGeniusMaterial;

                    // Attempt to get an existing custom property named "RMUOM".

                    var invRMUOMProperty = invCustomPropertySet.get_Item(pnRmUnit)
                    if (Information.Err().Number != 0)
                        // Failed to get the property, which means it doesn't exist so we'll create it.
                        invCustomPropertySet.Add("FT2", pnRmUnit);
                    else
                        // Got the property so update the value.
                        invRMUOMProperty.Value = "FT2";

                    // Request #3: sheet metal extent area and add to custom property "RMQTY"
                    var invFlatPattern = invSheetMetalComp.FlatPattern;

                    // Check to see if flat exists
                    if (!invFlatPattern == null)
                    {
                        // the extent of the face.
                        var oExtent = invFlatPattern.Body.RangeBox;

                        // Extract the width, length and area from the range.
                        var dLength = (oExtent.MaxPoint.X - oExtent.MinPoint.X);
                        var dWidth = (oExtent.MaxPoint.Y - oExtent.MinPoint.Y);
                        var dArea = dLength * dWidth;

                        // Convert these values into the document units.
                        // This will result in strings that are identical
                        // to the strings shown in the Extent dialog.
                        var oUOM = invDoc.UnitsOfMeasure;
                        var strWidth = oUOM.GetStringFromValue(dWidth, oUOM.GetStringFromType(oUOM.LengthUnits));
                        var strLength = oUOM.GetStringFromValue(dLength, oUOM.GetStringFromType(oUOM.LengthUnits));
                        var strArea = oUOM.GetStringFromValue(dArea, oUOM.GetStringFromType(oUOM.LengthUnits) + "^2");

                        // Add area to custom property set
                        // Attempt to get an existing custom property named "RMQTY".

                        var invRMQTYProperty = invCustomPropertySet.get_Item(pnRmQty)
                        if (Information.Err().Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add((dArea * cvArSqCm2SqFt), pnRmQty);
                        else
                            // Got the property so update the value.
                            invRMQTYProperty.Value = (dArea * cvArSqCm2SqFt);

                        // Add Width to custom property set
                        // Attempt to get an existing custom property named "Extent_Width".

                        var invWidthProperty = invCustomPropertySet.get_Item(pnWidth)
                        if (Information.Err().Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add(strWidth, pnWidth);
                        else
                            // Got the property so update the value.
                            invWidthProperty.Value = strWidth;

                        // Add Length to custom property set
                        // Attempt to get an existing custom property named "Extent_Length".

                        var invLengthProperty = invCustomPropertySet.get_Item(pnLength)
                        if (Information.Err().Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add(strLength, pnLength);
                        else
                            // Got the property so update the value.
                            invLengthProperty.Value = strLength;

                        // Add AreaDescription to custom property set
                        // Attempt to get an existing custom property named "Extent_Area".

                        var invAreaProperty = invCustomPropertySet.get_Item(pnArea)
                        if (Information.Err().Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add(strArea, pnArea);
                        else
                            // Got the property so update the value.
                            invAreaProperty.Value = strArea;
                    }

                    // Request #4: Change Cost Center iProperty. If BOMStructure = Normal, then Family = D-MTO, else if BOMStructure = Purchased then Family = D-PTS.

                    Family = invSheetMetalComp.BOMStructure switch
                    {
                        kNormalBOMStructure => "D-MTO",
                        kPurchasedBOMStructure => "D-PTS",
                        _ => Family
                    };

                    // the design tracking property set.
                    invDesignInfo = invDoc.PropertySets.get_Item("Design Tracking Properties");

                    // Update the Cost Center Property
                    invCostCenterProperty = invDesignInfo.get_Item(pnFamily);
                    invCostCenterProperty.Value = Family;
                    Family = Family; // Just put this in for a next line to run to (Ctrl-F8).
                }
                else
                {
                    // the Parts Component Definition
                    var invPartDocComp = invDoc.ComponentDefinition;

                    // Request #1: the Mass in Pounds and add to Custom Property GeniusMass
                    var invPartMass = double.Round(invPartDocComp.MassProperties.Mass * cvMassKg2LbM, 4);

                    // the custom property set.
                    var invCustomPartPropertySet = invDoc.PropertySets.get_Item("Inventor User Defined Properties");

                    // Attempt to get an existing custom property named "GeniusMass".

                    var invGeniusPartMassProperty = invCustomPartPropertySet.get_Item(pnMass);
                    if (Information.Err().Number != 0)
                        // Failed to get the property, which means it doesn't exist so we'll create it.
                        invCustomPartPropertySet.Add(invPartMass, pnMass);
                    else
                        // Got the property so update the value.
                        invGeniusPartMassProperty.Value = invPartMass;

                    // Request #2: Change Cost Center iProperty. If BOMStructure = Purchased and not content center, then Family = D-PTS, else Family = D-HDWR.
                    // Dim Family as String

                    if (invPartDocComp.BOMStructure == kPurchasedBOMStructure & !invPartDocComp.IsContentMember)
                        Family = "D-PTS";
                    else
                        Family = "D-HDWR";

                    // the design tracking property set.
                    // Dim invDesignInfo As PropertySet
                    invDesignInfo = invDoc.PropertySets.get_Item("Design Tracking Properties");

                    // Update the Cost Center Property
                    // Dim invCostCenterProperty As Property
                    invCostCenterProperty = invDesignInfo.get_Item(pnFamily);
                    invCostCenterProperty.Value = Family;
                } // Sheetmetal vs Part

                // Terminate the progress bar.
                invProgressBar.Close();
                // Visible, suppressed, excluded or Welds
            }

            Debug.Print(IterateAssyRevA0(invOcc.SubOccurrences, Level + 1)); // part or assembly
        }
    }
}