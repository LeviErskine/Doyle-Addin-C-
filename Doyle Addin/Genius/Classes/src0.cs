class SurroundingClass
{
    public void GeniusPropertiesUpdater()
    {
        int answer;

        answer = MsgBox("Are you sure you want to process this document? The process may require a few minutes depending on assembly size. Suppressed and excluded parts will not be processed", Constants.vbYesNo + Constants.vbQuestion, "Process Document Custom iProperties");

        if (answer == Constants.vbYes)
        {
            Document ActiveDoc;
            ActiveDoc = ThisApplication.ActiveDocument;

            if (ActiveDoc.DocumentType == kAssemblyDocumentObject)
            {

                // the active assembly.
                AssemblyDocument invAsmDoc;
                invAsmDoc = ThisApplication.ActiveDocument;
                // MsgBox ("Assembly Name: " & invAsmDoc.DisplayName)

                // Call IterateAssy(invAsmDoc.ComponentDefinition.Occurrences, 1)

                MsgBox("Process completed");
            }
            else
            {
                PartDocument invPartDoc;
                invPartDoc = ThisApplication.ActiveDocument;

                // Call IteratePart(invPartDoc)

                MsgBox("Process completed: Part " + invPartDoc.DisplayName + " processed");
            }
        }
        else
        {
        }
    }
    // 
    /// 

    public long IterateAssyRevA0(ComponentOccurrences Occurences, int Level)
    {
        // Iterate through the assembly
        ComponentOccurrence invOcc;
        long TotalStepCount;
        long CurrentStepCount;
        ProgressBar invProgressBar;
        PartDocument invDoc;
        SheetMetalComponentDefinition invSheetMetalComp;
        PropertySet invCustomPropertySet;
        double invSheetMetalMass;
        invGeniusMassProperty; string invGeniusMaterial;
        string invSheetMetalName;
        string invSheetMetalMaterial;
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invRMProperty As Property

 */
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invRMUOMProperty As Property

 */
        FlatPattern invFlatPattern;
        Box oExtent;
        double dLength;
        double dWidth;
        double dArea;
        UnitsOfMeasure oUOM;
        string strWidth;
        string strLength;
        string strArea;
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invRMQTYProperty As Property

 */
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invWidthProperty As Property

 */
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invLengthProperty As Property

 */
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invAreaProperty As Property

 */
        string Family;
        PropertySet invDesignInfo;
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invCostCenterProperty As Property

 */
        PartComponentDefinition invPartDocComp;
        double invPartMass;
        PropertySet invCustomPartPropertySet;
        ;/* Cannot convert LocalDeclarationStatementSyntax, System.ArgumentException: An item with the same key has already been added. Key: 0|0
   at System.Collections.Generic.Dictionary`2.TryInsert(TKey key, TValue value, InsertionBehavior behavior)
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.WithDelegateToParentAnnotation(SyntaxToken lastSourceToken, SyntaxToken destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 157
   at ICSharpCode.CodeConverter.Shared.TriviaConverter.PortConvertedTrivia[T](SyntaxNode sourceNode, T destination) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/Shared/TriviaConverter.cs:line 41
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.ConvertDeclaratorType(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 59
   at ICSharpCode.CodeConverter.CSharp.CommonConversions.SplitVariableDeclarations(VariableDeclaratorSyntax declarator) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommonConversions.cs:line 38
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitLocalDeclarationStatement(LocalDeclarationStatementSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 68
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
    Dim invGeniusPartMassProperty As Property

 */    // Dim Family as String
       // Dim invDesignInfo As PropertySet
       // Dim invCostCenterProperty As Property

        foreach (var invOcc in Occurences)
        {
            // MsgBox ("TYPE: " & invOcc.Definition.Type & vbNewLine & "VISIBLE: " & invOcc.Visible & vbNewLine & "NAME: " & invOcc.Name & vbNewLine & "Suboccurence: " & invOcc.SubOccurrences.Count & vbNewLine & "Occurence Type: " & invOcc.Definition.Occurrences.Type & vbNewLine & "BOMStructure: " & invOcc.BOMStructure)
            // Remove suppressed and excluded parts from the process
            if (invOcc.Definition.Type != kAssemblyComponentDefinitionObject & invOcc.Definition.Type != kWeldmentComponentDefinitionObject)
            {
                if (false & invOcc.Visible & !invOcc.Suppressed & !invOcc.Excluded & invOcc.Definition.Type != kWeldsComponentDefinitionObject)
                {
                    // -------------------------------'
                    // Create Progress Bar Information
                    // -------------------------------'

                    // Define Total Steps
                    TotalStepCount = Occurences.Count;

                    // Define Current Step
                    CurrentStepCount = CurrentStepCount + 1;

                    // Create a new ProgressBar object.
                    invProgressBar = ThisApplication.CreateProgressBar(true, TotalStepCount, "Progressing: ");

                    // the message for the progress bar
                    invProgressBar.Message = "Processing - " + invOcc.Name + " - " + CurrentStepCount + "/" + TotalStepCount;
                    invProgressBar.UpdateProgress();

                    // the active part document.
                    invDoc = invOcc.Definition.Document;

                    // -------------------'
                    // Check if SheetMetal'
                    // -------------------'
                    if (false & invDoc.SubType == guidSheetMetal)
                    {
                        invSheetMetalComp = invDoc.ComponentDefinition;

                        // the custom property set.
                        invCustomPropertySet = invDoc.PropertySets.Item("Inventor User Defined Properties");

                        // Request #1:  the Mass in Pounds and add to Custom Property GeniusMass
                        invSheetMetalMass = Round(invSheetMetalComp.MassProperties.Mass * cvMassKg2LbM, 4);

                        // Attempt to get an existing custom property named "GeniusMass".

                        invGeniusMassProperty = invCustomPropertySet.Item(pnMass);
                        if (Information.Err.Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add(invSheetMetalMass, pnMass);
                        else
                            // Got the property so update the value.
                            invGeniusMassProperty.Value = invSheetMetalMass;
                        // Request #2:  Genius SheetMetal by matching Style Name and Material. Add to Custom Property RM


                        invSheetMetalName = invSheetMetalComp.ActiveSheetMetalStyle.Name;

                        invSheetMetalMaterial = invSheetMetalComp.ActiveSheetMetalStyle.Material.Name;

                        // Map combination to corresponding Genius Part Number
                        if (invSheetMetalMaterial == "Stainless Steel")
                        {
                            if (invSheetMetalName == "18 GA")
                                invGeniusMaterial = "FS-48x96x0.048";
                            else if (invSheetMetalName == "14 GA")
                                invGeniusMaterial = "FS-60x120x0.075";
                            else if (invSheetMetalName == "13 GA")
                                invGeniusMaterial = "FS-60x97x0.09";
                            else if (invSheetMetalName == "12 GA")
                                invGeniusMaterial = "FS-60x120x0.105";
                            else if (invSheetMetalName == "10 GA")
                                invGeniusMaterial = "FS-60x144x0.135";
                            else if (invSheetMetalName == "3/16\"")
                                invGeniusMaterial = "FS-60x144x0.188";
                            else if (invSheetMetalName == "1/4\"")
                                invGeniusMaterial = "FS-60x144x0.25";
                            else if (invSheetMetalName == "5/16\"")
                                invGeniusMaterial = "FS-60x144x0.313";
                            else if (invSheetMetalName == "3/8\"")
                                invGeniusMaterial = "FS-60x144x0.375";
                            else if (invSheetMetalName == "1/2\"")
                                invGeniusMaterial = "FS-60x144x0.5";
                            else
                                invGeniusMaterial = "";
                        }
                        else if (invSheetMetalMaterial == "Steel, Mild")
                        {
                            if (invSheetMetalName == "14 GA")
                                invGeniusMaterial = "FM-60x144x0.075";
                            else if (invSheetMetalName == "12 GA")
                                invGeniusMaterial = "FM-60x144x0.105";
                            else if (invSheetMetalName == "10 GA")
                                invGeniusMaterial = "FM-60x144x0.135";
                            else if (invSheetMetalName == "3/16\"")
                                invGeniusMaterial = "FM-60x144x0.188";
                            else if (invSheetMetalName == "1/4\"")
                                invGeniusMaterial = "FM-60x144x0.25";
                            else if (invSheetMetalName == "5/16\"")
                                invGeniusMaterial = "FM-60x144x0.313";
                            else if (invSheetMetalName == "3/8\"")
                                invGeniusMaterial = "FM-60x144x0.375";
                            else if (invSheetMetalName == "1/2\"")
                                invGeniusMaterial = "FM-60x144x0.5";
                            else if (invSheetMetalName == "5/8\"")
                                invGeniusMaterial = "FM-60x144x0.625";
                            else if (invSheetMetalName == "3/4\"")
                                invGeniusMaterial = "FM-60x120x0.75";
                            else if (invSheetMetalName == "1\"")
                                invGeniusMaterial = "FM-48x120x1";
                            else
                                invGeniusMaterial = "";
                        }
                        else
                            invGeniusMaterial = ""; // Mapping of material

                        // Attempt to get an existing custom property named "RM".

                        invRMProperty = invCustomPropertySet.Item(pnRawMaterial);
                        if (Information.Err.Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add(invGeniusMaterial, pnRawMaterial);
                        else
                            // Got the property so update the value.
                            invRMProperty.Value = invGeniusMaterial;

                        // Attempt to get an existing custom property named "RMUOM".

                        invRMUOMProperty = invCustomPropertySet.Item(pnRmUnit);
                        if (Information.Err.Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPropertySet.Add("FT2", pnRmUnit);
                        else
                            // Got the property so update the value.
                            invRMUOMProperty.Value = "FT2";

                        // Request #3:  sheet metal extent area and add to custom property "RMQTY"
                        invFlatPattern = invSheetMetalComp.FlatPattern;

                        // Check to see if flat exists
                        if (!invFlatPattern == null)
                        {

                            // the extent of the face.
                            oExtent = invFlatPattern.Body.RangeBox;

                            // Extract the width, length and area from the range.
                            dLength = (oExtent.MaxPoint.X - oExtent.MinPoint.X);
                            dWidth = (oExtent.MaxPoint.Y - oExtent.MinPoint.Y);
                            dArea = dLength * dWidth;

                            // Convert these values into the document units.
                            // This will result in strings that are identical
                            // to the strings shown in the Extent dialog.
                            oUOM = invDoc.UnitsOfMeasure;
                            strWidth = oUOM.GetStringFromValue(dWidth, oUOM.GetStringFromType(oUOM.LengthUnits));
                            strLength = oUOM.GetStringFromValue(dLength, oUOM.GetStringFromType(oUOM.LengthUnits));
                            strArea = oUOM.GetStringFromValue(dArea, oUOM.GetStringFromType(oUOM.LengthUnits) + "^2");

                            // Add area to custom property set
                            // Attempt to get an existing custom property named "RMQTY".

                            invRMQTYProperty = invCustomPropertySet.Item(pnRmQty);
                            if (Information.Err.Number != 0)
                                // Failed to get the property, which means it doesn't exist so we'll create it.
                                invCustomPropertySet.Add((dArea * cvArSqCm2SqFt), pnRmQty);
                            else
                                // Got the property so update the value.
                                invRMQTYProperty.Value = (dArea * cvArSqCm2SqFt);

                            // Add Width to custom property set
                            // Attempt to get an existing custom property named "Extent_Width".

                            invWidthProperty = invCustomPropertySet.Item(pnWidth);
                            if (Information.Err.Number != 0)
                                // Failed to get the property, which means it doesn't exist so we'll create it.
                                invCustomPropertySet.Add(strWidth, pnWidth);
                            else
                                // Got the property so update the value.
                                invWidthProperty.Value = strWidth;

                            // Add Length to custom property set
                            // Attempt to get an existing custom property named "Extent_Length".

                            invLengthProperty = invCustomPropertySet.Item(pnLength);
                            if (Information.Err.Number != 0)
                                // Failed to get the property, which means it doesn't exist so we'll create it.
                                invCustomPropertySet.Add(strLength, pnLength);
                            else
                                // Got the property so update the value.
                                invLengthProperty.Value = strLength;

                            // Add AreaDescription to custom property set
                            // Attempt to get an existing custom property named "Extent_Area".

                            invAreaProperty = invCustomPropertySet.Item(pnArea);
                            if (Information.Err.Number != 0)
                                // Failed to get the property, which means it doesn't exist so we'll create it.
                                invCustomPropertySet.Add(strArea, pnArea);
                            else
                                // Got the property so update the value.
                                invAreaProperty.Value = strArea;
                        }

                        // Request #4: Change Cost Center iProperty. If BOMStructure = Normal, then Family = D-MTO, else if BOMStructure = Purchased then Family = D-PTS.

                        if (invSheetMetalComp.BOMStructure == kNormalBOMStructure)
                            Family = "D-MTO";
                        else if (invSheetMetalComp.BOMStructure == kPurchasedBOMStructure)
                            Family = "D-PTS";

                        // the design tracking property set.
                        invDesignInfo = invDoc.PropertySets.Item("Design Tracking Properties");

                        // Update the Cost Center Property
                        invCostCenterProperty = invDesignInfo.Item(pnFamily);
                        invCostCenterProperty.Value = Family;
                        Family = Family; // Just put this in for a next line to run to (Ctrl-F8).
                    }
                    else
                    {
                        // the Parts Component Definition
                        invPartDocComp = invDoc.ComponentDefinition;

                        // Request #1:  the Mass in Pounds and add to Custom Property GeniusMass
                        invPartMass = Round(invPartDocComp.MassProperties.Mass * cvMassKg2LbM, 4);

                        // the custom property set.
                        invCustomPartPropertySet = invDoc.PropertySets.Item("Inventor User Defined Properties");

                        // Attempt to get an existing custom property named "GeniusMass".

                        invGeniusPartMassProperty = invCustomPartPropertySet.Item(pnMass);
                        if (Information.Err.Number != 0)
                            // Failed to get the property, which means it doesn't exist so we'll create it.
                            invCustomPartPropertySet.Add(invPartMass, pnMass);
                        else
                            // Got the property so update the value.
                            invGeniusPartMassProperty.Value = invPartMass;

                        // Request #2: Change Cost Center iProperty. If BOMStructure = Purchased and not content center, then Family = D-PTS, else Family = D-HDWR.
                        // Dim Family as String

                        if (invPartDocComp.BOMStructure == kPurchasedBOMStructure & invPartDocComp.IsContentMember == false)
                            Family = "D-PTS";
                        else
                            Family = "D-HDWR";

                        // the design tracking property set.
                        // Dim invDesignInfo As PropertySet
                        invDesignInfo = invDoc.PropertySets.Item("Design Tracking Properties");

                        // Update the Cost Center Property
                        // Dim invCostCenterProperty As Property
                        invCostCenterProperty = invDesignInfo.Item(pnFamily);
                        invCostCenterProperty.Value = Family;
                    } // Sheetmetal vs Part

                    // Terminate the progress bar.
                    invProgressBar.Close();
                }
                else
                {
                } // Visible, suppressed, excluded or Welds
            }
            else
                Debug.Print(IterateAssyRevA0(invOcc.SubOccurrences, Level + 1)); // part or assembly
        }
    }
}