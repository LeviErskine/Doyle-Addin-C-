class SurroundingClass
{
    public Scripting.Dictionary dcCutTimePerimeter(Inventor.Document ad, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long incTop = 0)
    {
        Scripting.Dictionary rt;
        Inventor.Document ActiveDoc;
        Inventor.Document pt;
        Variant ky;

        // If dc Is Nothing Then
        // dcCutTimePerimeter = dcCutTimePerimeter(            ad, New Scripting.Dictionary, incTop        )
        // Else
        rt = new Scripting.Dictionary();

        {
            var withBlock = dcAiDocComponents(ad, dc, incTop);
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                rt.Addpt.PropertySets.Item(gnDesign).Item(pnPartNum).Value(null/* Conversion error: Set to default value for this argument */, fpPerimeterInch(pt));
            }
        }

        dcCutTimePerimeter = rt;
    }
    // Debug.Print dumpLsKeyVal(dcCutTimePerimeter(ThisApplication.ActiveDocument))

    public long mdl1g0f0()
    {
        Scripting.Dictionary dc;
        Variant ky;
        Inventor.Document ad;
        // Dim ps As Inventor.PropertySet
        Inventor.Property pr;

        dc = dcAssyDocComponents(ThisApplication.Documents.ItemByName(@"C:\Doyle_Vault\Designs\Misc\andrewT\02\02-weldmentStd-01.iam"));
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocument(withBlock.Item(ky));
                {
                    var withBlock1 = dcGeniusProps(ad);
                    if (withBlock1.Exists(pnRawMaterial))
                    {
                        pr = ad.PropertySets(gnCustom).Item(pnRawMaterial);
                        {
                            var withBlock2 = pr;
                            Debug.Print.Value();
                            ;/* Cannot convert MultiLineIfBlockSyntax, System.NotSupportedException: LikeExpression not supported!
   at ICSharpCode.CodeConverter.CSharp.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/SyntaxKindExtensions.cs:line 278
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1415
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitMultiLineIfBlock(MultiLineIfBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 353
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                        If .Value Like "FM-*" Then
                            With New fmTest1
                                If .AskAbout(ad) = vbYes Then
                                    With .ItemData
                                        '
                                        Stop
                                        pr.Value = .Item(pnRawMaterial)
                                    End With
                                End If
                            End With
                        End If

 */
                        }
                    }
                    else
                    {
                    }
                }
            }
        }
    }
    // Debug.Print cnGnsDoyle.Execute("select I.ItemID, I.Thickness, I.Item, I.Description1 from vgMfiItems as I where I.Family='DSHEET'").GetString

    public long mdl1g1f0()
    {
        {
            var withBlock = new fmTest1();
            withBlock.AskAbout(ThisApplication.ActiveDocument);
        }
    }

    public float mdl1g1f2(MSForms.Label lb) // , txt As String
    {
        long x0;
        float x1;
        long y0;
        float y1;
        MSForms.Control ct;

        ct = lb;
        {
            var withBlock = ct;
            // x0 = .Left
            // y0 = .Top
            x1 = withBlock.Width;
            y1 = withBlock.Height;
            {
                var withBlock1 = lb;
                // .Caption = txt
                withBlock1.AutoSize = true;
                withBlock1.AutoSize = false;
            }
            withBlock.Width = x1;
            mdl1g1f2 = withBlock.Height - y1;
        }
    }

    public float mdl1g1f3(MSForms.Control ct, float byX, float byY)
    {
        {
            var withBlock = ct;
            withBlock.Left = withBlock.Left + byX;
            withBlock.Top = withBlock.Top + byY;
        }
        mdl1g1f3 = Sqr(byX * byX + byY * byY);
    }

    /// For lack of a better place to put it, creating this node

    /// The following is a basic example of accessing Parameters,

    /// such as dimensions, from an Inventor Part Document.
    public long mdl1g1f1()
    {
        {
            var withBlock = aiDocPart(ThisApplication.ActiveDocument);
            // aiDocPart casts an Inventor Part Document
            // from its general Document reference, if valid.
            {
                var withBlock1 = withBlock.ComponentDefinition.Parameters.Item("Thickness");
                Debug.Print.ExposedAsProperty();
                Debug.Print.Value();
            }
        }
    }
    /// This example was written as a quick one-off to see how

    /// an Inventor Parameter like the Thickness setting for

    /// Sheet Metal Parts might have its "Export" status

    /// modified programmatically.

    public Variant aiPropVal(Inventor.Property pr, Variant ifNot = "")
    {
        if (pr == null)
            aiPropVal = ifNot;
        else
            aiPropVal = aiPropValAux(pr.Value, ifNot);
    }
    /// This example was written as a quick one-off to see how

    /// an Inventor Parameter like the Thickness setting for

    /// Sheet Metal Parts might have its "Export" status

    /// modified programmatically.

    public Variant aiPropValAux(Variant vl, Variant ifNot = "")
    {
        if (IsObject(vl))
        {
            if (vl == null)
                aiPropValAux = ifNot;
            else if (vl is stdole.StdPicture)
            {
                aiPropValAux = "<stdole.StdPicture>";
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }
            else
            {
                System.Diagnostics.Debugger.Break(); // and see what we need to do
                aiPropValAux = "<Object:" + TypeName(vl) + ">";
            }
        }
        else
            aiPropValAux = vl;
    }

    public Inventor.Property aiPropGnsItmFamily(Inventor.Document AiDoc)
    {
        if (AiDoc == null)
            aiPropGnsItmFamily = null/* TODO Change to default(_) if this is not a reference type */;
        else
            aiPropGnsItmFamily = AiDoc.PropertySets(gnDesign).Item(pnFamily);
    }

    public Inventor.Property aiPropShtMetalThickness(Inventor.PartDocument adPart)
    {
        if (adPart == null)
            aiPropShtMetalThickness = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            var withBlock = adPart;
            if (withBlock.SubType == guidSheetMetal)
            {
                if (smThicknessExposed(withBlock.ComponentDefinition))
                    aiPropShtMetalThickness = withBlock.PropertySets(gnCustom).Item(pnThickness);
                else
                    aiPropShtMetalThickness = null/* TODO Change to default(_) if this is not a reference type */;
            }
            else
                aiPropShtMetalThickness = null/* TODO Change to default(_) if this is not a reference type */;
        }
    }

    public long smThicknessExposed(Inventor.SheetMetalComponentDefinition smDef)
    {
        if (smDef.Parameters.IsExpressionValid(pnThickness, "in"))
            smThicknessExposed = parExposed(smDef.Parameters(pnThickness), 1);
        else
            System.Diagnostics.Debugger.Break();
    }

    public long parExposed(Inventor.Parameter par, long tryTo = 0)
    {
        // '  Check Inventor Parameter for exposure as Property.
        // '  Return 0 if not, unless caller requests exposure
        // '  (tryTo <> 0). Nonzero return indicates exposed
        // '  Parameter, with sign indicating initial status.
        // '  -1 indicates Parameter already exposed
        // '  1 indicates status change to expose it.
        // '  No provision is made for failure to expose,
        // '  nor to reverse exposure status.
        {
            var withBlock = par;
            if (withBlock.ExposedAsProperty)
                parExposed = -1;
            else if (tryTo)
            {
                withBlock.ExposedAsProperty = true;
                parExposed = 1 & parExposed(par);
            }
            else
                parExposed = 0;
        }
    }

    public Scripting.Dictionary dcGnsPropsListed(Inventor.Document ad, Variant ls, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long ifNone = 1)
    {
        /// dcGnsPropsListed --
        /// Return a Dictionary of any
        /// Properties in the supplied list
        /// from the "custom" PropertySet.
        /// 
        /// Missing Property names are addressed
        /// in one of three (present) ways,
        /// based on optional argument ifNone:
        /// 0 - do not add to Dictionary
        /// missing name is missing
        /// 1 - attempt to create. failure
        /// returns Nothing, which is
        /// still not added
        /// 2 - add Nothing to Dictionary
        /// under missing name
        /// 3 - attempt to create, adding
        /// Nothing for any failures
        /// (combines options 1 and 2)
        /// 
        Inventor.PropertySet ps;
        Inventor.Property pr;
        Variant ky;
        long mkNf; // try to make any not found
        long rtNf; // return Nothing for not found
        Variant wk;

        // rt = New Scripting.Dictionary
        ps = ad.PropertySets.Item(gnCustom);

        if (dc == null)
            dcGnsPropsListed = dcGnsPropsListed(ad, ls, new Scripting.Dictionary(), ifNone);
        else if (IsArray(ls))
        {
            mkNf = ifNone & 1; // IIf(ifNone = 1, 1, 0)
            rtNf = ifNone & 2; // IIf(ifNone = 2, 1, 0)
            /// originally used IIf construct
            /// to force mapping of exact values
            /// of ifNone to corresponding behaviors.
            /// 
            /// changed to bitcode matching once clear
            /// that each bit would map exclusively
            /// to a particular behavior, and could
            /// be combined with the other, if desired.
            /// 

            foreach (var ky in ls) // Array(pnMass, pnArea, pnWidth, pnLength, pnRawMaterial, pnRmQty, pnRmUnit) ', "SPEC01", "SPEC02", "SPEC03", "SPEC04", "SPEC05", "SPEC06", "SPEC07", "SPEC08", "SPEC16"'
            {
                pr = aiGetProp(ps, System.Convert.ToHexString(ky), mkNf);
                wk = Array(pr);

                if (pr == null)
                {
                    // if supposed to return Nothings

                    if (rtNf == 0)
                        wk = Array();
                }

                if (UBound(wk) < LBound(wk))
                {
                }
                else
                {
                    if (dc.Exists(ky))
                    {
                        dc.Remove(ky);
                        /// WARNING[2021.11.19]
                        /// This was added to permit
                        /// replacement of elements
                        /// already present under a
                        /// supplied key. It might
                        /// NOT be the best way to
                        /// address this situation.
                        /// Be prepared to correct
                        /// this with a more robust
                        /// solution in future.
                        /// Meanwhile, have a
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }

                    dc.Add(ky, pr);
                }
            }
            dcGnsPropsListed = dc; // rt
        }
        else if (VarType(ls) == Constants.vbString)
            dcGnsPropsListed = dcGnsPropsListed(ad, Array(ls), dc, ifNone);
        else
            System.Diagnostics.Debugger.Break();
    }

    public Scripting.Dictionary dcGnsPropsPart(Inventor.Document ad, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long ifNone = 1)
    {
        /// dcGnsPropsPart
        /// 
        /// REV[2021.11.18]:
        /// Added pnThickness to list
        /// of Properties to return.
        /// 
        dcGnsPropsPart = dcGnsPropsListed(ad, Array(pnMass, pnArea, pnWidth, pnLength, pnThickness, pnRawMaterial, pnRmQty, pnRmUnit), dc, ifNone);
    }

    public Scripting.Dictionary dcGnsPropsAssy(Inventor.Document ad, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long ifNone = 1)
    {
        // Dim rt As Scripting.Dictionary
        // Dim ps As Inventor.PropertySet
        // Dim pr As Inventor.Property
        // Dim ky As Variant

        dcGnsPropsAssy = dcGnsPropsListed(ad, Array(pnMass, "SPEC01", "SPEC02", "SPEC03", "SPEC04", "SPEC05", "SPEC06", "SPEC07", "SPEC08", "SPEC16"), dc, ifNone);
    }

    public Scripting.Dictionary dcProps4genius(Inventor.Document ad, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long Create = 1)
    {
        // Dim rt As Scripting.Dictionary
        Inventor.PropertySet ps;
        Inventor.Property pr;
        Variant ky;

        if (dc == null)
            dcProps4genius = dcProps4genius(ad, new Scripting.Dictionary(), Create);
        else
        {
            var withBlock = ad;
            if (withBlock.DocumentType == kAssemblyDocumentObject)
                dcProps4genius = dcGnsPropsAssy(ad, dc, Create);
            else if (withBlock.DocumentType == kPartDocumentObject)
                dcProps4genius = dcGnsPropsPart(ad, dc, Create);
            else
                dcProps4genius = dc;
        }
    }

    public Inventor.WorkPlanes mdl1g2f1(Inventor.Document ad)
    {
        mdl1g2f1 = aiDocPart(ad).ComponentDefinition.WorkPlanes;
    }

    public double mdl1g3f0(Inventor.Document ad)
    {
        double rt;

        switch (ad.DocumentType)
        {
            case object _ when kPartDocumentObject:
                {
                    rt = aiDocPart(ad).ComponentDefinition.MassProperties.Mass;
                    break;
                }

            case object _ when kAssemblyDocumentObject:
                {
                    rt = aiDocAssy(ad).ComponentDefinition.MassProperties.Mass;
                    break;
                }

            default:
                {
                    rt = 0#;
                    break;
                }
        }

        {
            var withBlock = ad.UnitsOfMeasure;
            mdl1g3f0 = withBlock.ConvertUnits(rt, kKilogramMassUnits, kLbMassMassUnits); // .MassUnits)
        }
    }

    public long mdl1g4f0()
    {
        long mx;
        long dx;

        {
            var withBlock = ThisApplication.CommandManager.ControlDefinitions;
            mx = withBlock.Count;
            for (dx = 1; dx <= mx; dx++)
            {
                {
                    var withBlock1 = withBlock.Item(dx);
                    ;/* Cannot convert MultiLineIfBlockSyntax, System.NotSupportedException: LikeExpression not supported!
   at ICSharpCode.CodeConverter.CSharp.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/SyntaxKindExtensions.cs:line 278
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1415
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitMultiLineIfBlock(MultiLineIfBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 353
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                If .InternalName Like "*ault*" Then
                    Debug.Print CStr(dx) & ": " & .InternalName & "/" & .DisplayName
                End If

 */
                }
            }
        }
    }

    public Scripting.Dictionary mdl1g5f0(Inventor.Document ad)
    {
        /// The purpose of this function is to return a Dictionary
        /// of Genius Family Inventor Properties
        /// for each component Document of an assembly
        /// or a single part Document.
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAiDocComponents(ad, new Scripting.Dictionary(), 1) // sc
       ;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocument(withBlock.Item(ky));
                    rt.Add.FullFileName(null/* Conversion error: Set to default value for this argument */, aiPropGnsItmFamily(withBlock1.PropertySets.Parent));
                }
            }
        }
        mdl1g5f0 = rt;
    }

    public Scripting.Dictionary mdl1g5f1(Inventor.Document ad)
    {
        /// This function calls mdl1g5f0 to retrieve a Dictionary
        /// of Genius Family Inventor Properties, and then
        /// transforms it into a Dictionary of Dictionaries
        /// grouped by Family Property Value
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Variant ky;
        string fm;
        Inventor.Property pr;

        rt = new Scripting.Dictionary();
        {
            var withBlock = mdl1g5f0(ad);
            foreach (var ky in withBlock.Keys)
            {
                pr = aiProperty(withBlock.Item(ky));
                fm = pr.Value;
                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(fm))
                        withBlock1.Add(fm, new Scripting.Dictionary());
                    dcOb(withBlock1.Item(fm)).Add(ky, pr);
                }
            }
        }
        mdl1g5f1 = rt;
    }
    // Debug.Print txDumpLs(mdl1g5f1(ThisApplication.ActiveDocument).Keys)
    // Debug.Print txDumpLs(dcOb(mdl1g5f1(ThisApplication.ActiveDocument).Item("")).Keys)

    public Scripting.Dictionary mdl1g5f2(Inventor.Document ad)
    {
        /// The purpose of this function is to return a Dictionary
        /// of Genius Family Inventor Properties
        /// for each component Document of an assembly
        /// or a single part Document.
        Scripting.Dictionary rt;
        fmTest1 fm;
        Variant ky;

        fm = new fmTest1();
        rt = new Scripting.Dictionary();
        {
            var withBlock = mdl1g5f1(ad);
            if (withBlock.Exists(""))
            {
                {
                    var withBlock1 = dcOb(withBlock.Item(""));
                    foreach (var ky in withBlock1.Keys)
                    {
                        {
                            var withBlock2 = aiProperty(withBlock1.Item(ky));
                            if (fm.AskAbout(withBlock2.Parent) == Constants.vbOK)
                                System.Diagnostics.Debugger.Break();
                            else
                                System.Diagnostics.Debugger.Break();
                            System.Diagnostics.Debugger.Break();
                        }
                    }
                }
            }
        }
        mdl1g5f2 = rt;
    }

    public Scripting.Dictionary mdl1g5f3(Inventor.AssemblyDocument ad)
    {
        /// Scan immediate members of Assembly document
        /// and group in Dictionary by declared Part Number
        /// and sub-grouped by Full Document Name.
        /// 
        /// (I wonder if an ADO Recordset wouldn't be a better choice?)
        /// 
        Inventor.ComponentOccurrence oc;
        Inventor.Document sd;
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        // Dim fm As fmTest1
        // Dim ky As Variant
        string pn;

        // fm = New fmTest1

        rt = new Scripting.Dictionary();
        foreach (var oc in ad.ComponentDefinition.Occurrences)
        {
            sd = oc.Definition.Document; // aiDocument()
            pn = sd.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
            {
                var withBlock = rt;
                if (withBlock.Exists(pn))
                    gp = dcAiDocsByFullDocName(sd, withBlock.Item(pn));
                else
                    withBlock.Add(pn, dcAiDocsByFullDocName(sd, new Scripting.Dictionary()));
            }
        }
        mdl1g5f3 = rt;
    }
    // Debug.Print txDumpLs(mdl1g5f3(ThisApplication.ActiveDocument).Keys)

    public Scripting.Dictionary mdl1g5f4(Scripting.Dictionary dc)
    {
        /// Transform keys from supplied Dictionary
        /// (expected from mdl1g5f3)
        /// into header/member indented form.
        Scripting.Dictionary rt;
        Variant ky;
        string dl;

        dl = Constants.vbNewLine + Constants.vbTab;
        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky + Constants.vbTab + Join(dcOb(withBlock.Item(ky)).Keys, Constants.vbNewLine + ky + Constants.vbTab), withBlock.Item(ky));
        }
        mdl1g5f4 = rt;
    }
    // Debug.Print txDumpLs(mdl1g5f4(mdl1g5f3(ThisApplication.ActiveDocument)).Keys)

    public Scripting.Dictionary dcAiDocsByFullDocName(Inventor.Document ad, Scripting.Dictionary dc)
    {
        /// Add supplied Inventor Document
        /// to supplied Dictionary
        /// under its Full Document Name
        /// (supports mdl1g5f3)
        string ky;

        ky = ad.FullDocumentName;
        if (dc == null)
            dcAiDocsByFullDocName = dcAiDocsByFullDocName(ad, new Scripting.Dictionary());
        else
        {
            {
                var withBlock = dc;
                if (withBlock.Exists(ky))
                    withBlock.Item(ky) = 1 + withBlock.Item(ky);
                else
                    withBlock.Add(ky, 1);
            }
            dcAiDocsByFullDocName = dc;
        }
    }

    public Scripting.Dictionary dcAssyDocsByPtNum(Inventor.AssemblyDocument ad)
    {
        /// Derived from mdl1g5f3
        /// 
        /// Scan immediate members of Assembly Document
        /// and collect source Documents in Dictionary,
        /// grouped by declared Part Number.
        /// 
        Inventor.ComponentOccurrence oc;
        Inventor.Document sd;
        Scripting.Dictionary rt;
        string pn;

        rt = new Scripting.Dictionary();
        foreach (var oc in ad.ComponentDefinition.Occurrences)
        {
            sd = oc.Definition.Document;
            pn = sd.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
            {
                var withBlock = rt;
                if (withBlock.Exists(pn))
                {
                    if (sd == withBlock.Item(pn))
                    {
                    }
                    else
                        System.Diagnostics.Debugger.Break();// and check it out
                }
                else
                    withBlock.Add(pn, sd);
            }
        }
        dcAssyDocsByPtNum = rt;
    }
    // Debug.Print txDumpLs(dcAssyDocsByPtNum(ThisApplication.ActiveDocument).Keys)

    public Scripting.Dictionary dcAiDocsByCompList(Scripting.Dictionary dc)
    {
        /// Derived from mdl1g5f4
        /// Transform keys from supplied Dictionary
        /// (expected from dcAssyDocsByPtNum)
        /// into tab-delimited list form.
        Scripting.Dictionary rt;
        Inventor.Document sd;
        Variant ky;
        string dl;

        dl = Constants.vbNewLine + Constants.vbTab;
        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                sd = aiDocument(withBlock.Item(ky));
                {
                    var withBlock1 = sd;
                    if (withBlock1.DocumentType == kAssemblyDocumentObject)
                        rt.Add(ky + Constants.vbTab + Join(Split(txDumpLs(mdl1g5f4(mdl1g5f3(sd)).Keys), Constants.vbNewLine), Constants.vbNewLine + ky + Constants.vbTab), sd);
                    else if (withBlock1.DocumentType == kPartDocumentObject)
                    {
                        // Stop
                        {
                            var withBlock2 = dcAiPropsInSet(sd.PropertySets.Item(gnCustom));
                            if (withBlock2.Exists(pnRawMaterial))
                            {
                                dl = Trim(aiProperty(withBlock2.Item(pnRawMaterial)).Value);
                                if (Strings.Len(dl) == 0)
                                    dl = "NO_RAW_STOCK" + Constants.vbTab + "<No Raw Stock Declared>";
                                else
                                {
                                    System.Diagnostics.Debugger.Break();
                                    {
                                        var withBlock3 = cnGnsDoyle.Execute(Join(Array("select Description1", "from vgMfiItems", "where Item = '" + dl + "';"), Constants.vbNewLine));
                                        if (withBlock3.BOF | withBlock3.EOF)
                                            dl = dl + Constants.vbTab + "<Stock Number Not Found>";
                                        else
                                            dl = dl + Constants.vbTab + withBlock3.Fields(0).Value;
                                    }
                                }
                            }
                            else
                                dl = "NO_RAW_STOCK" + Constants.vbTab + "<No Raw Stock Declared>";
                        }
                        rt.Add(ky + Constants.vbTab + dl, sd);
                    }
                    else
                        rt.Add(ky + Constants.vbTab + "(UNSUPPORTED DOCUMENT TYPE)", sd);
                }
            }
        }
        dcAiDocsByCompList = rt;
    }
    // Debug.Print txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)
    // send2clipBd txDumpLs(dcAiDocsByCompList(dcAssyDocsByPtNum(ThisApplication.ActiveDocument)).Keys)

    public ADODB.Recordset rsWinUpdHist()
    {
        /// Windows Update History
        WUApiLib.IUpdateHistoryEntry it;
        ADODB.Recordset rt;
        Variant ls;

        rt = rsNewWinUpdHist;
        ls = Array("ResultCode", "Operation", "Title", "Description", "Date");
        {
            var withBlock = new WUApiLib.UpdateSession() // .CreateUpdateSearcher
       ;
            {
                var withBlock1 = withBlock.CreateUpdateSearcher;
                foreach (var it in withBlock1.QueryHistory(0, withBlock1.GetTotalHistoryCount))
                {
                    {
                        var withBlock2 = it;
                        // Debug.Print .ResultCode, .Operation, .Title, .Description, .Date
                        rt.AddNew(ls, Array(withBlock2.ResultCode, withBlock2.Operation, withBlock2.Title, withBlock2.Description, withBlock2.Date));
                    }
                }
                rt.Filter = "";
            }
        }
        rsWinUpdHist = rt;
    }

    public ADODB.Recordset rsNewWinUpdHist()
    {
        ADODB.Recordset rt;
        rt = new ADODB.Recordset();
        {
            var withBlock = rt;
            {
                var withBlock1 = withBlock.Fields;
                // .Append "", adBigInt
                // .Append "", adVarChar, 1024
                withBlock1.Append("ResultCode", adBigInt);
                withBlock1.Append("Operation", adBigInt);
                withBlock1.Append("Title", adVarChar, 256);
                withBlock1.Append("Description", adVarChar, 1024);
                withBlock1.Append("Date", adDBDate);
            }
            withBlock.Open();
        }
        rsNewWinUpdHist = rt;
    }

    public ADODB.Recordset rsShtMtlCutPars(Inventor.Document ad, long incTop = 0)
    {
        /// Windows Update History
        ADODB.Recordset rt;
        Inventor.Document ActiveDoc;
        Inventor.Document pt;
        Variant ls;
        Variant ky;

        rt = rsNewShtMtlCutPars;
        ls = Array("Item", "Description", "Thickness", "Perimeter");

        {
            var withBlock = dcAiDocComponents(ad, null/* Conversion error: Set to default value for this argument */, incTop);
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                {
                    var withBlock1 = pt.PropertySets.Item(gnDesign);
                    rt.AddNew(ls, Array(withBlock1.Item(pnPartNum).Value, withBlock1.Item(pnDesc).Value, aiPropVal(aiPropShtMetalThickness(aiDocPart(pt)), -1), fpPerimeterInch(pt)));
                }
            }
            rt.Filter = "";
        }

        rsShtMtlCutPars = rt;
    }
    // send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , "|")
    // send2clipBd rsShtMtlCutPars(ThisApplication.ActiveDocument, 1).GetString(adClipString, , vbTab)

    public ADODB.Recordset rsNewShtMtlCutPars()
    {
        ADODB.Recordset rt;

        rt = new ADODB.Recordset();
        {
            var withBlock = rt;
            {
                var withBlock1 = withBlock.Fields;
                // .Append "", adBigInt
                // .Append "", adVarChar, 1024
                // .Append "Date", adDBDate
                // 
                withBlock1.Append("Item", adVarChar, 32, adFldKeyColumn);
                withBlock1.Append("Description", adVarChar, 128);
                withBlock1.Append("Thickness", adDouble);
                withBlock1.Append("Perimeter", adDouble);
            }
            withBlock.Open();
        }

        rsNewShtMtlCutPars = rt;
    }
}