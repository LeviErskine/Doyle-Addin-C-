class SurroundingClass
{
    public Variant m0g3f1(ADODB.Recordset rs)
    {
        Variant rt;
        Variant ar;
        Variant rw;
        long mx;
        long dx;
        long ct;
        long fd;

        {
            var withBlock = rs;
            withBlock.Filter = withBlock.Filter;
            if (withBlock.BOF | withBlock.EOF)
                rt = Array(Array("<NODATA>"));
            else
            {
                ct = withBlock.Fields.Count - 1;
                ar = Split(withBlock.GetString(adClipString, null/* Conversion error: Set to default value for this argument */, Constants.vbTab, Constants.vbVerticalTab), Constants.vbVerticalTab);
                mx = UBound(ar) - 1; // Last row is empty/blank
                ;/* Cannot convert RedimClauseSyntax, System.InvalidCastException: Unable to cast object of type 'Microsoft.CodeAnalysis.VisualBasic.Symbols.ExtendedErrorTypeSymbol' to type 'Microsoft.CodeAnalysis.IArrayTypeSymbol'.
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitRedimClause(RedimClauseSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 140
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
rt(mx, ct)

 */
                for (dx = 0; dx <= mx; dx++)
                {
                    rw = Split(ar(dx), Constants.vbTab);
                    for (fd = 0; fd <= ct; fd++)
                        rt(dx, fd) = rw(fd);
                }
            }
        }

        m0g3f1 = rt;
    }

    public string m0g3f2(Inventor.Document AiDoc)
    {
        Variant ar;
        long dx;
        Variant ky;

        {
            var withBlock = newFmTest1();
            withBlock.AskAbout(AiDoc);
            {
                var withBlock1 = withBlock.ItemData;
                if (withBlock1.Count > 0)
                {
                    ar = withBlock1.Keys;
                    for (dx = 0; dx <= UBound(ar); dx++) // Each ky In ar
                                                         // Debug.Print ky, .Item(ky)
                        ar(dx) = ar(dx) + "=" + withBlock1.Item(ar(dx));
                }
                else
                    ar = Array("<NODATA>");
                m0g3f2 = Join(ar, Constants.vbNewLine);
            }
        }
    }

    public long m0g2f1(Inventor.Document AiDoc)
    {
        m0g2f1 = newFmTest0().ft0g0f0(AiDoc.Thumbnail);
    }

    public long m0g2f2()
    {
        m0g2f2 = m0g2f1(ThisApplication.ActiveDocument);
    }

    public long m0g2f3()
    {
        Inventor.Document AiDoc;
        // Dim dc As Scripting.Dictionary
        Variant ky;

        // dc = New Scripting.Dictionary
        {
            var withBlock = new Scripting.Dictionary() // fmTest0
       ;
            foreach (var AiDoc in ThisApplication.Documents)
            {

                // .ft0g0f0 aiDoc.Thumbnail
                if (withBlock.Exists(AiDoc.FullFileName))
                    withBlock.Item(AiDoc.FullFileName) = 1 + withBlock.Item(AiDoc.FullFileName);
                else
                    withBlock.Add(AiDoc.FullFileName, 1);

                foreach (var ky in withBlock.Keys)
                {
                    if (withBlock.Item(ky) > 1)
                        Debug.Print.Item(ky);/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                }
            }
        }
    }

    public Scripting.Dictionary m0g1f0(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Inventor.DocumentTypeEnum tp;
        string ky;

        if (dc == null)
            rt = m0g1f0(AiDoc, new Scripting.Dictionary());
        else
        {
            {
                var withBlock = AiDoc;
                tp = withBlock.DocumentType;
                ky = withBlock.FullFileName;
            }

            rt = dc;
            if (rt.Exists(ky))
            {
            }
            else
            {
                rt.Add(ky, AiDoc);
                if (tp == kAssemblyDocumentObject)
                    rt = m0g1f0assy(AiDoc, dc);
                else if (tp == kPartDocumentObject)
                {
                }
                else
                {
                }
            }
        }

        m0g1f0 = rt;
    }
    // dc = m0g1f0(ThisApplication.ActiveDocument)
    // For Each ky In dc.Keys
    // Debug.Print aiDocument(dc.Item(ky)).PropertySets(gnDesign).Item(pnPartNum).Value, 'aiDocument(dc.Item(ky)).PropertySets(gnDesign).Item(pnMaterial).Value
    // Next

    public Scripting.Dictionary m0g1f0part(Inventor.PartDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;

        if (dc == null)
            rt = m0g1f0part(AiDoc, new Scripting.Dictionary());
        else
            rt = dc;

        m0g1f0part = rt;
    }

    public Scripting.Dictionary m0g1f0assy(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence aiOcc;

        if (dc == null)
            rt = m0g1f0assy(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;
            foreach (var aiOcc in AiDoc.ComponentDefinition.Occurrences)
                rt = m0g1f0(aiOcc.Definition.Document, rt);
        }

        m0g1f0assy = rt;
    }

    public Scripting.Dictionary dcAssyComp2A(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence invOcc;
        Inventor.ObjectTypeEnum tp;

        if (dc == null)
            rt = dcAssyComp2A(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;
            foreach (var invOcc in AiDoc.ComponentDefinition.Occurrences)
            {
                {
                    var withBlock = invOcc;
                    // Remove suppressed and excluded parts from the process
                    // Moved out here from inner checks
                    if (withBlock.Visible & !withBlock.Suppressed & !withBlock.Excluded)
                    {
                        tp = withBlock.Definition.Type;

                        if (tp != kAssemblyComponentDefinitionObjectAnd)
                            tp <> kWeldmentComponentDefinitionObjectThen
						If tp <> kWeldsComponentDefinitionObject Then
							rt = dcAddAiDoc(.Definition.Document, rt)
						End If
					Else 'assembly, check BOM Structure
						If .BOMStructure = kPurchasedBOMStructure Then 'it's purchased
							rt = dcAddAiDoc(.Definition.Document, rt)
						ElseIf .BOMStructure = kNormalBOMStructure Then 'we make it
							rt = dcAssyComp2A(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt)) 'NOT forgetting to add THIS document!
						ElseIf .BOMStructure = kInseparableBOMStructure Then 'maybe weldment?
							If tp = kWeldmentComponentDefinitionObject Then 'it is
								rt = dcAssyComp2A(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt))
							Else 'it's not
								Stop 'and see if we can figure out what its type is
							End If
						ElseIf .BOMStructure = kPhantomBOMStructure Then '"phantom" component
							'Gather its components, but NOT the document itself
							rt = dcAssyComp2A(.SubOccurrences, rt)
						Else 'not sure what we've got
							Stop 'and have a look at it
						End If
					End If 'part or assembly
				End If
			End With
		Next
	End If

	dcAssyComp2A = New Scripting.Dictionary
End Function

public Scripting.Dictionary dcAssyCompStops(Inventor.ComponentOccurrences Occurences, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, Scripting.Dictionary dcStops = null/* TODO Change to default(_) if this is not a reference type */)
{
                            /// Traverse the assembly,
                            /// including any/all subassemblies,
                            /// and collect all parts to be processed.
                            Scripting.Dictionary rt;
                            Inventor.ComponentOccurrence invOcc;
                            Inventor.ObjectTypeEnum tp;

                            if (dc == null)
                                rt = dcAssyCompStops(Occurences, new Scripting.Dictionary(), dcStops);
                            else
                            {
                                rt = dc;
                                foreach (var invOcc in Occurences)
                                {
                                    {
                                        var withBlock = invOcc;
                                        if (dcStops.Exists(aiDocument(withBlock.Definition.Document).PropertySets.Item(gnDesign).Item(pnPartNum).Value))
                                        {
                                        }

                                        // Remove suppressed and excluded parts from the process
                                        // Moved out here from inner checks
                                        if (withBlock.Visible & !withBlock.Suppressed & !withBlock.Excluded)
                                        {
                                            tp = withBlock.Definition.Type;

                                            // MsgBox Join(Array("TYPE: " & tp,"VISIBLE: " & .Visible,"NAME: " & .Name,"Suboccurence: " & .SubOccurrences.Count,"Occurence Type: " & .Definition.Occurrences.Type,"BOMStructure: " & .BOMStructure), vbNewLine)

                                            if (tp != kAssemblyComponentDefinitionObjectAnd)
                                                tp <> kWeldmentComponentDefinitionObjectThen
						'(moved suppression/exclusion check OUTSIDE)
						If tp <> kWeldsComponentDefinitionObject Then

							' rt = dcAddAiDoc(aiDocument(.Definition.Document), rt)
							''' Recasting by aiDocument not likely necessary here.
							''' Revised to following:
							rt = dcAddAiDoc(.Definition.Document, rt)

						End If 'inVisible, suppressed, excluded or Welds

					Else 'assembly, check BOM Structure
						If .BOMStructure = kPurchasedBOMStructure Then 'it's purchased
							'Just add it to the Dictionary
							rt = dcAddAiDoc(.Definition.Document, rt)
						ElseIf .BOMStructure = kNormalBOMStructure Then 'we make it
							'Gather its components
							rt = dcAssyCompStops(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt), dcStops) 'NOT forgetting to add THIS document!
						ElseIf .BOMStructure = kInseparableBOMStructure Then 'maybe weldment?
							If tp = kWeldmentComponentDefinitionObject Then 'it is
								'Treat it like an assembly
								rt = dcAssyCompStops(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt), dcStops)
							Else 'it's not
								Stop 'and see if we can figure out what its type is
							End If
						ElseIf .BOMStructure = kPhantomBOMStructure Then '"phantom" component
							'Gather its components, but NOT the document itself
							rt = dcAssyCompStops(.SubOccurrences, rt, dcStops)
						Else 'not sure what we've got
							Stop 'and have a look at it
						End If
					End If 'part or assembly
				End If
			End With
		Next
	End If
	dcAssyCompStops = rt
End Function
 */
                                        }
                                    }
                                }
                            }
                        }

public Scripting.Dictionary dcStopAt(string tx, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
                            Scripting.Dictionary rt;

                            if (dc == null)
                                rt = dcStopAt(tx, new Scripting.Dictionary());
                            else
                                rt = dc;

                            {
                                var withBlock = rt;
                                if (!withBlock.Exists(tx))
                                    withBlock.Add(tx, tx);
                            }

                            dcStopAt = rt;
                        }

    public Scripting.Dictionary m0g0f0()
    {
        Scripting.Dictionary rt;
        Variant tx;

        rt = new Scripting.Dictionary();
        foreach (var tx in Array("29-197A", "29-197B", "29-633", "29-634", "29-637", "29-638", "29-647", "29-648", "29-650", "29-651", "29-652", "HD182"))
            rt = dcStopAt((tx), rt);
        m0g0f0 = rt;
    }

    public Scripting.Dictionary m0g4f0(Scripting.Dictionary dcIn)
{
    // '  g4 currently reserved for development
    // '  relating to identification of purchased parts
    // '  by reference to Vault file path, as well as
    // '  BOM Structure and Cost Center settings
    // '
    // '  f0 separates members of incoming Dictionary
    // '  into design subgroups, including "doyle" and "purchased",
    // '  as indicated by the fourth component of each Vault path
    // '
    Scripting.Dictionary rt;
    Scripting.Dictionary sd;
    Variant ky;
    Variant ls;
    long mx;
    long dx;
    string g0;
    string g1;
    string fp;
    string fn;
    string rf;

    rt = new Scripting.Dictionary();
    {
        var withBlock = dcIn;
        foreach (var ky in withBlock.Keys)
        {
            ls = Split(ky, @"\");
            mx = UBound(ls);

            fn = ls(mx);
            g0 = ls(3);
            g1 = ls(4);
            fp = "";
            for (dx = 5; dx <= mx - 1; dx++)
                fp = fp + @"\" + ls(dx);
            fp = Strings.Mid(fp, 2);
            rf = Join(Array(fn, g1, fp), Constants.vbTab);

            {
                var withBlock1 = rt;
                // Stop
                if (withBlock1.Exists(g0))
                    sd = withBlock1.Item(g0);
                else
                {
                    sd = new Scripting.Dictionary();
                    withBlock1.Add(g0, sd);
                }

                {
                    var withBlock2 = sd;
                    if (withBlock2.Exists(rf))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock2.Add(rf, dcIn.Item(ky));
                }
            }
        }
    }
    m0g4f0 = rt;
}


public Scripting.Dictionary m0g4f1(Scripting.Dictionary dcIn)
{
    // '  g4 currently reserved for development
    // '  relating to identification of purchased parts
    // '  by reference to Vault file path, as well as
    // '  BOM Structure and Cost Center settings
    // '
    // '  f1 scans documents in the "purchased" design group
    // '  for unexpected settings, specifically:
    // '  BOMStructure should be
    // '      kPurchasedBOMStructure (51973 / 0xCB05)
    // '  Design Tracking Property "Cost Center"
    // '      should be either "D-PTS" or "R-PTS"
    // '
    // '  Any documents not matching these parameters
    // '  are dropped into one or both subDictionaries
    // '  in the returned Dictionary, according to issue
    // '
    Scripting.Dictionary rt;
    Scripting.Dictionary rtBom;
    Scripting.Dictionary rtFam;
    Scripting.Dictionary sd;
    Inventor.Document ivDoc;
    Inventor.ComponentDefinition CpDef;
    Inventor.Property prFam;
    Variant ky;

    rt = new Scripting.Dictionary();
    rtBom = new Scripting.Dictionary();
    rtFam = new Scripting.Dictionary();
    rt.Add("bom", rtBom);
    rt.Add("fam", rtFam);
    sd = dcIn.Item("purchased");
    {
        var withBlock = sd;
        foreach (var ky in withBlock.Keys)
        {
            ivDoc = withBlock.Item(ky);
            CpDef = aiCompDefOf(ivDoc);

            if (CpDef == null)
                System.Diagnostics.Debugger.Break();
            else
            {
                var withBlock1 = CpDef;
                if (withBlock1.BOMStructure != kPurchasedBOMStructure)
                    rtBom.Add(ky, ivDoc);

                prFam = aiDocument(withBlock1.Document).PropertySets.Item(gnDesign).Item(pnFamily);

                if (prFam.Value != "D-PTS" | prFam.Value != "R-PTS")
                    rtFam.Add(ky, ivDoc);
            }
        }
    }
    m0g4f1 = rt;
}


public Scripting.Dictionary m0g4f1fixBom(Scripting.Dictionary dcIn, long RiverView = 0)
{
    // '  g4 currently reserved for development
    // '  relating to identification of purchased parts
    // '  by reference to Vault file path, as well as
    // '  BOM Structure and Cost Center settings
    // '
    // '  f1fixBom purports to fix incorrect BOM
    // '  Structure settings in purchased parts
    // '
    Scripting.Dictionary rt;
    Scripting.Dictionary sd;
    // Dim ivDoc As Inventor.Document
    Inventor.ComponentDefinition CpDef;
    Variant ky;

    sd = dcIn.Item("bom");
    rt = new Scripting.Dictionary();
    {
        var withBlock = sd;
        foreach (var ky in withBlock.Keys)
        {
            CpDef = aiCompDefOf(withBlock.Item(ky));
            if (CpDef == null)
            {
                System.Diagnostics.Debugger.Break();
                rt.Add(ky, 0);
            }
            else
            {
                var withBlock1 = CpDef;
                withBlock1.BOMStructure = kPurchasedBOMStructure;
                rt.Add(ky, IIf(withBlock1.BOMStructure == kPurchasedBOMStructure, 1, 0));
            }
        }
    }
    m0g4f1fixBom = rt;
}


public Scripting.Dictionary m0g4f1fixFam(Scripting.Dictionary dcIn, long RiverView = 0)
{
    // '  g4 currently reserved for development
    // '  relating to identification of purchased parts
    // '  by reference to Vault file path, as well as
    // '  BOM Structure and Cost Center settings
    // '
    // '  f1fixFam purports to fix incorrect
    // '  Family settings in purchased parts
    // '
    Scripting.Dictionary rt;
    Scripting.Dictionary sd;
    Inventor.Document ivDoc;
    Inventor.ComponentDefinition CpDef;
    Variant ky;
    string txFam;

    txFam = Interaction.IIf(RiverView, "R", "D") + "-PTS";
    sd = dcIn.Item("fam");

    rt = new Scripting.Dictionary();
    {
        var withBlock = sd;
        foreach (var ky in withBlock.Keys)
        {
            ivDoc = withBlock.Item(ky);
            {
                var withBlock1 = ivDoc.PropertySets.Item(gnDesign).Item(pnFamily);
                withBlock1.Value = txFam;
                rt.Add(ky, IIf(withBlock1.Value == txFam, 1, 0));
            }

            CpDef = aiCompDefOf(withBlock.Item(ky));
            if (CpDef == null)
            {
                System.Diagnostics.Debugger.Break();
                rt.Add(ky, 0);
            }
            else
            {
                var withBlock1 = CpDef;
                withBlock1.BOMStructure = kPurchasedBOMStructure;
                rt.Add(ky, IIf(withBlock1.BOMStructure == kPurchasedBOMStructure, 1, 0));
            }
        }
    }
    m0g4f1fixFam = rt;
}

End Class
                    }
                }
            }
        }
    }
}