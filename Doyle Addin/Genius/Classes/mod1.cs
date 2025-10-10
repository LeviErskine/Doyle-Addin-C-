class SurroundingClass
{
    public Scripting.Dictionary m1g0f0()
    {
        Scripting.Dictionary rt;
        Variant ky;
        long gt0;
        long eq0;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAiSheetMetal(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument)));
            gt0 = 0;
            eq0 = 0;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocPart(withBlock.Item(ky));
                    {
                        var withBlock2 = vcChkFlatPat(aiCompDefShtMetal(withBlock1.ComponentDefinition));
                        if (Abs(withBlock2.X * withBlock2.Y * withBlock2.Z) > 0)
                        {
                            Debug.Print.X(*withBlock2.Y * withBlock2.Z, ky);
                            // Stop
                            gt0 = 1 + gt0;
                        }
                        else
                            // Stop
                            eq0 = 1 + eq0;
                    }
                }
            }
            Debug.Print(gt0, eq0);
        }
        m1g0f0 = rt;
    }

    public Scripting.Dictionary m1g0f1()
    {
        Scripting.Dictionary rt;
        Variant ky;
        long gt0;
        long eq0;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAssyDocComponents(ThisApplication.ActiveDocument, null/* Conversion error: Set to default value for this argument */, 1);
            gt0 = 0;
            eq0 = 0;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocument(withBlock.Item(ky));
                    if (withBlock1.DocumentInterests.HasInterest(guidDesignAccl))
                        System.Diagnostics.Debugger.Break();
                }
            }
            Debug.Print(gt0, eq0);
        }
        m1g0f1 = rt;
    }

    public Inventor.Vector vcDiaBox(Inventor.Box bx)
    {
        // '  Given a Box Object
        // '  containing diametrically opposed
        // '  MinPoint and MaxPoint Objects,
        // '  return the Vector from Min to Max
        {
            var withBlock = bx;
            vcDiaBox = withBlock.MinPoint.VectorTo(withBlock.MaxPoint);
        }
    }

    public Inventor.Vector vc2flatPat(Inventor.SheetMetalComponentDefinition df)
    {
        // '  From an Inventor Sheet Metal Component Definition,
        // '  obtain a Vector representing the translation
        // '  of the Folded Model's bounding box diagonal
        // '  to that of the Flat Pattern.
        // '
        Inventor.Vector rt;

        {
            var withBlock = df;
            rt = vcDiaBox(withBlock.RangeBox);
            if (withBlock.HasFlatPattern)
                rt.SubtractVector(vcDiaBox(withBlock.FlatPattern.RangeBox));
            else
                rt.SubtractVector(rt);
        }
        // '
        // '  If they're the same point, this vector should
        // '  have zero length, however, this is NOT proof
        // '  positive of an invalid Flat Pattern. A valid
        // '  flat piece, with no folds, should produce
        // '  the same result.
        // '
        // '  A good follow-up check would probably be to
        // '  compare the Flat Pattern diagonal vector's
        // '  Z component to the model's Thickness

        vc2flatPat = rt;
    }
    // Debug.Print vc2flatPat(aiCompDefShtMetal(aiDocPart(ThisApplication.Documents(2)).ComponentDefinition)).Length

    public Inventor.Vector vcCubicThickness(Inventor.SheetMetalComponentDefinition df)
    {
        double hk;

        hk = df.Thickness.Value;
        {
            var withBlock = ThisApplication.TransientGeometry;
            vcCubicThickness = withBlock.CreateVector(hk, hk, hk);
        }
    }

    public Inventor.Vector vcChkFlatPat(Inventor.SheetMetalComponentDefinition df)
    {
        // '  From an Inventor Sheet Metal Component Definition,
        // '  subtract a vector of cubic thickness from the
        // '  diagonal vector of either its Flat Pattern,
        // '  if available, or otherwise, the Folded Model.
        // '
        Inventor.Vector rt;

        {
            var withBlock = df;
            if (withBlock.HasFlatPattern)
                rt = vcDiaBox(withBlock.FlatPattern.RangeBox);
            else
                rt = vcDiaBox(withBlock.RangeBox);
        }
        rt.SubtractVector(vcCubicThickness(df));
        // '
        // '  If the model is a valid sheet metal part,
        // '  one of the dimensions of its flat pattern's
        // '  bounding box diagonal should either equal
        // '  the defined sheet metal thickness,
        // '  or fall very close. At least, in theory.
        // '
        // '  Plan at this point is to try to determine
        // '  just how often this bears out.
        // '  While this HAS failed one pretest, the Flat
        // '  Pattern of that model includes features;
        // '  a relatively infrequent occurrence, and
        // '  quite possibly one that can throw off
        // '  the boundaries.

        vcChkFlatPat = rt;
    }
    // Debug.Print vcChkFlatPat(aiCompDefShtMetal(aiDocPart(ThisApplication.Documents(2)).ComponentDefinition)).Length

    public Variant m1tst0()
    {
        Debug.Print(iFacAssy(aiCompDefAssy(aiDocAssy(aiCompDefAssy(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition).Occurrences(1).Definition.Document).ComponentDefinition)));
    }

    public Variant m1tst1()
    {
        Variant ky;
        Inventor.Property pr;
        // Dim bm As Inventor.BOMStructureEnum

        foreach (var ky in Filter(dcAssyCompAndSub(aiDocDefAssy(ThisApplication.ActiveDocument).Occurrences).Keys, "(DC)"))
        {
            {
                var withBlock = aiDocPart(ThisApplication.Documents.ItemByName((ky)));
                pr = withBlock.PropertySets("Design Tracking Properties").Item("Cost Center");
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    if (withBlock1.BOMStructure != kPurchasedBOMStructure)
                    {
                        withBlock1.BOMStructure = kPurchasedBOMStructure;
                        pr.Value = "D-PTS";
                        Debug.Print(pr.Value, withBlock1.BOMStructure, ky);
                    }
                }
            }
        }
    }

    public Inventor.iAssemblyFactory iFacAssy(Inventor.AssemblyComponentDefinition ob)
    {
        {
            var withBlock = ob;
            if (withBlock.iAssemblyFactory == null)
            {
                if (withBlock.iAssemblyMember == null)
                    iFacAssy = null/* TODO Change to default(_) if this is not a reference type */;
                else
                    iFacAssy = withBlock.iAssemblyMember.ParentFactory;
            }
            else
                iFacAssy = withBlock.iAssemblyFactory;
        }
    }

    public Inventor.Document aiOccDoc(Inventor.ComponentOccurrence ob)
    {
        aiOccDoc = ob.Definition.Document;
    }

    public Inventor.AssemblyComponentDefinition aiDocDefAssy(Inventor.AssemblyDocument ob)
    {
        aiDocDefAssy = ob.ComponentDefinition;
    }

    public Scripting.Dictionary dcAssyComponentsImmediate(Inventor.AssemblyDocument ob)
    {
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence oc;
        Inventor.Document cp;
        string fn;

        rt = new Scripting.Dictionary();
        foreach (var oc in ob.ComponentDefinition.Occurrences)
        {
            // aiDocument(oc.Definition.Document).FullFileName
            cp = oc.Definition.Document;
            fn = cp.FullFileName;
            {
                var withBlock = rt;
                if (!withBlock.Exists(fn))
                    withBlock.Add(fn, cp);
            }
        }
        dcAssyComponentsImmediate = rt;
    }
    // For Each pn In Array(pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnArea, pnLength, pnWidth, pnThickness): Debug.Print pn & "=" & aiDocument(ThisApplication.ActiveDocument).PropertySets.Item(gnCustom).Item((pn)).Value: Next

    public long iSyncPartFactory(Inventor.PartDocument pd)
    {
        // '
        // '  Backport iPart Member Properties to parent Factory
        // '
        Scripting.Dictionary dcCols;
        Scripting.Dictionary dcRows;
        Inventor.iPartTableRow tr;
        Inventor.PropertySet ps;
        Inventor.Property pr;
        long rt;
        VbMsgBoxResult ck;

        rt = 0;
        {
            var withBlock = pd.ComponentDefinition;
            ps = aiDocument(withBlock.Document).PropertySets.Item(gnCustom);
            if (withBlock.iPartMember == null)
            {
            }
            else
            {
                var withBlock1 = withBlock.iPartMember;
                dcRows = dcIPartTbRows(withBlock1.ParentFactory.TableRows);
                if (dcRows.Exists(pd.DisplayName))
                {
                    dcCols = dcIPartTbCols(withBlock1.ParentFactory.TableColumns);


                    Information.Err.Clear();
                    tr = withBlock1.Row;

                    /// REV[2022.03.23.1624]
                    /// adding "pickup" attempt to capture
                    /// and recover from errors encountered
                    /// in trying to retrieve .Row directly
                    if (Information.Err.Number == 0)
                    {
                    }
                    else
                    {
                        tr = dcRows.Item(pd.DisplayName);
                        if (tr.MemberName == pd.DisplayName)
                            Information.Err.Clear();
                        else if (tr.PartName == pd.DisplayName)
                            Information.Err.Clear();
                        else
                            System.Diagnostics.Debugger.Break();
                    }

                    if (Information.Err.Number == 0)
                    {
                        foreach (var pr in ps)
                        {
                            if (dcCols.Exists(pr.Name))
                            {
                                Debug.Print(pr.Name);
                                {
                                    var withBlock2 = tr.Item(dcCols.Item(pr.Name));
                                    Debug.Print("  ");
                                    if (pr.Value == withBlock2.Value)
                                        // Stop 'No change necessary
                                        Debug.Print(" (NO CHANGE)");
                                    else
                                    {
                                        withBlock2.Value = pr.Value;
                                        if (Information.Err.Number == 0)
                                        {
                                            rt = 1 + rt;

                                            // ' The update invalidated the object
                                            // ' We'll have to grab it again
                                            {
                                                var withBlock3 = tr.Item(dcCols.Item(pr.Name));
                                                Debug.Print(" -> ");
                                            }
                                        }
                                        else
                                            Debug.Print(" <!!ERROR!!> Couldn't Change");
                                    }
                                }
                            }
                            else
                            {
                            }
                        }
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                    else
                    {
                        Debug.Print("=== CAN'T SYNC IFACTORY ===");
                        Debug.Print("   Failed to access Row");
                        Debug.Print("   for Member " + pd.DisplayName);
                        Debug.Print("   of Factory " + aiDocument(withBlock1.ParentFactory.Parent).DisplayName);
                        Debug.Print("=== PLEASE CHECK PARENT ===");
                        Debug.Print("====== FACTORY TABLE ======");
                        Debug.Print("Error 0x" + Hex(Information.Err.Number) + "(");
                        Debug.Print("    " + Information.Err.Description);
                        Debug.Print("===========================");
                        ck = MsgBox(Join(Array("" + "iPart Member " + pd.DisplayName, "in Factory" + aiDocument(withBlock1.ParentFactory.Parent).DisplayName, "could not be accessed for updates.", "", "Its Row might still be present", "but somehow unavailable.", "", "Please review iPart Factory."), Constants.vbNewLine), Constants.vbOKCancel, "ERROR ACCESSING MEMBER ROW!");
                        if (ck == Constants.vbCancel)
                            System.Diagnostics.Debugger.Break();
                    }


                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                }
                else
                {
                    Debug.Print("==== CAN'T FIND MEMBER ====");
                    Debug.Print("   Failed to locate Row");
                    Debug.Print("   for Member " + pd.DisplayName);
                    Debug.Print("   of Factory " + aiDocument(withBlock1.ParentFactory.Parent).DisplayName);
                    Debug.Print("=== PLEASE CHECK PARENT ===");
                    Debug.Print("====== FACTORY TABLE ======");
                    // Debug.Print "Error 0x" & Hex$(Err.Number) & "("; CStr(Err.Number) & ")"
                    // Debug.Print "    " & Err.Description
                    // Debug.Print "==========================="
                    ck = MsgBox(Join(Array("" + "iPart Member " + pd.DisplayName, "could not be located in Factory", aiDocument(withBlock1.ParentFactory.Parent).DisplayName, "", "Its Row might have been removed", "or separated from the main table.", "", "Please review iPart Factory Table."), Constants.vbNewLine), Constants.vbOKCancel, "WARNING!! MEMBER NOT FOUND!");
                    if (ck == Constants.vbCancel)
                        System.Diagnostics.Debugger.Break();
                }
            }
        }
    }

    public long iSyncAssyFactory(Inventor.AssemblyDocument pd)
    {
        // '
        // '  Backport iPart Member Properties to parent Factory
        // '
        Scripting.Dictionary dcCols;
        Scripting.Dictionary dcRows;
        Inventor.iAssemblyTableRow tr;
        Inventor.PropertySet ps;
        Inventor.Property pr;
        long rt;

        rt = 0;
        {
            var withBlock = pd.ComponentDefinition;
            ps = aiDocument(withBlock.Document).PropertySets.Item(gnCustom);
            if (withBlock.iAssemblyMember == null)
            {
            }
            else
            {
                var withBlock1 = withBlock.iAssemblyMember;
                dcRows = dcIAssyTbRows(withBlock1.ParentFactory.TableRows);
                if (dcRows.Exists(pd.DisplayName))
                {
                    dcCols = dcIAssyTbCols(withBlock1.ParentFactory.TableColumns);
                    tr = withBlock1.Row;
                    foreach (var pr in ps)
                    {
                        if (dcCols.Exists(pr.Name))
                        {
                            Debug.Print(pr.Name);
                            {
                                var withBlock2 = tr.Item(dcCols.Item(pr.Name));
                                Debug.Print("  ");
                                if (pr.Value == withBlock2.Value)
                                    // Stop 'No change necessary
                                    Debug.Print(" (NO CHANGE)");
                                else
                                {
                                    withBlock2.Value = pr.Value;
                                    if (Information.Err.Number == 0)
                                    {
                                        rt = 1 + rt;

                                        // ' The update invalidated the object
                                        // ' We'll have to grab it again
                                        {
                                            var withBlock3 = tr.Item(dcCols.Item(pr.Name));
                                            Debug.Print(" -> ");
                                        }
                                    }
                                    else
                                        Debug.Print(" <!!ERROR!!> Couldn't Change");
                                }
                            }
                        }
                        else
                        {
                        }
                    }
                }
                else
                    System.Diagnostics.Debugger.Break();
            }
        }
    }

    public Scripting.Dictionary dcColumnsIPart(Inventor.PartDocument pd)
    {
        // '  Retrieve Dictionary of iPart Factory Table Columns
        // '  If supplied Part Document is NOT an iPart Factory
        // '  OR Member, returned Dictionary will be empty
        // '
        {
            var withBlock = pd.ComponentDefinition;
            if (withBlock.iPartMember == null)
            {
                if (withBlock.iPartFactory == null)
                    dcColumnsIPart = new Scripting.Dictionary();
                else
                    dcColumnsIPart = dcIPartTbCols(withBlock.iPartFactory.TableColumns);
            }
            else
                dcColumnsIPart = dcIPartTbCols(withBlock.iPartMember.ParentFactory.TableColumns);
        }
    }

    public Scripting.Dictionary dcColumnsIAssy(Inventor.AssemblyDocument pd)
    {
        // '  Retrieve Dictionary of iAssembly Factory Table Columns
        // '  If supplied Assembly Document is NOT an iAssembly Factory
        // '  OR Member, returned Dictionary will be empty
        // '
        {
            var withBlock = pd.ComponentDefinition;
            if (withBlock.iAssemblyMember == null)
            {
                if (withBlock.iAssemblyFactory == null)
                    dcColumnsIAssy = new Scripting.Dictionary();
                else
                    dcColumnsIAssy = dcIAssyTbCols(withBlock.iAssemblyFactory.TableColumns);
            }
            else
                dcColumnsIAssy = dcIAssyTbCols(withBlock.iAssemblyMember.ParentFactory.TableColumns);
        }
    }

    public Scripting.Dictionary dcIPartTbCols(Inventor.iPartTableColumns ls)
    {
        Scripting.Dictionary rt;
        Inventor.iPartTableColumn it;

        rt = new Scripting.Dictionary();

        foreach (var it in ls)
        {
            rt.Add(it.DisplayHeading, it.Index);
            if (Information.Err.Number)
            {
                rt.Add(it.FormattedHeading, it.Index);
                Information.Err.Clear();
            }
        }

        dcIPartTbCols = rt;
    }

    public Scripting.Dictionary dcIPartTbRows(Inventor.iPartTableRows ls)
    {
        Scripting.Dictionary rt;
        Inventor.iPartTableRow it;
        long ck;

        rt = new Scripting.Dictionary();

        /// REV[2022.08.05.1003]
        /// added error trap code to more gracefully handle
        /// errors accessing iPart/iAssembly factory table
        /// (replicated changes also to dcIAssyTbRows)

        ck = ls.Count; // to trigger potential error

        if (Information.Err.Number == 0)
        {
            foreach (var it in ls)
            {
                /// REV[2022.03.23.1618]
                /// replacing Index of iPartTableRow
                /// with iPartTableRow itself, so it
                /// can just be pulled directly out
                /// of the Dictionary by the client
                /// process. if it needs the Index,
                /// it can just get it itself, right?
                rt.Add(it.MemberName, it); rt.Add(it.PartName, it); Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // debug landing
            }
        }
        else
        {
            Debug.Print("ERROR " + System.Convert.ToHexString(Information.Err.Number) + " (" + Hex(Information.Err.Number) + ")" + Constants.vbNewLine + Information.Err.Description);
            Debug.Print(Join(Array("Could not access Table Rows", "for member of iPart factory."), Constants.vbNewLine));
            // Stop
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }



        dcIPartTbRows = rt;
    }

    public Scripting.Dictionary dcIAssyTbCols(Inventor.iAssemblyTableColumns ls)
    {
        Scripting.Dictionary rt;
        Inventor.iAssemblyTableColumn it;

        rt = new Scripting.Dictionary();

        foreach (var it in ls)
        {
            rt.Add(it.DisplayHeading, it.Index);
            if (Information.Err.Number)
            {
                rt.Add(it.FormattedHeading, it.Index);
                Information.Err.Clear();
            }
        }

        dcIAssyTbCols = rt;
    }

    public Scripting.Dictionary dcIAssyTbRows(Inventor.iAssemblyTableRows ls)
    {
        Scripting.Dictionary rt;
        Inventor.iAssemblyTableRow it;
        long ck;

        rt = new Scripting.Dictionary();

        /// REV[2022.08.05.1003]
        /// added error trap code to more gracefully handle
        /// errors accessing iPart/iAssembly factory table
        /// (replicated from dcIPartTbRows)

        ck = ls.Count; // to trigger potential error

        if (Information.Err.Number == 0)
        {
            foreach (var it in ls)
            {
                rt.Add(it.MemberName, it.Index);
                rt.Add(it.DocumentName, it.Index);
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // debug landing
            }
        }
        else
        {
            Debug.Print("ERROR " + System.Convert.ToHexString(Information.Err.Number) + " (" + Hex(Information.Err.Number) + ")" + Constants.vbNewLine + Information.Err.Description);
            Debug.Print(Join(Array("Could not access Table Rows", "for member of iAssembly factory."), Constants.vbNewLine));
            System.Diagnostics.Debugger.Break();
        }



        dcIAssyTbRows = rt;
    }

    public Inventor.iPartTableColumn m1g1f2(Variant vr)
    {
        m1g1f2 = vr;
    }

    public long m1g1f3(Inventor.AssemblyDocument ad)
    {
        Inventor.ComponentOccurrence oc;
        long rt;

        rt = 0;
        foreach (var oc in ad.ComponentDefinition.Occurrences) // (1)
        {
            {
                var withBlock = oc.Definition;
                // Debug.Print Hex$(aiDocument(.Document).Type)
                // Stop
                if (aiDocument(withBlock.Document).DocumentType == kPartDocumentObject)
                    rt = rt + iSyncPartFactory(aiDocPart(withBlock.Document));
                else
                {
                }
            }
        }
        m1g1f3 = rt;
    }

    public long m1g1f4()
    {
        m1g1f4 = m1g1f3(aiDocAssy(ThisApplication.ActiveDocument));
    }

    public Scripting.Dictionary m1g1f5(Inventor.PartDocument pd)
    {
        // '  Retrieve Dictionary of Custom Properties
        // '
        Scripting.Dictionary rt;
        Inventor.PropertySet psMember;
        Inventor.PropertySet psFactry;
        // Dim dcMember As Scripting.Dictionary
        Scripting.Dictionary dcFactry;
        Inventor.Property pr;

        rt = new Scripting.Dictionary();
        {
            var withBlock = pd;
            psMember = withBlock.PropertySets.Item(gnCustom);
            {
                var withBlock1 = withBlock.ComponentDefinition;
                if (!withBlock1.iPartMember == null)
                {
                    {
                        var withBlock2 = withBlock1.iPartMember;
                        {
                            var withBlock3 = aiDocument(withBlock2.ParentFactory.Parent);
                            psFactry = withBlock3.PropertySets.Item(gnCustom);
                            dcFactry = dcAiPropsInSet(psFactry);
                        }

                        foreach (var pr in psMember)
                        {
                            if (!dcFactry.Exists(pr.Name))
                                // rt.Add pr.Name, pr
                                rt = dcWithProp(psFactry, pr.Name, pr.Value, rt);
                        }
                    }
                }
            }
        }
        m1g1f5 = rt;
    }

    public string m1g1f5t0()
    {
        m1g1f5t0 = Join(m1g1f5(aiDocPart(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences.Item(1).Definition.Document)).Keys, Constants.vbNewLine);
    }
}