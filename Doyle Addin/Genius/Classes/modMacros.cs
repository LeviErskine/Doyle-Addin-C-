class SurroundingClass
{
    public Scripting.Dictionary gnsPropSetAll()
    {
        gnsPropSetAll = nuDcPopulator().Setting("Part Number", "Item").Setting("Description", "Description1").Setting("Cost Center", "Family").Setting("GeniusMass", "Weight").Setting("Thickness", "Thickness").Setting("Extent_Area", "Diameter").Setting("Extent_Width", "Width").Setting("Extent_Length", "Length").Setting("RM", "Stock").Setting("RMQTY", "QuantityInConversionUnit").Setting("RMUNIT", "ConversionUnit").Setting("OFFTHK", "OFFTHK").Dictionary;
    }

    public Scripting.Dictionary gnsPropSetItem()
    {
        gnsPropSetItem = nuDcPopulator().Setting("Part Number", "Item").Setting("Cost Center", "Family").Setting("GeniusMass", "Weight").Setting("Thickness", "Thickness").Setting("Extent_Area", "Diameter").Setting("Extent_Width", "Width").Setting("Extent_Length", "Length").Setting("OFFTHK", "OFFTHK").Dictionary;
    }

    public Scripting.Dictionary gnsPropSetBomRaw()
    {
        /// gnsPropSetBomRaw -- Property Names for Genius BOM
        /// !!!NOT READY!!! Just dup'd from gnsPropSetItem
        /// Needs adjustment to BOM Column/Field names
        /// 
        gnsPropSetBomRaw = nuDcPopulator().Setting("Part Number", "Product").Setting("RM", "Item").Setting("RMQTY", "QuantityInConversionUnit").Setting("RMUNIT", "ConversionUnit").Dictionary;
    }

    public Scripting.Dictionary gnsPropsCurrent(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */, Scripting.Dictionary dcProps = null/* TODO Change to default(_) if this is not a reference type */, long incTop = 0, long inclPhantom = 0)
    {
        // Dim rf As Scripting.Dictionary
        Scripting.Dictionary rt;
        // Dim ActiveDoc As Document
        Variant ky;
        // Dim dx As Long

        if (AiDoc == null)
            gnsPropsCurrent = gnsPropsCurrent(aiDocActive(), dcProps, incTop, inclPhantom);
        else if (dcProps == null)
            gnsPropsCurrent = gnsPropsCurrent(AiDoc, gnsPropSetAll(), incTop, inclPhantom); // gnsPropSetItem
        else
        {
            // rf = dcProps 'gnsPropSetItem() 'dcOfIdent(Array("Part Number", "Cost Center","GeniusMass", "Extent_Area","Extent_Width", "Extent_Length","RM", "RMQTY", "RMUNIT","OFFTHK"))

            /// Collect Components for Processing
            /// (retained from Sub UpdateGeniusProperties_2023_0406_pre)
            /// NOTE[2021.08.09]:
            /// Function dcRemapByPtNum previously
            /// revised to address Key collisions
            /// in a crude manner. See that function
            /// for details.
            /// 
            rt = dcRemapByPtNum(dcAiDocComponents(AiDoc, null/* Conversion error: Set to default value for this argument */, incTop));
            // NOTE: incTop here is used to indicate
            // whether to include top level assembly.
            // This decision probably still needs
            // to be made, but is not really
            // addressed at the moment.

            /// Retrieve the full Component Collection
            {
                var withBlock = rt;
                foreach (var ky in withBlock.Keys)
                    /// Genius Properties for Item
                    /// Probably DON'T want to replace
                    /// Property Objects with their
                    /// values at this point, so that
                    /// they can be updated without
                    /// having to retrieve them again.
                    withBlock.Item(ky) = dcKeysInCommon(dcOfPropsInAiDoc(withBlock.Item(ky)), dcProps, 1);// dcProps replaces rf

                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }

            gnsPropsCurrent = rt;
        }
    }

    public void testGnsPropsCurrent()
    {
        Inventor.Document md;
        Scripting.Dictionary mPrpGnsAll;
        Scripting.Dictionary mPrpGnsItm;
        Scripting.Dictionary mPrpBomRaw;
        Scripting.Dictionary dcPr;
        Scripting.Dictionary dcVl;
        Scripting.Dictionary dcGn;
        Scripting.Dictionary dcGnDx;
        Scripting.Dictionary dcBomDx;
        Scripting.Dictionary nd;
        Scripting.Dictionary ck;
        VbMsgBoxResult goAhead;
        string txOut;
        Variant ky;

        md = aiDocActive();
        mPrpGnsAll = gnsPropSetAll(); // gnsPropSetItem
                                      // mPrpBomRaw = gnsPropSetBomRaw()

        /// Retrieve all Documents'
        /// Genius Property Objects...
        dcPr = gnsPropsCurrent(md, mPrpGnsAll);

        /// ...and their current Values
        dcVl = new Scripting.Dictionary();
        {
            var withBlock = dcPr;
            foreach (var ky in withBlock.Keys)
                // .Item(ky) = dcPropVals(dcOb(.Item(ky)))
                dcVl.Add(ky, dcPropVals(dcOb(withBlock.Item(ky))));
        }
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        if (false)
        {
            send2clipBdWin10(ConvertToJson(dcVl, Constants.vbTab));
            System.Diagnostics.Debugger.Break();
        }

        dcGnDx = dcRecSetDcDx4json(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(q1g1x2(md)))));
        {
            var withBlock = dcGnDx;
            dcGn = dcOb(withBlock.Item("Item"));
            {
                var withBlock1 = dcOb(withBlock.Item(""));
                foreach (var ky in dcGn.Keys)
                    dcGn.Item(ky) = withBlock1.Item(dcGn.Item(ky)(0));
            }
        }
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        if (false)
        {
            send2clipBdWin10(ConvertToJson(dcGn, Constants.vbTab));
            System.Diagnostics.Debugger.Break();
        }

        /// This is to extract BOM from Assembly
        // bomInfoBkDn(bomViewStruct(aiDocAssy(

        dcBomDx = dcRecSetDcDx4json(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(q1g2x2(md)))));
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        if (false)
        {
            send2clipBdWin10(ConvertToJson(dcBomDx, Constants.vbTab));
            System.Diagnostics.Debugger.Break();
        }

        nd = new Scripting.Dictionary();
        {
            var withBlock = dcPr;
            foreach (var ky in withBlock.Keys)
            {
                ck = dcKeysMissing(mPrpGnsAll, dcOb(withBlock.Item(ky)));
                if (ck.Count > 0)
                    nd.Add(ky, ck);
            }
        }
        // Debug.Print ConvertToJson(nd, vbTab)
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing

        /// Index the Dictionary here
        /// (might be temporary)
        // dcPr = dcRecSetDcDx4json(dcDxFromRecSetDc(dcPr))

        /// Dump to JSON text format
        txOut = ConvertToJson(dcRecSetDcDx4json(dcDxFromRecSetDc(dcVl)), Constants.vbTab);  // dcPr
                                                                                            // Debug.Print txOut

        goAhead = MsgBox(Join(Array("Assembly Name:", md.DisplayName, "Process Completed", "", "Copy report text", "(JSON format)", "to Clipboard?", "", "(Cancel for Debug)"), Constants.vbNewLine), Constants.vbYesNoCancel, "Update Complete");
        if (goAhead == Constants.vbCancel)
            System.Diagnostics.Debugger.Break();
        else if (goAhead == Constants.vbYes)
            send2clipBdWin10(txOut);
    }

    public void ExposeAllSheetMetalThicknesses()
    {
        Inventor.PartDocument pt;
        Inventor.Parameter tk;
        Variant ky;
        long ct;
        long xp;
        long nc;

        {
            var withBlock = dcAiSheetMetal(dcAiPartDocs(dcAiDocComponents(ThisApplication.ActiveDocument)));
            ct = 0;
            xp = 0;
            nc = 0;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiCompDefShtMetal(aiDocPart(withBlock.Item(ky)).ComponentDefinition);
                    if (withBlock1.BOMStructure == kNormalBOMStructure)
                    {
                        ct = 1 + ct;


                        Information.Err.Clear();

                        /// REV[2023.01.18.1626]
                        /// added check for iPart member
                        /// to check exposure of Thickness
                        /// parameter of its Parent Factory
                        /// rather than its own.
                        /// 
                        /// this seeks to avoid errors
                        /// that now seem to arise when
                        /// attempting to set exposure
                        /// on the members themselves.
                        /// 
                        if (withBlock1.IsiPartMember)
                        {
                            // Stop
                            {
                                var withBlock2 = aiCompDefShtMetal(aiDocPart(withBlock1.iPartMember.ParentFactory.Parent));
                                tk = withBlock2.Thickness;
                            }
                        }
                        else
                            tk = withBlock1.Thickness;

                        if (Information.Err.Number)
                        {
                            Debug.Print("!ERROR!: " + ky);
                            tk = null/* TODO Change to default(_) if this is not a reference type */;
                            nc = 1 + nc;
                            Information.Err.Clear();
                        }
                        else
                        {
                            var withBlock2 = tk;
                            if (withBlock2.ExposedAsProperty)
                                Debug.Print("NOCHNGE: " + ky);
                            else if (aiCompDefShtMetal(withBlock2.Parent.Parent).IsiPartMember)
                            {
                                // Stop
                                if (aiDocPart(aiCompDefShtMetal(withBlock2.Parent.Parent).iPartMember.ParentFactory.Parent).ComponentDefinition.Parameters.Item(pnThickness).ExposedAsProperty)
                                    Debug.Print("NOCHNGE: " + ky);
                                else
                                {
                                    nc = 1 + nc;
                                    Debug.Print("FAILED!: " + ky);
                                }
                            }
                            else
                            {
                                withBlock2.ExposedAsProperty = true;
                                if (withBlock2.ExposedAsProperty)
                                {
                                    xp = 1 + xp;
                                    Debug.Print("EXPOSED: " + ky);
                                }
                                else
                                {
                                    nc = 1 + nc;
                                    Debug.Print("FAILED!: " + ky);
                                }
                            }
                        }
                    }
                }
            }
            if (xp + nc > 0)
                MsgBox(Join(Array("Found " + System.Convert.ToHexString(ct) + " components.", "Thickness already exposed on " + System.Convert.ToHexString(ct - xp - nc), "   Exposed additional " + System.Convert.ToHexString(xp), "   Failed to expose " + System.Convert.ToHexString(nc)), Constants.vbNewLine), Constants.vbOKOnly, "Sheet Metal Processed");
            else
                MsgBox(Join(Array("Thickness already exposed", "on " + System.Convert.ToHexString(ct) + " components."), Constants.vbNewLine), Constants.vbOKOnly, "No Change Required");
        }
    }

    public void AddProps4Genius()
    {
        Variant ky;
        Inventor.Property pr;

        {
            var withBlock = dcProps4genius(ThisApplication.ActiveDocument);
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiProperty(obOf(withBlock.Item(ky)));
                    Debug.Print.Parent.Name(+":" + withBlock1.Name + "=" + withBlock1.Value);
                }
            }
        }
    }
    // For Each itm In ActiveDocsComponents(ThisApplication): Debug.Print aiDocument(obOf(itm)).FullFileName: Next

    public void MakeViewImageFiles()
    {
        Debug.Print(d0g9f0());
    }

    public void iParts_GenerateAll()
    {
        Inventor.PartDocument oDoc;
        Inventor.iPartFactory oFactory;
        string sFile;
        long iCount;
        long i;
        long mx;
        long dx;
        long bk;

        // oDoc = AskUser4aiDoc(, dcOf_iAll_Factories(ThisApplication.Documents.VisibleDocuments))

        oDoc = AskUser4aiDoc(null/* Conversion error: Set to default value for this argument */, dcOf_iPartFactories());

        if (oDoc == null)
        {
        }
        else if (oDoc.ComponentDefinition.IsiPartFactory == true)
        {
            sFile = oDoc.FullFileName;

            oFactory = oDoc.ComponentDefinition.iPartFactory;

            // With oFactory
            mx = oFactory.TableRows.Count;
            dx = 1;
            do
            {
                bk = 1 + mx - dx;
                if (bk > 10)
                    bk = 10;
                iCount = 0;
                do
                {
                    ThisApplication.StatusBarText = System.Convert.ToHexString(dx) + "/" + System.Convert.ToHexString(mx) + ": " + oFactory.TableRows.Item(dx).MemberName;
                    // Member File creation
                    // .CreateMember dx
                    // disabled for testing
                    MsgBox(oFactory.TableRows.Item(dx).MemberName, Constants.vbOKOnly, "Member " + System.Convert.ToHexString(dx) + "/" + System.Convert.ToHexString(mx));
                    dx = dx + 1;
                    iCount = iCount + 1;
                }
                while (iCount < bk);
                if (dx <= mx)
                {
                    oDoc.Close();
                    oDoc = ThisApplication.Documents.Open(sFile);
                    oFactory = oDoc.ComponentDefinition.iPartFactory;
                    iCount = 0;
                }
            }
            while (!dx > mx); // Next
        }
    }
}