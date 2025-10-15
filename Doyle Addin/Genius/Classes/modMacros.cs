using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class modMacros
{
    public Dictionary gnsPropSetAll()
    {
        return nuDcPopulator().Setting("Part Number", "Item").Setting("Description", "Description1")
            .Setting("Cost Center", "Family").Setting("GeniusMass", "Weight").Setting("Thickness", "Thickness")
            .Setting("Extent_Area", "Diameter").Setting("Extent_Width", "Width").Setting("Extent_Length", "Length")
            .Setting("RM", "Stock").Setting("RMQTY", "QuantityInConversionUnit").Setting("RMUNIT", "ConversionUnit")
            .Setting("OFFTHK", "OFFTHK").Dictionary;
    }

    public Dictionary gnsPropSetItem()
    {
        return nuDcPopulator().Setting("Part Number", "Item").Setting("Cost Center", "Family")
            .Setting("GeniusMass", "Weight").Setting("Thickness", "Thickness").Setting("Extent_Area", "Diameter")
            .Setting("Extent_Width", "Width").Setting("Extent_Length", "Length").Setting("OFFTHK", "OFFTHK").Dictionary;
    }

    public Dictionary gnsPropSetBomRaw()
    {
        // gnsPropSetBomRaw -- Property Names for Genius BOM
        // !!!NOT READY!!! Just dup'd from gnsPropSetItem
        // Needs adjustment to BOM Column/Field names
        // 
        return nuDcPopulator().Setting("Part Number", "Product").Setting("RM", "Item")
            .Setting("RMQTY", "QuantityInConversionUnit").Setting("RMUNIT", "ConversionUnit").Dictionary;
    }

    public Dictionary gnsPropsCurrent(Document AiDoc = null, Dictionary dcProps = null, long incTop = 0,
        long inclPhantom = 0)
    {
        while (true)
        {
            // Dim rf As Scripting.Dictionary
            // Dim ActiveDoc As Document

            // Dim dx As Long

            if (AiDoc == null)
            {
                AiDoc = aiDocActive();
                continue;
            }

            if (dcProps == null)
            {
                dcProps = gnsPropSetAll();
                continue;
            }

            // rf = dcProps 'gnsPropSetItem() 'dcOfIdent(new [] {"Part Number", "Cost Center","GeniusMass", "Extent_Area","Extent_Width", "Extent_Length","RM", "RMQTY", "RMUNIT","OFFTHK"))
            // Collect Components for Processing
            // (retained from Sub UpdateGeniusProperties_2023_0406_pre)
            // NOTE[2021.08.09]:
            // Function dcRemapByPtNum previously
            // revised to address Key collisions
            // crudely. See that function
            // for details.
            // 
            var rt = dcRemapByPtNum(dcAiDocComponents(AiDoc, null , incTop));
            // NOTE: incTop here is used to indicate
            // whether to include top level assembly.
            // This decision probably still needs
            // to be made, but is not really
            // addressed at the moment.
            // Retrieve the full Component Collection
            {
                foreach (var ky in rt.Keys)
                    // Genius Properties for Item
                    // Probably DON'T want to replace
                    // Property Objects with their
                    // values at this point, so that
                    // they can be updated without
                    // having to retrieve them again.
                    rt.get_Item(ky) =
                        dcKeysInCommon(dcOfPropsInAiDoc(rt.get_Item(ky)), dcProps, 1); // dcProps replaces rf

                Debug.Print(""); // Breakpoint Landing
            }

            return rt;

            break;
        }
    }

    public void testGnsPropsCurrent()
    {
        Dictionary mPrpGnsItm;
        Dictionary mPrpBomRaw;
        Dictionary dcGn;

        var md = aiDocActive();
        var mPrpGnsAll = gnsPropSetAll(); // gnsPropSetItem
        // mPrpBomRaw = gnsPropSetBomRaw()
        // Retrieve all Documents'
        // Genius Property Objects...
        var dcPr = gnsPropsCurrent(md, mPrpGnsAll);

        // ...and their current Values
        var dcVl = new Dictionary();
        {
            foreach (var ky in dcPr.Keys)
                // .get_Item(ky) = dcPropVals(dcOb(.get_Item(ky)))
                dcVl.Add(ky, dcPropVals(dcOb(dcPr.get_Item(ky))));
        }
        Debug.Print(""); // Breakpoint Landing
        if (false)
        {
            send2clipBdWin10(ConvertToJson(dcVl, Constants.vbTab));
            Debugger.Break();
        }

        var dcGnDx = dcRecSetDcDx4json(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(q1g1x2(md)))));
        {
            dcGn = dcOb(dcGnDx.get_Item("Item"));
            {
                var withBlock1 = dcOb(dcGnDx.get_Item(""));
                foreach (var ky in dcGn.Keys)
                    dcGn.get_Item(ky) = withBlock1.get_Item(dcGn.get_Item(ky)(0));
            }
        }
        Debug.Print(""); // Breakpoint Landing
        if (false)
        {
            send2clipBdWin10(ConvertToJson(dcGn, Constants.vbTab));
            Debugger.Break();
        }

        // This is to extract BOM from Assembly
        // bomInfoBkDn(bomViewStruct(aiDocAssy(
        var dcBomDx = dcRecSetDcDx4json(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(q1g2x2(md)))));
        Debug.Print(""); // Breakpoint Landing
        if (false)
        {
            send2clipBdWin10(ConvertToJson(dcBomDx, Constants.vbTab));
            Debugger.Break();
        }

        var nd = new Dictionary();
        {
            foreach (var ky in dcPr.Keys)
            {
                Dictionary ck = dcKeysMissing(mPrpGnsAll, dcOb(dcPr.get_Item(ky)));
                if (ck.Count > 0)
                    nd.Add(ky, ck);
            }
        }
        // Debug.Print ConvertToJson(nd, vbTab)
        Debug.Print(""); // Breakpoint Landing
        // Index the Dictionary here
        // (might be temporary)
        // dcPr = dcRecSetDcDx4json(dcDxFromRecSetDc(dcPr))
        // Dump to JSON text format
        var txOut = ConvertToJson(dcRecSetDcDx4json(dcDxFromRecSetDc(dcVl)), Constants.vbTab); // dcPr

        // Debug.Print txOut
        VbMsgBoxResult goAhead = MessageBox.Show(Join(new[]
        {
            "Assembly Name:", md.DisplayName, "Process Completed", "", "Copy report text", "(JSON format)",
            "to Clipboard?", "", "(Cancel for Debug)"
        }, Constants.vbCrLf), Constants.vbYesNoCancel, "Update Complete");
        if (goAhead == Constants.vbCancel)
            Debugger.Break();
        else if (goAhead == Constants.vbYes)
            send2clipBdWin10(txOut);
    }

    public void ExposeAllSheetMetalThicknesses()
    {
        PartDocument pt;

        {
            var withBlock = dcAiSheetMetal(dcAiPartDocs(dcAiDocComponents(ThisApplication.ActiveDocument)));
            long ct = 0;
            long xp = 0;
            long nc = 0;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiCompDefShtMetal(aiDocPart(withBlock.get_Item(ky)).ComponentDefinition);
                    if (withBlock1.BOMStructure != kNormalBOMStructure) continue;
                    ct = 1 + ct;

                    Information.Err().Clear();

                    // REV[2023.01.18.1626]
                    // added check for iPart member
                    // to check exposure of Thickness
                    // parameter of its Parent Factory
                    // rather than its own.
                    // 
                    // this seeks to avoid errors
                    // that now seem to arise when
                    // attempting to set exposure
                    // on the members themselves.
                    // 
                    Parameter tk;
                    if (withBlock1.IsiPartMember)
                    {
                        // Stop
                        {
                            var withBlock2 =
                                aiCompDefShtMetal(aiDocPart(withBlock1.iPartMember.ParentFactory.Parent));
                            tk = withBlock2.Thickness;
                        }
                    }
                    else
                        tk = withBlock1.Thickness;

                    if (Information.Err().Number)
                    {
                        Debug.Print("!ERROR!: " + ky);
                        tk = null;
                        nc = 1 + nc;
                        Information.Err().Clear();
                    }
                    else
                    {
                        var withBlock2 = tk;
                        if (withBlock2.ExposedAsProperty)
                            Debug.Print("NOCHNGE: " + ky);
                        else if (aiCompDefShtMetal(withBlock2.Parent.Parent).IsiPartMember)
                        {
                            // Stop
                            if (aiDocPart(aiCompDefShtMetal(withBlock2.Parent.Parent).iPartMember.ParentFactory
                                    .Parent).ComponentDefinition.Parameters.get_Item(pnThickness).ExposedAsProperty)
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

            if (xp + nc > 0)
                MessageBox.Show(Join(new[]
                    {
                        "Found " + Convert.ToString(ct) + " components.",
                        "Thickness already exposed on " + Convert.ToString(ct - xp - nc),
                        " Exposed additional " + Convert.ToHexString(xp), " Failed to expose " + Convert.ToString(nc)
                    },
                    Constants.vbCrLf), Constants.vbOKOnly, "Sheet Metal Processed");
            else
                MessageBox.Show(Join(new[]
                    {
                        "Thickness already exposed",
                        "on " + Convert.ToHexString(ct) + " components."
                    }, Constants.vbCrLf), Constants.vbOKOnly,
                    "No Change Required");
        }
    }

    public void AddProps4Genius()
    {
        Property pr;

        {
            var withBlock = dcProps4genius(ThisApplication.ActiveDocument);
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiProperty(obOf(withBlock.get_Item(ky)));
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
        long i;

        // oDoc = AskUser4aiDoc(, dcOf_iAll_Factories(ThisApplication.Documents.VisibleDocuments))
        PartDocument oDoc = AskUser4aiDoc(null, dcOf_iPartFactories());

        if (oDoc == null)
        {
        }
        else if (oDoc.ComponentDefinition.IsiPartFactory == true)
        {
            var sFile = oDoc.FullFileName;

            var oFactory = oDoc.ComponentDefinition.iPartFactory;

            // With oFactory
            long mx = oFactory.TableRows.Count;
            long dx = 1;
            do
            {
                var bk = 1 + mx - dx;
                if (bk > 10)
                    bk = 10;
                long iCount = 0;
                do
                {
                    ThisApplication.StatusBarText = Convert.ToHexString(dx) + "/" + Convert.ToHexString(mx) + ": " +
                                                    oFactory.TableRows.get_Item(dx).MemberName;
                    // Member File creation
                    // .CreateMember dx
                    // disabled for testing
                    MessageBox.Show(oFactory.TableRows.get_Item(dx).MemberName, Constants.vbOKOnly,
                        "Member " + Convert.ToHexString(dx) + "/" + Convert.ToHexString(mx));
                    dx = dx + 1;
                    iCount = iCount + 1;
                } while (iCount < bk);

                if (dx > mx) continue;
                oDoc.Close();
                oDoc = ThisApplication.Documents.Open(sFile);
                oFactory = oDoc.ComponentDefinition.iPartFactory;
                iCount = 0;
            } while (!dx > mx); // Next
        }
    }
}