using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class mod0
{
    public dynamic m0g3f1(ADODB.Recordset rs)
    {
        {
            rs.Filter = rs.Filter;
            if (rs.BOF | rs.EOF)
                dynamic rt = new[] { new[] { "<NODATA>" } };
            else
            {
                long ct = withBlock.Fields.Count - 1;
                dynamic ar = Split(withBlock.GetString(adClipString, null, Constants.vbTab, Constants.vbVerticalTab),
                    Constants.vbVerticalTab);
                long mx = UBound(ar) - 1;
                rt(mx, ct);
                long dx;
                for (dx = 0; dx <= mx; dx++)
                {
                    dynamic rw = Split(ar(dx), Constants.vbTab);
                    long fd;
                    for (fd = 0; fd <= ct; fd++)
                        rt(dx, fd) = rw(fd);
                }
            }
        }

        return rt;
    }

    public string m0g3f2(Document AiDoc)
    {
        dynamic ky;

        {
            var withBlock = newFmTest1();
            withBlock.AskAbout(AiDoc);
            {
                var withBlock1 = withBlock.ItemData;
                dynamic ar;
                if (withBlock1.Count > 0)
                {
                    ar = withBlock1.Keys;
                    for (long dx = 0; dx <= UBound(ar); dx++) // Each ky In ar
                        // Debug.Print ky, .get_Item(ky)
                        ar(dx) = ar(dx) + "=" + withBlock1.get_Item(ar(dx));
                }
                else
                    ar = new[] { "<NODATA>" };

                return Join(ar, Constants.vbCrLf);
            }
        }
    }

    public long m0g2f1(Document AiDoc)
    {
        return newFmTest0().ft0g0f0(AiDoc.Thumbnail);
    }

    public long m0g2f2()
    {
        return m0g2f1(ThisApplication.ActiveDocument);
    }

    public long m0g2f3()
    {
        // Dim dc As Scripting.Dictionary

        // dc = New Scripting.Dictionary
        {
            var withBlock = new Dictionary() // fmTest0
                ;
            foreach (Document AiDoc in ThisApplication.Documents)
            {
                // .ft0g0f0 aiDoc.Thumbnail
                if (withBlock.Exists(AiDoc.FullFileName))
                    withBlock.get_Item(AiDoc.FullFileName) = 1 + withBlock.get_Item(AiDoc.FullFileName);
                else
                    withBlock.Add(AiDoc.FullFileName, 1);

                foreach (var ky in withBlock.Keys)
                {
                    if (withBlock.get_Item(ky) > 1)
                        Debug.Print.get_Item(ky);
                }
            }
        }
    }

    public Dictionary m0g1f0(Document AiDoc, Dictionary dc = null)
    {
        Dictionary rt;

        if (dc == null)
            rt = m0g1f0(AiDoc, new Dictionary());
        else
        {
            string ky;
            DocumentTypeEnum tp;
            {
                tp = AiDoc.DocumentType;
                ky = AiDoc.FullFileName;
            }

            rt = dc;
            if (rt.Exists(ky))
            {
            }
            else
            {
                rt.Add(ky, AiDoc);
                switch (tp)
                {
                    case kAssemblyDocumentObject:
                        rt = m0g1f0assy(AiDoc, dc);
                        break;
                    case kPartDocumentObject:
                    case kUnknownDocumentObject:
                    case kDrawingDocumentObject:
                    case kPresentationDocumentObject:
                    case kDesignElementDocumentObject:
                    case kForeignModelDocumentObject:
                    case kSATFileDocumentObject:
                    case kNoDocument:
                    case kNestingDocument:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        return rt;
    }
    // dc = m0g1f0(ThisApplication.ActiveDocument)
    // For Each ky In dc.Keys
    // Debug.Print aiDocument(dc.get_Item(ky)).PropertySets(gnDesign).get_Item(pnPartNum).Value, 'aiDocument(dc.get_Item(ky)).PropertySets(gnDesign).get_Item(pnMaterial).Value
    // Next

    public Dictionary m0g1f0part(PartDocument AiDoc, Dictionary dc = null)
    {
        Dictionary rt;

        rt = dc ?? m0g1f0part(AiDoc, new Dictionary());

        return rt;
    }

    public Dictionary m0g1f0assy(AssemblyDocument AiDoc, Dictionary dc = null)
    {
        Dictionary rt;
        ComponentOccurrence aiOcc;

        if (dc == null)
            rt = m0g1f0assy(AiDoc, new Dictionary());
        else
        {
            rt = AiDoc.ComponentDefinition.Occurrences.Cast<object>()
                .Aggregate(dc, (current, aiOcc) => m0g1f0(aiOcc.Definition.Document, current));
        }

        return rt;
    }

    public Dictionary dcAssyComp2A(AssemblyDocument AiDoc, Dictionary dc = null)
    {
        Dictionary rt;
        ObjectTypeEnum tp;

        if (dc == null)
            rt = dcAssyComp2A(AiDoc, new Dictionary());
        else
        {
            rt = dc;
            foreach (ComponentOccurrence invOcc in AiDoc.ComponentDefinition.Occurrences)
            {
                {
                    // Remove suppressed and excluded parts from the process
                    // Moved out here from inner checks
                    if (invOcc.Visible & !invOcc.Suppressed & !invOcc.Excluded)
                    {
                        tp = invOcc.Definition.Type;

                        if (tp != kAssemblyComponentDefinitionObjectAnd)
                            tp<> kWeldmentComponentDefinitionObjectThen
                        if tp<>
                            kWeldsComponentDefinitionObject Then
                            rt = dcAddAiDoc(.Definition.Document, rt)
                        end if
                        else //assembly, check BOM Structure
                        If.BOMStructure = kPurchasedBOMStructure Then //it's purchased
                            rt = dcAddAiDoc(.Definition.Document, rt)
                        elseif.BOMStructure = kNormalBOMStructure Then //we make it
                            rt = dcAssyComp2A(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt)
                            ) //NOT forgetting to add THIS document!
                        elseif.BOMStructure = kInseparableBOMStructure Then //maybe weldment?
                        if tp = kWeldmentComponentDefinitionObject Then //it is
                            rt = dcAssyComp2A(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt))
                        else //it's not
                        Stop //and see if we can figure out what its type is
                            End if
                        elseif.BOMStructure = kPhantomBOMStructure Then //"phantom" component
                            //Gather its components, but NOT the document itself
                            rt = dcAssyComp2A(.SubOccurrences, rt)
                        else //not sure what we've got
                        Stop //and have a look at it
                            end if
                        end if //part or assembly
                        end if
                        end with
                        next
                            end if

                        return new Dictionary
                    }
                }
            }
        }
    }

    public Dictionary dcAssyCompStops(ComponentOccurrences Occurences, Dictionary dc = null, Dictionary dcStops = null)
    {
        // Traverse the assembly,
        // including any/all subassemblies,
        // and collect all parts to be processed.
        Dictionary rt;
        ComponentOccurrence invOcc;
        ObjectTypeEnum tp;

        if (dc == null)
            rt = dcAssyCompStops(Occurences, new Dictionary(), dcStops);
        else
        {
            rt = dc;
            foreach (var invOcc in Occurences)
            {
                {
                    if (dcStops.Exists(aiDocument(invOcc.Definition.Document).PropertySets.get_Item(gnDesign)
                            .get_Item(pnPartNum).Value))
                    {
                    }

                    // Remove suppressed and excluded parts from the process
                    // Moved out here from inner checks
                    if (invOcc.Visible & !invOcc.Suppressed & !invOcc.Excluded)
                    {
                        tp = invOcc.Definition.Type;

                        // MessageBox.Show Join(new [] {"TYPE: " & tp,"VISIBLE: " & .Visible,"NAME: " & .Name,"Suboccurence: " & .SubOccurrences.Count,"Occurence Type: " & .Definition.Occurrences.Type,"BOMStructure: " & .BOMStructure), vbCrLf)

                        if (tp != kAssemblyComponentDefinitionObjectAnd)
                            tp<> kWeldmentComponentDefinitionObjectThen
                        //(moved suppression / exclusion check OUTSIDE)
                        If tp<>
                        kWeldsComponentDefinitionObject Then

                        // rt = dcAddAiDoc(aiDocument(.Definition.Document), rt)
                        // Recasting by aiDocument not likely necessary here.
                        // Revised to following:
                        rt = dcAddAiDoc(.Definition.Document, rt)

                        End If //inVisible, suppressed, excluded or Welds
                            Else //assembly, check BOM Structure
                        If.BOMStructure = kPurchasedBOMStructure Then //it's purchased
                        //Just add it to the Dictionary
                            rt = dcAddAiDoc(.Definition.Document, rt)
                        ElseIf.BOMStructure = kNormalBOMStructure Then //we make it
                        //Gather its components
                            rt = rt(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt), dcStops) //NOT forgetting to add THIS document!
                                ElseIf.BOMStructure =
                            kInseparableBOMStructure Then //maybe weldment ? If tp =
                            kWeldmentComponentDefinitionObject Then //it is
                            //Treat it like an assembly
                            rt = rt(.SubOccurrences, dcAddAiDoc(.Definition.Document, rt), dcStops)
                        Else //it's not
                        Stop //and see if we can figure out what its type is
                            End If
                        ElseIf.BOMStructure = kPhantomBOMStructure Then //"phantom" component
                        //Gather its components, but NOT the document itself
                            rt = rt(.SubOccurrences, rt, dcStops)
                        Else //not sure what we've got
                        Stop //and have a look at it
                        End If
                        End If //part or assembly
                            End If
                            End With
                            Next
                        End
                        
                            
                    }
                }
            }
        }
    }

    public Dictionary dcStopAt(string tx, Dictionary dc = null)
    {
        var rt = dc ?? dcStopAt(tx, new Dictionary());

        {
            if (!rt.Exists(tx))
                rt.Add(tx, tx);
        }

        return rt;
    }

    public Dictionary m0g0f0()
    {
        var rt = new Dictionary();
        return new[]
        {
            "29-197A", "29-197B", "29-633", "29-634", "29-637", "29-638", "29-647", "29-648",
            "29-650", "29-651", "29-652", "HD182"
        }.Aggregate(rt, (current, tx) => dcStopAt((tx), current));
    }

    public Dictionary m0g4f0(Dictionary dcIn)
    {
        // ' g4 currently reserved for development
        // ' relating to identification of purchased parts
        // ' by reference to Vault file path, as well as
        // ' BOM Structure and Cost Center settings
        // '
        // ' f0 separates members of incoming Dictionary
        // ' into design subgroups, including "doyle" and "purchased",
        // ' as indicated by the fourth component of each Vault path
        // '
        var rt = new Dictionary();
        {
            foreach (var ky in dcIn.Keys)
            {
                dynamic ls = Split(ky, @"\");
                long mx = UBound(ls);

                string fn = ls(mx);
                string g0 = ls(3);
                string g1 = ls(4);
                var fp = "";
                for (long dx = 5; dx <= mx - 1; dx++)
                    fp = fp + @"\" + ls(dx);
                fp = Strings.Mid(fp, 2);
                var rf = Join(new[] { fn, g1, fp }, Constants.vbTab);

                {
                    var withBlock1 = rt;
                    // Stop
                    Dictionary sd;
                    if (withBlock1.Exists(g0))
                        sd = withBlock1.get_Item(g0);
                    else
                    {
                        sd = new Dictionary();
                        withBlock1.Add(g0, sd);
                    }

                    {
                        var withBlock2 = sd;
                        if (withBlock2.Exists(rf))
                            Debugger.Break();
                        else
                            withBlock2.Add(rf, dcIn.get_Item(ky));
                    }
                }
            }

            ;
        }
        return rt;
    }


    public Dictionary m0g4f1(Dictionary dcIn)
    {
        // ' g4 currently reserved for development
        // ' relating to identification of purchased parts
        // ' by reference to Vault file path, as well as
        // ' BOM Structure and Cost Center settings
        // '
        // ' f1 scans documents in the "purchased" design group
        // ' for unexpected settings, specifically:
        // ' BOMStructure should be
        // ' kPurchasedBOMStructure (51973 / 0xCB05)
        // ' Design Tracking Property "Cost Center"
        // ' should be either "D-PTS" or "R-PTS"
        // '
        // ' Any documents not matching these parameters
        // ' are dropped into one or both subDictionaries
        // ' in the returned Dictionary, according to issue
        // '

        var rt = new Dictionary();
        var rtBom = new Dictionary();
        var rtFam = new Dictionary();
        rt.Add("bom", rtBom);
        rt.Add("fam", rtFam);
        Dictionary sd = dcIn.get_Item("purchased");
        {
            foreach (var ky in sd.Keys)
            {
                Document ivDoc = sd.get_Item(ky);
                var CpDef = aiCompDefOf(ivDoc);

                if (CpDef == null)
                    Debugger.Break();
                else
                {
                    if (CpDef.BOMStructure != kPurchasedBOMStructure)
                        rtBom.Add(ky, ivDoc);

                    var prFam = aiDocument(CpDef.Document).PropertySets.get_Item(gnDesign).get_Item(pnFamily);

                    if (prFam.Value != "D-PTS" | prFam.Value != "R-PTS")
                        rtFam.Add(ky, ivDoc);
                }
            }
        }
        return rt;
    }


    public Dictionary m0g4f1fixBom(Dictionary dcIn, long RiverView = 0)
    {
        // ' g4 currently reserved for development
        // ' relating to identification of purchased parts
        // ' by reference to Vault file path, as well as
        // ' BOM Structure and Cost Center settings
        // '
        // ' f1fixBom purports to fix incorrect BOM
        // ' Structure settings in purchased parts
        // '
        // Dim ivDoc As Inventor.Document

        Dictionary sd = dcIn.get_Item("bom");
        var rt = new Dictionary();
        {
            foreach (var ky in sd.Keys)
            {
                ComponentDefinition CpDef = aiCompDefOf(sd.get_Item(ky));
                if (CpDef == null)
                {
                    Debugger.Break();
                    rt.Add(ky, 0);
                }
                else
                {
                    CpDef.BOMStructure = kPurchasedBOMStructure;
                    rt.Add(ky, IIf(CpDef.BOMStructure == kPurchasedBOMStructure, 1, 0));
                }
            }
        }
        return rt;
    }


    public Dictionary m0g4f1fixFam(Dictionary dcIn, long RiverView = 0)
    {
        // ' g4 currently reserved for development
        // ' relating to identification of purchased parts
        // ' by reference to Vault file path, as well as
        // ' BOM Structure and Cost Center settings
        // '
        // ' f1fixFam purports to fix incorrect
        // ' Family settings in purchased parts
        // '

        var txFam = Interaction.IIf(RiverView, "R", "D") + "-PTS";
        Dictionary sd = dcIn.get_Item("fam");

        var rt = new Dictionary();
        {
            foreach (var ky in sd.Keys)
            {
                Document ivDoc = sd.get_Item(ky);
                {
                    var withBlock1 = ivDoc.PropertySets.get_Item(gnDesign).get_Item(pnFamily);
                    withBlock1.Value = txFam;
                    rt.Add(ky, IIf(withBlock1.Value == txFam, 1, 0));
                }

                ComponentDefinition CpDef = aiCompDefOf(sd.get_Item(ky));
                if (CpDef == null)
                {
                    Debugger.Break();
                    rt.Add(ky, 0);
                }
                else
                {
                    CpDef.BOMStructure = kPurchasedBOMStructure;
                    rt.Add(ky, IIf(CpDef.BOMStructure == kPurchasedBOMStructure, 1, 0));
                }
            }
        }
        return rt;
    }
}