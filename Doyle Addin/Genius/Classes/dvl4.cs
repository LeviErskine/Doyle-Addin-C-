using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class dvl4
{
    public static Dictionary d4g2f1(Document AiDoc = null)
    {
        // d4g2f1 -- not sure what
        // but looks like something to do
        // with grouping items by family
        // 
        Dictionary rt;

        if (AiDoc == null)
            rt = d4g2f1(aiDocActive());
        else
        {
            var withBlock = dcAiDocCompSetsByPtNum(AiDoc);
            // Stop
            if (withBlock.Exists(1))
                rt = dcFrom2Fields(
                    cnGnsDoyle().Execute("select ls.it, ISNULL(i.Family, '') fm from " +
                                         sqlValuesFromDc(dcOb(withBlock.get_Item(1)), "ls", "it") +
                                         " left join vgMfiItems i on ls.it = i.Item"), "it", "fm");
            else
            {
                rt = new Dictionary();
                Debugger.Break();
            }
        }

        return rt;
    }

    public static Dictionary d4g0f1(Document AiDoc, bool incTop = false)
    {
        // Dim gn As Scripting.Dictionary

        var rt = new Dictionary();
        {
            var withBlock = dcOfDcOfDcByPlurality(dcAiDocSetsByPtNum(dcAiDocComponents(AiDoc, null, incTop)));
            if (withBlock.Exists(2))
                // ky = TypeName( userChoiceFromDc( dcNewIfNone( dcOb(.get_Item(2)) )))
                ky = nuSelFromDict(dcNewIfNone(dcOb(withBlock.get_Item(2)))).GetReply();

            if (withBlock.Exists(""))
                Debugger.Break();

            Debugger.Break();
            // Debug.Print txDumpLs(dcNewIfNone(dcOb(obOf(.get_Item(1)))).Keys)
            Dictionary gp = dcNewIfNone(dcOb(obOf(withBlock.get_Item(1))));
            {
                var withBlock1 = dcNewIfNone(dcOb(obOf(withBlock.get_Item(1)))) // dcAiDocComponents(AiDoc, , incTop)
                    ;
                foreach (var ky in withBlock1.Keys)
                {
                    Document pt = aiDocument(withBlock1.get_Item(ky)); // .PartDocument
                    if (pt == null)
                    {
                    }
                    else
                        rt.Add(ky, dcGnsPtProps_Rev20220830_inProg(pt));
                }
            }
        }

        // gn = dcDxFromRecSetDc(dcFromAdoRS(' cnGnsDoyle().Execute(q1g1x2(AiDoc)')))
        // With gn
        // End With

        return rt;
    }

    public static Dictionary dcOfGnsProps(Document invDoc, Dictionary dc = null)
    {
        // dcOfGnsProps
        // 

        var rt = dcNewIfNone(dc);

        var dcPr = dcOfPropsInAiDoc(invDoc);
        {
            foreach (var ky in new[]
                     {
                         pnPartNum, pnFamily, pnDesc, pnMaterial, pnStockNum, pnCatWebLink, pnMass, pnRawMaterial,
                         pnRmQty, pnRmUnit, pnThickness, pnLength, pnWidth, pnArea
                     })
            {
                rt.Add(ky, dcPr.Exists(ky) ? dcPr.get_Item(ky) : null);
            }

            // NOTE[2022.09.16.1024]
            // extraction of Categories XML text
            // expected to move to content center
            // processing in dcGnsValFromContentCtr
            // If .Exists("Categories") Then
            if (Len(obAiProp(dcPr.get_Item("Categories")).Value) <= 0) return rt;
            rt.Add("Categories", dcPr.get_Item("Categories"));
            // rt.Add "Parameters", dcAiDocParVals(invDoc)
            Debug.Print(""); // Breakpoint Landing: Content Center
        }

        return rt;
    }

    public static Dictionary dcGnsValFromContentCtr(PartComponentDefinition CpDef, Dictionary dc = null)
    {
        // dcGnsValFromContentCtr
        // 

        Parameter pr;

        var rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            Dictionary wk;
            {
                {
                    var withBlock1 = aiDocPart(CpDef.Document).PropertySets;
                    string catXml = withBlock1.get_Item(gnDesign).get_Item("Categories").Value;

                    wk = dcAiPropsInSet(withBlock1.get_Item(guidPrSetCLib));
                }
            }

            {
                foreach (var ky in
                         wk.Keys) // new [] {"Member FileName", "Family","Standard", "Size Designation","Categories")
                    // If .Exists(ky) Then
                    rt.Add(ky, obAiProp(wk.get_Item(ky)).Value);
            }
        }

        return rt;
    }

    public static Dictionary dcGnsValFromPartCompDef(PartComponentDefinition CpDef, Dictionary dc = null)
    {
        // dcGnsValFromPartCompDef
        // 

        var rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            {
                // part of general ComponentDefinition
                // rt.Add "bomStruct", .BOMStructure

                if (CpDef.IsContentMember)
                {
                    Debug.Print(""); // Breakpoint Landing: Content Member
                    // Stop 'and look at this one
                    // rt.Add "ContentMem", 1
                    rt.Add("ccPropVals", dcGnsValFromContentCtr(CpDef));
                }

                {
                    var withBlock1 = CpDef.MassProperties;
                    rt.Add(pnMass, Round(cvMassKg2LbM * withBlock1.Mass, 4));
                }

                if (CpDef.IsiPartMember)
                {
                    rt.Add("isIPartMem", 1);
                    rt.Add("iPartFactory", CpDef.iPartMember.ReferencedDocumentDescriptor.FullDocumentName);
                }

                // part of general ComponentDefinition
                // With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
                // rt.Add "dimsModel", .Dictionary()
                // End With

                // part of general ComponentDefinition
                // MIGHT have some use
                // for this one in future
                // With .BOMQuantity
                // '.BaseUnits
                // End With

                // also possible, but unsure
                {
                    var withBlock1 = CpDef.Parameters;
                }
            }

            {
                var withBlock = dcFlatPatVals(aiCompDefShtMetal(CpDef), dcDotted());
                if (withBlock.Count > 2)
                    rt.Add("flatPat", dcUnDotted(withBlock.get_Item(".")));
            }
        }

        return rt;
    }

    public static Dictionary dcGnsValFromAssyCompDef(AssemblyComponentDefinition CpDef, Dictionary dc = null)
    {
        // dcGnsValFromAssyCompDef
        // 

        var rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            {
                // part of general ComponentDefinition
                // rt.Add "bomStruct", .BOMStructure

                {
                    var withBlock1 = CpDef.MassProperties;
                    rt.Add(pnMass, Round(cvMassKg2LbM * withBlock1.Mass, 4));
                }

                if (CpDef.IsiAssemblyMember)
                {
                    rt.Add("isIAssyMem", 1);
                    rt.Add("iAssyFactory", CpDef.iAssemblyMember.ReferencedDocumentDescriptor.FullDocumentName);
                }

                // part of general ComponentDefinition
                // With nuAiBoxData().UsingBox(.RangeBox).UsingInches()
                // rt.Add "dimsModel", .Dictionary()
                // End With

                // part of general ComponentDefinition
                // MIGHT have some use
                // for this one in future
                // With .BOMQuantity
                // '.BaseUnits
                // End With

                // also possible, but unsure
                {
                    var withBlock1 = CpDef.Parameters;
                }
            }

            {
                var withBlock = dcFlatPatVals(aiCompDefShtMetal(CpDef), dcDotted());
                if (withBlock.Count > 2)
                    rt.Add("flatPat", dcUnDotted(withBlock.get_Item(".")));
            }
        }

        return rt;
    }

    public static Dictionary dcGnsValFromGenCompDef(ComponentDefinition CpDef, Dictionary dc = null)
    {
        // dcGnsValFromGenCompDef
        // 

        var rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            rt.Add("bomStruct", CpDef.BOMStructure);
            {
                var withBlock1 = nuAiBoxData().UsingBox(CpDef.RangeBox).UsingInches();
                rt.Add("dimsModel", withBlock1.Dictionary());
            }
        }

        return dcGnsValFromAssyCompDef(aiCompDefAssy(CpDef), dcGnsValFromPartCompDef(aiCompDefPart(CpDef), rt));
    }

    public static Dictionary dcGnsValGeneral(Document AiDoc, Dictionary dc = null)
    {
        // dcGnsValGeneral
        // 

        var rt =
            // Dim dcPr As Scripting.Dictionary
            // Dim ky As dynamic
            dcNewIfNone(dc);

        if (AiDoc == null)
        {
        }
        else
        {
            {
                {
                    var withBlock1 = AiDoc.PropertySets.get_Item(gnDesign);
                    rt.Add(pnPartNum, withBlock1.get_Item(pnPartNum).Value);
                }

                rt.Add("subType", AiDoc.SubType);
            }

            rt = dcGnsValFromGenCompDef(aiCompDefOf(AiDoc), rt);
        }

        // dcPr = dcOfPropsInAiDoc(AiDoc)
        // With dcPr
        // For Each ky In new [] {' pnPartNum, pnFamily, pnDesc,' pnMaterial, pnStockNum, pnCatWebLink,' pnMass, pnRawMaterial, pnRmQty, pnRmUnit,' pnThickness, pnLength, pnWidth, pnArea' )
        // If .Exists(ky) Then
        // rt.Add ky, .get_Item(ky)
        // Else
        // rt.Add ky, Nothing
        // End If
        // Next
        // End With

        return rt;
    }

    public static Dictionary dcGnsPtProps_Rev20220830_inProg(Document AiDoc, Dictionary dc = null) // .PartDocument
    {
        // dcGnsPtProps_Rev20220830_inProg
        // 
        var rt =
            // Dim dcPr As Scripting.Dictionary
            // Dim dcVl As Scripting.Dictionary
            // Dim ky As dynamic
            // rt = dcNewIfNone(dc)
            dcGnsValGeneral(AiDoc, dcNewIfNone(dc));

        {
            var withBlock = dcOfGnsProps(AiDoc, dcDotted());
            if (withBlock.Count > 2)
            {
                rt.Add("props", dcUnDotted(withBlock.get_Item(".")));
                rt.Add("propVals", dcPropVals(rt.get_Item("props")));
                if (withBlock.Exists("Categories"))
                    rt.Add("Parameters", dcAiDocParVals(AiDoc));
            }
        }

        if (AiDoc == null)
        {
        }
        else
            // dcPr = dcOfGnsProps(AiDoc) 'dcOfPropsInAiDoc
            // dcVl = dcAiPropValsFromDc(dcPr)

        {
            {
                var withBlock1 = AiDoc.PropertySets.get_Item(gnDesign);
            }
        } // AiDoc Is Nothing

        // Stop
        // Call iSyncPartFactory(AiDoc)
        // rt = dcVl

        return rt;
    }

    public string sqlValuesFromDc(Dictionary dc, string vw = "ls", string it = "it")
    {
        while (true)
        {
            // sqlValuesFromDc - generate SQL
            // "VALUES" clause from Keys
            // of supplied Dictionary.
            // result is a relation
            // of one attribute
            // '
            // VALUES clause must end
            // with an AS phrase naming
            // the relation and all
            // attributes.
            // '
            // in this function, the
            // names default to "ls"
            // (list) for the relation,
            // and "it" (item) for
            // its one attribute
            // 
            if (dc == null)
            {
                dc = new Dictionary();
                vw = "ls";
                it = "it";
            }
            else
                return "(values ('" + Join(dc.Keys, "'), ('") + "')) as ls(it)";
        }
    }

    public static Dictionary dcAiDocCompSetsByPtNum(Document AiDoc, long incTop = 0)
    {
        // dcAiDocCompSetsByPtNum -- formerly d4g3f1
        // 
        // 
        // 
        var dc =
            // Dim ct As Long
            // ct = 1 'to include main assembly (for now)
            // now disabled in favor of input parameter incTop
            // dcAiDocSetsByPtNum replaces dcRemapByPtNum
            dcOfDcOfDcByPlurality(dcAiDocSetsByPtNum(dcAiDocComponents(AiDoc,
                null, incTop))); // incTop replaces ct

        // dcOfDcOfDcByItemCount() removed from original lineup
        // of dcOfDcOfDcByPlurality(dcOfDcOfDcByItemCount(dcAiDocSetsByPtNum(
        // dcOfDcOfDcByPlurality calls dcOfDcOfDcByItemCount internally
        // as part of its normal processing.
        return dc;
    }

    public static Dictionary dcAiDocSetsByPtNum(Dictionary dc)
    {
        // dcAiDocSetsByPtNum -- formerly d4g3f2
        // 
        // Returns Dictionary of Dictionaries
        // of Inventor Documents keyed on
        // associated Part Numbers.
        // 
        // Derived from dcRemapByPtNum, this
        // variation collects all models of a
        // given Part Number into a secondary
        // Dictionary, under each Document's
        // file name. Ideally, Part Numbers
        // should map one-to-one to Documents,
        // so each sub Dictionary should
        // contain only one entry.
        // 
        // However, as it IS possible for more
        // than one model to represent the same
        // Part, more than one Document might
        // in fact have the same Part Number.
        // 
        // Therefore, it may sometimes prove
        // necessary to take additional steps
        // in properly identifying which model
        // (or models) to process in preparation
        // for Genius.
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document pt = aiDocument(dc.get_Item(ky));
                var pn = Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));

                // REV[2022.05.17.1536]
                // removing check/handling
                // for blank/null part number.
                // client process can deal with that.
                Dictionary gp;
                {
                    if (rt.Exists(pn))
                    {
                    }
                    else
                        rt.Add(pn, new Dictionary());

                    gp = rt.get_Item(pn);
                }

                {
                    if (gp.Exists(ky))
                        Debugger.Break(); // this should NOT happen
                    else
                        gp.Add(ky, pt);
                }
            }
        }

        return rt;
    }

    public static Dictionary dcOfDcOfDcByItemCount(Dictionary dc)
    {
        // dcOfDcOfDcByItemCount -- formerly d4g3f3
        // 
        // subdivide supplied Dictionary
        // of Dictionaries into groups
        // by Count of members.
        // 
        // result is a 3rd-order Dictionary,
        // that is, a Dictionary (1)
        // keyed by member count
        // of Dictionaries (2)
        // keyed by a shared key
        // of yet more Dictionaries (3)
        // keyed to some unique value
        // 

        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                Dictionary gp = dc.get_Item(ky);
                long ct = gp.Count;

                Dictionary xp;
                {
                    if (!rt.Exists(ct))
                        rt.Add(ct, new Dictionary());

                    xp = rt.get_Item(ct);
                }

                {
                    if (xp.Exists(ky))
                        Debugger.Break(); // this should NOT happen
                    else
                        xp.Add(ky, gp);
                }
            }
        }

        return rt;
    }

    public static Dictionary dcOfDcOfDcByPlurality(Dictionary dc)
    {
        // dcOfDcOfDcByPlurality -- formerly d4g3f4
        // 
        // given a 2nd-order Dictionary (NOT 3rd)
        // as supplied by dcAiDocSetsByPtNum (NOT dcOfDcOfDcByItemCount),
        // return a reorganized version as follows:
        // 
        // under key 1: a Dictionary of all
        // Dictionaries having only one member.
        // this should form the bulk of the
        // supplied Dictionary's content.
        // 
        // under key 2: a Dictionary of Dictionaries
        // having more than one member. these
        // "plurals" might require additional
        // review and/or processing to resolve
        // ambiguities, conflicts, etc.
        // 
        // under key "" (blank string): the Dictionary,
        // if present, of members with no assigned
        // part or item number. this should almost
        // NEVER arise, but again, might require
        // special processing to resolve issues.
        // 
        // '''
        // 

        long ct;

        var rt = new Dictionary();

        // to avoid modifying supplied Dictionary,
        // generate a copy to work with directly.
        var gp = dcCopy(dc);
        {
            if (gp.Exists(""))
            {
                // of blank "numbered"
                // items moved over...

                rt.Add("", gp.get_Item(""));
                gp.Remove("");
            }
        }
        // before grouping by member counts:

        gp = dcOfDcOfDcByItemCount(gp);
        {
            // prep the "singles" return Dictionary
            var xp = new Dictionary();
            rt.Add(1, xp);
            if (gp.Exists(1))
            {
                // the (one, single) member of each
                // Dictionary under this one
                {
                    var withBlock1 = dcOb(gp.get_Item(1));
                    foreach (var ky in withBlock1.Keys)
                    {
                        {
                            var withBlock2 = dcOb(withBlock1.get_Item(ky));
                            xp.Add(ky, withBlock2.Items(0));
                        }
                    }
                }
                gp.Remove(1);
            }
            // at this point, any remaining members
            // should be "plural" Dictionaries
            // containing more than one member.

            // THESE are to be combined into one
            // "plural" Dictionary to be returned.
            xp = new Dictionary();
            // DO NOT add to return Dictionary yet!

            foreach (var ky in gp.Keys)
                // this step generates a NEW
                // Dictionary at each iteration.
                xp = dcKeysCombined(xp, dcOb(gp.get_Item(ky)), 1);

            // NOW, we can add the final result
            // to the return Dictionary...
            if (xp.Count > 0)
                // ...ASSUMING any are
                // left to add, of course!
                rt.Add(2, xp);
        }

        // Stop 'because NOT sure this thing
        // is ready for prime time...

        // disabling the following section completely
        // it should not be needed, as all three parts
        // of the return Dictionary should be in place
        // at the end of the preceding With block.

        // this SHOULD add Dictionaries for all
        // Part Numbers with more than one Document,
        // in a single Dictionary of "plurals",
        // but need to check it out for sure yet.
        // Hence the preceding Stop

        // rt.Add 2, dcKeysMissing(gp, rt.get_Item(1))

        // With dc: For Each ky In .Keys
        // gp = .get_Item(ky)
        // ct = gp.Count
        // 
        // With rt
        // If Not .Exists(ct) Then
        // .Add ct, New Scripting.Dictionary
        // End If
        // 
        // xp = .get_Item(ct)
        // End With
        // 
        // With xp
        // If .Exists(ky) Then
        // Stop 'this should NOT happen
        // Else
        // .Add ky, gp
        // End If: End With
        // Next: End With

        return rt;
    }

    public static Dictionary d4g3f5from2(Dictionary dc)
    {
        // d4g3f5from2
        // 
        // Returns Dictionary of Dictionaries
        // of Inventor Documents keyed on
        // associated Part Numbers.
        // 
        // Derived from dcRemapByPtNum, this
        // variation collects all models of a
        // given Part Number into a secondary
        // Dictionary, under each Document's
        // file name. Ideally, Part Numbers
        // should map one-to-one to Documents,
        // so each sub Dictionary should
        // contain only one entry.
        // 
        // However, as it IS possible for more
        // than one model to represent the same
        // Part, more than one Document might
        // in fact have the same Part Number.
        // 
        // Therefore, it may sometimes prove
        // necessary to take additional steps
        // in properly identifying which model
        // (or models) to process in preparation
        // for Genius.
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document pt = aiDocument(dc.get_Item(ky));

                // pt.co

                var pn = Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));

                // REV[2022.05.17.1536]
                // removing check/handling
                // for blank/null part number.
                // client process can deal with that.
                Dictionary gp;
                {
                    if (rt.Exists(pn))
                    {
                    }
                    else
                        rt.Add(pn, new Dictionary());

                    gp = rt.get_Item(pn);
                }

                {
                    if (gp.Exists(ky))
                        Debugger.Break(); // this should NOT happen
                    else
                        gp.Add(ky, pt);
                }
            }
        }

        return rt;
    }

    public static dynamic d4g3f5b(ComponentDefinition cd)
    {
    }

    public static dynamic d4g3f5a(ComponentDefinition cd) // Inventor.iAssemblyTableCell
    {
        // aiCompDefOf

        switch (cd)
        {
            case null:
                return null;
            case AssemblyComponentDefinition:
            {
                {
                    var withBlock = aiCompDefAssy(cd);
                    if (withBlock.IsiAssemblyMember)
                    {
                    }
                    else if (withBlock.IsiAssemblyFactory)
                    {
                    }
                    else if (withBlock.IsModelStateMember)
                    {
                    }
                    else if (withBlock.IsModelStateFactory)
                    {
                    }
                }
                break;
            }
            case PartComponentDefinition:
                break;
        }
    }

    public static Dictionary gnsUpdtAll_iFact(ComponentDefinition cd)
    {
        return cd switch
        {
            PartComponentDefinition => gnsUpdtAll_iPart(cd),
            AssemblyComponentDefinition => gnsUpdtAll_iAssy(cd),
            _ => new Dictionary()
        };
    }

    public static Dictionary gnsUpdtAll_iAssy(AssemblyComponentDefinition cd)
    {
        iAssemblyFactory fc;

        var rt = new Dictionary();

        if (cd == null)
        {
        }
        else if (cd.IsiAssemblyFactory)
        {
            {
                var withBlock = cd.iAssemblyFactory;
                AssemblyDocument md = withBlock.Parent.Parent;

                {
                    var withBlock1 = withBlock.TableColumns // .get_Item()
                        ;
                }

                // note initial DefaultRow
                var r0 = withBlock.DefaultRow;

                foreach (iAssemblyTableRow rw in withBlock.TableRows)
                {
                    withBlock.DefaultRow = rw;
                    rt.Add.DefaultRow.MemberName(null /* Conversion error: Set to default value for this argument */,
                        dcOfDcAiPropVals(dcGeniusProps(md)));
                }

                // restore initial DefaultRow
                withBlock.DefaultRow = r0;
            }
        }

        return rt;
    }

    public static Dictionary gnsUpdtAll_iPart(PartComponentDefinition cd)
    {
        iPartFactory fc;

        var rt = new Dictionary();

        if (cd == null)
        {
        }
        else if (cd.IsiPartFactory)
        {
            {
                var withBlock = cd.iPartFactory;
                PartDocument md = withBlock.Parent;

                {
                    var withBlock1 = withBlock.TableColumns // .get_Item()
                        ;
                }

                // note initial DefaultRow
                var r0 = withBlock.DefaultRow;

                foreach (iPartTableRow rw in withBlock.TableRows)
                {
                    withBlock.DefaultRow = rw;
                    rt.Add.DefaultRow.MemberName(null, dcAiPropValsFromDc(dcGeniusProps(md)));
                    DoEvents();
                    md.Save();
                }

                // restore initial DefaultRow
                withBlock.DefaultRow = r0;
            }
        }

        return rt;
    }

    public static Dictionary d4g3f6pt(PartComponentDefinition cd)
    {
        var rt = new Dictionary();
        if (cd == null)
        {
        }
        else if (cd.IsiPartFactory)
        {
            Debugger.Break();
            {
                var withBlock = cd.iPartFactory.TableColumns;
                long mx = withBlock.Count;
                Debugger.Break();
                for (long dx = 1; dx <= mx; dx++)
                {
                    var ck = withBlock.get_Item(dx);
                    {
                        rt.Add.Heading(null,
                            nuDcPopulator().Setting("dx", dx).Setting("dh", ck.DisplayHeading)
                                .Setting("fh", ck.FormattedHeading).Setting("dt", ck.ReferencedDataType)
                                .Setting("ob", ck.ReferencedObject).Setting("ot", TypeName(ck.ReferencedObject))
                                .Dictionary);
                        // 
                        // ).Setting("hd", .Heading'
                        Debugger.Break();
                    }
                }
            }
        }
        else if (cd.IsiPartMember)
        {
            Debugger.Break();
            rt = d4g3f6pt(aiDocPart(cd.iPartMember.ParentFactory.Parent).ComponentDefinition);
        }

        return rt;
    }

    public static Dictionary d4g3f6as(AssemblyComponentDefinition cd)
    {
        var rt = new Dictionary();
        if (cd == null)
        {
        }
        else if (cd.IsiAssemblyFactory)
        {
            Debugger.Break();
            {
                var withBlock = cd.iAssemblyFactory.TableColumns;
                long mx = withBlock.Count;
                Debugger.Break();
                for (long dx = 1; dx <= mx; dx++)
                {
                    var ck = withBlock.get_Item(dx);
                    Debugger.Break();
                }
            }
        }
        else if (cd.IsiAssemblyMember)
            Debugger.Break();

        return rt;
    }

    internal static readonly string[] Item = new[]
        { "Index", "Key", "CustomColumn", "DisplayHeading", "FormattedHeading", "ReferencedDataType" };

    public static Dictionary d4g3f7pt(PartComponentDefinition cd)
    {
        // Dim md As Inventor.PartDocument
        iPartFactory fc;
        // Dim r0 As Inventor.iPartTableRow
        // Dim df As Long
        var rt = new Dictionary();
        rt.Add("", new Dictionary());
        Dictionary hd = rt.get_Item("");

        if (cd == null)
            fc = null;
        else if (cd.IsiPartFactory)
            fc = cd.iPartFactory;
        else if (cd.IsiPartMember)
            fc = cd.iPartMember.ParentFactory;
        else
            fc = null;

        if (fc == null)
        {
            {
                iPartFactory withBlock = null;
                // md = .Parent '.Parent

                hd.Add("", Item);
                foreach (iPartTableColumn co in withBlock.TableColumns)
                {
                    {
                        var withBlock1 = co;
                        // .get_Item()
                        {
                            hd.Add.Heading(null, new[]
                            {
                                co.Index, co.Key, co.CustomColumn, co.DisplayHeading,
                                co.FormattedHeading, co.ReferencedDataType
                            });
                        }
                    }
                }

                // note initial DefaultRow
                // r0 = .DefaultRow

                foreach (iPartTableRow rw in withBlock.TableRows)
                {
                    {
                        // df = rw Is r0
                        rt.Add.Index(null, new[] { rw.Index, rw.MemberName, rw.PartName, rw });
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary d4g4f0(Document AiDoc = null)
    {
        // d4g4f0 -- rebuilding Sub Update_Genius_Properties
        // more or less from the ground up
        // 
        // Dim invProgressBar As Inventor.ProgressBar

        // Dim fc As gnsIfcAiDoc
        Dictionary rt;
        // Dim goAhead As VbMsgBoxResult
        // Dim ActiveDoc As Document
        // Dim txOut As String
        // Dim kyPt As dynamic
        long ct;

        // Dim dx As Long

        // Dim fm As fmIfcTest05A

        // NOTE[2022.06.01.1441]
        // adding check for supplied Document
        // call for selection if none
        if (AiDoc == null)
            // rt = d4g4f0(userChoiceFromDc(dcAiDocsVisible(), aiDocActive()))
            rt = d4g4f0(aiDocActive());
        else
        {
            // NOTE[2022.06.01.1442]
            // disabling/skipping user checks for now
            // this isn't the purpose of this whole mess.
            // Confirm User Request
            // to process active Document
            // goAhead = MessageBox.Show(' Join(new [] {' "Are you sure you want to process this document?",' "The process may require a few minutes depending on assembly size.",' "Suppressed and excluded parts will not be processed."' ), " "),' vbYesNo + vbQuestion,' "Process Document Custom iProperties"')
            // If goAhead = vbYes Then
            // 
            // Else
            // End If

            // NOTE[2022.06.01.1444]
            // see Update_Genius_Properties REV[2022.05.24.0956]
            {
                var withBlock = dcAiDocCompSetsByPtNum(AiDoc, ct) // ActiveDoc
                    ;
                if (withBlock.Exists(""))
                    Debugger.Break(); // for now

                if (withBlock.Exists(2))
                {
                    // THIS situation IS known to occur,
                    // if not TERRIBLY frequently, so a
                    // handler here is a good idea.
                    // 
                    {
                        var withBlock1 = nuDcPopulator(withBlock.get_Item(2)) // d4g4f4(dcOb(.get_Item(2)))
                            ;
                        Debug.Print(MessageBox.Show(msg_2022_0603_1127(withBlock1.Dictionary),
                            "Duplicate Part Numbers!", MessageBoxButtons.OK));
                    }
                }

                // and HERE is the step which ACTUALLY
                // replaces the prior version above.
                // Key 1 is guaranteed to be present
                // in the Dictionary returned, so no
                // need to check for it here.
                Dictionary dc = dcOb(withBlock.get_Item(1));
            }
            // 
            // 

            // NOTE[2022.06.01.1502]
            // this section expected to be
            // exported to its own function
            // NOTE[2022.06.02.0906]
            // (follow-up) original code
            // extracted to functions
            // dcOfKeys2match and d4g4f1
            var mt = dcOfKeys2match(new[]
                { pnFamily, pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnWidth, pnLength, pnArea, pnThickness });
            // pnFamily replaces "Cost Center"
            // pnMass replaces "GeniusMass"
            // pnRawMaterial replaces "RM"
            // pnRmQty replaces "RMQTY"
            // pnRmUnit replaces "RMUNIT"
            // pnWidth replaces "Extent_Width"
            // pnLength replaces "Extent_Length"
            // pnArea replaces "Extent_Area"
            // pnThickness replaces "Thickness"

            // rt = d4g4f1(dc, mt)

            rt = new Dictionary();
            {
                var withBlock = d4g4f1(dc, mt);
                foreach (var ky in withBlock.Keys)
                {
                    Dictionary wk;
                    {
                        if (!rt.Exists(ky))
                            rt.Add(ky, new Dictionary());

                        wk = rt.get_Item(ky);
                    }

                    {
                        var withBlock1 = dcOb(withBlock.get_Item(ky));
                        foreach (var k2 in withBlock1.Keys)
                            wk.Add(k2, obAiProp(withBlock1.get_Item(k2)).Value);
                    }
                }
            }
        }

        ;
        // 
        return rt;
    }

    public static Dictionary d4g4f1(Dictionary dc, Dictionary rf)
    {
        // d4g4f1 -- returns a Dictionary of Dictionaries
        // copied from supplied Dictionary dc,
        // but with only those Keys matching those
        // found in supplied 'reference' Dictionary rf
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
                rt.Add(ky, dcKeysInCommon(dcOfPropsInAiDoc(aiDocument(dc.get_Item(ky))), rf, 1));
        }

        return rt;
    }

    public static PartDocument d4g4f2(Dictionary dc)
    {
        // d4g4f2 -- given a Dictionary of Part Documents
        // return the first Content Center Member found
        // (if none found, return Nothing)
        // 

        PartDocument pt;

        {
            foreach (var ky in dc.Keys)
            {
                if (pt != null) continue;
                pt = aiDocPart(aiDocument(dc.get_Item(ky)));
                if (!pt != null) continue;
                if (!pt.ComponentDefinition().IsContentMember)
                    pt = null;
            }
        }

        return pt;
    }

    public static Dictionary d4g4f3(Dictionary dc)
    {
        // d4g4f3 -- given a Dictionary of Part Document
        // Dictionaries, return a subset containing
        // only Content Center Members
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                PartDocument pt = d4g4f2(dc.get_Item(ky));
                if (!pt == null)
                    rt.Add(ky, pt);
            }
        }

        return rt;
    }

    public static Dictionary d4g4f4(Dictionary dc)
    {
        // d4g4f4 -- given a Dictionary of Part Document
        // Dictionaries, return a subset dropping
        // any with Content Center Members
        // 
        return dcKeysMissing(dc, d4g4f3(dc));
    }

    public static Dictionary d4g4f5() // dc As Scripting.Dictionary'''
    {
        // d4g4f5 -- given a Dictionary of Part Document
        // Dictionaries, return a subset dropping
        // any with Content Center Members
        // 
        // Dim dcRb As Scripting.Dictionary
        dynamic ob;

        var rt = new Dictionary();
        // ThisApplication.UserInterfaceManager.RibbonState
        {
            var withBlock = ThisApplication.UserInterfaceManager;
            foreach (Ribbon rb in withBlock.Ribbons)
            {
                var dcRb = new Dictionary();
                {
                    rt.Add.InternalName(null, dcRb);
                    foreach (RibbonTab tb in rb.RibbonTabs) // .QuickAccessControls
                    {
                        var dcTb = new Dictionary();
                        {
                            dcRb.Add.InternalName(null, dcTb);
                            foreach (RibbonPanel rp in tb.RibbonPanels)
                            {
                                var dcRp = new Dictionary();
                                {
                                    var withBlock3 = rp;
                                    dcTb.Add.InternalName(null, dcRp);
                                    Debugger.Break();
                                }
                            }
                        }
                    }
                }
            }
        }
        return rt;
    }

    public static ComponentDefinition compDefOfPart(PartDocument AiDoc)
    {
        return (ComponentDefinition)AiDoc?.ComponentDefinition;
    }

    public static ComponentDefinition compDefOfAssy(AssemblyDocument AiDoc)
    {
        return (ComponentDefinition)AiDoc?.ComponentDefinition;
    }

    public static ComponentDefinition compDefOf(Document AiDoc)
    {
        var rt = compDefOfPart(aiDocPart(AiDoc)) ?? compDefOfAssy(aiDocAssy(AiDoc));

        // AiDoc.FullFileName
        return rt;
    }

    public static string famOfAiDoc(Document AiDoc)
    {
        // 

        string pf;

        VbMsgBoxResult ck;

        if (AiDoc == null)
            return "";
        {
            // NOTE!!! ONLY use this for Assemblies!
            // will disable until better set up
            // With nuDcPopulator(' ).Setting("doyle", "D"' ).Setting("riverview", "R"').Matching(Split(.FullDocumentName, "\"))
            // If .Count = 1 Then
            // pf = .get_Item(.Keys(0))
            // Else
            // pf = ""
            // End If
            // End With

            {
                var withBlock1 = AiDoc.PropertySets.get_Item(gnDesign);
                string mdFam = withBlock1.get_Item(pnFamily).Value;
                string itNum = withBlock1.get_Item(pnPartNum).Value;
                var gnFam = famInGenius(itNum);
            }
        }

        {
            var withBlock = compDefOf(AiDoc);
            if (withBlock.BOMStructure == kPurchasedBOMStructure)
                var sf = "PTS";
        }
    }

    public static string famInGenius(string itNum)
    {
        string gnFam;

        {
            var withBlock = cnGnsDoyle().Execute("select Family from vgMfiItems where Item = '" + itNum + "';");
            if (withBlock.BOF | withBlock.EOF)
                gnFam = "";
            else
                gnFam = Split(withBlock.GetString(adClipString, null, "", "", ""), Constants.vbCr)(0);
        }

        return gnFam;
    }

    public static string famIfValid(string mdFam)
    {
        {
            var withBlock = cnGnsDoyle().Execute(Join(new[]
            {
                "select ISNULL(f.Family, '') Family", "from (values ('" + mdFam + "')) as i(f)",
                "left join vgMfiFamilies f on i.f = f.Family"
            }));
            if (withBlock.BOF | withBlock.EOF)
                return "";
            return withBlock.Fields("Family").Value;
        }
    }

    public static string famVsGenius(string itNum, string mdFam = "")
    {
        // get current family from
        // Genius, if it has one
        var gnFam = famInGenius(itNum);

        if (Strings.Len(gnFam) == 0)
            return mdFam; // so just use the model's
        // first, verify model family
        var ckFam = famIfValid(mdFam);
        // if not in Genius...

        if (Strings.Len(ckFam) == 0)
            return gnFam;
        if (gnFam == ckFam)
            return ckFam;
        VbMsgBoxResult ck = MessageBox.Show(Join(new[]
            {
                "Item " + itNum, "Model Part Family " + ckFam + " differs", "from Genius Part Family " + gnFam, "",
                "Change Model to match Genius?", "", "(click [CANCEL] to debug)"
            }, Constants.vbCrLf),
            Constants.vbYesNoCancel + Constants.vbQuestion, "Use Genius Family?");

        if (ck == Constants.vbCancel)
            Debugger.Break(); // to debug
        else if (ck == Constants.vbYes)
            return gnFam;
        else
            return ckFam;
    }

    public static string msg_2022_0603_1127(Dictionary dc)
    {
        // msg_2022_0603_1127
        // 

        var rt = "";

        var cc = d4g4f3(dc);
        var rm = dcKeysMissing(dc, cc);

        if (rm.Count > 0)
            rt = Join(new[]
            {
                rt, "The following Part Numbers are", "assigned to more than one Model:", "",
                Constants.vbTab + Join(rm.Keys, Constants.vbCrLf + Constants.vbTab), ""
            }, Constants.vbCrLf);

        if (cc.Count > 0)
            rt = Join(new[]
            {
                rt, "These duplicated Part Numbers are", "associated with at least one Content",
                "Center Member, which cannot be modified:", "",
                Constants.vbTab + Join(cc.Keys, Constants.vbCrLf + Constants.vbTab), ""
            }, Constants.vbCrLf);

        rt = Join(new[] { rt, "These will not be processed.", "" }, Constants.vbCrLf);

        return
            rt; // Join(new [] {"The following Part Numbers are","assigned to more than one Model:","",vbTab & Join(d4g4f4(.Dictionary).Keys, vbCrLf & vbTab),"","These duplicated Part Numbers are","associated with at least one Content","Center Member, which cannot be modified:","",vbTab & Join(cc.Keys, vbCrLf & vbTab),"","These will not be processed.",""), vbCrLf)
    }

    public static Dictionary askUserForPartMatl(PartDocument AiDoc, Dictionary dc = null)
    {
        // askUserForPartMatl -- Prompt User
        // for Part Family and Material
        // Selection, returning result
        // in Dictionary
        // 
        // REV[2022.08.29.1621]
        // add optional Dictionary parameter to which
        // data from this function can be added.
        // (see also askUserForMatlQty)
        // 
        Dictionary rt;

        if (dc == null)
            rt = askUserForPartMatl(AiDoc, new Dictionary());
        else
        {
            rt = dc; // New Scripting.Dictionary

            {
                if (rt.Exists(pnFamily))
                    rt.Remove(pnFamily);
                if (rt.Exists(pnRawMaterial))
                    rt.Remove(pnRawMaterial);
                if (AiDoc == null)
                {
                    rt.Add(pnFamily, "");
                    rt.Add(pnRawMaterial, "");
                }
                else
                {
                    var withBlock1 = newFmTest1();
                    aiBoxData bd = nuAiBoxData().UsingInches.SortingDims(AiDoc.ComponentDefinition.RangeBox);
                    VbMsgBoxResult ck = withBlock1.AskAbout(AiDoc,
                        "No Stock Found! Please Review" + Constants.vbCrLf + Constants.vbCrLf + bd.Dump(0));

                    if (ck == Constants.vbYes)
                    {
                        // Stop 'because this will
                        // override supplied Dictionary!
                        // rt =
                        {
                            var withBlock2 = withBlock1.ItemData();
                            rt.Add(pnFamily, withBlock2.get_Item(pnFamily));
                            rt.Add(pnRawMaterial, withBlock2.get_Item(pnRawMaterial));
                        }
                    }
                    else
                    {
                        // rt = New Scripting.Dictionary
                        Debugger.Break();

                        {
                            var withBlock2 = AiDoc.PropertySets;
                            string tx = withBlock2.get_Item(gnDesign).get_Item(pnFamily).Value;
                            rt.Add(pnFamily, tx);
                            Information.Err().Clear();
                            tx = withBlock2.get_Item(gnCustom).get_Item(pnRawMaterial).Value;
                            if (Information.Err().Number)
                                tx = "";

                            rt.Add(pnRawMaterial, tx);
                        }
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary askUserForMatlQty(PartDocument AiDoc, Dictionary dc = null)
    {
        // askUserForMatlQty -- Prompt User
        // for Material Quantity and Units,
        // returning result in Dictionary
        // 
        // REV[2022.08.29.1624]
        // add optional Dictionary parameter to which
        // data from this function can be added.
        // (see also askUserForPartMatl)
        // 
        Dictionary rt;
        VbMsgBoxResult ck;
        aiBoxData bd;
        string tx;

        if (dc == null)
            rt = askUserForMatlQty(AiDoc, new Dictionary());
        else
        {
            rt = dc; // New Scripting.Dictionary
            {
                if (rt.Exists(pnRmQty))
                    rt.Remove(pnRmQty);
                if (rt.Exists(pnRmUnit))
                    rt.Remove(pnRmUnit);
            }

            if (AiDoc == null)
            {
                rt.Add(pnRmQty, 0);
                rt.Add(pnRmUnit, "");
            }
            else
            {
                var withBlock = nu_fmIfcMatlQty01().SeeUserWithPart(AiDoc);
                // following copied from dcGeniusPropsPartRev20180530 line 1632~?
                if (withBlock.Exists(pnRmQty))
                    rt.Add(pnRmQty, withBlock.get_Item(pnRmQty));

                if (withBlock.Exists(pnRmUnit))
                    rt.Add(pnRmUnit, withBlock.get_Item(pnRmUnit));
            }
        }

        return rt;
    }

    public static Dictionary askUserForPartMatlUpdate(PartDocument AiDoc)
    {
        // askUserForPartMatlUpdate --
        // Attempt to update Part Document
        // material Properties from results
        // of askUserForPartMatl
        // (Family and Material Selection)
        // and askUserForMatlQty
        // (Material Quantity and Units)
        // Return Dictionary of results
        // 
        // NOTE[2022.08.29.1627]
        // want to separate user data collection
        // from property updates in this function.
        // further review/development called for.
        // 
        Dictionary dcWk;
        Property pr;

        var dcPr = dcOfPropsInAiDoc(AiDoc);
        VbMsgBoxResult ck = Constants.vbOK;
        if (!dcPr.Exists(pnRawMaterial))
        {
            ck = MessageBox.Show(Join(new[]
                {
                    "Custom Property " + pnRawMaterial + ",", "used to identify Raw Material,",
                    "is not yet present in this model.", "", "Go ahead and create it?"
                }, Constants.vbCrLf),
                Constants.vbYesNo + Constants.vbQuestion, "Required Property Missing!");

            if (ck == Constants.vbYes)
            {
                {
                    var withBlock = AiDoc.PropertySets.get_Item(gnCustom);
                    foreach (var ky in new[] { pnRawMaterial, pnRmQty, pnRmUnit })
                    {
                    }

                    Information.Err().Clear();
                    pr = withBlock.Add("", ky); // pnRawMaterial
                    if (Information.Err().Number == 0)
                    {
                        dcPr.Add(ky, pr);
                        ck = Constants.vbOK;
                    }
                    else
                        ck = Constants.vbAbort;
                }
            }
            else
                ck = Constants.vbOK;
        }

        if (ck != Constants.vbOK)
            ck = MessageBox.Show(Join(new[]
                {
                    "Custom Property " + pnRawMaterial + ",", "was not created! Raw Material", "will not be saved!"
                },
                Constants.vbCrLf), Constants.vbOKCancel & Constants.vbExclamation, "Property Not Created!");

        var rt = new Dictionary();

        if (ck == Constants.vbOK)
        {
            // REV[2022.08.29.1616]
            // condense two nearly identical With blocks
            // into one, combining results of part material
            // and material quantity data collections.
            // NOTE: this required additional REVs
            // to askUserForPartMatl (nee d4g1f1)
            // and askUserForMatlQty (nee d4g1f3)
            // to accept optional Dictionary to receive
            // data points collected by each function.
            {
                var withBlock = askUserForMatlQty(AiDoc, askUserForPartMatl(AiDoc));
                foreach (var ky in withBlock.Keys)
                {
                    {
                        pr = dcPr.Exists(ky) ? (Property)dcPr.get_Item(ky) : null;
                    }

                    if (pr != null) continue;
                    if (Len(Trim(withBlock.get_Item(ky))) > 0)
                    {
                        if (pr.Value != withBlock.get_Item(ky))
                        {
                            Information.Err().Clear();
                            // Stop 'so we can make sure this works
                            pr.Value = withBlock.get_Item(ky);
                            Debug.Print(""); // Breakpoint Landing
                            // DON'T try to step at pr.Value
                            if (Information.Err().Number)
                                Debugger.Break();
                        }
                    }

                    rt.Add(ky, pr.Value);
                }
            }
        }

        return rt;
    }

    public static Dictionary dcGeniusPropsPartRev20180530_ck(PartDocument invDoc, Dictionary dc = null)
    {
        while (true)
        {
            // 
            // NOTICE TO DEVELOPER [2021.11.12]
            // '''
            // 
            // This function definition was restored
            // from a prior copy of this project
            // (VB-000-1002_2021-1001.ipt)
            // to restore current "normal" operation
            // of the Genius Properties Update macro.
            // The prior development version was
            // retained for reference, renamed to
            // dcGeniusPropsPartRev20180530_ck_broken
            // 
            // One minor revision was made to this
            // restored version to retain improved
            // generation of Genius Mass data.
            // Additional changes should be kept
            // to a MINIMUM to maintain correct
            // operation going forward, and any
            // desired changes implemented through
            // some form of "shim"
            // 
            // '''
            // REV[2022.01.21.1351]
            // Added following two Dictionaries
            // to collect settings already in Genius
            // to add a layer of separation
            // to FlatPattern data collection
            // (might not want to use for Properties
            // so don't update immediately)

            // '
            // '
            // ADDED[2021.03.11] to simplify access
            // to Part Number of Model, since it's
            // requested several times in function
            // '
            // ADDED[2021.03.11] to further
            // simplify access to Part Number
            // UPDATE[2018.05.30]:
            // Rename variable Family to nmFamily
            // to minimize confusion between code
            // and comment text in searches.
            // Also add variable mtFamily
            // for raw material Family name
            string qtUnit;

            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            var rt = dc;

            {
                // Property Sets
                PropertySet aiPropsUser;
                PropertySet aiPropsDesign;
                {
                    var withBlock1 = invDoc.PropertySets;
                    aiPropsUser = withBlock1.Add(gnCustom);
                    aiPropsDesign = withBlock1.Add(gnDesign);
                }

                // Custom Properties
                var prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1); // pnRawMaterial
                var prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1); // pnRmUnit
                var prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1); // pnRmQty

                // Part Number and Family properties
                // are from Design, NOT Custom set
                var prPartNum = aiGetProp(aiPropsDesign, pnPartNum); // pnPartNum
                // ADDED 2021.03.11
                string pnModel = prPartNum.Value;
                var prFamily = aiGetProp(aiPropsDesign, pnFamily);

                // We should check HERE for possibly misidentified purchased parts
                // UPDATE[2018.02.06]: Using new UserForm; see below
                VbMsgBoxResult ck;
                BOMStructureEnum bomStruct;
                string nmFamily;
                {
                    var withBlock1 = invDoc.ComponentDefinition;
                    // Request #1: the Mass in Pounds
                    // and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        // Update [2021.11.12]
                        // Round mass to nearest ten-thousandth
                        // to try to match expected Genius value.
                        // This should reduce or minimize reported
                        // discrepancies during ETM process.
                        rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                    }

                    // BOM Structure type, correcting if appropriate,
                    // and prepare Family value for part, if purchased.
                    // 
                    ck = Constants.vbNo;
                    // UPDATE[2018.05.31]: Combined both InStr checks
                    // by addition to generate a single test for > 0
                    // If EITHER string match succeeds, the total
                    // SHOULD exceed zero, so this SHOULD work.
                    if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") + InStr(1,
                            "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0)
                        // UPDATE[2018.02.06]: Using same
                        // new UserForm as noted above.
                        ck = newFmTest2().AskAbout(invDoc,
                            null /* Conversion error: Set to default value for this argument */,
                            "Is this a Purchased Part?" + Constants.vbCrLf + "(Cancel to debug)");

                    // Check process below replaces duplicate check/responses above.
                    if (ck == Constants.vbCancel)
                        Debugger.Break();
                    else if (ck == Constants.vbYes)
                    {
                        if (withBlock1.BOMStructure != kPurchasedBOMStructure)
                        {
                            withBlock1.BOMStructure = kPurchasedBOMStructure;
                            bomStruct = Information.Err().Number == 0
                                ? withBlock1.BOMStructure
                                : kPurchasedBOMStructure;
                        }
                        else
                            bomStruct = withBlock1.BOMStructure; // to make sure this is captured
                    }
                    else
                        bomStruct = withBlock1.BOMStructure; // to make sure this is captured

                    // Request #2: Change Cost Center iProperty.
                    // If BOMStructure = Purchased and not content center,
                    // then Family = D-PTS, else Family = D-HDWR.
                    // 
                    // UPDATE[2018-05-30]: Value produced here
                    // will now be held for later processing,
                    // more toward the end of this function.
                    if (bomStruct == kPurchasedBOMStructure)
                    {
                        nmFamily = withBlock1.IsContentMember ? "D-HDWR" : "D-PTS";
                    }
                    else
                        nmFamily = "";
                }
                // At this point, nmFamily SHOULD be set
                // to a non-blank value if Item is purchased.
                // We should be able to check this later on,
                // if Item BOMStructure is NOT Normal

                switch (bomStruct)
                {
                    // Request #4: Change Cost Center iProperty.
                    // If BOMStructure = Normal, then Family = D-MTO,
                    // else if BOMStructure = Purchased then Family = D-PTS.
                    case kNormalBOMStructure:
                    {
                        // REV[2022.01.28.1014]
                        // Added initial raw material capture
                        // to check against Genius
                        // HOLD![2022.01.28.1046]
                        // commenting out again
                        // probably best below, still
                        string pnStock = prRawMatl.Value;
                        // REV[2022.02.08.1304]
                        // restored, to obtain any
                        // value already defined.
                        // MIGHT need moved further down,
                        // but hold off on that for now.

                        // REV[2022.01.17.1123]
                        // Start adding code to capture
                        // any raw material items for
                        // part already in Genius.
                        // REV[2022.01.21.1357]
                        // Separated capture from With statement
                        // into new Dictionary dynamic in order
                        // to check and use it further down,
                        // as well as passing it to nuSelFromDict
                        // to handle multiple line items
                        // REV[2022.01.31.1008]
                        // Restored assignment of dcFromAdoRS
                        // result to Dictionary dynamic dcIn,
                        // in order to pass it to other
                        // functions, as needed.
                        // 
                        var dcIn = dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)));
                        // Debug.Print ConvertToJson(dcDxFromRecSetDc(dcIn), vbTab)
                        // Stop
                        // dcIn = dcOb(dcDxFromRecSetDc(dcIn).get_Item(pnRawMaterial))
                        if (dcIn.Count > 0)
                        {
                            {
                                var withBlock1 = dcOb(dcDxFromRecSetDc(dcIn).get_Item(pnRawMaterial));
                                // REV[2022.01.28.1336]
                                // Added code to collect captured
                                // dcIn = New Scripting.Dictionary

                                // REV[2022.01.28.0857]
                                // Added code to collect captured
                                // material item number, asking user
                                // to select from list if more than one.
                                if (withBlock1.Count > 0)
                                {
                                    if (Strings.Len(pnStock) > 0)
                                    {
                                        // some material already assigned
                                        if (withBlock1.Exists(pnStock))
                                        {
                                        }
                                        else
                                            // so forget current value (for now)
                                            pnStock = "";
                                    }

                                    if (Strings.Len(pnStock) == 0)
                                        // grab first material item found
                                        // Stop
                                        // pnStock = dcOb(.get_Item(.Keys(0))).get_Item(pnRawMaterial)
                                        pnStock = withBlock1.Keys(0);

                                    // and use it for the default...
                                    if (withBlock1.Count > 1)
                                    {
                                        Debugger.Break(); // because selection is going
                                        // to be a lot more complicated.
                                        // (just look at that pnStock
                                        // assignment up there!)

                                        pnStock = nuSelector().GetReply(withBlock1.Keys, pnStock);

                                        Debugger.Break(); // to make sure things are okay
                                    }
                                }

                                // REV[2022.01.28.0903]
                                // Separated Dictionary capture
                                // from Count check
                                if (Strings.Len(pnStock) > 0)
                                {
                                    if (Len(Convert.ToHexString(prRawMatl.Value)) == 0)
                                    {
                                    }
                                    else if (pnStock == prRawMatl.Value)
                                    {
                                    }
                                    else
                                    {
                                        Debug.Print("=== CURRENT GENIUS MATERIAL DATA ===");
                                        // Debug.Print dumpLsKeyVal(dcIn, ":" & vbTab)
                                        ck = newFmTest2().AskAbout(invDoc, "Raw Material " + prRawMatl.Value,
                                            "does not match " + pnStock);
                                        if (ck == Constants.vbCancel)
                                            Debugger.Break(); // to check things out
                                        else if (ck == Constants.vbNo)
                                            // NOTE[2022.02.08.1359]
                                            // DO NOT DISABLE this instance
                                            // of the pnStock assignment!
                                            pnStock = prRawMatl.Value;
                                    }

                                    // REV[2022.01.28.1448]
                                    // Changed data extraction process here
                                    // to work with form returned from dcFromAdoRS
                                    // 
                                    // NOTE! This is !!!TEMPORARY!!!
                                    // Implemented during run time,
                                    // some truly insane acrobatics were required
                                    // to make it work without resetting the run.
                                    // This code, including the With statement
                                    // above, MUST be rewritten as soon as feasible!
                                    // 
                                    // Stop 'because we're doing to need to do something different
                                    // Debug.Print ConvertToJson(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).get_Item(pnRawMaterial), vbTab)
                                    // Debug.Print ConvertToJson(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).get_Item(pnRawMaterial)).get_Item(pnStock), vbTab)
                                    // Debug.Print ConvertToJson(dcOb(.get_Item(dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).get_Item(pnRawMaterial)).get_Item(pnStock)).Keys(0))), vbTab)
                                    // dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).get_Item(pnRawMaterial)).get_Item(pnStock)).Keys(0)
                                    // Stop

                                    if (withBlock1.Exists(pnStock))
                                    {
                                        dcIn = dcOb(dcIn.get_Item(dcOb(withBlock1.get_Item(pnStock)).Keys(0)));
                                        // This is DEFINITELY going to need a rework!
                                        // But that will need a new function, most likely

                                        // deactivated the version below
                                        // to be superceded by the one above
                                        // dcIn = dcOb(.get_Item(dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).get_Item(pnRawMaterial)).get_Item(pnStock)).Keys(0)))

                                        // original version, also deactivated
                                        // for obvious reasons
                                        // dcIn = .get_Item(pnStock)

                                        Debug.Print(""); // Breakpoint Landing
                                    }
                                    else
                                        Debugger.Break(); // because we've got a REAL problem here!
                                }
                                else
                                    dcIn = new Dictionary();
                            }
                        }

                        {
                            if (dcIn.Count == 0)
                            {
                                dcIn.Add("Ord", 0);
                                dcIn.Add("RM", "");
                                dcIn.Add("MtFamily", "");
                                dcIn.Add("RMQTY", 0);
                                dcIn.Add("RMUNIT", "");
                            }
                        }

                        // ----------------------------------------------------'
                        aiBoxData bd;
                        if (invDoc.SubType == guidSheetMetal)
                        {
                            // ----------------------------------------------------'
                            // NOTE[2018-05-31]: At this point, we MAY wish
                            // to check for a valid flat pattern,
                            // and otherwise attempt to verify
                            // an actual sheet metal design.
                            // 

                            // REV[2022.01.28.0903]
                            // HERE is where things start to get interesting
                            // Before processing Part as sheet metal,
                            // want to make sure it's supposed to be.
                            // 
                            // FIRST, check what Genius had to say
                            string mtFamily;
                            {
                                mtFamily = dcIn.Exists("MtFamily") ? (string)dcIn.get_Item("MtFamily") : "";
                            }

                            if (Strings.Len(mtFamily) == 0)
                                ck = Constants.vbRetry;
                            else if (mtFamily == "DSHEET")
                                ck = Constants.vbYes;
                            else
                                ck = Constants.vbNo;

                            // REV[2022.01.31.1335]
                            // Move flat pattern collection out here
                            // from inside the next If-Then block
                            Dictionary dcFP;
                            if (ck == Constants.vbNo)
                                dcFP = new Dictionary();
                            else
                            {
                                dcFP = dcFlatPatVals(invDoc.ComponentDefinition);
                                // try to get flat pattern data
                                // WITHOUT mucking up Properties!
                                // Want to avoid dirtying file with
                                // changes until absolutely necessary)

                                if (dcFP.Exists(pnThickness))
                                {
                                    pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                    dcFP.Add(pnRawMaterial, pnStock);
                                }
                            }

                            Debug.Print(""); // Breakpoint Landing
                            if (false)
                                Debug.Print(ConvertToJson(new[]
                                {
                                    dcIn, dcFP
                                }, Constants.vbTab));
                            if (ck == Constants.vbRetry)
                            {
                                // so let's see what the flat pattern can tell us
                                if (dcFP.Exists("mtFamily"))
                                {
                                    if (dcFP.get_Item("mtFamily") == "DSHEET")
                                    {
                                        if (dcFP
                                            .Exists("OFFTHK"))
                                        {
                                            Debugger.Break();
                                            ck = newFmTest2().AskAbout(invDoc,
                                                "This Part: ",
                                                "might not be sheet metal. " + Constants.vbCrLf + Constants.vbCrLf +
                                                "Is it in fact sheet metal?");
                                            if (ck == Constants.vbCancel)
                                            {
                                                ck =
                                                    Constants.vbRetry;
                                                Debugger.Break(); // to debug
                                            }
                                        }
                                        else ck = Constants.vbYes;
                                    }
                                    else if (dcFP.get_Item("mtFamily") == "D-BAR") ck = Constants.vbNo;
                                    else ck = Constants.vbRetry;
                                }
                                else ck = Constants.vbRetry;
                            }

                            if (ck == Constants.vbRetry)
                            {
                                Debug.Print(ConvertToJson(new[]
                                {
                                    dcIn, dcFP
                                }, Constants.vbTab));
                                Debugger.Break(
                                ); // so we can figure out what to do next.
                            }

                            // Request #3:
                            // sheet metal extent area
                            // and add to custom property "RMQTY"

                            // REV[2022.01.28.1556]
                            // change if-then-else sequence
                            // to check ck instead of dcIn
                            if (ck == Constants.vbYes) rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                            else if (ck == Constants.vbRetry) rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                            else if (ck == Constants.vbNo)
                            {
                            }
                            else
                                // material type detection SHOULD produce
                                // one of the three preceding values
                                Debugger.Break(); // and check it out

                            // NOTE[2018-05-30]:
                            // Raw Material Quantity value
                            // SHOULD be set upon return
                            // We may need to review the process
                            // to find an appropriate place
                            // to set for NON sheet metal

                            // Moved to start of block to check for NON sheet metal

                            // NOTE: THIS call might best be combined somehow
                            // with the flat pattern prop pickup above.
                            // Note especially that if dcFlatPatProps
                            // FINDS NO .FlatPattern, then there should
                            // BE NO sheet metal part number!
                            if (prRawMatl == null)
                            {
                                if (rt.Exists("OFFTHK"))
                                {
                                    // NOTE[2021.12.10]:
                                    // Believe this OFFTHK property is meant
                                    // to capture "Sheet Metal" Parts that
                                    // aren't actually Sheet Metal.
                                    // This check might be needed further down.
                                    // UPDATE[2018.05.30]:
                                    // Restoring original key check
                                    // and adding code for debug
                                    // Previously changed to "~OFFTHK"
                                    // to avoid this block and its issues.
                                    // (Might re-revert if not prepped to fix now)
                                    Debug.Print(aiProperty(rt.get_Item("OFFTHK")).Value);
                                    Debugger.Break(); // because we're going to need to do something with this.
                                    pnStock = ""; // Originally the ONLY line in this block.
                                    // A more substantial response is required here.
                                    if (0) Debugger.Break(); // (just a skipover)
                                }
                                else
                                {
                                    Debugger.Break(); // because we don't know IF this is sheet metal yet
                                    pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                                }
                            }
                            else
                            {
                                // ' ACTION ADVISED[2018.09.14]:
                                // ' pnStock can probably be set
                                // ' to prRawMatl.Value and THEN
                                // ' checked for length to see
                                // ' if lookup needed.
                                // ' This might also allow us to check
                                // ' for machined or other non-sheet
                                // ' metal parts.

                                // REV[2021.12.17]: sanity check
                                // Add sanity check to make sure
                                // any existing sheet metal stock
                                // number matches model specs
                                if (Len(prRawMatl.Value) > 0)
                                {
                                    // we need to check it
                                    if (Strings.Len(pnStock) == 0)
                                        // REV[2022.01.28.1445]:
                                        // Placed this pnStock stock assignment
                                        // inside this If-Then block to prevent
                                        // overriding value from Genius
                                        pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                                    // NOTE[2021.12.17@15:32]:
                                    // copied this up from
                                    // NOTE[2021.12.17@15:32]
                                    // for use in sanity check

                                    // NOTE[2021.12.17]:
                                    // This section simply warns the user
                                    // that the current raw material does
                                    // not match the recommended default,
                                    // and offers an opportunity to fix it.
                                    // 
                                    // This is yet another quick and dirty
                                    // "solution" that should be revised
                                    // NOTE[2022.01.05]:
                                    // Adding check for null recommendation.
                                    // Do NOT believe user should be offered
                                    // opportunity to overwrite any current
                                    // part number with a BLANK one. Believe
                                    // the option to CLEAR is somewhere below.
                                    if (Strings.Len(pnStock) > 0)
                                    {
                                        if (pnStock != prRawMatl.Value)
                                        {
                                            // Stop

                                            // NOTE[2022.01.03]:
                                            // Following text SHOULD no longer
                                            // be needed. Verify function of
                                            // fmTest2 following, and when good,
                                            // disable and/or remove this block.
                                            Debug.Print("!!! NOTICE !!!");
                                            Debug.Print("Recommended Sheet Metal Stock (" + pnStock + ")");
                                            Debug.Print("does not match current Stock (" + prRawMatl.Value + ")");
                                            Debug.Print("");
                                            Debug.Print("To continue with no change, just press [F5]. Otherwise,");
                                            Debug.Print("press [ENTER] on the following line first to change:");
                                            Debug.Print("prRawMatl.Value = \"" + pnStock + "\"");
                                            Debug.Print("");

                                            // NOTE[2022.01.03]:
                                            // Now using fmTest2(?) to prompt
                                            // user as in other checks (above?)
                                            ck = newFmTest2().AskAbout(invDoc,
                                                "Suggest Sheet Metal change from" + Constants.vbCrLf + prRawMatl.Value +
                                                " to" + Constants.vbCrLf + pnStock + " for", "Change it?");
                                            if (ck == Constants.vbCancel) Debugger.Break(); // to check things out
                                            else if (ck == Constants.vbYes)
                                                // Stop
                                                prRawMatl.Value = pnStock;
                                        }
                                    }
                                }
                                else if (Strings.Len(pnStock) > 0)
                                    // go ahead and assign material
                                    prRawMatl.Value = pnStock;

                                if (Len(prRawMatl.Value) > 0)
                                {
                                    if (rt.Exists("OFFTHK"))
                                    {
                                        // Stop 'and verify raw material item
                                        // NOTE[2021.12.13]:
                                        // OFFTHK property check added
                                        // to catch sheet metal already
                                        // assigned by accident.
                                        ck = newFmTest2().AskAbout(invDoc, "Assigned Raw Material " + prRawMatl.Value,
                                            "Clear it?");
                                        if (ck == Constants.vbCancel) Debugger.Break(); // to check things out
                                        else if (ck == Constants.vbYes) prRawMatl.Value = "";
                                    }

                                    if (pnStock == prRawMatl.Value)
                                        // no need to assign it again
                                        Debug.Print(""); // Breakpoint Landing
                                    else
                                    {
                                        Debug.Print(ConvertToJson(new[] { pnStock, prRawMatl.Value }));
                                        Debugger.Break(); // before we do something stupid!
                                        pnStock = prRawMatl.Value;
                                    }

                                    // The following With block copied and modified [2021.03.11]
                                    // from elsewhere in this function as a temporary measure
                                    // to address a stopping situation later in the function.
                                    // See comment below for details.
                                    // 
                                    {
                                        var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " +
                                                                              "where Item='" + pnStock + "';");
                                        if (withBlock1.BOF | withBlock1.EOF)
                                        {
                                            if (pnStock != "0")
                                                Debugger.Break(); // because Material value likely invalid
                                            // REV[2022.02.08.1413]
                                            // reinstated interruption here
                                            // because at this point, pnStock
                                            // has likely already been assigned
                                            // to prRawMatl, so changing it here
                                            // is NOT likely to be productive.
                                            // this section will likely need
                                            // reconsideration, revision,
                                            // and/or possibly removal.
                                            // UPDATE[2021.12.10]:
                                            // added this check for OFFTHK
                                            // to avoid blindly adding sheet
                                            // metal stock to a "sheet metal"
                                            // part that isn't actually meant
                                            // to be made of sheet metal.
                                            if (rt.Exists("OFFTHK"))
                                                // actual Sheet Metal, so just clear this:
                                                pnStock = "";
                                            else
                                            {
                                                pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                                Debug.Print(""); // Breakpoint Landing
                                            }
                                        }
                                    }
                                }
                                else if (rt.Exists("OFFTHK"))
                                    // UPDATE[2021.12.10]:
                                    // another OFFTHK check added to avoid
                                    // adding sheet metal stock by mistake.
                                    pnStock = "";
                                else pnStock = ptNumShtMetal(withBlock.ComponentDefinition);

                                if (Strings.Len(pnStock) == 0)
                                {
                                    // UPDATE[2018.05.30]:
                                    // Pulling ALL code/text from this section
                                    // to get rid of excessive cruft.
                                    // 
                                    // In fact, reversing logic to go directly
                                    // to User Prompt if no stock identified
                                    // 
                                    // IN DOUBLE FACT, hauling this WHOLE MESS
                                    // RIGHT UP after initial pnStock assignment
                                    // to prompt user IMMEDIATELY if no stock found
                                    {
                                        var withBlock1 = newFmTest1();
                                        if (invDoc.ComponentDefinition.Document != invDoc) Debugger.Break();
                                        bd = nuAiBoxData().UsingInches
                                            .SortingDims(invDoc.ComponentDefinition.RangeBox);
                                        ck = withBlock1.AskAbout(invDoc,
                                            "No Stock Found! Please Review" + Constants.vbCrLf + Constants.vbCrLf +
                                            bd.Dump(0));
                                        if (ck == Constants.vbYes)
                                        {
                                            // UPDATE[2018.05.30]:
                                            // Pulling some extraneous commented code
                                            // from here and beginning of block
                                            {
                                                var withBlock2 = withBlock1.ItemData;
                                                if (withBlock2.Exists(pnFamily))
                                                {
                                                    nmFamily = withBlock2.get_Item(pnFamily);
                                                    Debug.Print(pnFamily + "=" + nmFamily);
                                                }

                                                if (withBlock2.Exists(pnRawMaterial))
                                                {
                                                    pnStock = withBlock2.get_Item(pnRawMaterial);
                                                    Debug.Print(pnRawMaterial + "=" + pnStock);
                                                }
                                            }
                                            if (0) Debugger.Break(); // Use this for a debugging shim
                                        }
                                    }
                                }
                                else if (Left(pnStock, 2) == "LG")
                                {
                                    Debug.Print(pnModel + ": PROBABLE LAGGING [" + pnStock + "]");
                                    Debug.Print(" TRY TO VERIFY. IF CHANGE REQUIRED,");
                                    Debug.Print(" FILL IN NEW VALUE FOR pnStock BELOW, ");
                                    Debug.Print(" AND PRESS ENTER ON THE LINE. WHEN ");
                                    Debug.Print(" READY, PRESS [F5] TO CONTINUE.");
                                    Debug.Print(" pnStock = \"" + pnStock + "\"");
                                    Debugger.Break();
                                }

                                if (Strings.Len(pnStock) > 0)
                                {
                                    // do we look for a Raw Material Family!
                                    {
                                        var withBlock1 = cnGnsDoyle().Execute(
                                            "select Family, Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16 " +
                                            "from vgMfiItems " + "where Item='" + pnStock + "';");
                                        if (withBlock1.BOF | withBlock1.EOF)
                                            Debugger.Break(); // because Material value likely invalid
                                        else
                                        {
                                            {
                                                var withBlock2 = withBlock1.Fields;
                                                mtFamily = withBlock2.get_Item("Family").Value;
                                            }

                                            // UPDATE[2021.06.18]:
                                            // New pre-check for Material Item
                                            // in Purchased Parts Family.
                                            // VERY basic handler simply
                                            // maps Material Family to D-BAR
                                            // to force extra processing below.
                                            // Further refinement VERY much needed!
                                            if (mtFamily)
                                            {
                                                // Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                                Debug.Print(pnModel + "[" + prRmQty.Value + qtUnit + "*" + pnStock +
                                                            ": " + aiPropsDesign(pnDesc).Value + "]");
                                                Debugger.Break(); // FULL Stop!
                                            }
                                            else
                                                switch (mtFamily)
                                                {
                                                    case "D-PTS":
                                                        nmFamily = "D-RMT";
                                                        Debugger.Break(); // NOT SO FAST!
                                                        mtFamily = "D-BAR";
                                                        break;
                                                    case "R-PTS":
                                                        nmFamily = "R-RMT";
                                                        Debugger.Break(); // NOT SO FAST!
                                                        mtFamily = "D-BAR";
                                                        break;
                                                }

                                            switch (mtFamily)
                                            {
                                                case "DSHEET":
                                                    // We should be okay. This is sheet metal stock
                                                    nmFamily = "D-RMT";
                                                    qtUnit = "FT2";
                                                    break;
                                                case "D-BAR":
                                                {
                                                    // UPDATE[2021.06.18]:
                                                    // Added check for Part Family already set
                                                    // to more properly handle new situation (above)
                                                    if (Strings.Len(nmFamily) == 0) nmFamily = "R-RMT";
                                                    else Debug.Print(""); // Breakpoint Landing

                                                    // UPDATE[2022.01.11]:
                                                    // Adding Do..Loop Until to following section
                                                    // to allow user to retry setting material
                                                    // quantity and units. This change made in
                                                    // conjunction with new prompt form (below).
                                                    // NOTE! This is FIRST instance of revision
                                                    // Search on UPDATE text above to locate
                                                    // the other in this function
                                                    qtUnit = prRmUnit.Value; // "IN"
                                                    ck = Constants.vbCancel;
                                                    do
                                                    {
                                                        // 'may want function here
                                                        // UPDATE[2018.05.30]: As noted above
                                                        // Will keep Stop for now
                                                        // pending further review,
                                                        // hopefully soon
                                                        // Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                                        // Debug.Print CDbl(dcIn.get_Item(pnRmQty))
                                                        // UPDATE[2021.03.11]: Replaced
                                                        // aiPropsDesign.get_Item(pnPartNum)
                                                        // with prPartNum (and now pnModel)
                                                        // since it's used in several places

                                                        // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                                        // Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                                        // Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."

                                                        // REV[2022.02.08.1511]
                                                        // replaced boilerplate above with new version below
                                                        // in hopes of better presenting change options
                                                        // in a more compact and accessible form.
                                                        Debug.Print(
                                                            "===== CHECK AND VERIFY RAW MATERIAL QUANTITY =====");
                                                        Debug.Print(" If change required, place new values at end");
                                                        Debug.Print(" of lines below for prRmQty.Value and qtUnit.");
                                                        Debug.Print(" Press [ENTER] on each line to be changed.");
                                                        Debug.Print(" Press [F5] when ready to continue.");
                                                        Debug.Print("----- " + pnModel + " [" + prRawMatl.Value +
                                                                    "]: " + aiPropsDesign(pnDesc).Value + " -----");
                                                        // Debug.Print ""

                                                        // REV[2022.02.09.0923]
                                                        // replication of REV[2022.02.09.0919]
                                                        // from section below: prep to replace
                                                        // old dimension dump operation with more
                                                        // compact call to aiBoxData's Dump method
                                                        if (true)
                                                        {
                                                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                                            {
                                                                var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                                Debug.PrintRound(
                                                                    (withBlock2.MaxPoint.X - withBlock2.MinPoint.X) /
                                                                    (double)cvLenIn2cm, 4);
                                                            }
                                                        }

                                                        {
                                                            var withBlock2 = nuAiBoxData().UsingInches()
                                                                .UsingBox(invDoc.ComponentDefinition.RangeBox);
                                                            Debug.Print.Dump(0);
                                                        }
                                                        // Stop 'and check output against prior version

                                                        // REV[2022.02.08.1446]
                                                        // removed block of Debug.Print lines
                                                        // disabled now for some time, as they
                                                        // do not seem to have been missed.
                                                        Debug.Print("prRmQty.Value = ");
                                                        if (dcIn.Exists(pnRmQty)) Debug.Print("In Genius: ");
                                                        Debug.Print("");
                                                        Debug.Print("qtUnit = \"");
                                                        if (dcIn.Exists(pnRmUnit)) Debug.Print("In Genius: ");
                                                        if (dcIn.get_Item(pnRmUnit) != "IN")
                                                            Debug.Print(" ( or try IN )");
                                                        Debug.Print("");
                                                        // Debug.Print "qtUnit = ""IN"""
                                                        // Debug.Print ""
                                                        // Debug.Print ""
                                                        // Debug.Print ""
                                                        Debugger.Break(); // Actually, we might NOT need to stop here
                                                        // if bar stock is already selected,
                                                        // because quantities would presumably
                                                        // have been established already.
                                                        // Any D-BAR handler probably needs
                                                        // to be implemented in prior section(s)
                                                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                        ck = newFmTest2().AskAbout(invDoc,
                                                            "Raw Material Quantity is now " +
                                                            Convert.ToHexString(prRmQty.Value) + qtUnit + " for",
                                                            "If this is okay, click [YES]. Otherwise," +
                                                            Constants.vbCrLf + "click [NO] or [CANCEL] to fix.");
                                                    } while (!ck == Constants.vbYes)/ because we might
                                                        want a D - BAR handler;
                                                    // UPDATE[2022.01.11]:
                                                    // This is the terminal end of the
                                                    // Do..Loop Until block noted above
                                                    rt = dcAddProp(prRmQty, rt);
                                                    Debug.Print(""); // Landing line for debugging. Do not disable.
                                                    break;
                                                }
                                                default:
                                                    nmFamily = "";
                                                    qtUnit = ""; // may want function here
                                                    // UPDATE[2018.05.30]: As noted above
                                                    // However, might need more handling here.
                                                    Debugger.Break(); // because we don't know WHAT to do with it
                                                    break;
                                            }
                                        }
                                    }
                                }
                                else if (0) Debugger.Break(); // and regroup
                            }

                            {
                                if (Len(Trim(prRawMatl.Value)) > 0)
                                {
                                    if (pnStock != prRawMatl.Value)
                                    {
                                        // Debug.Print "Raw Stock Selection"
                                        // Debug.Print " Current : " & prRawMatl.Value
                                        // Debug.Print " Proposed: " & pnStock
                                        // Stop 'because we might not want to change existing stock setting
                                        // if
                                        ck = MessageBox.Show(Join(new[]
                                        {
                                            "Raw Stock Change Suggested", " for Item " + pnModel, "",
                                            " Current : " + prRawMatl.Value, " Proposed: " + pnStock, "", "Change It?",
                                            ""
                                        }, Constants.vbCrLf), Constants.vbYesNo, pnModel + " Stock");
                                        // "Change Raw Material?"
                                        // "Suggested Sheet Metal"
                                        if (ck == Constants.vbYes) prRawMatl.Value = pnStock;
                                    }
                                }
                                else prRawMatl.Value = pnStock;
                            }
                            rt = dcAddProp(prRawMatl, rt);
                            {
                                if (Len(prRmUnit.Value) > 0)
                                {
                                    if (Strings.Len(qtUnit) > 0)
                                    {
                                        if (prRmUnit.Value != qtUnit)
                                        {
                                            // Stop 'and check both so we DON'T
                                            // automatically "fix" the RMUNIT value
                                            ck = newFmTest2().AskAbout(invDoc,
                                                null /* Conversion error: Set to default value for this argument */,
                                                "Raw Material " + prRawMatl.Value);
                                            if (ck == Constants.vbCancel) Debugger.Break();
                                            else if (ck == Constants.vbYes) prRmUnit.Value = qtUnit;
                                            if (0) Debugger.Break(); // Ctrl-9 here to skip changing
                                        }
                                    }
                                }
                                else prRmUnit.Value = qtUnit;
                            }
                            rt = dcAddProp(prRmUnit, rt);
                            // rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                            // Plan to remove commented line above,
                            // superceded by the one above that
                            Debug.Print(""); // Another landing line
                        }
                        else
                        {
                            // --------------------------------------------'
                            // [2018.07.31 by AT]
                            // Duped following block from above
                            // to mod for material assignment
                            // to non-sheet metal part.
                            // 
                            // Except, this isn't enough.
                            // Also need the code to add
                            // Stock PN to Attribute RM.
                            // That's a whole 'nother
                            // block of code, and likely
                            // best consolidated.
                            {
                                var withBlock1 = newFmTest1();
                                if (invDoc.ComponentDefinition.Document != invDoc) Debugger.Break();

                                // [2018.07.31 by AT]
                                // Added the following to try to
                                // preselect non-sheet metal stock
                                // .dbFamily.Value = "D-BAR"
                                // .lbxFamily.Value = "D-BAR"
                                // Doesn't quite do it.
                                // With New aiBoxData
                                // bd = nuAiBoxData().UsingInches.UsingBox(invDoc.ComponentDefinition.RangeBox)
                                bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                                // End With
                                ck = withBlock1.AskAbout(invDoc,
                                    "Please Select Stock for Machined Part" + Constants.vbCrLf + Constants.vbCrLf +
                                    bd.Dump(0));
                                if (ck == Constants.vbYes)
                                {
                                    // UPDATE[2018.05.30]:
                                    // Pulling some extraneous commented code
                                    // from here and beginning of block
                                    {
                                        var withBlock2 = withBlock1.ItemData;
                                        if (withBlock2.Exists(pnFamily))
                                        {
                                            nmFamily = withBlock2.get_Item(pnFamily);
                                            Debug.Print(pnFamily + "=" + nmFamily);
                                        }

                                        if (withBlock2.Exists(pnRawMaterial))
                                        {
                                            pnStock = withBlock2.get_Item(pnRawMaterial);
                                            Debug.Print(pnRawMaterial + "=" + pnStock);
                                        }
                                    }
                                    if (0) Debugger.Break(); // Use this for a debugging shim
                                }
                            }
                            // 
                            // 
                            // 
                            // The following If block is copied
                            // wholesale from sheet metal section above.
                            // Some changes (to be) made to accommodate
                            // machined or other non-sheet metal stock.
                            // 
                            // Ultimately, whole mess to require refactor.
                            // 
                            if (Strings.Len(pnStock) > 0)
                            {
                                // do we look for a Raw Material Family!

                                // This enclosing With block should NOT be necessary

                                // since the newFmTest1 above takes care of collecting

                                // the Stock Family along with the Stock itself
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " +
                                                                          "where Item='" + pnStock + "';");
                                    if (withBlock1.BOF | withBlock1.EOF)
                                        Debugger.Break(); // because Material value likely invalid
                                    else
                                    {
                                        var withBlock2 = withBlock1.Fields;
                                        mtFamily = withBlock2.get_Item("Family").Value;
                                    }
                                }
                                // These closing statements moved up from below following If block

                                // 

                                // mtFamily = nmFamily 'to force "correct" behavior of following section
                                if (mtFamily == "DSHEET")
                                {
                                    Debugger.Break(); // because we should NOT be doing Sheet Metal in this section.
                                    // This might require further investigation and/or development, if encountered.
                                    // We should be okay. This is sheet metal stock
                                    nmFamily = "D-RMT";
                                    qtUnit = "FT2";
                                }
                                else if (mtFamily == "D-BAR")
                                {
                                    // UPDATE[2022.01.11]:
                                    // Adding Do..Loop Until to following section
                                    // to allow user to retry setting material
                                    // quantity and units. This change made in
                                    // conjunction with new prompt form (below).
                                    // NOTE! This is SECOND instance of revision
                                    // Search on UPDATE text above to locate
                                    // the other in this function
                                    nmFamily = "R-RMT";
                                    qtUnit = prRmUnit.Value; // "IN"
                                    ck = Constants.vbCancel;
                                    do
                                    {
                                        // Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                                        // UPDATE[2021.03.11]: Replaced
                                        // aiPropsDesign.get_Item(pnPartNum)
                                        // as noted above
                                        // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                        // Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                        // Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."

                                        // REV[2022.02.08.1521]
                                        // replaced boilerplate above with new version below
                                        // as per REV[2022.02.08.1511]
                                        Debug.Print("===== CHECK AND VERIFY RAW MATERIAL QUANTITY =====");
                                        Debug.Print(" If change required, place new values at end");
                                        Debug.Print(" of lines below for prRmQty.Value and qtUnit.");
                                        Debug.Print(" Press [ENTER] on each line to be changed.");
                                        Debug.Print(" Press [F5] when ready to continue.");
                                        Debug.Print("----- " + pnModel + " [" + prRawMatl.Value + "]: " +
                                                    aiPropsDesign(pnDesc).Value + " -----");
                                        // Debug.Print ""

                                        // REV[2022.02.09.0919]
                                        // prep to replace old dimension dump
                                        // operation with more compact call
                                        // to aiBoxData's Dump method
                                        if (true)
                                        {
                                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                            // REV[2022.02.09.0904]
                                            // replicated With block from other section
                                            // to replace original "sprawled out" version
                                            // of Print statement hastily generated
                                            // during run time.
                                            {
                                                var withBlock1 = invDoc.ComponentDefinition.RangeBox;
                                                Debug.PrintRound(
                                                    (withBlock1.MaxPoint.X - withBlock1.MinPoint.X) /
                                                    (double)cvLenIn2cm, 4);
                                            }
                                        }

                                        {
                                            var withBlock1 = nuAiBoxData().UsingInches()
                                                .UsingBox(invDoc.ComponentDefinition.RangeBox);
                                            Debug.Print.Dump(0);
                                        }
                                        // Stop 'and check output against prior version

                                        // REV[2022.02.08.1446]
                                        // removed block of Debug.Print lines
                                        // disabled now for some time, as they
                                        // do not seem to have been missed.
                                        Debug.Print("prRmQty.Value = ");
                                        if (dcIn.Exists(pnRmQty)) Debug.Print("In Genius: ");
                                        Debug.Print("");
                                        Debug.Print("qtUnit = \"");
                                        if (dcIn.Exists(pnRmUnit)) Debug.Print("In Genius: ");
                                        Debug.Print(" ( or try IN )");

                                        // REV[2022.02.08.1525]
                                        // replaced boilerplate below with new version
                                        // above in like manner to REV[2022.02.08.1446]
                                        // and also per REV[2022.02.08.1511]

                                        // Debug.Print "qtUnit = ""IN"""
                                        // Debug.Print ""
                                        // Debug.Print ""
                                        // Debug.Print ""
                                        // Debug.Print ""
                                        // Debug.Print "PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED."
                                        // Debug.Print "PRESS ENTER/RETURN TWICE. THEN CONTINUE."
                                        // Debug.Print ""
                                        // Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                                        // Debug.Print "qtUnit = ""IN"""
                                        Debug.Print("");
                                        Debugger.Break(); // Actually, we might NOT need to stop here
                                        // if bar stock is already selected,
                                        // because quantities would presumably
                                        // have been established already.
                                        // Any D-BAR handler probably needs
                                        // to be implemented in prior section(s)
                                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                        ck = newFmTest2().AskAbout(invDoc,
                                            "Raw Material Quantity is now " + Convert.ToHexString(prRmQty.Value) +
                                            qtUnit + " for",
                                            "If this is okay, click [YES]. Otherwise," + Constants.vbCrLf +
                                            "click [NO] or [CANCEL] to fix.");
                                    } while (!ck == Constants.vbYes)/ because we might want a D - BAR handler;
                                    // UPDATE[2022.01.11]:
                                    // This is the terminal end of the
                                    // Do..Loop Until block noted above
                                    rt = dcAddProp(prRmQty, rt);
                                    Debug.Print(""); // Landing line for debugging. Do not disable.
                                }
                                else
                                {
                                    nmFamily = "";
                                    qtUnit = ""; // may want function here
                                    // UPDATE[2018.05.30]: As noted above
                                    // However, might need more handling here.
                                    Debugger.Break(); // because we don't know WHAT to do with it
                                }
                            }
                            else if (0) Debugger.Break(); // and regroup

                            // NOTE[2022.01.07.1004]:
                            // Another check for null recommendation.
                            // (SEE NOTE[2022.01.05] elsewhere in this function)
                            // Again, don't want user accidentally
                            // clearing an existing part number.
                            if (Strings.Len(pnStock) > 0)
                            {
                                {
                                    if (Len(Trim(prRawMatl.Value)) > 0)
                                    {
                                        if (pnStock != prRawMatl.Value)
                                        {
                                            // Debug.Print "Raw Stock Selection"
                                            // Debug.Print " Current : " & prRawMatl.Value
                                            // Debug.Print " Proposed: " & pnStock
                                            // Stop 'because we might not want to change existing stock setting
                                            // if
                                            ck = MessageBox.Show(Join(new[]
                                                {
                                                    "Raw Stock Change Suggested", " Current : " + prRawMatl.Value,
                                                    " Proposed: " + pnStock, "", "Change It?", ""
                                                }, Constants.vbCrLf),
                                                Constants.vbYesNo, "Change Raw Material?");
                                            // "Suggested Sheet Metal"
                                            if (ck == Constants.vbCancel) Debugger.Break();
                                            else if (ck == Constants
                                                         .vbYes) prRawMatl.Value = pnStock;
                                        }
                                    }
                                    else prRawMatl.Value = pnStock;
                                }
                            }

                            rt = dcAddProp(prRawMatl, rt);
                            {
                                if (Len(prRmUnit.Value) > 0)
                                {
                                    if (Strings.Len(qtUnit) > 0)
                                    {
                                        if (prRmUnit.Value != qtUnit)
                                        {
                                            // Stop 'and check both so we DON'T
                                            // automatically "fix" the RMUNIT value
                                            ck = newFmTest2().AskAbout(invDoc,
                                                null /* Conversion error: Set to default value for this argument */,
                                                "Raw Material " + prRawMatl.Value);
                                            if (ck == Constants.vbCancel) Debugger.Break();
                                            else if (ck == Constants.vbYes) prRmUnit.Value = qtUnit;
                                            if (0) Debugger.Break(); // Ctrl-9 here to skip changing
                                        }
                                    }
                                }
                                else prRmUnit.Value = qtUnit;
                            }
                            rt = dcAddProp(prRmUnit, rt);
                        } // Sheetmetal vs Part

                        break;
                    }
                    case kPurchasedBOMStructure:
                    {
                        // As mentioned above, nmFamily
                        // SHOULD be set at this point
                        if (Strings.Len(nmFamily) == 0)
                        {
                            if (1) Debugger.Break(); // because we might
                            // need to check out the situation
                            nmFamily = "D-PTS"; // by default
                        }

                        break;
                    }
                    case kPhantomBOMStructure:
                    {
                        // REV[2022.01.17.1135]
                        // Adding a crude handler for Phantom
                        // Part Documents. Since they shouldn't
                        // have subcomponents to promote, they
                        // shouldn't have that BOM structure.
                        // User intervention might be required.
                        ck = newFmTest2().AskAbout(invDoc, "For some reason, THIS Item is marked Phantom:",
                            "Is this okay? (Click [NO] OR [CANCEL] if not)");
                        if (ck == Constants.vbYes)
                        {
                        }
                        else Debugger.Break();

                        break;
                    }
                    case kDefaultBOMStructure:
                    case kReferenceBOMStructure:
                    case kInseparableBOMStructure:
                    case kVariesBOMStructure:
                    default:
                    {
                        // REV[2022.01.17.1138]
                        // Adding another handler to catch Part
                        // Documents with an unexpected BOM Structure. Since they shouldn't
                        // have subcomponents to promote, they
                        // shouldn't have that BOM structure.
                        // User intervention might be required.
                        ck = newFmTest2().AskAbout(invDoc, "The following Item has an unhandled BOM Structure:",
                            "Skip it? (Click [NO] OR [CANCEL] to review)");
                        if (ck == Constants.vbYes)
                        {
                        }
                        else Debugger.Break(); // and let User decide what to do with it.

                        Debugger.Break(); // (extraneous; disable/remove whenever)
                        break;
                    }
                }

                // the design tracking property set,
                // and update the Cost Center Property
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
                else if (Strings.Len(nmFamily) > 0)
                {
                    prFamily.Value = nmFamily;
                    if (Information.Err().Number)
                    {
                        Debug.Print("CHGFAIL[FAMILY]{'" + prFamily.Value + "' -> '" + nmFamily + "'}: " +
                                    invDoc.DisplayName + " (" + invDoc.FullDocumentName + ")");
                        if (MessageBox.Show("Couldn't Change Family" + vbCrLf,
                                Constants.vbYesNo | Constants.vbDefaultButton2, invDoc.DisplayName) ==
                            Constants.vbYes) Debugger.Break();
                    }

                    rt = dcAddProp(prFamily, rt);
                }
            }
            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            return rt;
        }
    }
    break;
}

}