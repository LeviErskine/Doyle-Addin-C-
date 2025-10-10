class dvl4
{
    public Scripting.Dictionary d4g2f1(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// d4g2f1 -- not sure what
        /// but looks like something to do
        /// with grouping items by family
        /// 
        Scripting.Dictionary rt;

        if (AiDoc == null)
            rt = d4g2f1(aiDocActive());
        else
        {
            var withBlock = dcAiDocCompSetsByPtNum(AiDoc);
            // Stop
            if (withBlock.Exists(1))
                rt = dcFrom2Fields(cnGnsDoyle().Execute("select ls.it, ISNULL(i.Family, '') fm from " + sqlValuesFromDc(dcOb(withBlock.Item(1)), "ls", "it") + " left join vgMfiItems i on ls.it = i.Item"), "it", "fm");
            else
            {
                rt = new Scripting.Dictionary();
                System.Diagnostics.Debugger.Break();
            }
        }

        d4g2f1 = rt;
    }

    public Scripting.Dictionary d4g0f1(Inventor.Document AiDoc, long incTop = 0)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        // Dim gn As Scripting.Dictionary
        Variant ky;
        Inventor.Document pt; // .PartDocument

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcOfDcOfDcByPlurality(dcAiDocSetsByPtNum(dcAiDocComponents(AiDoc, null/* Conversion error: Set to default value for this argument */, incTop)));
            if (withBlock.Exists(2))
                // ky = TypeName(                userChoiceFromDc(                dcNewIfNone(                dcOb(.Item(2))            )))
                ky = nuSelFromDict(dcNewIfNone(dcOb(withBlock.Item(2)))).GetReply();

            if (withBlock.Exists(""))
                System.Diagnostics.Debugger.Break();

            System.Diagnostics.Debugger.Break();
            // Debug.Print txDumpLs(dcNewIfNone(dcOb(obOf(.Item(1)))).Keys)
            gp = dcNewIfNone(dcOb(obOf(withBlock.Item(1))));
            {
                var withBlock1 = dcNewIfNone(dcOb(obOf(withBlock.Item(1)))) // dcAiDocComponents(AiDoc, , incTop)
       ;
                foreach (var ky in withBlock1.Keys)
                {
                    pt = aiDocument(withBlock1.Item(ky)); // aiDocPart()
                    if (pt == null)
                    {
                    }
                    else
                        rt.Add(ky, dcGnsPtProps_Rev20220830_inProg(pt));
                }
            }
        }

        // gn = dcDxFromRecSetDc(dcFromAdoRS('    cnGnsDoyle().Execute(q1g1x2(AiDoc)')))
        // With gn
        // End With


        d4g0f1 = rt;
    }

    public Scripting.Dictionary dcOfGnsProps(Inventor.Document invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcOfGnsProps
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcPr;
        Variant ky;

        rt = dcNewIfNone(dc);

        dcPr = dcOfPropsInAiDoc(invDoc);
        {
            var withBlock = dcPr;
            foreach (var ky in Array(pnPartNum, pnFamily, pnDesc, pnMaterial, pnStockNum, pnCatWebLink, pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnThickness, pnLength, pnWidth, pnArea))
            {
                if (withBlock.Exists(ky))
                    rt.Add(ky, withBlock.Item(ky));
                else
                    rt.Add(ky, null/* TODO Change to default(_) if this is not a reference type */);
            }

            /// NOTE[2022.09.16.1024]
            /// extraction of Categories XML text
            /// expected to move to content center
            /// processing in dcGnsValFromContentCtr
            // If .Exists("Categories") Then
            if (Len(obAiProp(withBlock.Item("Categories")).Value) > 0)
            {
                rt.Add("Categories", withBlock.Item("Categories"));
                // rt.Add "Parameters", dcAiDocParVals(invDoc)
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing: Content Center
            }
            else
            {
            }
        }

        dcOfGnsProps = rt;
    }

    public Scripting.Dictionary dcGnsValFromContentCtr(Inventor.PartComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsValFromContentCtr
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Variant ky;

        string catXml;
        Inventor.Parameter pr;

        rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            {
                var withBlock = CpDef;
                {
                    var withBlock1 = aiDocPart(withBlock.Document).PropertySets;
                    catXml = withBlock1.Item(gnDesign).Item("Categories").Value;

                    wk = dcAiPropsInSet(withBlock1.Item(guidPrSetCLib));
                }
            }

            {
                var withBlock = wk;
                foreach (var ky in withBlock.Keys) // Array("Member FileName", "Family","Standard", "Size Designation","Categories")
                                                   // If .Exists(ky) Then
                    rt.Add(ky, obAiProp(withBlock.Item(ky)).Value);
            }
        }

        dcGnsValFromContentCtr = rt;
    }

    public Scripting.Dictionary dcGnsValFromPartCompDef(Inventor.PartComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsValFromPartCompDef
        /// 
        Scripting.Dictionary rt;

        rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            {
                var withBlock = CpDef;
                // part of general ComponentDefinition
                // rt.Add "bomStruct", .BOMStructure

                if (withBlock.IsContentMember)
                {
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing: Content Member
                                                                                 // Stop 'and look at this one
                                                                                 // rt.Add "ContentMem", 1
                    rt.Add("ccPropVals", dcGnsValFromContentCtr(CpDef));
                }

                {
                    var withBlock1 = withBlock.MassProperties;
                    rt.Add(pnMass, Round(cvMassKg2LbM * withBlock1.Mass, 4));
                }

                if (withBlock.IsiPartMember)
                {
                    rt.Add("isIPartMem", 1);
                    rt.Add("iPartFactory", withBlock.iPartMember.ReferencedDocumentDescriptor.FullDocumentName);
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
                    var withBlock1 = withBlock.Parameters;
                }
            }

            {
                var withBlock = dcFlatPatVals(aiCompDefShtMetal(CpDef), dcDotted());
                if (withBlock.Count > 2)
                    rt.Add("flatPat", dcUnDotted(withBlock.Item(".")));
                else
                {
                }
            }
        }

        dcGnsValFromPartCompDef = rt;
    }

    public Scripting.Dictionary dcGnsValFromAssyCompDef(Inventor.AssemblyComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsValFromAssyCompDef
        /// 
        Scripting.Dictionary rt;

        rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            {
                var withBlock = CpDef;
                // part of general ComponentDefinition
                // rt.Add "bomStruct", .BOMStructure

                {
                    var withBlock1 = withBlock.MassProperties;
                    rt.Add(pnMass, Round(cvMassKg2LbM * withBlock1.Mass, 4));
                }

                if (withBlock.IsiAssemblyMember)
                {
                    rt.Add("isIAssyMem", 1);
                    rt.Add("iAssyFactory", withBlock.iAssemblyMember.ReferencedDocumentDescriptor.FullDocumentName);
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
                    var withBlock1 = withBlock.Parameters;
                }
            }

            {
                var withBlock = dcFlatPatVals(aiCompDefShtMetal(CpDef), dcDotted());
                if (withBlock.Count > 2)
                    rt.Add("flatPat", dcUnDotted(withBlock.Item(".")));
                else
                {
                }
            }
        }

        dcGnsValFromAssyCompDef = rt;
    }

    public Scripting.Dictionary dcGnsValFromGenCompDef(Inventor.ComponentDefinition CpDef, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsValFromGenCompDef
        /// 
        Scripting.Dictionary rt;

        rt = dcNewIfNone(dc);

        if (CpDef == null)
        {
        }
        else
        {
            var withBlock = CpDef;
            rt.Add("bomStruct", withBlock.BOMStructure);
            {
                var withBlock1 = nuAiBoxData().UsingBox(withBlock.RangeBox).UsingInches();
                rt.Add("dimsModel", withBlock1.Dictionary());
            }
        }

        dcGnsValFromGenCompDef = dcGnsValFromAssyCompDef(aiCompDefAssy(CpDef), dcGnsValFromPartCompDef(aiCompDefPart(CpDef), rt));
    }

    public Scripting.Dictionary dcGnsValGeneral(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsValGeneral
        /// 
        Scripting.Dictionary rt;
        // Dim dcPr As Scripting.Dictionary
        // Dim ky As Variant

        rt = dcNewIfNone(dc);

        if (AiDoc == null)
        {
        }
        else
        {
            {
                var withBlock = AiDoc;
                {
                    var withBlock1 = withBlock.PropertySets.Item(gnDesign);
                    rt.Add(pnPartNum, withBlock1.Item(pnPartNum).Value);
                }

                rt.Add("subType", withBlock.SubType);
            }

            rt = dcGnsValFromGenCompDef(aiCompDefOf(AiDoc), rt);
        }

        // dcPr = dcOfPropsInAiDoc(AiDoc)
        // With dcPr
        // For Each ky In Array('        pnPartNum, pnFamily, pnDesc,'        pnMaterial, pnStockNum, pnCatWebLink,'        pnMass, pnRawMaterial, pnRmQty, pnRmUnit,'        pnThickness, pnLength, pnWidth, pnArea'    )
        // If .Exists(ky) Then
        // rt.Add ky, .Item(ky)
        // Else
        // rt.Add ky, Nothing
        // End If
        // Next
        // End With

        dcGnsValGeneral = rt;
    }

    public Scripting.Dictionary dcGnsPtProps_Rev20220830_inProg(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */) // .PartDocument
    {
        /// dcGnsPtProps_Rev20220830_inProg
        /// 
        Scripting.Dictionary rt;
        // Dim dcPr As Scripting.Dictionary
        // Dim dcVl As Scripting.Dictionary
        // Dim ky As Variant

        // rt = dcNewIfNone(dc)
        rt = dcGnsValGeneral(AiDoc, dcNewIfNone(dc));

        {
            var withBlock = dcOfGnsProps(AiDoc, dcDotted());
            if (withBlock.Count > 2)
            {
                rt.Add("props", dcUnDotted(withBlock.Item(".")));
                rt.Add("propVals", dcPropVals(rt.Item("props")));
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
            var withBlock = AiDoc;
            {
                var withBlock1 = withBlock.PropertySets.Item(gnDesign);
            }
        } // AiDoc Is Nothing

        // Stop
        // Call iSyncPartFactory(AiDoc)
        // rt = dcVl

        dcGnsPtProps_Rev20220830_inProg = rt;
    }

    public string sqlValuesFromDc(Scripting.Dictionary dc, string vw = "ls", string it = "it")
    {
        /// sqlValuesFromDc - generate SQL
        /// "VALUES" clause from Keys
        /// of supplied Dictionary.
        /// result is a relation
        /// of one attribute
        /// '
        /// VALUES clause must end
        /// with an AS phrase naming
        /// the relation and all
        /// attributes.
        /// '
        /// in this function, the
        /// names default to "ls"
        /// (list) for the relation,
        /// and "it" (item) for
        /// its one attribute
        /// 
        if (dc == null)
            sqlValuesFromDc = sqlValuesFromDc(new Scripting.Dictionary());
        else
            sqlValuesFromDc = "(values ('" + Join(dc.Keys, "'), ('") + "')) as ls(it)";
    }

    public Scripting.Dictionary dcAiDocCompSetsByPtNum(Inventor.Document AiDoc, long incTop = 0)
    {
        /// dcAiDocCompSetsByPtNum -- formerly d4g3f1
        /// 
        /// 
        /// 
        Scripting.Dictionary dc;
        // Dim ct As Long

        // ct = 1 'to include main assembly (for now)
        // now disabled in favor of input parameter incTop

        // dcAiDocSetsByPtNum replaces dcRemapByPtNum
        dc = dcOfDcOfDcByPlurality(dcAiDocSetsByPtNum(dcAiDocComponents(AiDoc, null/* Conversion error: Set to default value for this argument */, incTop))); // incTop replaces ct
                                                                                                                                                              // dcOfDcOfDcByItemCount() removed from original lineup
                                                                                                                                                              // of dcOfDcOfDcByPlurality(dcOfDcOfDcByItemCount(dcAiDocSetsByPtNum(
                                                                                                                                                              // dcOfDcOfDcByPlurality calls dcOfDcOfDcByItemCount internally
                                                                                                                                                              // as part of its normal processing.

        dcAiDocCompSetsByPtNum = dc;
    }

    public Scripting.Dictionary dcAiDocSetsByPtNum(Scripting.Dictionary dc)
    {
        /// dcAiDocSetsByPtNum -- formerly d4g3f2
        /// 
        /// Returns Dictionary of Dictionaries
        /// of Inventor Documents keyed on
        /// associated Part Numbers.
        /// 
        /// Derived from dcRemapByPtNum, this
        /// variation collects all models of a
        /// given Part Number into a secondary
        /// Dictionary, under each Document's
        /// file name. Ideally, Part Numbers
        /// should map one-to-one to Documents,
        /// so each sub Dictionary should
        /// contain only one entry.
        /// 
        /// However, as it IS possible for more
        /// than one model to represent the same
        /// Part, more than one Document might
        /// in fact have the same Part Number.
        /// 
        /// Therefore, it may sometimes prove
        /// necessary to take additional steps
        /// in properly identifying which model
        /// (or models) to process in preparation
        /// for Genius.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Inventor.Document pt;
        Variant ky;
        string pn;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                pn = System.Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));

                /// REV[2022.05.17.1536]
                /// removing check/handling
                /// for blank/empty part number.
                /// client process can deal with that.
                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(pn))
                    {
                    }
                    else
                        withBlock1.Add(pn, new Scripting.Dictionary());

                    gp = withBlock1.Item(pn);
                }

                {
                    var withBlock1 = gp;
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break(); // this should NOT happen
                    else
                        withBlock1.Add(ky, pt);
                }
            }
        }

        dcAiDocSetsByPtNum = rt;
    }

    public Scripting.Dictionary dcOfDcOfDcByItemCount(Scripting.Dictionary dc)
    {
        /// dcOfDcOfDcByItemCount -- formerly d4g3f3
        /// 
        /// subdivide supplied Dictionary
        /// of Dictionaries into groups
        /// by Count of members.
        /// 
        /// result is a 3rd-order Dictionary,
        /// that is, a Dictionary (1)
        /// keyed by member count
        /// of Dictionaries (2)
        /// keyed by a shared key
        /// of yet more Dictionaries (3)
        /// keyed to some unique value
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Scripting.Dictionary xp;
        Variant ky;
        long ct;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                gp = withBlock.Item(ky);
                ct = gp.Count;

                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(ct))
                        withBlock1.Add(ct, new Scripting.Dictionary());

                    xp = withBlock1.Item(ct);
                }

                {
                    var withBlock1 = xp;
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break(); // this should NOT happen
                    else
                        withBlock1.Add(ky, gp);
                }
            }
        }

        dcOfDcOfDcByItemCount = rt;
    }

    public Scripting.Dictionary dcOfDcOfDcByPlurality(Scripting.Dictionary dc)
    {
        /// dcOfDcOfDcByPlurality -- formerly d4g3f4
        /// 
        /// given a 2nd-order Dictionary (NOT 3rd)
        /// as supplied by dcAiDocSetsByPtNum (NOT dcOfDcOfDcByItemCount),
        /// return a reorganized version as follows:
        /// 
        /// under key 1: a Dictionary of all
        /// Dictionaries having only one member.
        /// this should form the bulk of the
        /// supplied Dictionary's content.
        /// 
        /// under key 2: a Dictionary of Dictionaries
        /// having more than one member. these
        /// "plurals" might require additional
        /// review and/or processing to resolve
        /// ambiguities, conflicts, etc.
        /// 
        /// under key "" (blank string): the Dictionary,
        /// if present, of members with no assigned
        /// part or item number. this should almost
        /// NEVER arise, but again, might require
        /// special processing to resolve issues.
        /// 
        /// '''
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Scripting.Dictionary xp;
        Variant ky;
        long ct;

        rt = new Scripting.Dictionary();

        /// to avoid modifying supplied Dictionary,
        /// generate a copy to work with directly.
        gp = dcCopy(dc);
        {
            var withBlock = gp;
            if (withBlock.Exists(""))
            {
                // of blank "numbered"
                // items moved over...

                rt.Add("", withBlock.Item(""));
                withBlock.Remove("");
            }
        }
        // before grouping by member counts:

        gp = dcOfDcOfDcByItemCount(gp);
        {
            var withBlock = gp;
            // prep the "singles" return Dictionary
            xp = new Scripting.Dictionary();
            rt.Add(1, xp);
            if (withBlock.Exists(1))
            {
                // the (one, single) member of each
                // Dictionary under this one
                {
                    var withBlock1 = dcOb(withBlock.Item(1));
                    foreach (var ky in withBlock1.Keys)
                    {
                        {
                            var withBlock2 = dcOb(withBlock1.Item(ky));
                            xp.Add(ky, withBlock2.Items(0));
                        }
                    }
                }
                withBlock.Remove(1);
            }
            else
            {
            }
            // at this point, any remaining members
            // should be "plural" Dictionaries
            // containing more than one member.

            // THESE are to be combined into one
            // "plural" Dictionary to be returned.
            xp = new Scripting.Dictionary();
            // DO NOT add to return Dictionary yet!

            foreach (var ky in withBlock.Keys)
                // this step generates a NEW
                // Dictionary at each iteration.
                xp = dcKeysCombined(xp, dcOb(withBlock.Item(ky)), 1);

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

        // rt.Add 2, dcKeysMissing(gp, rt.Item(1))

        // With dc: For Each ky In .Keys
        // gp = .Item(ky)
        // ct = gp.Count
        // 
        // With rt
        // If Not .Exists(ct) Then
        // .Add ct, New Scripting.Dictionary
        // End If
        // 
        // xp = .Item(ct)
        // End With
        // 
        // With xp
        // If .Exists(ky) Then
        // Stop 'this should NOT happen
        // Else
        // .Add ky, gp
        // End If: End With
        // Next: End With

        dcOfDcOfDcByPlurality = rt;
    }

    public Scripting.Dictionary d4g3f5from2(Scripting.Dictionary dc)
    {
        /// d4g3f5from2
        /// 
        /// Returns Dictionary of Dictionaries
        /// of Inventor Documents keyed on
        /// associated Part Numbers.
        /// 
        /// Derived from dcRemapByPtNum, this
        /// variation collects all models of a
        /// given Part Number into a secondary
        /// Dictionary, under each Document's
        /// file name. Ideally, Part Numbers
        /// should map one-to-one to Documents,
        /// so each sub Dictionary should
        /// contain only one entry.
        /// 
        /// However, as it IS possible for more
        /// than one model to represent the same
        /// Part, more than one Document might
        /// in fact have the same Part Number.
        /// 
        /// Therefore, it may sometimes prove
        /// necessary to take additional steps
        /// in properly identifying which model
        /// (or models) to process in preparation
        /// for Genius.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Inventor.Document pt;
        Variant ky;
        string pn;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));

                // pt.co

                pn = System.Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));

                /// REV[2022.05.17.1536]
                /// removing check/handling
                /// for blank/empty part number.
                /// client process can deal with that.
                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(pn))
                    {
                    }
                    else
                        withBlock1.Add(pn, new Scripting.Dictionary());

                    gp = withBlock1.Item(pn);
                }

                {
                    var withBlock1 = gp;
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break(); // this should NOT happen
                    else
                        withBlock1.Add(ky, pt);
                }
            }
        }

        d4g3f5from2 = rt;
    }

    public Variant d4g3f5b(Inventor.ComponentDefinition cd)
    {
    }

    public object d4g3f5a(Inventor.ComponentDefinition cd) // Inventor.iAssemblyTableCell
    {
        // aiCompDefOf

        if (cd == null)
            d4g3f5a = null;
        else if (cd is Inventor.AssemblyComponentDefinition)
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
        }
        else if (cd is Inventor.PartComponentDefinition)
        {
        }
    }

    public Scripting.Dictionary gnsUpdtAll_iFact(Inventor.ComponentDefinition cd)
    {
        if (cd is Inventor.PartComponentDefinition)
            gnsUpdtAll_iFact = gnsUpdtAll_iPart(cd);
        else if (cd is Inventor.AssemblyComponentDefinition)
            gnsUpdtAll_iFact = gnsUpdtAll_iAssy(cd);
        else
            gnsUpdtAll_iFact = new Scripting.Dictionary();
    }

    public Scripting.Dictionary gnsUpdtAll_iAssy(Inventor.AssemblyComponentDefinition cd)
    {
        Scripting.Dictionary rt;
        Inventor.AssemblyDocument md;
        Inventor.iAssemblyFactory fc;
        Inventor.iAssemblyTableRow rw;
        Inventor.iAssemblyTableRow r0;

        rt = new Scripting.Dictionary();

        if (cd == null)
        {
        }
        else if (cd.IsiAssemblyFactory)
        {
            {
                var withBlock = cd.iAssemblyFactory;
                md = withBlock.Parent.Parent;

                {
                    var withBlock1 = withBlock.TableColumns // .Item()
       ;
                }

                // note initial DefaultRow
                r0 = withBlock.DefaultRow;

                foreach (var rw in withBlock.TableRows)
                {
                    withBlock.DefaultRow = rw;
                    rt.Add.DefaultRow.MemberName(null/* Conversion error: Set to default value for this argument */, dcOfDcAiPropVals(dcGeniusProps(md)));
                }

                // restore initial DefaultRow
                withBlock.DefaultRow = r0;
            }
        }
        else
        {
        }

        gnsUpdtAll_iAssy = rt;
    }

    public Scripting.Dictionary gnsUpdtAll_iPart(Inventor.PartComponentDefinition cd)
    {
        Scripting.Dictionary rt;
        Inventor.PartDocument md;
        Inventor.iPartFactory fc;
        Inventor.iPartTableRow rw;
        Inventor.iPartTableRow r0;

        rt = new Scripting.Dictionary();

        if (cd == null)
        {
        }
        else if (cd.IsiPartFactory)
        {
            {
                var withBlock = cd.iPartFactory;
                md = withBlock.Parent; // .Parent

                {
                    var withBlock1 = withBlock.TableColumns // .Item()
       ;
                }

                // note initial DefaultRow
                r0 = withBlock.DefaultRow;

                foreach (var rw in withBlock.TableRows)
                {
                    withBlock.DefaultRow = rw;
                    rt.Add.DefaultRow.MemberName(null/* Conversion error: Set to default value for this argument */, dcAiPropValsFromDc(dcGeniusProps(md))); DoEvents();
                    md.Save();
                }

                // restore initial DefaultRow
                withBlock.DefaultRow = r0;
            }
        }
        else
        {
        }

        gnsUpdtAll_iPart = rt;
    }

    public Scripting.Dictionary d4g3f6pt(Inventor.PartComponentDefinition cd)
    {
        Scripting.Dictionary rt;
        long mx;
        long dx;
        Inventor.iPartTableColumn ck;

        rt = new Scripting.Dictionary();
        if (cd == null)
        {
        }
        else if (cd.IsiPartFactory)
        {
            System.Diagnostics.Debugger.Break();
            {
                var withBlock = cd.iPartFactory.TableColumns;
                mx = withBlock.Count;
                System.Diagnostics.Debugger.Break();
                for (dx = 1; dx <= mx; dx++)
                {
                    ck = withBlock.Item(dx);
                    {
                        var withBlock1 = ck;
                        rt.Add.Heading(null/* Conversion error: Set to default value for this argument */, nuDcPopulator().Setting("dx", dx).Setting("dh", withBlock1.DisplayHeading).Setting("fh", withBlock1.FormattedHeading).Setting("dt", withBlock1.ReferencedDataType).Setting("ob", withBlock1.ReferencedObject).Setting("ot", TypeName(withBlock1.ReferencedObject)).Dictionary);
                        // 
                        // ).Setting("hd", .Heading'
                        System.Diagnostics.Debugger.Break();
                    }
                }
            }
        }
        else if (cd.IsiPartMember)
        {
            System.Diagnostics.Debugger.Break();
            rt = d4g3f6pt(aiDocPart(cd.iPartMember.ParentFactory.Parent).ComponentDefinition);
        }
        else
        {
        }

        d4g3f6pt = rt;
    }

    public Scripting.Dictionary d4g3f6as(Inventor.AssemblyComponentDefinition cd)
    {
        Scripting.Dictionary rt;
        long mx;
        long dx;
        Inventor.iAssemblyTableColumn ck;

        rt = new Scripting.Dictionary();
        if (cd == null)
        {
        }
        else if (cd.IsiAssemblyFactory)
        {
            System.Diagnostics.Debugger.Break();
            {
                var withBlock = cd.iAssemblyFactory.TableColumns;
                mx = withBlock.Count;
                System.Diagnostics.Debugger.Break();
                for (dx = 1; dx <= mx; dx++)
                {
                    ck = withBlock.Item(dx);
                    System.Diagnostics.Debugger.Break();
                }
            }
        }
        else if (cd.IsiAssemblyMember)
            System.Diagnostics.Debugger.Break();
        else
        {
        }

        d4g3f6as = rt;
    }

    public Scripting.Dictionary d4g3f7pt(Inventor.PartComponentDefinition cd)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary hd;
        // Dim md As Inventor.PartDocument
        Inventor.iPartFactory fc;
        Inventor.iPartTableColumn co;
        Inventor.iPartTableRow rw;
        // Dim r0 As Inventor.iPartTableRow
        // Dim df As Long

        rt = new Scripting.Dictionary();
        rt.Add("", new Scripting.Dictionary());
        hd = rt.Item("");


        if (cd == null)
            fc = null/* TODO Change to default(_) if this is not a reference type */;
        else if (cd.IsiPartFactory)
            fc = cd.iPartFactory;
        else if (cd.IsiPartMember)
            fc = cd.iPartMember.ParentFactory;
        else
            fc = null/* TODO Change to default(_) if this is not a reference type */;

        if (!fc == null)
        {
            {
                var withBlock = fc;
                // md = .Parent '.Parent

                hd.Add("", Array("Index", "Key", "CustomColumn", "DisplayHeading", "FormattedHeading", "ReferencedDataType"));
                foreach (var co in withBlock.TableColumns)
                {
                    {
                        var withBlock1 = co;
                        // .Item()
                        {
                            var withBlock2 = co;
                            hd.Add.Heading(null/* Conversion error: Set to default value for this argument */, Array(withBlock2.Index, withBlock2.Key, withBlock2.CustomColumn, withBlock2.DisplayHeading, withBlock2.FormattedHeading, withBlock2.ReferencedDataType));
                        }
                    }
                }

                // note initial DefaultRow
                // r0 = .DefaultRow

                foreach (var rw in withBlock.TableRows)
                {
                    {
                        var withBlock1 = rw;
                        // df = rw Is r0
                        rt.Add.Index(null/* Conversion error: Set to default value for this argument */, Array(withBlock1.Index, withBlock1.MemberName, withBlock1.PartName, rw));
                    }
                }
            }
        }

        d4g3f7pt = rt;
    }

    public Scripting.Dictionary d4g4f0(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// d4g4f0 -- rebuilding Sub Update_Genius_Properties
        /// more or less from the ground up
        /// 
        // Dim invProgressBar As Inventor.ProgressBar

        // Dim fc As gnsIfcAiDoc
        Scripting.Dictionary dc;
        Scripting.Dictionary mt;
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        // Dim goAhead As VbMsgBoxResult
        // Dim ActiveDoc As Document
        // Dim txOut As String
        Variant ky;
        Variant k2;
        // Dim kyPt As Variant
        long ct;

        // Dim dx As Long

        // Dim fm As fmIfcTest05A

        /// NOTE[2022.06.01.1441]
        /// adding check for supplied Document
        /// call for selection if none
        if (AiDoc == null)
            // rt = d4g4f0(userChoiceFromDc(dcAiDocsVisible(), aiDocActive()))
            rt = d4g4f0(aiDocActive());
        else
        {
            /// NOTE[2022.06.01.1442]
            /// disabling/skipping user checks for now
            /// this isn't the purpose of this whole mess.
            /// Confirm User Request
            /// to process active Document
            // goAhead = MsgBox('    Join(Array('        "Are you sure you want to process this document?",'        "The process may require a few minutes depending on assembly size.",'        "Suppressed and excluded parts will not be processed."'    ), " "),'    vbYesNo + vbQuestion,'    "Process Document Custom iProperties"')
            // If goAhead = vbYes Then
            // 
            // Else
            // End If

            /// NOTE[2022.06.01.1444]
            /// see Update_Genius_Properties REV[2022.05.24.0956]
            {
                var withBlock = dcAiDocCompSetsByPtNum(AiDoc, ct) // ActiveDoc
       ;
                if (withBlock.Exists(""))
                    System.Diagnostics.Debugger.Break();// for now

                if (withBlock.Exists(2))
                {
                    // THIS situation IS known to occur,
                    // if not TERRIBLY frequently, so a
                    // handler here is a good idea.
                    // 
                    {
                        var withBlock1 = nuDcPopulator(withBlock.Item(2)) // d4g4f4(dcOb(.Item(2)))
       ;
                        Debug.Print(MsgBox(msg_2022_0603_1127(withBlock1.Dictionary), Constants.vbOKOnly | Constants.vbInformation, "Duplicate Part Numbers!"));
                    }
                }

                // and HERE is the step which ACTUALLY
                // replaces the prior version above.
                // Key 1 is guaranteed to be present
                // in the Dictionary returned, so no
                // need to check for it here.
                dc = dcOb(withBlock.Item(1));
            }
            /// 
            /// 

            /// NOTE[2022.06.01.1502]
            /// this section expected to be
            /// exported to its own function
            /// NOTE[2022.06.02.0906]
            /// (follow-up) original code
            /// extracted to functions
            /// dcOfKeys2match and d4g4f1
            mt = dcOfKeys2match(Array(pnFamily, pnMass, pnRawMaterial, pnRmQty, pnRmUnit, pnWidth, pnLength, pnArea, pnThickness));
            // pnFamily       replaces "Cost Center"
            // pnMass         replaces "GeniusMass"
            // pnRawMaterial  replaces "RM"
            // pnRmQty        replaces "RMQTY"
            // pnRmUnit       replaces "RMUNIT"
            // pnWidth        replaces "Extent_Width"
            // pnLength       replaces "Extent_Length"
            // pnArea         replaces "Extent_Area"
            // pnThickness    replaces "Thickness"

            // rt = d4g4f1(dc, mt)

            rt = new Scripting.Dictionary();
            {
                var withBlock = d4g4f1(dc, mt);
                foreach (var ky in withBlock.Keys)
                {
                    {
                        var withBlock1 = rt;
                        if (!withBlock1.Exists(ky))
                            withBlock1.Add(ky, new Scripting.Dictionary());

                        wk = withBlock1.Item(ky);
                    }

                    {
                        var withBlock1 = dcOb(withBlock.Item(ky));
                        foreach (var k2 in withBlock1.Keys)
                            wk.Add(k2, obAiProp(withBlock1.Item(k2)).Value);
                    }
                }
            }
        }
        /// 
        d4g4f0 = rt;
    }

    public Scripting.Dictionary d4g4f1(Scripting.Dictionary dc, Scripting.Dictionary rf)
    {
        /// d4g4f1 -- returns a Dictionary of Dictionaries
        /// copied from supplied Dictionary dc,
        /// but with only those Keys matching those
        /// found in supplied 'reference' Dictionary rf
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcKeysInCommon(dcOfPropsInAiDoc(aiDocument(withBlock.Item(ky))), rf, 1));
        }

        d4g4f1 = rt;
    }

    public Inventor.PartDocument d4g4f2(Scripting.Dictionary dc)
    {
        /// d4g4f2 -- given a Dictionary of Part Documents
        /// return the first Content Center Member found
        /// (if none found, return Nothing)
        /// 
        Variant ky;
        Inventor.PartDocument pt;

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                if (pt == null)
                {
                    pt = aiDocPart(aiDocument(withBlock.Item(ky)));
                    if (!pt == null)
                    {
                        if (!pt.ComponentDefinition().IsContentMember)
                            pt = null/* TODO Change to default(_) if this is not a reference type */;
                    }
                }
            }
        }

        d4g4f2 = pt;
    }

    public Scripting.Dictionary d4g4f3(Scripting.Dictionary dc)
    {
        /// d4g4f3 -- given a Dictionary of Part Document
        /// Dictionaries, return a subset containing
        /// only Content Center Members
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Inventor.PartDocument pt;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = d4g4f2(withBlock.Item(ky));
                if (!pt == null)
                    rt.Add(ky, pt);
            }
        }

        d4g4f3 = rt;
    }

    public Scripting.Dictionary d4g4f4(Scripting.Dictionary dc)
    {
        /// d4g4f4 -- given a Dictionary of Part Document
        /// Dictionaries, return a subset dropping
        /// any with Content Center Members
        /// 
        d4g4f4 = dcKeysMissing(dc, d4g4f3(dc));
    }

    public Scripting.Dictionary d4g4f5() // dc As Scripting.Dictionary'''
    {
        /// d4g4f5 -- given a Dictionary of Part Document
        /// Dictionaries, return a subset dropping
        /// any with Content Center Members
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcRb;
        Scripting.Dictionary dcTb;
        Scripting.Dictionary dcRp;
        // Dim dcRb As Scripting.Dictionary
        Inventor.Ribbon rb;
        Inventor.RibbonTab tb;
        Inventor.RibbonPanel rp;
        object ob;

        rt = new Scripting.Dictionary();
        // ThisApplication.UserInterfaceManager.RibbonState
        {
            var withBlock = ThisApplication.UserInterfaceManager;
            foreach (var rb in withBlock.Ribbons)
            {
                dcRb = new Scripting.Dictionary();
                {
                    var withBlock1 = rb;
                    rt.Add.InternalName(null/* Conversion error: Set to default value for this argument */, dcRb);
                    foreach (var tb in withBlock1.RibbonTabs) // .QuickAccessControls
                    {
                        dcTb = new Scripting.Dictionary();
                        {
                            var withBlock2 = tb;
                            dcRb.Add.InternalName(null/* Conversion error: Set to default value for this argument */, dcTb);
                            foreach (var rp in withBlock2.RibbonPanels)
                            {
                                dcRp = new Scripting.Dictionary();
                                {
                                    var withBlock3 = rp;
                                    dcTb.Add.InternalName(null/* Conversion error: Set to default value for this argument */, dcRp);
                                    System.Diagnostics.Debugger.Break();
                                }
                            }
                        }
                    }
                }
            }
        }
        d4g4f5 = rt;
    }

    public Inventor.ComponentDefinition compDefOfPart(Inventor.PartDocument AiDoc)
    {
        if (AiDoc == null)
            compDefOfPart = null/* TODO Change to default(_) if this is not a reference type */;
        else
            compDefOfPart = AiDoc.ComponentDefinition;
    }

    public Inventor.ComponentDefinition compDefOfAssy(Inventor.AssemblyDocument AiDoc)
    {
        if (AiDoc == null)
            compDefOfAssy = null/* TODO Change to default(_) if this is not a reference type */;
        else
            compDefOfAssy = AiDoc.ComponentDefinition;
    }

    public Inventor.ComponentDefinition compDefOf(Inventor.Document AiDoc)
    {
        Inventor.ComponentDefinition rt;

        rt = compDefOfPart(aiDocPart(AiDoc));
        if (rt == null)
            rt = compDefOfAssy(aiDocAssy(AiDoc));

        // AiDoc.FullFileName

        compDefOf = rt;
    }

    public string famOfAiDoc(Inventor.Document AiDoc)
    {
        /// 
        string itNum;
        string mdFam;
        string gnFam;

        string pf;
        string sf;

        VbMsgBoxResult ck;

        if (AiDoc == null)
            famOfAiDoc = "";
        else
        {
            {
                var withBlock = AiDoc;
                /// NOTE!!! ONLY use this for Assemblies!
                /// will disable until better set up
                // With nuDcPopulator('    ).Setting("doyle", "D"'    ).Setting("riverview", "R"').Matching(Split(.FullDocumentName, "\"))
                // If .Count = 1 Then
                // pf = .Item(.Keys(0))
                // Else
                // pf = ""
                // End If
                // End With

                {
                    var withBlock1 = withBlock.PropertySets.Item(gnDesign);
                    mdFam = withBlock1.Item(pnFamily).Value;
                    itNum = withBlock1.Item(pnPartNum).Value;
                    gnFam = famInGenius(itNum);
                }
            }

            {
                var withBlock = compDefOf(AiDoc);
                if (withBlock.BOMStructure == kPurchasedBOMStructure)
                    sf = "PTS";
            }
        }
    }

    public string famInGenius(string itNum)
    {
        string gnFam;

        {
            var withBlock = cnGnsDoyle().Execute("select Family from vgMfiItems where Item = '" + itNum + "';");
            if (withBlock.BOF | withBlock.EOF)
                gnFam = "";
            else
                gnFam = Split(withBlock.GetString(adClipString, null/* Conversion error: Set to default value for this argument */, "", "", ""), Constants.vbCr)(0);
        }

        famInGenius = gnFam;
    }

    public string famIfValid(string mdFam)
    {
        {
            var withBlock = cnGnsDoyle().Execute(Join(Array("select ISNULL(f.Family, '') Family", "from (values ('" + mdFam + "')) as i(f)", "left join vgMfiFamilies f on i.f = f.Family")));
            if (withBlock.BOF | withBlock.EOF)
                famIfValid = "";
            else
                famIfValid = withBlock.Fields("Family").Value;
        }
    }

    public string famVsGenius(string itNum, string mdFam = "")
    {
        string ckFam;
        string gnFam;
        VbMsgBoxResult ck;

        /// get current family from
        /// Genius, if it has one
        gnFam = famInGenius(itNum);

        if (Strings.Len(gnFam) == 0)
            famVsGenius = mdFam; // so just use the model's
        else
        {
            /// first, verify model family
            ckFam = famIfValid(mdFam);
            /// if not in Genius...

            if (Strings.Len(ckFam) == 0)
                famVsGenius = gnFam;
            else if (gnFam == ckFam)
                famVsGenius = ckFam;
            else
            {
                ck = MsgBox(Join(Array("Item " + itNum, "Model Part Family " + ckFam + " differs", "from Genius Part Family " + gnFam, "", "Change Model to match Genius?", "", "(click [CANCEL] to debug)"), Constants.vbNewLine), Constants.vbYesNoCancel + Constants.vbQuestion, "Use Genius Family?");

                if (ck == Constants.vbCancel)
                    System.Diagnostics.Debugger.Break(); // to debug
                else if (ck == Constants.vbYes)
                    famVsGenius = gnFam;
                else
                    famVsGenius = ckFam;
            }
        }
    }

    public string msg_2022_0603_1127(Scripting.Dictionary dc)
    {
        /// msg_2022_0603_1127
        /// 
        Scripting.Dictionary cc;
        Scripting.Dictionary rm;
        string rt;

        rt = "";

        cc = d4g4f3(dc);
        rm = dcKeysMissing(dc, cc);

        if (rm.Count > 0)
            rt = Join(Array(rt, "The following Part Numbers are", "assigned to more than one Model:", "", Constants.vbTab + Join(rm.Keys, Constants.vbNewLine + Constants.vbTab), ""), Constants.vbNewLine);

        if (cc.Count > 0)
            rt = Join(Array(rt, "These duplicated Part Numbers are", "associated with at least one Content", "Center Member, which cannot be modified:", "", Constants.vbTab + Join(cc.Keys, Constants.vbNewLine + Constants.vbTab), ""), Constants.vbNewLine);

        rt = Join(Array(rt, "These will not be processed.", ""), Constants.vbNewLine);

        msg_2022_0603_1127 = rt; // Join(Array("The following Part Numbers are","assigned to more than one Model:","",vbTab & Join(d4g4f4(.Dictionary).Keys, vbNewLine & vbTab),"","These duplicated Part Numbers are","associated with at least one Content","Center Member, which cannot be modified:","",vbTab & Join(cc.Keys, vbNewLine & vbTab),"","These will not be processed.",""), vbNewLine)
    }

    public Scripting.Dictionary askUserForPartMatl(Inventor.PartDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// askUserForPartMatl -- Prompt User
        /// for Part Family and Material
        /// Selection, returning result
        /// in Dictionary
        /// 
        /// REV[2022.08.29.1621]
        /// add optional Dictionary parameter to which
        /// data from this function can be added.
        /// (see also askUserForMatlQty)
        /// 
        Scripting.Dictionary rt;
        VbMsgBoxResult ck;
        aiBoxData bd;
        string tx;

        if (dc == null)
            rt = askUserForPartMatl(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc; // New Scripting.Dictionary

            {
                var withBlock = rt;
                if (withBlock.Exists(pnFamily))
                    withBlock.Remove(pnFamily);
                if (withBlock.Exists(pnRawMaterial))
                    withBlock.Remove(pnRawMaterial);
                if (AiDoc == null)
                {
                    withBlock.Add(pnFamily, "");
                    withBlock.Add(pnRawMaterial, "");
                }
                else
                {
                    var withBlock1 = newFmTest1();
                    bd = nuAiBoxData().UsingInches.SortingDims(AiDoc.ComponentDefinition.RangeBox);
                    ck = withBlock1.AskAbout(AiDoc, "No Stock Found! Please Review" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                    if (ck == Constants.vbYes)
                    {
                        // Stop 'because this will
                        // override supplied Dictionary!
                        // rt =
                        {
                            var withBlock2 = withBlock1.ItemData();
                            rt.Add(pnFamily, withBlock2.Item(pnFamily));
                            rt.Add(pnRawMaterial, withBlock2.Item(pnRawMaterial));
                        }
                    }
                    else
                    {
                        // rt = New Scripting.Dictionary
                        System.Diagnostics.Debugger.Break();

                        {
                            var withBlock2 = AiDoc.PropertySets;
                            tx = withBlock2.Item(gnDesign).Item(pnFamily).Value;
                            rt.Add(pnFamily, tx);
                            Information.Err.Clear();
                            tx = withBlock2.Item(gnCustom).Item(pnRawMaterial).Value;
                            if (Information.Err.Number)
                                tx = "";

                            rt.Add(pnRawMaterial, tx);
                        }
                    }
                }
            }
        }

        askUserForPartMatl = rt;
    }

    public Scripting.Dictionary askUserForMatlQty(Inventor.PartDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// askUserForMatlQty -- Prompt User
        /// for Material Quantity and Units,
        /// returning result in Dictionary
        /// 
        /// REV[2022.08.29.1624]
        /// add optional Dictionary parameter to which
        /// data from this function can be added.
        /// (see also askUserForPartMatl)
        /// 
        Scripting.Dictionary rt;
        VbMsgBoxResult ck;
        aiBoxData bd;
        string tx;

        if (dc == null)
            rt = askUserForMatlQty(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc; // New Scripting.Dictionary
            {
                var withBlock = rt;
                if (withBlock.Exists(pnRmQty))
                    withBlock.Remove(pnRmQty);
                if (withBlock.Exists(pnRmUnit))
                    withBlock.Remove(pnRmUnit);
            }

            if (AiDoc == null)
            {
                rt.Add(pnRmQty, 0);
                rt.Add(pnRmUnit, "");
            }
            else
            {
                var withBlock = nu_fmIfcMatlQty01().SeeUserWithPart(AiDoc);
                /// following copied from dcGeniusPropsPartRev20180530 line 1632~?
                if (withBlock.Exists(pnRmQty))
                    rt.Add(pnRmQty, withBlock.Item(pnRmQty));
                else
                {
                }

                if (withBlock.Exists(pnRmUnit))
                    rt.Add(pnRmUnit, withBlock.Item(pnRmUnit));
                else
                {
                }
            }
        }

        askUserForMatlQty = rt;
    }

    public Scripting.Dictionary askUserForPartMatlUpdate(Inventor.PartDocument AiDoc)
    {
        /// askUserForPartMatlUpdate --
        /// Attempt to update Part Document
        /// material Properties from results
        /// of askUserForPartMatl
        /// (Family and Material Selection)
        /// and askUserForMatlQty
        /// (Material Quantity and Units)
        /// Return Dictionary of results
        /// 
        /// NOTE[2022.08.29.1627]
        /// want to separate user data collection
        /// from property updates in this function.
        /// further review/development called for.
        /// 
        Scripting.Dictionary dcPr;
        Scripting.Dictionary dcWk;
        Scripting.Dictionary rt;
        Inventor.Property pr;
        VbMsgBoxResult ck;
        Variant ky;

        dcPr = dcOfPropsInAiDoc(AiDoc);
        ck = Constants.vbOK;
        if (!dcPr.Exists(pnRawMaterial))
        {
            ck = MsgBox(Join(Array("Custom Property " + pnRawMaterial + ",", "used to identify Raw Material,", "is not yet present in this model.", "", "Go ahead and create it?"), Constants.vbNewLine), Constants.vbYesNo + Constants.vbQuestion, "Required Property Missing!");

            if (ck == Constants.vbYes)
            {
                {
                    var withBlock = AiDoc.PropertySets.Item(gnCustom);
                    foreach (var ky in Array(pnRawMaterial, pnRmQty, pnRmUnit))
                    {
                    }

                    Information.Err.Clear();
                    pr = withBlock.Add("", ky); // pnRawMaterial
                    if (Information.Err.Number == 0)
                    {
                        dcPr.Add(ky, pr); ck = Constants.vbOK;
                    }
                    else
                        ck = Constants.vbAbort;
                }
            }
            else
                ck = Constants.vbOK;
        }

        if (ck != Constants.vbOK)
            ck = MsgBox(Join(Array("Custom Property " + pnRawMaterial + ",", "was not created! Raw Material", "will not be saved!"), Constants.vbNewLine), Constants.vbOKCancel + Constants.vbExclamation, "Property Not Created!");

        rt = new Scripting.Dictionary();

        if (ck == Constants.vbOK)
        {
            /// REV[2022.08.29.1616]
            /// condense two nearly identical With blocks
            /// into one, combining results of part material
            /// and material quantity data collections.
            /// NOTE: this required additional REVs
            /// to askUserForPartMatl (nee d4g1f1)
            /// and askUserForMatlQty (nee d4g1f3)
            /// to accept optional Dictionary to receive
            /// data points collected by each function.
            {
                var withBlock = askUserForMatlQty(AiDoc, askUserForPartMatl(AiDoc));
                foreach (var ky in withBlock.Keys)
                {
                    {
                        var withBlock1 = dcPr;
                        if (withBlock1.Exists(ky))
                            pr = withBlock1.Item(ky);
                        else
                            pr = null/* TODO Change to default(_) if this is not a reference type */;
                    }

                    if (!pr == null)
                    {
                        if (Len(Trim(withBlock.Item(ky))) > 0)
                        {
                            if (pr.Value != withBlock.Item(ky))
                            {
                                Information.Err.Clear();
                                // Stop 'so we can make sure this works
                                pr.Value = withBlock.Item(ky);
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                                                                             // DON'T try to step at pr.Value
                                if (Information.Err.Number)
                                    System.Diagnostics.Debugger.Break();
                            }
                        }
                        rt.Add(ky, pr.Value);
                    }
                }
            }
        }

        askUserForPartMatlUpdate = rt;
    }

    public Scripting.Dictionary dcGeniusPropsPartRev20180530_ck(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// 
        /// NOTICE TO DEVELOPER [2021.11.12]
        /// '''
        /// 
        /// This function definition was restored
        /// from a prior copy of this project
        /// (VB-000-1002_2021-1001.ipt)
        /// to restore current "normal" operation
        /// of the Genius Properties Update macro.
        /// The prior development version was
        /// retained for reference, renamed to
        /// dcGeniusPropsPartRev20180530_ck_broken
        /// 
        /// One minor revision was made to this
        /// restored version to retain improved
        /// generation of Genius Mass data.
        /// Additional changes should be kept
        /// to a MINIMUM to maintain correct
        /// operation going forward, and any
        /// desired changes implemented through
        /// some form of "shim"
        /// 
        /// '''
        Scripting.Dictionary rt;
        /// REV[2022.01.21.1351]
        /// Added following two Dictionaries
        Scripting.Dictionary dcIn;
        /// to collect settings already in Genius
        Scripting.Dictionary dcFP;
        /// to add a layer of separation
        /// to FlatPattern data collection
        /// (might not want to use for Properties
        /// so don't update immediately)

        // '
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;
        // '
        Inventor.Property prPartNum; // pnPartNum
        /// ADDED[2021.03.11] to simplify access
        /// to Part Number of Model, since it's
        /// requested several times in function
        Inventor.Property prFamily;
        Inventor.Property prRawMatl; // pnRawMaterial
        Inventor.Property prRmUnit; // pnRmUnit
        Inventor.Property prRmQty; // pnRmQty
                                   // '
        string pnModel;
        /// ADDED[2021.03.11] to further
        /// simplify access to Part Number
        string nmFamily;
        string mtFamily;
        /// UPDATE[2018.05.30]:
        /// Rename variable Family to nmFamily
        /// to minimize confusion between code
        /// and comment text in searches.
        /// Also add variable mtFamily
        /// for raw material Family name
        string pnStock;
        string qtUnit;
        Inventor.BOMStructureEnum bomStruct;
        VbMsgBoxResult ck;
        aiBoxData bd;

        if (dc == null)
            dcGeniusPropsPartRev20180530_ck = dcGeniusPropsPartRev20180530_ck(invDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = invDoc;
                // Property Sets
                {
                    var withBlock1 = withBlock.PropertySets;
                    aiPropsUser = withBlock1.Item(gnCustom);
                    aiPropsDesign = withBlock1.Item(gnDesign);
                }

                // Custom Properties
                prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1);
                prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1);
                prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1);

                // Part Number and Family properties
                // are from Design, NOT Custom set
                prPartNum = aiGetProp(aiPropsDesign, pnPartNum);
                // ADDED 2021.03.11
                pnModel = prPartNum.Value;
                prFamily = aiGetProp(aiPropsDesign, pnFamily);

                /// We should check HERE for possibly misidentified purchased parts
                /// UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// Request #1:  the Mass in Pounds
                    /// and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        /// Update [2021.11.12]
                        /// Round mass to nearest ten-thousandth
                        /// to try to match expected Genius value.
                        /// This should reduce or minimize reported
                        /// discrepancies during ETM process.
                        rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                    }

                    /// BOM Structure type, correcting if appropriate,
                    /// and prepare Family value for part, if purchased.
                    /// 
                    ck = Constants.vbNo;
                    /// UPDATE[2018.05.31]: Combined both InStr checks
                    /// by addition to generate a single test for > 0
                    /// If EITHER string match succeeds, the total
                    /// SHOULD exceed zero, so this SHOULD work.
                    if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0)
                        /// UPDATE[2018.02.06]: Using same
                        /// new UserForm as noted above.
                        ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Is this a Purchased Part?" + Constants.vbNewLine + "(Cancel to debug)");

                    /// Check process below replaces duplicate check/responses above.
                    if (ck == Constants.vbCancel)
                        System.Diagnostics.Debugger.Break();
                    else if (ck == Constants.vbYes)
                    {
                        if (withBlock1.BOMStructure != kPurchasedBOMStructure)
                        {
                            withBlock1.BOMStructure = kPurchasedBOMStructure;
                            if (Information.Err.Number == 0)
                                bomStruct = withBlock1.BOMStructure;
                            else
                                bomStruct = kPurchasedBOMStructure;
                        }
                        else
                            bomStruct = withBlock1.BOMStructure;// to make sure this is captured
                    }
                    else
                        bomStruct = withBlock1.BOMStructure;// to make sure this is captured

                    // Request #2: Change Cost Center iProperty.
                    // If BOMStructure = Purchased and not content center,
                    // then Family = D-PTS, else Family = D-HDWR.
                    // 
                    // UPDATE[2018-05-30]: Value produced here
                    // will now be held for later processing,
                    // more toward the end of this function.
                    if (bomStruct == kPurchasedBOMStructure)
                    {
                        if (withBlock1.IsContentMember)
                            nmFamily = "D-HDWR";
                        else
                            nmFamily = "D-PTS";
                    }
                    else
                        nmFamily = "";
                }
                /// At this point, nmFamily SHOULD be set
                /// to a non-blank value if Item is purchased.
                /// We should be able to check this later on,
                /// if Item BOMStructure is NOT Normal

                // Request #4: Change Cost Center iProperty.
                // If BOMStructure = Normal, then Family = D-MTO,
                // else if BOMStructure = Purchased then Family = D-PTS.
                if (bomStruct == kNormalBOMStructure)
                {

                    /// REV[2022.01.28.1014]
                    /// Added initial raw material capture
                    /// to check against Genius
                    /// HOLD![2022.01.28.1046]
                    /// commenting out again
                    /// probably best below, still
                    pnStock = prRawMatl.Value;
                    /// REV[2022.02.08.1304]
                    /// restored, to obtain any
                    /// value already defined.
                    /// MIGHT need moved further down,
                    /// but hold off on that for now.

                    /// REV[2022.01.17.1123]
                    /// Start adding code to capture
                    /// any raw material items for
                    /// part already in Genius.
                    /// REV[2022.01.21.1357]
                    /// Separated capture from With statement
                    /// into new Dictionary object in order
                    /// to check and use it further down,
                    /// as well as passing it to nuSelFromDict
                    /// to handle multiple line items
                    /// REV[2022.01.31.1008]
                    /// Restored assignment of dcFromAdoRS
                    /// result to Dictionary Object dcIn,
                    /// in order to pass it to other
                    /// functions, as needed.
                    /// 
                    dcIn = dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)));
                    // Debug.Print ConvertToJson(dcDxFromRecSetDc(dcIn), vbTab)
                    // Stop
                    // dcIn = dcOb(dcDxFromRecSetDc(dcIn).Item(pnRawMaterial))
                    if (dcIn.Count > 0)
                    {
                        {
                            var withBlock1 = dcOb(dcDxFromRecSetDc(dcIn).Item(pnRawMaterial));
                            /// REV[2022.01.28.1336]
                            /// Added code to collect captured
                            // dcIn = New Scripting.Dictionary


                            /// REV[2022.01.28.0857]
                            /// Added code to collect captured
                            /// material item number, asking user
                            /// to select from list if more than one.
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
                                    // pnStock = dcOb(.Item(.Keys(0))).Item(pnRawMaterial)
                                    pnStock = withBlock1.Keys(0);

                                // and use it for the default...
                                if (withBlock1.Count > 1)
                                {
                                    System.Diagnostics.Debugger.Break(); // because selection is going
                                                                         // to be a lot more complicated.
                                                                         // (just look at that pnStock
                                                                         // assignment up there!)

                                    pnStock = nuSelector().GetReply(withBlock1.Keys, pnStock);

                                    System.Diagnostics.Debugger.Break(); // to make sure things are okay
                                }
                            }
                            else
                            {
                            }

                            /// REV[2022.01.28.0903]
                            /// Separated Dictionary capture
                            /// from Count check
                            if (Strings.Len(pnStock) > 0)
                            {
                                if (Len(System.Convert.ToHexString(prRawMatl.Value)) == 0)
                                {
                                }
                                else if (pnStock == prRawMatl.Value)
                                {
                                }
                                else
                                {
                                    Debug.Print("=== CURRENT GENIUS MATERIAL DATA ===");
                                    // Debug.Print dumpLsKeyVal(dcIn, ":" & vbTab)
                                    ck = newFmTest2().AskAbout(invDoc, "Raw Material " + prRawMatl.Value, "does not match " + pnStock);
                                    if (ck == Constants.vbCancel)
                                        System.Diagnostics.Debugger.Break(); // to check things out
                                    else if (ck == Constants.vbNo)
                                        /// NOTE[2022.02.08.1359]
                                        /// DO NOT DISABLE this instance
                                        /// of the pnStock assignment!
                                        pnStock = prRawMatl.Value;
                                }

                                /// REV[2022.01.28.1448]
                                /// Changed data extraction process here
                                /// to work with form returned from dcFromAdoRS
                                /// 
                                /// NOTE! This is !!!TEMPORARY!!!
                                /// Implemented during run time,
                                /// some truly insane acrobatics were required
                                /// to make it work without resetting the run.
                                /// This code, including the With statement
                                /// above, MUST be rewritten as soon as feasible!
                                /// 
                                // Stop 'because we're doing to need to do something different
                                // Debug.Print ConvertToJson(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial), vbTab)
                                // Debug.Print ConvertToJson(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock), vbTab)
                                // Debug.Print ConvertToJson(dcOb(.Item(dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0))), vbTab)
                                // dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0)
                                // Stop

                                if (withBlock1.Exists(pnStock))
                                {
                                    dcIn = dcOb(dcIn.Item(dcOb(withBlock1.Item(pnStock)).Keys(0)));
                                    // This is DEFINITELY going to need a rework!
                                    // But that will need a new function, most likely

                                    // deactivated the version below
                                    // to be superceded by the one above
                                    // dcIn = dcOb(.Item(dcOb(dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)))).Item(pnRawMaterial)).Item(pnStock)).Keys(0)))

                                    // original version, also deactivated
                                    // for obvious reasons
                                    // dcIn = .Item(pnStock)

                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                }
                                else
                                    System.Diagnostics.Debugger.Break();// because we've got a REAL problem here!
                            }
                            else
                                dcIn = new Scripting.Dictionary();
                        }
                    }

                    {
                        var withBlock1 = dcIn;
                        if (withBlock1.Count == 0)
                        {
                            withBlock1.Add("Ord", 0);
                            withBlock1.Add("RM", "");
                            withBlock1.Add("MtFamily", "");
                            withBlock1.Add("RMQTY", 0);
                            withBlock1.Add("RMUNIT", "");
                        }
                    }

                    // ----------------------------------------------------'
                    if (withBlock.SubType == guidSheetMetal)
                    {
                        // ----------------------------------------------------'
                        /// NOTE[2018-05-31]: At this point, we MAY wish
                        /// to check for a valid flat pattern,
                        /// and otherwise attempt to verify
                        /// an actual sheet metal design.
                        /// 

                        /// REV[2022.01.28.0903]
                        /// HERE is where things start to get interesting
                        /// Before processing Part as sheet metal,
                        /// want to make sure it's supposed to be.
                        /// 
                        /// FIRST, check what Genius had to say
                        {
                            var withBlock1 = dcIn;
                            if (withBlock1.Exists("MtFamily"))
                                mtFamily = withBlock1.Item("MtFamily");
                            else
                                mtFamily = "";
                        }

                        if (Strings.Len(mtFamily) == 0)
                            ck = Constants.vbRetry;
                        else if (mtFamily == "DSHEET")
                            ck = Constants.vbYes;
                        else
                            ck = Constants.vbNo;

                        /// REV[2022.01.31.1335]
                        /// Move flat pattern collection out here
                        /// from inside the next If-Then block
                        if (ck == Constants.vbNo)
                            dcFP = new Scripting.Dictionary();
                        else
                        {
                            dcFP = dcFlatPatVals(withBlock.ComponentDefinition);
                            /// try to get flat pattern data
                            /// WITHOUT mucking up Properties!
                            /// Want to avoid dirtying file with
                            /// changes until absolutely necessary)

                            if (dcFP.Exists(pnThickness))
                            {
                                pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                dcFP.Add(pnRawMaterial, pnStock);
                            }
                        }
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        if (false)
                            Debug.Print(ConvertToJson(Array(dcIn, dcFP), Constants.vbTab));

                        if (ck == Constants.vbRetry)
                        {
                            /// so let's see what the flat pattern can tell us

                            if (dcFP.Exists("mtFamily"))
                            {
                                if (dcFP.Item("mtFamily") == "DSHEET")
                                {
                                    if (dcFP.Exists("OFFTHK"))
                                    {
                                        System.Diagnostics.Debugger.Break();
                                        ck = newFmTest2().AskAbout(invDoc, "This Part: ", "might not be sheet metal. " + Constants.vbNewLine + Constants.vbNewLine + "Is it in fact sheet metal?");
                                        if (ck == Constants.vbCancel)
                                        {
                                            ck = Constants.vbRetry;
                                            System.Diagnostics.Debugger.Break(); // to debug
                                        }
                                    }
                                    else
                                        ck = Constants.vbYes;
                                }
                                else if (dcFP.Item("mtFamily") == "D-BAR")
                                    ck = Constants.vbNo;
                                else
                                    ck = Constants.vbRetry;
                            }
                            else
                                ck = Constants.vbRetry;
                        }

                        if (ck == Constants.vbRetry)
                        {
                            Debug.Print(ConvertToJson(Array(dcIn, dcFP), Constants.vbTab));
                            System.Diagnostics.Debugger.Break(); // so we can figure out what to do next.
                        }

                        // Request #3:
                        // sheet metal extent area
                        // and add to custom property "RMQTY"

                        /// REV[2022.01.28.1556]
                        /// change if-then-else sequence
                        /// to check ck instead of dcIn
                        if (ck == Constants.vbYes)
                            rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                        else if (ck == Constants.vbRetry)
                            rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                        else if (ck == Constants.vbNo)
                        {
                        }
                        else
                            // material type detection SHOULD produce
                            // one of the three preceding values

                            System.Diagnostics.Debugger.Break();// and check it out

                        /// NOTE[2018-05-30]:
                        /// Raw Material Quantity value
                        /// SHOULD be set upon return
                        /// We may need to review the process
                        /// to find an appropriate place
                        /// to set for NON sheet metal

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
                                /// NOTE[2021.12.10]:
                                /// Believe this OFFTHK property is meant
                                /// to capture "Sheet Metal" Parts that
                                /// aren't actually Sheet Metal.
                                /// This check might be needed further down.
                                /// UPDATE[2018.05.30]:
                                /// Restoring original key check
                                /// and adding code for debug
                                /// Previously changed to "~OFFTHK"
                                /// to avoid this block and its issues.
                                /// (Might re-revert if not prepped to fix now)
                                Debug.Print(aiProperty(rt.Item("OFFTHK")).Value);
                                System.Diagnostics.Debugger.Break(); // because we're going to need to do something with this.

                                pnStock = ""; // Originally the ONLY line in this block.
                                              // A more substantial response is required here.

                                if (0)
                                    System.Diagnostics.Debugger.Break(); // (just a skipover)
                            }
                            else
                            {
                                System.Diagnostics.Debugger.Break(); // because we don't know IF this is sheet metal yet
                                pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                            }
                        }
                        else
                        {
                            // '  ACTION ADVISED[2018.09.14]:
                            // '  pnStock can probably be set
                            // '  to prRawMatl.Value and THEN
                            // '  checked for length to see
                            // '  if lookup needed.
                            // '  This might also allow us to check
                            // '  for machined or other non-sheet
                            // '  metal parts.

                            /// REV[2021.12.17]: sanity check
                            /// Add sanity check to make sure
                            /// any existing sheet metal stock
                            /// number matches model specs
                            if (Len(prRawMatl.Value) > 0)
                            {
                                // we need to check it

                                if (Strings.Len(pnStock) == 0)
                                    /// REV[2022.01.28.1445]:
                                    /// Placed this pnStock stock assignment
                                    /// inside this If-Then block to prevent
                                    /// overriding value from Genius
                                    pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                                /// NOTE[2021.12.17@15:32]:
                                /// copied this up from
                                /// NOTE[2021.12.17@15:32]
                                /// for use in sanity check

                                /// NOTE[2021.12.17]:
                                /// This section simply warns the user
                                /// that the current raw material does
                                /// not match the recommended default,
                                /// and offers an opportunity to fix it.
                                /// 
                                /// This is yet another quick and dirty
                                /// "solution" that should be revised
                                /// NOTE[2022.01.05]:
                                /// Adding check for empty recommendation.
                                /// Do NOT believe user should be offered
                                /// opportunity to overwrite any current
                                /// part number with a BLANK one. Believe
                                /// the option to CLEAR is somewhere below.
                                if (Strings.Len(pnStock) > 0)
                                {
                                    if (pnStock != prRawMatl.Value)
                                    {
                                        // Stop

                                        /// NOTE[2022.01.03]:
                                        /// Following text SHOULD no longer
                                        /// be needed. Verify function of
                                        /// fmTest2 following, and when good,
                                        /// disable and/or remove this block.
                                        Debug.Print("!!! NOTICE !!!");
                                        Debug.Print("Recommended Sheet Metal Stock (" + pnStock + ")");
                                        Debug.Print("does not match current Stock (" + prRawMatl.Value + ")");
                                        Debug.Print();
                                        Debug.Print("To continue with no change, just press [F5]. Otherwise,");
                                        Debug.Print("press [ENTER] on the following line first to change:");
                                        Debug.Print("prRawMatl.Value = \"" + pnStock + "\"");
                                        Debug.Print();

                                        /// NOTE[2022.01.03]:
                                        /// Now using fmTest2(?) to prompt
                                        /// user as in other checks (above?)
                                        ck = newFmTest2().AskAbout(invDoc, "Suggest Sheet Metal change from" + Constants.vbNewLine + prRawMatl.Value + " to" + Constants.vbNewLine + pnStock + " for", "Change it?");
                                        if (ck == Constants.vbCancel)
                                            System.Diagnostics.Debugger.Break(); // to check things out
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
                                    /// NOTE[2021.12.13]:
                                    /// OFFTHK property check added
                                    /// to catch sheet metal already
                                    /// assigned by accident.
                                    ck = newFmTest2().AskAbout(invDoc, "Assigned Raw Material " + prRawMatl.Value, "Clear it?");
                                    if (ck == Constants.vbCancel)
                                        System.Diagnostics.Debugger.Break(); // to check things out
                                    else if (ck == Constants.vbYes)
                                        prRawMatl.Value = "";
                                }


                                if (pnStock == prRawMatl.Value)
                                    // no need to assign it again
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                else
                                {
                                    Debug.Print(ConvertToJson(Array(pnStock, prRawMatl.Value))); System.Diagnostics.Debugger.Break(); // before we do something stupid!
                                    pnStock = prRawMatl.Value;
                                }

                                /// The following With block copied and modified [2021.03.11]
                                /// from elsewhere in this function as a temporary measure
                                /// to address a stopping situation later in the function.
                                /// See comment below for details.
                                /// 
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                                    if (withBlock1.BOF | withBlock1.EOF)
                                    {
                                        if (pnStock != "0")
                                            System.Diagnostics.Debugger.Break();// because Material value likely invalid
                                        /// REV[2022.02.08.1413]
                                        /// reinstated interruption here
                                        /// because at this point, pnStock
                                        /// has likely already been assigned
                                        /// to prRawMatl, so changing it here
                                        /// is NOT likely to be productive.
                                        /// this section will likely need
                                        /// reconsideration, revision,
                                        /// and/or possibly removal.
                                        /// UPDATE[2021.12.10]:
                                        /// added this check for OFFTHK
                                        /// to avoid blindly adding sheet
                                        /// metal stock to a "sheet metal"
                                        /// part that isn't actually meant
                                        /// to be made of sheet metal.
                                        if (rt.Exists("OFFTHK"))
                                            // actual Sheet Metal, so just clear this:
                                            pnStock = "";
                                        else
                                        {
                                            pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                        }
                                    }
                                    else
                                    {
                                    }
                                }
                            }
                            else if (rt.Exists("OFFTHK"))
                                /// UPDATE[2021.12.10]:
                                /// another OFFTHK check added to avoid
                                /// adding sheet metal stock by mistake.
                                pnStock = "";
                            else
                                pnStock = ptNumShtMetal(withBlock.ComponentDefinition);

                            if (Strings.Len(pnStock) == 0)
                            {
                                /// UPDATE[2018.05.30]:
                                /// Pulling ALL code/text from this section
                                /// to get rid of excessive cruft.
                                /// 
                                /// In fact, reversing logic to go directly
                                /// to User Prompt if no stock identified
                                /// 
                                /// IN DOUBLE FACT, hauling this WHOLE MESS
                                /// RIGHT UP after initial pnStock assignment
                                /// to prompt user IMMEDIATELY if no stock found
                                {
                                    var withBlock1 = newFmTest1();
                                    if (!(invDoc.ComponentDefinition.Document == invDoc))
                                        System.Diagnostics.Debugger.Break();

                                    bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                                    ck = withBlock1.AskAbout(invDoc, "No Stock Found! Please Review" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                                    if (ck == Constants.vbYes)
                                    {
                                        /// UPDATE[2018.05.30]:
                                        /// Pulling some extraneous commented code
                                        /// from here and beginning of block
                                        {
                                            var withBlock2 = withBlock1.ItemData;
                                            if (withBlock2.Exists(pnFamily))
                                            {
                                                nmFamily = withBlock2.Item(pnFamily);
                                                Debug.Print(pnFamily + "=" + nmFamily);
                                            }

                                            if (withBlock2.Exists(pnRawMaterial))
                                            {
                                                pnStock = withBlock2.Item(pnRawMaterial);
                                                Debug.Print(pnRawMaterial + "=" + pnStock);
                                            }
                                        }
                                        if (0)
                                            System.Diagnostics.Debugger.Break(); // Use this for a debugging shim
                                    }
                                }
                            }
                            else if (Left(pnStock, 2) == "LG")
                            {
                                Debug.Print(pnModel + ": PROBABLE LAGGING [" + pnStock + "]");
                                Debug.Print("  TRY TO VERIFY. IF CHANGE REQUIRED,");
                                Debug.Print("  FILL IN NEW VALUE FOR pnStock BELOW, ");
                                Debug.Print("  AND PRESS ENTER ON THE LINE. WHEN ");
                                Debug.Print("  READY, PRESS [F5] TO CONTINUE.");
                                Debug.Print("  pnStock = \"" + pnStock + "\"");
                                System.Diagnostics.Debugger.Break();
                            }

                            if (Strings.Len(pnStock) > 0)
                            {
                                // do we look for a Raw Material Family!

                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family, Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16 " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                                    if (withBlock1.BOF | withBlock1.EOF)
                                        System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                    else
                                    {
                                        {
                                            var withBlock2 = withBlock1.Fields;
                                            mtFamily = withBlock2.Item("Family").Value;
                                        }

                                        /// UPDATE[2021.06.18]:
                                        /// New pre-check for Material Item
                                        /// in Purchased Parts Family.
                                        /// VERY basic handler simply
                                        /// maps Material Family to D-BAR
                                        /// to force extra processing below.
                                        /// Further refinement VERY much needed!
                                        if (mtFamily) /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                        {
                                            // Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                            Debug.Print(pnModel + "[" + prRmQty.Value + qtUnit + "*" + pnStock + ": " + aiPropsDesign(pnDesc).Value + "]"); System.Diagnostics.Debugger.Break(); // FULL Stop!
                                        }
                                        else if (mtFamily == "D-PTS")
                                        {
                                            nmFamily = "D-RMT";
                                            System.Diagnostics.Debugger.Break(); // NOT SO FAST!
                                            mtFamily = "D-BAR";
                                        }
                                        else if (mtFamily == "R-PTS")
                                        {
                                            nmFamily = "R-RMT";
                                            System.Diagnostics.Debugger.Break(); // NOT SO FAST!
                                            mtFamily = "D-BAR";
                                        }

                                        if (mtFamily == "DSHEET")
                                        {
                                            // We should be okay. This is sheet metal stock
                                            nmFamily = "D-RMT";
                                            qtUnit = "FT2";
                                        }
                                        else if (mtFamily == "D-BAR")
                                        {
                                            /// UPDATE[2021.06.18]:
                                            /// Added check for Part Family already set
                                            /// to more properly handle new situation (above)
                                            if (Strings.Len(nmFamily) == 0)
                                                nmFamily = "R-RMT";
                                            else
                                                Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing

                                            /// UPDATE[2022.01.11]:
                                            /// Adding Do..Loop Until to following section
                                            /// to allow user to retry setting material
                                            /// quantity and units. This change made in
                                            /// conjunction with new prompt form (below).
                                            /// NOTE! This is FIRST instance of revision
                                            /// Search on UPDATE text above to locate
                                            /// the other in this function
                                            qtUnit = prRmUnit.Value; // "IN"
                                            ck = Constants.vbCancel;
                                            do
                                            {

                                                // 'may want function here
                                                /// UPDATE[2018.05.30]: As noted above
                                                /// Will keep Stop for now
                                                /// pending further review,
                                                /// hopefully soon
                                                // Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                                // Debug.Print CDbl(dcIn.Item(pnRmQty))
                                                /// UPDATE[2021.03.11]: Replaced
                                                /// aiPropsDesign.Item(pnPartNum)
                                                /// with prPartNum (and now pnModel)
                                                /// since it's used in several places

                                                // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                                // Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                                // Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."

                                                /// REV[2022.02.08.1511]
                                                /// replaced boilerplate above with new version below
                                                /// in hopes of better presenting change options
                                                /// in a more compact and accessible form.

                                                Debug.Print("===== CHECK AND VERIFY RAW MATERIAL QUANTITY =====");
                                                Debug.Print("  If change required, place new values at end");
                                                Debug.Print("  of lines below for prRmQty.Value and qtUnit.");
                                                Debug.Print("  Press [ENTER] on each line to be changed.");
                                                Debug.Print("  Press [F5] when ready to continue.");
                                                Debug.Print("----- " + pnModel + " [" + prRawMatl.Value + "]: " + aiPropsDesign(pnDesc).Value + " -----");
                                                // Debug.Print ""

                                                /// REV[2022.02.09.0923]
                                                /// replication of REV[2022.02.09.0919]
                                                /// from section below: prep to replace
                                                /// old dimension dump operation with more
                                                /// compact call to aiBoxData's Dump method
                                                if (true)
                                                {
                                                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                                    {
                                                        var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                        Debug.PrintRound((withBlock2.MaxPoint.X - withBlock2.MinPoint.X) / (double)cvLenIn2cm, 4);
                                                    }
                                                }

                                                {
                                                    var withBlock2 = nuAiBoxData().UsingInches().UsingBox(invDoc.ComponentDefinition.RangeBox);
                                                    Debug.Print.Dump(0);
                                                }
                                                // Stop 'and check output against prior version

                                                /// REV[2022.02.08.1446]
                                                /// removed block of Debug.Print lines
                                                /// disabled now for some time, as they
                                                /// do not seem to have been missed.
                                                Debug.Print("prRmQty.Value = ");
                                                if (dcIn.Exists(pnRmQty))
                                                    Debug.Print("In Genius: ");
                                                Debug.Print();
                                                Debug.Print("qtUnit = \"");
                                                if (dcIn.Exists(pnRmUnit))
                                                    Debug.Print("In Genius: ");
                                                if (dcIn.Item(pnRmUnit) != "IN")
                                                    Debug.Print(" ( or try IN )");
                                                Debug.Print();
                                                // Debug.Print "qtUnit = ""IN"""
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                System.Diagnostics.Debugger.Break();                                        /// Actually, we might NOT need to stop here
                                                                                                                            /// if bar stock is already selected,
                                                                                                                            /// because quantities would presumably
                                                                                                                            /// have been established already.
                                                                                                                            /// Any D-BAR handler probably needs
                                                                                                                            /// to be implemented in prior section(s)
                                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                ck = newFmTest2().AskAbout(invDoc, "Raw Material Quantity is now " + System.Convert.ToHexString(prRmQty.Value) + qtUnit + " for", "If this is okay, click [YES]. Otherwise," + Constants.vbNewLine + "click [NO] or [CANCEL] to fix.");
                                            }
                                            while (!ck == Constants.vbYes)/* TODO ERROR: Skipped SkippedTokensTrivia */// because we might want a D-BAR handler
    ;
                                            /// UPDATE[2022.01.11]:
                                            /// This is the terminal end of the
                                            /// Do..Loop Until block noted above

                                            rt = dcAddProp(prRmQty, rt);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                                        }
                                        else
                                        {
                                            nmFamily = "";
                                            qtUnit = ""; // may want function here
                                            /// UPDATE[2018.05.30]: As noted above
                                            /// However, might need more handling here.
                                            System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                                        }
                                    }
                                }
                            }
                            else if (0)
                                System.Diagnostics.Debugger.Break();// and regroup
                        }

                        {
                            var withBlock1 = prRawMatl;
                            if (Len(Trim(withBlock1.Value)) > 0)
                            {
                                if (pnStock != withBlock1.Value)
                                {
                                    // Debug.Print "Raw Stock Selection"
                                    // Debug.Print "  Current : " & prRawMatl.Value
                                    // Debug.Print "  Proposed: " & pnStock
                                    // Stop 'because we might not want to change existing stock setting
                                    // if
                                    ck = MsgBox(Join(Array("Raw Stock Change Suggested", "  for Item " + pnModel, "", "  Current : " + prRawMatl.Value, "  Proposed: " + pnStock, "", "Change It?", ""), Constants.vbNewLine), Constants.vbYesNo, pnModel + " Stock");
                                    // "Change Raw Material?"
                                    // "Suggested Sheet Metal"
                                    if (ck == Constants.vbYes)
                                        withBlock1.Value = pnStock;
                                }
                            }
                            else
                                withBlock1.Value = pnStock;
                        }
                        rt = dcAddProp(prRawMatl, rt);

                        {
                            var withBlock1 = prRmUnit;
                            if (Len(withBlock1.Value) > 0)
                            {
                                if (Strings.Len(qtUnit) > 0)
                                {
                                    if (withBlock1.Value != qtUnit)
                                    {
                                        // Stop 'and check both so we DON'T
                                        // automatically "fix" the RMUNIT value

                                        ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Raw Material " + prRawMatl.Value);
                                        if (ck == Constants.vbCancel)
                                            System.Diagnostics.Debugger.Break();
                                        else if (ck == Constants.vbYes)
                                            withBlock1.Value = qtUnit;
                                        if (0)
                                            System.Diagnostics.Debugger.Break(); // Ctrl-9 here to skip changing
                                    }
                                }
                            }
                            else
                                withBlock1.Value = qtUnit;
                        }
                        rt = dcAddProp(prRmUnit, rt);
                        // rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                        /// Plan to remove commented line above,
                        /// superceded by the one above that
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Another landing line
                    }
                    else
                    {
                        // --------------------------------------------'
                        /// [2018.07.31 by AT]
                        /// Duped following block from above
                        /// to mod for material assignment
                        /// to non-sheet metal part.
                        /// 
                        /// Except, this isn't enough.
                        /// Also need the code to add
                        /// Stock PN to Attribute RM.
                        /// That's a whole 'nother
                        /// block of code, and likely
                        /// best consolidated.
                        {
                            var withBlock1 = newFmTest1();
                            if (!(invDoc.ComponentDefinition.Document == invDoc))
                                System.Diagnostics.Debugger.Break();

                            /// [2018.07.31 by AT]
                            /// Added the following to try to
                            /// preselect non-sheet metal stock
                            // .dbFamily.Value = "D-BAR"
                            // .lbxFamily.Value = "D-BAR"
                            /// Doesn't quite do it.
                            // With New aiBoxData
                            // bd = nuAiBoxData().UsingInches.UsingBox(invDoc.ComponentDefinition.RangeBox)
                            bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                            // End With

                            ck = withBlock1.AskAbout(invDoc, "Please Select Stock for Machined Part" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                            if (ck == Constants.vbYes)
                            {
                                /// UPDATE[2018.05.30]:
                                /// Pulling some extraneous commented code
                                /// from here and beginning of block
                                {
                                    var withBlock2 = withBlock1.ItemData;
                                    if (withBlock2.Exists(pnFamily))
                                    {
                                        nmFamily = withBlock2.Item(pnFamily);
                                        Debug.Print(pnFamily + "=" + nmFamily);
                                    }

                                    if (withBlock2.Exists(pnRawMaterial))
                                    {
                                        pnStock = withBlock2.Item(pnRawMaterial);
                                        Debug.Print(pnRawMaterial + "=" + pnStock);
                                    }
                                }
                                if (0)
                                    System.Diagnostics.Debugger.Break(); // Use this for a debugging shim
                            }
                        }
                        /// 
                        /// 
                        /// 
                        /// The following If block is copied
                        /// wholesale from sheet metal section above.
                        /// Some changes (to be) made to accommodate
                        /// machined or other non-sheet metal stock.
                        /// 
                        /// Ultimately, whole mess to require refactor.
                        /// 
                        if (Strings.Len(pnStock) > 0)
                        {
                            // do we look for a Raw Material Family!

                            /// This enclosing With block should NOT be necessary

                            /// since the newFmTest1 above takes care of collecting

                            /// the Stock Family along with the Stock itself
                            {
                                var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                                if (withBlock1.BOF | withBlock1.EOF)
                                    System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                else
                                {
                                    var withBlock2 = withBlock1.Fields;
                                    mtFamily = withBlock2.Item("Family").Value;
                                }
                            }
                            /// These closing statements moved up from below following If block

                            /// 

                            // mtFamily = nmFamily 'to force "correct" behavior of following section
                            if (mtFamily == "DSHEET")
                            {
                                System.Diagnostics.Debugger.Break(); // because we should NOT be doing Sheet Metal in this section.
                                                                     // This might require further investigation and/or development, if encountered.
                                                                     // We should be okay. This is sheet metal stock
                                nmFamily = "D-RMT";
                                qtUnit = "FT2";
                            }
                            else if (mtFamily == "D-BAR")
                            {
                                /// UPDATE[2022.01.11]:
                                /// Adding Do..Loop Until to following section
                                /// to allow user to retry setting material
                                /// quantity and units. This change made in
                                /// conjunction with new prompt form (below).
                                /// NOTE! This is SECOND instance of revision
                                /// Search on UPDATE text above to locate
                                /// the other in this function
                                nmFamily = "R-RMT";
                                qtUnit = prRmUnit.Value; // "IN"
                                ck = Constants.vbCancel;
                                do
                                {
                                    // Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                                    /// UPDATE[2021.03.11]: Replaced
                                    /// aiPropsDesign.Item(pnPartNum)
                                    /// as noted above
                                    // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                                    // Debug.Print "THEN SELECT LENGTH FROM THE FOLLOWING SPANS,"
                                    // Debug.Print "AND ENTER AT END OF prRmQty LINE BELOW."

                                    /// REV[2022.02.08.1521]
                                    /// replaced boilerplate above with new version below
                                    /// as per REV[2022.02.08.1511]

                                    Debug.Print("===== CHECK AND VERIFY RAW MATERIAL QUANTITY =====");
                                    Debug.Print("  If change required, place new values at end");
                                    Debug.Print("  of lines below for prRmQty.Value and qtUnit.");
                                    Debug.Print("  Press [ENTER] on each line to be changed.");
                                    Debug.Print("  Press [F5] when ready to continue.");
                                    Debug.Print("----- " + pnModel + " [" + prRawMatl.Value + "]: " + aiPropsDesign(pnDesc).Value + " -----");
                                    // Debug.Print ""

                                    /// REV[2022.02.09.0919]
                                    /// prep to replace old dimension dump
                                    /// operation with more compact call
                                    /// to aiBoxData's Dump method
                                    if (true)
                                    {
                                        Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                        /// REV[2022.02.09.0904]
                                        /// replicated With block from other section
                                        /// to replace original "sprawled out" version
                                        /// of Print statement hastily generated
                                        /// during run time.
                                        {
                                            var withBlock1 = invDoc.ComponentDefinition.RangeBox;
                                            Debug.PrintRound((withBlock1.MaxPoint.X - withBlock1.MinPoint.X) / (double)cvLenIn2cm, 4);
                                        }
                                    }

                                    {
                                        var withBlock1 = nuAiBoxData().UsingInches().UsingBox(invDoc.ComponentDefinition.RangeBox);
                                        Debug.Print.Dump(0);
                                    }
                                    // Stop 'and check output against prior version

                                    /// REV[2022.02.08.1446]
                                    /// removed block of Debug.Print lines
                                    /// disabled now for some time, as they
                                    /// do not seem to have been missed.
                                    Debug.Print("prRmQty.Value = ");
                                    if (dcIn.Exists(pnRmQty))
                                        Debug.Print("In Genius: ");
                                    Debug.Print();
                                    Debug.Print("qtUnit = \"");
                                    if (dcIn.Exists(pnRmUnit))
                                        Debug.Print("In Genius: ");
                                    Debug.Print(" ( or try IN )");

                                    /// REV[2022.02.08.1525]
                                    /// replaced boilerplate below with new version
                                    /// above in like manner to REV[2022.02.08.1446]
                                    /// and also per REV[2022.02.08.1511]

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
                                    System.Diagnostics.Debugger.Break();                                        /// Actually, we might NOT need to stop here
                                                                                                                /// if bar stock is already selected,
                                                                                                                /// because quantities would presumably
                                                                                                                /// have been established already.
                                                                                                                /// Any D-BAR handler probably needs
                                                                                                                /// to be implemented in prior section(s)
                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                    ck = newFmTest2().AskAbout(invDoc, "Raw Material Quantity is now " + System.Convert.ToHexString(prRmQty.Value) + qtUnit + " for", "If this is okay, click [YES]. Otherwise," + Constants.vbNewLine + "click [NO] or [CANCEL] to fix.");
                                }
                                while (!ck == Constants.vbYes)/* TODO ERROR: Skipped SkippedTokensTrivia */// because we might want a D-BAR handler
;
                                /// UPDATE[2022.01.11]:
                                /// This is the terminal end of the
                                /// Do..Loop Until block noted above

                                rt = dcAddProp(prRmQty, rt);
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                            }
                            else
                            {
                                nmFamily = "";
                                qtUnit = ""; // may want function here
                                /// UPDATE[2018.05.30]: As noted above
                                /// However, might need more handling here.
                                System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                            }
                        }
                        else if (0)
                            System.Diagnostics.Debugger.Break();// and regroup


                        /// NOTE[2022.01.07.1004]:
                        /// Another check for empty recommendation.
                        /// (SEE NOTE[2022.01.05] elsewhere in this function)
                        /// Again, don't want user accidentally
                        /// clearing an existing part number.
                        if (Strings.Len(pnStock) > 0)
                        {
                            {
                                var withBlock1 = prRawMatl;
                                if (Len(Trim(withBlock1.Value)) > 0)
                                {
                                    if (pnStock != withBlock1.Value)
                                    {
                                        // Debug.Print "Raw Stock Selection"
                                        // Debug.Print "  Current : " & prRawMatl.Value
                                        // Debug.Print "  Proposed: " & pnStock
                                        // Stop 'because we might not want to change existing stock setting
                                        // if
                                        ck = MsgBox(Join(Array("Raw Stock Change Suggested", "  Current : " + prRawMatl.Value, "  Proposed: " + pnStock, "", "Change It?", ""), Constants.vbNewLine), Constants.vbYesNo, "Change Raw Material?");
                                        // "Suggested Sheet Metal"
                                        if (ck == Constants.vbCancel)
                                            System.Diagnostics.Debugger.Break();
                                        else if (ck == Constants.vbYes)
                                            withBlock1.Value = pnStock;
                                    }
                                }
                                else
                                    withBlock1.Value = pnStock;
                            }
                        }
                        rt = dcAddProp(prRawMatl, rt);

                        {
                            var withBlock1 = prRmUnit;
                            if (Len(withBlock1.Value) > 0)
                            {
                                if (Strings.Len(qtUnit) > 0)
                                {
                                    if (withBlock1.Value != qtUnit)
                                    {
                                        // Stop 'and check both so we DON'T
                                        // automatically "fix" the RMUNIT value

                                        ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Raw Material " + prRawMatl.Value);
                                        if (ck == Constants.vbCancel)
                                            System.Diagnostics.Debugger.Break();
                                        else if (ck == Constants.vbYes)
                                            withBlock1.Value = qtUnit;
                                        if (0)
                                            System.Diagnostics.Debugger.Break(); // Ctrl-9 here to skip changing
                                    }
                                }
                            }
                            else
                                withBlock1.Value = qtUnit;
                        }
                        rt = dcAddProp(prRmUnit, rt);
                    } // Sheetmetal vs Part
                }
                else if (bomStruct == kPurchasedBOMStructure)
                {
                    /// As mentioned above, nmFamily
                    /// SHOULD be set at this point
                    if (Strings.Len(nmFamily) == 0)
                    {
                        if (1)
                            System.Diagnostics.Debugger.Break(); // because we might
                                                                 // need to check out the situation
                        nmFamily = "D-PTS"; // by default
                    }
                }
                else if (bomStruct == kPhantomBOMStructure)
                {
                    /// REV[2022.01.17.1135]
                    /// Adding a crude handler for Phantom
                    /// Part Documents. Since they shouldn't
                    /// have subcomponents to promote, they
                    /// shouldn't have that BOM structure.
                    /// User intervention might be required.
                    ck = newFmTest2().AskAbout(invDoc, "For some reason, THIS Item is marked Phantom:", "Is this okay? (Click [NO] OR [CANCEL] if not)");
                    if (ck == Constants.vbYes)
                    {
                    }
                    else
                        System.Diagnostics.Debugger.Break();
                }
                else
                {
                    /// REV[2022.01.17.1138]
                    /// Adding another handler to catch Part
                    /// Documents with an unexpected BOM Structure. Since they shouldn't
                    /// have subcomponents to promote, they
                    /// shouldn't have that BOM structure.
                    /// User intervention might be required.
                    ck = newFmTest2().AskAbout(invDoc, "The following Item has an unhandled BOM Structure:", "Skip it? (Click [NO] OR [CANCEL] to review)");
                    if (ck == Constants.vbYes)
                    {
                    }
                    else
                        System.Diagnostics.Debugger.Break();// and let User decide what to do with it.
                    System.Diagnostics.Debugger.Break(); // (extraneous; disable/remove whenever)
                }

                // the design tracking property set,
                // and update the Cost Center Property
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
                else if (Strings.Len(nmFamily) > 0)
                {
                    prFamily.Value = nmFamily;
                    if (Information.Err.Number)
                    {
                        Debug.Print("CHGFAIL[FAMILY]{'" + prFamily.Value + "' -> '" + nmFamily + "'}: " + invDoc.DisplayName + " (" + invDoc.FullDocumentName + ")");
                        if (MsgBox("Couldn't Change Family" + vbNewLine, Constants.vbYesNo | Constants.vbDefaultButton2, invDoc.DisplayName) == Constants.vbYes)
                            System.Diagnostics.Debugger.Break();
                    }
                    else
                    {
                    }

                    rt = dcAddProp(prFamily, rt);
                }
            }

            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            dcGeniusPropsPartRev20180530_ck = rt;
        }
    }
}