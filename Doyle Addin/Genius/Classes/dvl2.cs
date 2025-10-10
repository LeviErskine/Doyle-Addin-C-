class dvl2
{
    public Scripting.Dictionary dcGnsMatlSpecPairings()
    {
        /// dcGnsMatlSpecPairings -- Genius Raw Material Spec Relations
        /// Return a Dictionary of Dictionaries
        /// keyed to each Specification value
        /// found in ANY Spec field of any
        /// Raw Material Item, each listing
        /// all OTHER Spec values found in
        /// conjunction with each value.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Scripting.Dictionary dcVl;
        Scripting.Dictionary dcAl;
        Variant kyVl;
        Variant dxAl;
        string kyAl;

        rt = new Scripting.Dictionary();

        wk = dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_MatlSpecXref())));
        if (wk == null)
        {
        }
        else if (wk.Exists("val"))
        {
            dcVl = wk.Item("val");
            {
                var withBlock = dcVl;
                foreach (var kyVl in withBlock.Keys)
                    rt.Add(kyVl, new Scripting.Dictionary());

                foreach (var kyVl in withBlock.Keys)
                {
                    dcAl = rt.Item(kyVl);
                    {
                        var withBlock1 = dcOb(withBlock.Item(kyVl));
                        foreach (var dxAl in withBlock1.Keys)
                        {
                            kyAl = dcOb(withBlock1.Item(dxAl)).Item("also");
                            {
                                var withBlock2 = rt;
                                if (withBlock2.Exists(kyAl))
                                    dcAl.Add(kyAl, withBlock2.Item(kyAl));
                                else
                                    System.Diagnostics.Debugger.Break();// because something went wrong
                            }
                        }
                    }
                }
            }
        }
        else
        {
        }

        dcGnsMatlSpecPairings = rt;
    }

    public Scripting.Dictionary dcOfDcWithXrefsDep1st(Scripting.Dictionary dc, Scripting.Dictionary wk = null/* TODO Change to default(_) if this is not a reference type */, string pt = "#")
    {
        /// dcOfDcWithXrefsDep1st
        /// Replace rudundant / recursive Dictionary
        /// Objects in hierarchical Dictionary structure
        /// '
        /// This is a depth-first implementation, which
        /// might locate an initial Dictionary reference
        /// deep inside an early branch before finding a
        /// shallower instance that might be preferable.
        /// '
        /// A breadth-first implementation might be preferred.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary ck;
        Variant ar;
        Variant ky;
        string sp;

        if (wk == null)
            rt = dcOfDcWithXrefsDep1st(dc, new Scripting.Dictionary());
        else
        {
            rt = new Scripting.Dictionary();
            {
                var withBlock = dc;
                foreach (var ky in withBlock.Keys)
                {
                    ar = Array(withBlock.Item(ky));
                    ck = dcOb(obOf(ar(0)));

                    if (ck == null)
                    {
                    }
                    else if (wk.Exists(ck))
                        ar = Array(wk.Item(ck));
                    else
                    {
                        /// prep new $ref path
                        sp = pt + "/" + System.Convert.ToHexString(ky);

                        /// add new $ref to wk
                        {
                            var withBlock1 = nuDcPopulator().Setting("$ref", sp);
                            wk.Add(ck, withBlock1.Dictionary);
                        }

                        /// go ahead and process
                        /// subdictionary
                        ar = Array(dcOfDcWithXrefsDep1st(ck, wk, sp));
                    }

                    rt.Add(ky, ar(0));
                }
            }

            dcOfDcWithXrefsDep1st = rt;
        }

        dcOfDcWithXrefsDep1st = rt;
    }

    public Scripting.Dictionary dcOfDcWithXrefsBrd1st(Scripting.Dictionary dc, Scripting.Dictionary wk = null/* TODO Change to default(_) if this is not a reference type */, string pt = "#")
    {
        /// dcOfDcWithXrefsBrd1st
        /// Replace rudundant / recursive Dictionary
        /// Objects in hierarchical Dictionary structure
        /// '
        /// This is a depth-first implementation, which
        /// might locate an initial Dictionary reference
        /// deep inside an early branch before finding a
        /// shallower instance that might be preferable.
        /// '
        /// A breadth-first implementation might be preferred.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary ck;
        Scripting.Dictionary ls;
        Variant ar;
        Variant ky;
        string sp;

        if (wk == null)
            rt = dcOfDcWithXrefsBrd1st(dc, new Scripting.Dictionary());
        else
        {
            /// create returned Dictionary
            rt = new Scripting.Dictionary();

            /// create local working
            /// Dictionary of Dictionaries
            ls = new Scripting.Dictionary();

            /// being processing
            /// supplied Dictionary
            {
                var withBlock = dc;
                /// first pass: collect and process
                /// all sub Dictionary Objects
                foreach (var ky in withBlock.Keys)
                {
                    ck = dcOb(obOf(withBlock.Item(ky)));

                    if (!ck == null)
                    {
                        if (wk.Exists(ck))
                            /// add existing $ref Dictionary
                            /// to Dictionary list. thinking
                            /// recursion should NOT be an issue
                            ls.Add(ky, wk.Item(ck));
                        else
                        {
                            /// add new Dictionary to list
                            /// for subsequent recursion
                            ls.Add(ky, withBlock.Item(ky));

                            /// prep new $ref path
                            sp = pt + "/" + System.Convert.ToHexString(ky);

                            {
                                var withBlock1 = wk;
                                /// add new $ref Dictionary
                                withBlock1.Add(ck, new Scripting.Dictionary());

                                /// add path to Dictionary
                                dcOb(withBlock1.Item(ck)).Add("$ref", sp);
                            }
                        }
                    }
                }

                foreach (var ky in withBlock.Keys)
                {
                    if (ls.Exists(ky))
                        rt.Add(ky, dcOfDcWithXrefsBrd1st(ls.Item(ky), wk, pt + "/" + System.Convert.ToHexString(ky)));
                    else
                        rt.Add(ky, withBlock.Item(ky));
                }
            }
        }

        dcOfDcWithXrefsBrd1st = rt;
    }

    public Scripting.Dictionary dcGnsMatlSpecPairings4json()
    {
        /// dcGnsMatlSpecPairings4json -- check on dcGnsMatlSpecPairings
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant k2;

        rt = dcGnsMatlSpecPairings();

        {
            var withBlock = rt;
            foreach (var ky in withBlock.Keys())
            {
                // .Item(ky) = Join(dcOb(.Item(ky)).Keys)
                {
                    var withBlock1 = dcOb(withBlock.Item(ky));
                    foreach (var k2 in withBlock1.Keys())
                        withBlock1.Item(k2) = Join(dcOb(withBlock1.Item(k2)).Keys);
                }
            }
        }

        dcGnsMatlSpecPairings4json = rt;
    }

    public Scripting.Dictionary dcSpecSubsetWith(string txSpec, Scripting.Dictionary inDc)
    {
        Scripting.Dictionary rt;

        if (inDc.Exists(txSpec))
            dcSpecSubsetWith = dcKeysInCommon(inDc, dcOb(inDc.Item(txSpec)), 1);
        else
            dcSpecSubsetWith = new Scripting.Dictionary();
    }

    public Scripting.Dictionary dcSpecSubsetWithAll(Scripting.Dictionary dcSpec, Scripting.Dictionary inDc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = inDc;
        foreach (var ky in dcSpec.Keys)
            rt = dcSpecSubsetWith(System.Convert.ToHexString(ky), rt);
        dcSpecSubsetWithAll = rt;
    }

    public Scripting.Dictionary dcSpecSetFromUser()
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary dc;
        fmSelectorList fm;
        string nx;

        rt = new Scripting.Dictionary();
        dc = dcGnsMatlSpecPairings();
        // Debug.Print Join(dc.Keys)

        do
        {
            fm = nuSelFromDict(dc);
            nx = fm.GetReply(null/* Conversion error: Set to default value for this argument */, "");
            if (Strings.Len(nx) > 0)
            {
                rt.Add(nx, nx);
                dc = dcSpecSubsetWith(nx, dc);
                if (dc.Count == 0)
                    nx = "";
            }
        }
        while (Strings.Len(nx) > 0);

        // Stop
        dcSpecSetFromUser = rt;
    }

    public Scripting.Dictionary d2g3f1(Inventor.PartDocument Part, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// d2g3f1 -- Return Dictionary
        /// of relevant Part Properties
        /// and information for use in
        /// Genius data extraction
        /// 
        Scripting.Dictionary rt;
        string txPartNum;

        if (dc == null)
            rt = d2g3f1(Part, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = Part;
                {
                    var withBlock1 = withBlock.PropertySets.Item(gnDesign);
                    rt.Add(pnPartNum, withBlock1.Item(pnPartNum));
                    rt.Add(pnFamily, withBlock1.Item(pnFamily));
                }
            }

            rt = dcGnsInfoCompDef(aiCompDefOf(Part), rt);
        }

        d2g3f1 = rt;
    }

    public Scripting.Dictionary dcGnsInfoAiDocBase(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsInfoAiDocBase (formerly d2g3f1a)
        /// Return Dictionary of Document Properties
        /// and information relevant to Genius
        /// for data extraction
        /// 
        Scripting.Dictionary rt;
        string txPartNum;

        if (dc == null)
            rt = dcGnsInfoAiDocBase(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = AiDoc;
                {
                    var withBlock1 = withBlock.PropertySets.Item(gnDesign);
                    rt.Add(pnPartNum, withBlock1.Item(pnPartNum));
                    rt.Add(pnFamily, withBlock1.Item(pnFamily));
                }

                if (false)
                {
                    rt.Add("subType", withBlock.SubType); rt.Add("docType", withBlock.DocumentType);
                    rt.Add("dsbType", withBlock.DocumentSubType.DocumentSubTypeID);
                }
            }

            rt = dcGnsInfoCompDef(aiCompDefOf(AiDoc), rt);
        }

        dcGnsInfoAiDocBase = rt;
    }

    public Scripting.Dictionary dcGnsInfoCompDef(Inventor.ComponentDefinition CpDef, Scripting.Dictionary dcWkg = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsInfoCompDef -- Generate and/or populate
        /// Dictionary (new or supplied) with data for
        /// Genius from supplied ComponentDefinition.
        /// This is the "generic" variant, which dispatches
        /// the supplied ComponentDefinition to a variant
        /// more specific to its Class. (some Class
        /// variants remain to be implemented)
        /// Note that this function follows the convention
        /// of a recursive call with a new Dictionary
        /// object when none is supplied. Duplication
        /// of the basic function structure should ensure
        /// this pattern is followed by all specialized
        /// variants. While this should not usually be
        /// necessary under normal usage (dispatch to
        /// specialized variants from here), it should
        /// help accommodate the possibility of direct
        // '      calls from other client functions.
        /// 
        Scripting.Dictionary rt;

        rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDef(CpDef, new Scripting.Dictionary());
        else if (CpDef == null)
        {
        }
        else
        {
            {
                var withBlock = CpDef // .ComponentDefinition
;
                rt.Add("bomStr", withBlock.BOMStructure); if (withBlock.BOMStructure == kNormalBOMStructure)
                    rt.Add("Type", "M");
                if (withBlock.BOMStructure == kPurchasedBOMStructure)
                    rt.Add("Type", "R");
                {
                    var withBlock1 = nuAiBoxData().UsingBox(withBlock.RangeBox) // .SortingDims
       ;
                    {
                        var withBlock2 = withBlock1.UsingInches() // WARNING[2021.12.15]
      ;
                        /// Forcing inch conversion MAY lead
                        /// to issues in future development.
                        /// It is absolutely ESSENTIAL that
                        /// unit measurement be tracked and
                        /// kept consistent throughout the
                        /// entire management process.
                        rt.Add(pnLength, Round(withBlock2.SpanX, 6));
                        rt.Add(pnWidth, Round(withBlock2.SpanY, 6));
                        rt.Add("Height", Round(withBlock2.SpanZ, 6));
                    }
                }
            }

            if (CpDef is Inventor.SheetMetalComponentDefinition)
                rt = dcGnsInfoCompDefShtMtl(CpDef, rt);
            else if (CpDef is Inventor.WeldmentComponentDefinition)
            {
                System.Diagnostics.Debugger.Break(); // using general Assembly handler
                rt = dcGnsInfoCompDefAssy(CpDef, rt);
            }
            else if (CpDef is Inventor.WeldsComponentDefinition)
            {
                System.Diagnostics.Debugger.Break(); // using general Assembly handler
                rt = dcGnsInfoCompDefAssy(CpDef, rt);
            }
            else if (CpDef is Inventor.PartComponentDefinition)
                rt = dcGnsInfoCompDefPart(CpDef, rt);
            else if (CpDef is Inventor.AssemblyComponentDefinition)
                rt = dcGnsInfoCompDefAssy(CpDef, rt);
            else
            {
            }
        }

        dcGnsInfoCompDef = rt;
    }

    public Scripting.Dictionary dcGnsInfoCompDefShtMtl(Inventor.SheetMetalComponentDefinition CpDef, Scripting.Dictionary dcWkg = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsInfoCompDefShtMtl -- Generate and/or populate Dictionary
        /// (new or supplied) with data for Genius
        /// from supplied ComponentDefinition.
        /// This is the Assembly variant.
        /// '
        /// 
        Scripting.Dictionary rt;
        double wd; // width
        double lg; // length
        double ht; // height
        double tk; // thickness
        double ar; // area
        double ck; // check height vs thickness

        Scripting.Dictionary rm;
        string s6;
        Variant ky;

        rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefShtMtl(CpDef, new Scripting.Dictionary());
        else
        {
            rt = dcGnsInfoCompDefPart(CpDef, rt);
            {
                var withBlock = rt;
                if (!withBlock.Exists("SPEC06"))
                {
                    System.Diagnostics.Debugger.Break();
                    withBlock.Add("SPEC06", steelSpec6("", 1));
                }
                s6 = withBlock.Item("SPEC06");

                if (!withBlock.Exists("RMLIST"))
                    withBlock.Add("RMLIST", new Scripting.Dictionary());
                rm = dcOb(withBlock.Item("RMLIST"));
            }

            {
                var withBlock = CpDef;
                tk = withBlock.Thickness.Value / (double)cvLenIn2cm;
                // NOTE conversion to Inches from Centimeters.
                // keep in mind we're grabbing Thickness HERE
                // and will use Height (below) in an effort
                // to validate the Flat Pattern, and determine
                // this Part is MEANT to be Sheet Metal.

                if (withBlock.HasFlatPattern)
                {
                    {
                        var withBlock1 = withBlock.FlatPattern;
                        {
                            var withBlock2 = nuAiBoxData().UsingBox(withBlock1.RangeBox);
                            {
                                var withBlock3 = withBlock2.UsingInches();
                                ht = Round(withBlock3.SpanZ, 6);
                                // remember, Height here is meant
                                // to verify Sheet Metal Part

                                lg = Round(withBlock3.SpanX, 6);
                                wd = Round(withBlock3.SpanY, 6);
                            }
                        }
                    }

                    ck = Round(Abs(ht - tk), 6);
                    if (ck > 0.002)
                    {
                        {
                            var withBlock1 = dcFlatPatSpansByVertices(withBlock.FlatPattern);
                            if (ht > withBlock1.Item("Z"))
                            {
                                ht = withBlock1.Item("Z");

                                Debug.Print.Item("X"); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                Debug.Print.Item("Y"); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }
                        }
                    }
                    else
                    {
                    }

                    if (Round(Abs(ht - tk), 6) > 0)
                    {
                    }

                    ar = lg * wd; // .SpanX * .SpanY
                }
                else
                {
                    {
                        var withBlock1 = rt;
                        if (withBlock1.Exists(pnLength))
                            lg = withBlock1.Item(pnLength);

                        if (withBlock1.Exists(pnWidth))
                            wd = withBlock1.Item(pnWidth);

                        if (withBlock1.Exists("Height"))
                            ht = withBlock1.Item("Height");
                    }

                    ar = 0; // STOPGAP[2021.12.08]
                }
                /// At this point, we should have either
                /// likely dimensions of the flat pattern, OR
                /// the original dimensions of the part itself.
                /// 
                /// The next step is to determine whether they
                /// are consistent with a valid sheet metal part.
                /// If not, it's likely a structural one.
                /// 
                /// The key criterion is how closely the height
                /// dimension matches the given thickness.
                /// 
                ck = Round(Abs(ht - tk), 6);
                if (ck > 0.002)
                {
                }
                else
                {
                }

                /// REV[2021.12.15]:
                /// add material option collection
                /// specific to sheet metal
                {
                    var withBlock1 = dcGnsMatlOps(dcCtOfEach(Array(tk, lg, wd, ht)), s6);
                    foreach (var ky in withBlock1.Keys)
                    {
                        if (!rm.Exists(ky))
                            rm.Add(ky, withBlock1.Item(ky));
                    }
                }

                /// 
                /// 
                {
                    var withBlock1 = rt;
                    // first, remove any previous
                    // dimensional values
                    if (withBlock1.Exists(pnLength))
                        withBlock1.Remove(pnLength);
                    if (withBlock1.Exists(pnWidth))
                        withBlock1.Remove(pnWidth);
                    if (withBlock1.Exists("Height"))
                        withBlock1.Remove("Height");
                    // (not sure this is the best way
                    // but going to try it for now)

                    withBlock1.Add(pnThickness, tk);
                    withBlock1.Add(pnLength, lg);
                    withBlock1.Add(pnWidth, wd);
                    withBlock1.Add(pnArea, ar);
                    withBlock1.Add("Height", ht);
                }
            }
        }

        dcGnsInfoCompDefShtMtl = rt;
    }

    public Scripting.Dictionary dcFlatPatSpansByVertices(Inventor.FlatPattern smFlat)
    {
        /// dcFlatPatSpansByVertices -- get extents of
        /// Sheet Metal Flat Pattern
        /// from a scan of its Vertices.
        /// this is a last resort,
        /// in case an erroneous Z span
        /// reported from the Range Box
        /// fails to match Thickness.
        /// 
        Scripting.Dictionary rt;
        Inventor.Vertex vx;
        double xmn;
        double xmx;
        double ymn;
        double ymx;
        double zmn;
        double zmx;

        rt = new Scripting.Dictionary();

        if (!smFlat.Body == null)
        {
            {
                var withBlock = smFlat.Body // .Vertices'.RangeBox
;
                foreach (var vx in withBlock.Vertices)
                {
                    {
                        var withBlock1 = vx.Point;
                        if (withBlock1.X < xmn)
                            xmn = withBlock1.X;
                        if (withBlock1.X > xmx)
                            xmx = withBlock1.X;
                        if (withBlock1.Y < ymn)
                            ymn = withBlock1.Y;
                        if (withBlock1.Y > ymx)
                            ymx = withBlock1.Y;
                        if (withBlock1.Z < zmn)
                            zmn = withBlock1.Z;
                        if (withBlock1.Z > zmx)
                            zmx = withBlock1.Z;
                    }
                }
            }
        }

        {
            var withBlock = rt;
            withBlock.Add("X", xmx - xmn);
            withBlock.Add("Y", ymx - ymn);
            withBlock.Add("Z", zmx - zmn);
        }

        dcFlatPatSpansByVertices = rt;
    }

    public Scripting.Dictionary dcGnsInfoCompDefPart(Inventor.PartComponentDefinition CpDef, Scripting.Dictionary dcWkg = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsInfoCompDefPart -- Generate and/or populate Dictionary
        /// (new or supplied) with data for Genius
        /// from supplied ComponentDefinition.
        /// This is the general Part variant.
        /// It's grown somewhat in complexity
        /// since development's begun.
        /// Here is rough flow map:
        /// (inside component definition)
        /// - stage supplied Dictionary for return
        /// (start a new one, if none supplied
        /// one USUALLY should be)
        /// - get Mass -- don't add to Dictionary yet
        /// (no data should be added to Dictionary
        /// until all data are collected and verified)
        /// - get Active Material, and its name
        /// - use this to set target Spec 6
        /// (inside returning Dictionary)
        /// - collect length, width, and height
        /// dimensions from Dictionary
        /// (this is why one should be supplied)
        /// - collect raw material candidate items
        /// from Genius into a Recordset, using
        /// an SQL query generated from collected
        /// dimensions, and material Spec 6
        /// - generate Dictionary of candidates
        /// from the Recordset, keyed on item names
        /// - add data to Dictionary:
        /// - mass
        /// - material name
        /// - spec 6
        /// - Dictionary of raw material
        /// item candidates
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk; // may be temporary
        Scripting.Dictionary d2; // may be temporary
        Inventor.MaterialAsset mt;
        string mtName;
        string s6;
        double ms; // mass
        Variant ky;
        double ck;

        rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefPart(CpDef, new Scripting.Dictionary());
        else
        {
            {
                var withBlock = CpDef;
                {
                    var withBlock1 = withBlock.MassProperties;
                    ms = Round(withBlock1.Mass * cvMassKg2LbM, 4);
                }

                // ptNumShtMetal
                mt = aiDocPart(withBlock.Document).ActiveMaterial;
                if (mt == null)
                    mtName = "";
                else
                    mtName = mt.DisplayName;
                s6 = steelSpec6(mtName);
            }

            {
                var withBlock = rt;
                wk = new Scripting.Dictionary();
                foreach (var ky in Array(pnLength, pnWidth, "Height"))
                {
                    if (withBlock.Exists(ky))
                        wk.Add(ky, Round(System.Convert.ToDouble(withBlock.Item(ky)), 6));
                }
                wk = dcCtOfEach(wk.Items);
                if (wk.Count == 0)
                    wk.Add(0.075, 1);
                // another kludge to trap an error
                // which should NOT occur as long as
                // a prepared Dictionary is supplied.

                /// Here is where we'll attempt to collect
                /// raw material Item candidates from Genius
                /// 

                // present setup -- plan to change
                // "select d.v "
                // Debug.Print "from (values (" & txDumpLs(wk.Keys, "), (") & "))"
                // " as d(v)"

                // future proposal -- counts occurrences
                // "select d.v, d.c "
                // Debug.Print "from (values (" & dumpLsKeyVal(wk, ", ", "), (") & "))"
                // " as d(v, c)"

                wk = dcGnsMatlOps(wk, s6);
                /// REV[2021.12.15]:
                /// preceding line replaces With block below,
                /// moving Genius material options request
                /// to function dcGnsMatlOps, so it can be
                /// called from other functions, like,
                /// again, dcGnsInfoCompDefShtMtl
                // With cnGnsDoyle()
                // Dim rs As ADODB.Recordset
                // 
                // 'wk.RemoveAll
                // 
                // 
                // Err.Clear
                // rs = .Execute('    sqlOf_GnsMatlOptions('        s6, wk.Keys'    ))
                // 
                // If Err.Number = 0 Then
                // With dcFromAdoRS(rs, "")  wk =
                // For Each ky In .Keys
                // d2 = dcOb(.Item(ky))
                // If d2 Is Nothing Then
                // Stop
                // Else
                // wk.Add d2.Item("Item"), d2
                // End If
                // 'Stop
                // ''' ENDOFDAY[2021.12.08]:
                // '''     Need to setup process of remapping
                // '''     raw material Items from Genius
                // '''     to their Item names
                // Next: End With
                // 
                // rs.Close
                // Else
                // Stop
                // Err.Clear
                // End If
                // 
                // 
                // .Close
                // End With

                // .Add pnRawMaterial, wk

                withBlock.Add(pnMass, ms);
                withBlock.Add(pnMaterial, mtName);
                withBlock.Add("SPEC06", s6);

                // If False Then
                // not quite ready for this one yet
                withBlock.Add("RMLIST", wk);
                // End If
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }
        }

        dcGnsInfoCompDefPart = rt;
    }

    public Scripting.Dictionary dcGnsInfoCompDefAssy(Inventor.AssemblyComponentDefinition CpDef, Scripting.Dictionary dcWkg = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsInfoCompDefAssy -- Generate and/or populate Dictionary
        /// (new or supplied) with data for Genius
        /// from supplied ComponentDefinition.
        /// This is the general Assembly variant.
        /// '
        /// 
        Scripting.Dictionary rt;

        rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefAssy(CpDef, new Scripting.Dictionary());
        else
        {
            var withBlock = CpDef;
            {
                var withBlock1 = withBlock.MassProperties;
                rt.Add(pnMass, Round(withBlock1.Mass * cvMassKg2LbM, 4));
            }
        }

        dcGnsInfoCompDefAssy = rt;
    }

    public Scripting.Dictionary dcGnsInfoCompDefTBD(Inventor.ComponentDefinition CpDef, Scripting.Dictionary dcWkg = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGnsInfoCompDefTBD (formerly d2g4f2zz)
        /// Generate and/or populate Dictionary
        /// (new or supplied) with data for Genius
        /// from supplied ComponentDefinition.
        /// This is the <TBD> variant. (formerly <zz>)
        /// Use it as a template for others.
        /// (be sure to modify comments accordingly)
        /// 
        Scripting.Dictionary rt;

        rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefTBD(CpDef, new Scripting.Dictionary());
        else
        {
            var withBlock = CpDef;
        }

        dcGnsInfoCompDefTBD = rt;
    }

    public Scripting.Dictionary dcGnsInfoSQLitem(string Item)
    {
        /// dcGnsInfoSQLitem -- Return a Dictionary
        /// of Part data from Genius
        /// for the indicated Item
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary mt;
        ADODB.Recordset rs;
        Variant ky;

        {
            var withBlock = cnGnsDoyle();
            Information.Err.Clear();
            rs = withBlock.Execute(sqlOf_GnsPartInfo(Item)); // sqlOf_ASDF
            if (Information.Err.Number == 0)
                rt = dcFromAdoRSrow(rs, "");
            else
            {
                Debug.Print(Information.Err.Number);
                Debug.Print(Information.Err.Description);
                System.Diagnostics.Debugger.Break();
            }

            rs = withBlock.Execute(sqlOf_GnsPartMatl(Item)); // sqlOf_ASDF
            if (Information.Err.Number == 0)
            {
                mt = dcFromAdoRS(rs, "");
                {
                    var withBlock1 = mt;
                    if (withBlock1.Count > 0)
                    {
                        if (withBlock1.Count > 1)
                        {
                            System.Diagnostics.Debugger.Break(); // to handle multiple raw materials
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }
                        else
                        {
                            var withBlock2 = dcOb(withBlock1.Item(withBlock1.Keys(0)));
                            foreach (var ky in withBlock2.Keys)
                            {
                                if (rt.Exists(ky))
                                    System.Diagnostics.Debugger.Break(); // to deal with collision
                                else
                                    rt.Add(ky, withBlock2.Item(ky));
                            }
                        }
                    }
                }
            }
            else
            {
                Debug.Print(Information.Err.Number);
                Debug.Print(Information.Err.Description);
                System.Diagnostics.Debugger.Break();
            }


            withBlock.Close();
        }

        dcGnsInfoSQLitem = rt;
    }

    public Scripting.Dictionary d2g3f4(Inventor.PartDocument Part, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// d2g3f4 -- Return a Dictionary
        /// of Properties and info from
        /// Inventor Part Document for
        /// Genius Interface.
        /// 
        Scripting.Dictionary rt;
        // Dim rs As ADODB.Recordset

        rt = dcProps4genius(Part, d2g3f1(Part, dc), 0);

        d2g3f4 = rt;
    }

    public Scripting.Dictionary d2g3f5(Inventor.Document AiDoc)
    {
        /// d2g3f5 -- Gather Dictionaries of Inventor
        /// Properties and Genius info from supplied
        /// Document for correlation and potential
        /// revision.
        /// REV[2021.12.15]:
        /// Parameter Part renamed to AiDoc, with Class
        /// changed from PartDocument to the more general
        /// Document, as it would appear all supporting
        /// functions will accept and work with it.
        /// 
        Scripting.Dictionary dcPt;
        // base Genius info and inherent Properties
        Scripting.Dictionary dcPr;
        // base + custom Genius Properties
        Scripting.Dictionary dcVlAi;
        // values of all collected Properties
        Scripting.Dictionary dcVlPr;
        // values of all collected Properties
        Scripting.Dictionary dcGn;
        // information from Genius database
        Scripting.Dictionary rt;
        Variant ky;

        dcPt = dcGnsInfoAiDocBase(AiDoc);
        dcVlAi = dcMapAiProps2vals(dcPt);
        /// REV[2021.12.16]:
        /// additional value Dictionary
        /// collects values ONLY from
        /// the inherent document data
        /// and Properties, which gets
        /// overridden at the next step:

        dcPr = dcProps4genius(AiDoc, dcCopy(dcPt), 2);
        /// REV[2021.12.15] argument 2 replaces 0 in order
        /// to generate references to missing Properties
        /// (see dcGnsPropsListed) without trying to
        /// create them. That way, client functions may
        /// be made aware of Properties that need created.
        /// 
        /// Modifications to those functions might be needed
        /// to be prepared for missing Properties, whose
        /// names will map to Nothing (void references)
        /// UPDATE[2021.12.16]:
        /// just happened today. created new function
        /// blankIfNoValElseSelf to address this issue.
        /// see dcMapAiProps2vals for application
        /// 
        // = dcProps4genius(AiDoc, d2g3f1(AiDoc, dc), 0)
        // = d2g3f4(AiDoc)

        dcVlPr = dcMapAiProps2vals(dcPr); // dcPt
        {
            var withBlock = dcVlPr;
            foreach (var ky in Array(pnThickness, pnWidth, pnLength, pnArea, pnRmQty)) // '' THIS IS A KLUDGE
            {
                // to temporarily "fix" an issue with Width,
                // Length, and Area values that don't match
                // up between Inventor Properties and Genius,
                // even when their numeric values are equal.
                // 
                // A more thorough review and revision
                // will probably be needed eventually.
                // 
                // REV[2021.12.07]:
                // Thickness is also affected
                // and has been added to the list.
                // 
                /// REV[2021.12.16]:
                /// added pnRmQty to list as quick and
                /// dirty method to force a blank value
                /// to zero, and prevent an error in
                /// the correction code after the loop.
                /// 
                /// This must be the sort of cruft Joel
                /// Spolsky was talking about in that
                /// essay of his. Still not a justified
                /// defense for crud programming, which
                /// this is, let's face it! Right there
                /// with you, Brando!
                /// 
                if (withBlock.Exists(ky))
                    withBlock.Item(ky) = Val(Split("0" + System.Convert.ToHexString(withBlock.Item(ky)), " ")(0));
            }

            if (withBlock.Exists(pnRmQty))
                withBlock.Item(pnRmQty) = Round(withBlock.Item(pnRmQty), 8);
        }

        dcGn = dcGnsInfoSQLitem(dcVlPr.Item(pnPartNum));

        rt = new Scripting.Dictionary();
        {
            var withBlock = rt;
            withBlock.Add("aiVal", dcVlAi);
            withBlock.Add("inv", dcVlPr);
            withBlock.Add("gns", dcGn);
            withBlock.Add("prp", dcPr);
        }

        d2g3f5 = rt;
    }

    public Scripting.Dictionary d2g3f5as(Inventor.AssemblyDocument Assy, long ThisToo = 0)
    {
        /// d2g3f5as -- Assembly counterpart to d2g3f5
        /// not sure what's actually to be done with it yet.
        /// probably just remove it; d2g3f5 can handle both.
        /// 
        Scripting.Dictionary dc;

        dc = dcRemapByPtNum(dcAiDocComponents(Assy, null/* Conversion error: Set to default value for this argument */, ThisToo));
    }

    public Variant dcMapAiProps2vals(Scripting.Dictionary dc, long Flags = 0)
    {
        /// dcMapAiProps2vals --
        /// Return a Dictionary
        /// containing the Values of
        /// any Inventor Properties
        /// in supplied Dictionary,
        /// with all other members
        /// returned as they are.
        /// 
        /// related functions:
        /// dcOfDcAiPropVals
        /// dcAiPropValsFromDc
        /// dcOfPropsInAiDoc
        /// 
        Scripting.Dictionary rt;
        Inventor.Property pr;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, blankIfNoValElseSelf(valIfAiPropElseSelf(withBlock.Item(ky))));
        }
        dcMapAiProps2vals = rt;
    }

    public Variant valIfAiPropElseSelf(Variant vl)
    {
        /// valIfAiPropElseSelf --
        /// Return the Value of any
        /// supplied Inventor Property.
        /// Any other type of argument
        /// should be returned directly.
        /// 
        Inventor.Property pr;

        if (IsObject(vl))
        {
            pr = aiProperty(obOf(vl));
            if (pr == null)
                valIfAiPropElseSelf = vl;
            else
                valIfAiPropElseSelf = pr.Value;
        }
        else
            valIfAiPropElseSelf = vl;
    }

    public Variant blankIfNoValElseSelf(Variant vl)
    {
        /// blankIfNoValElseSelf --
        /// Return the Value of any
        /// supplied Inventor Property.
        /// Any other type of argument
        /// should be returned directly.
        /// 
        Inventor.Property pr;

        if (IsObject(vl))
        {
            if (obOf(vl) == null)
                blankIfNoValElseSelf = "";
            else
                blankIfNoValElseSelf = vl;
        }
        else if (IsNull(vl))
            blankIfNoValElseSelf = "";
        else if (IsEmpty(vl))
            blankIfNoValElseSelf = "";
        else
            blankIfNoValElseSelf = vl;
    }

    public Scripting.Dictionary d2g3f7(Inventor.Document AiDoc)
    {
        /// d2g3f7 --
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        // rt = New Scripting.Dictionary
        {
            var withBlock = d2g3f5(AiDoc);
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            rt = dcTreeReKeyedInPlc("src1", "gns", dcTreeReKeyedInPlc("src0", "inv", dcWBQbyCmpResult(dcCmpTextOf2dc(withBlock.Item("inv"), withBlock.Item("gns")))));

            {
                var withBlock1 = rt;
            }
            rt.Add("prp", withBlock.Item("prp"));
            rt.Add("doc", AiDoc);
        }
        d2g3f7 = rt;
    }

    public Scripting.Dictionary d2g3f8(Inventor.Document AiDoc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// d2g3f8 --
        /// 
        Scripting.Dictionary rt;
        // Dim ck As Inventor.Document
        Variant ky;

        if (AiDoc == null)
        {
            {
                var withBlock = ThisApplication;
                if (withBlock.ActiveDocument == null)
                    System.Diagnostics.Debugger.Break();
                else
                    rt = d2g3f8(withBlock.ActiveDocument);
            }
        }
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(dcAiDocComponents(AiDoc));
                /// 
                {
                    var withBlock1 = withBlock.dcIn() // Parts
       ;
                    foreach (var ky in withBlock1.Keys)
                        rt.Add(ky, d2g3f7(aiDocPart(obOf(withBlock1.Item(ky)))));
                }

                {
                    var withBlock1 = withBlock.dcOut() // Assemblies
       ;
                }
            }
        }

        d2g3f8 = rt;
    }

    public Scripting.Dictionary dcTreeMembersWithKey(Variant tg, Scripting.Dictionary dc, Scripting.Dictionary wk = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcTreeMembersWithKey (formerly d2g5f1)
        /// Given a Dictionary that might contain
        /// other Dictionaries, check it and any
        /// sub Dictionaries for target key (tg)
        /// and return a Dictionary of those
        /// Dictionaries containing it, each
        /// keyed to the number already found.
        /// This should ensure a unique key
        /// for each match found, with no
        /// need to track any other keys.
        /// 
        /// The ultimate goal of this function is to
        /// support a Key Find/Replace operation
        /// across a hierarchy of Dictionaries.
        /// 
        /// This is initially and specifically to map
        /// comparison keys "src0" and "src1" to
        /// the names of sources they represent.
        /// 
        /// This is of course the 'Find' component
        /// of the ultimate product
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        if (wk == null)
            rt = dcTreeMembersWithKey(tg, dc, new Scripting.Dictionary());
        else
        {
            rt = wk;
            if (!dc == null)
            {
                {
                    var withBlock = dc;
                    if (withBlock.Exists(tg))
                    {
                        {
                            var withBlock1 = rt;
                            withBlock1.Add.Count(null/* Conversion error: Set to default value for this argument */, dc);
                        }
                    }

                    foreach (var ky in withBlock.Keys)
                        rt = dcTreeMembersWithKey(tg, dcOb(obOf(withBlock.Item(ky))), rt);
                }
            }
        }
        dcTreeMembersWithKey = rt;
    }

    public Scripting.Dictionary dcTreeMemWithReplcmt(Variant rp, Scripting.Dictionary dc)
    {
        /// dcTreeMemWithReplcmt (formerly d2g5f2)
        /// Given a Dictionary of Dictionaries,
        /// check for any Dictionary containing target
        /// replacement Key rp, and return a Dictionary
        /// containing any results.
        /// 
        /// This is a check for potential Key collisions.
        /// The Dictionary returned should be empty.
        /// 
        /// This is presently accomplished by first calling
        /// dcTreeMembersWithKey against the supplied Dictionary,
        /// which is normally expected to be the result
        /// of a PRIOR call to dcTreeMembersWithKey using the target
        /// key to be replaced.
        /// 
        /// It is therefore possible that the supplied
        /// Dictionary might contain replacement key rp,
        /// and thus be included in the local result.
        /// That Dictionary should NOT be included
        /// in the FINAL result returned.
        /// 
        /// It is therefore necessary to scan the result
        /// of the local dcTreeMembersWithKey call, and remove it,
        /// if found.
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = dcTreeMembersWithKey(rp, dc);
        {
            var withBlock = rt;
            foreach (var ky in withBlock.Keys)
            {
                if (dcOb(obOf(withBlock.Item(ky))) == dc)
                {
                    System.Diagnostics.Debugger.Break();
                    withBlock.Remove(ky);
                }
            }
        }
        dcTreeMemWithReplcmt = rt;
    }

    public Scripting.Dictionary dcTreeReKeyedInPlc(Variant tg, Variant rp, Scripting.Dictionary dc)
    {
        /// dcTreeReKeyedInPlc (formerly d2g5f3)
        /// Given a target Key tg (to be replaced),
        /// a replacement Key rp, and a Dictionary
        /// that includes other Dictionaries,
        /// attempt to replace all instances of the
        /// target Key with the replacement Key in
        /// all Dictionaries within the hierarchy.
        /// 
        /// Note that this is a DESTRUCTIVE replacement
        /// operation. A preferable option might be
        /// to generate a NEW hierarchical Dictionary
        /// replicating the original, with the desired
        /// key substitution. Will consider that for
        /// a later implementation.
        /// 
        /// Note also that error checking/handling
        /// in this implementation is presently minimal.
        /// A more robust process should also be considered.
        /// 
        Scripting.Dictionary wk;
        Scripting.Dictionary ck;
        Variant ky;

        wk = dcTreeMembersWithKey(tg, dc);
        ck = dcTreeMemWithReplcmt(rp, wk);

        if (ck.Count > 0)
            System.Diagnostics.Debugger.Break();
        else
        {
            var withBlock = wk;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcOb(obOf(withBlock.Item(ky)));
                    /// A Dictionary object is assumed, here.
                    /// Though typically risky in a With block,
                    /// it SHOULD be guaranteed here,
                    /// so no error should occur.
                    /// Don't be surprised if it does, though.
                    if (withBlock1.Exists(rp))
                        System.Diagnostics.Debugger.Break(); // because this
                    else
                    {
                        // a proper error handler might
                        // be desired here in future
                        // 
                        // for now, keep disabled

                        // note order of operations here
                        withBlock1.Add(rp, withBlock1.Item(tg));
                        withBlock1.Remove(tg);
                    }
                }
            }
        }

        dcTreeReKeyedInPlc = dc;
    }

    public Variant userChoiceFromDc(object dcAs = Scripting.Dictionary == null/* TODO Change to default(_) if this is not a reference type */, Variant ifNone = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// userChoiceFromDc (formerly d2g3f2)
        /// Request User Selection from
        /// a Dictionary of options.
        /// 
        /// A list of Dictionary Keys is
        /// presented to the user. After
        /// User selects a Key, matching
        /// Item is returned for use.
        /// 
        VbMsgBoxResult ck;
        string msNoSel;
        Variant rp;
        Variant rt;

        /// REV[2023.05.17.1304]
        /// add ifNone processing to present
        /// User with information on default
        /// option(s), if supplied

        Information.Err.Clear();
        msNoSel = System.Convert.ToHexString(ifNone);

        if (Information.Err.Number == 0)
        {
            if (Strings.Len(msNoSel) > 0)
                msNoSel = "Use default value (" + msNoSel + ")?";
        }
        else
        {
            msNoSel = "";
            Information.Err.Clear();

            if (IsObject(ifNone))
            {
                if (!ifNone == null)
                    msNoSel = Join(Array("Use default " + TypeName(ifNone) + " Object?", "(Object details not available)"), Constants.vbNewLine);
            }
            else
                System.Diagnostics.Debugger.Break();
        }


        if (Strings.Len(msNoSel) > 0)
            msNoSel = Constants.vbNewLine + msNoSel;
        // msNoSel = Join(Array("User selection was requested", "with no available options!", msNoSel), vbNewLine)

        if (dc == null)
            rt = Array(userChoiceFromDc(dcAiDocsVisible()));
        else if (dc.Count > 0)
        {
            rp = nuSelFromDict(dc).GetReply();
            // , , , , ,, Join(Array("No option selected!",msNoSel), vbNewLine)
            if (dc.Exists(rp))
                rt = Array(dc.Item(rp));
            else
                rt = Array(ifNone);
        }
        else
        {
            ck = MsgBox(Join(Array("User selection was requested", "with no available options!"), Constants.vbNewLine), Constants.vbOKOnly, "No Options!");
            /// IIf(Len(msNoSel) > 0,vbYesNo, vbOKOnly),msNoSel
            /// 
            if (ck == Constants.vbNo)
                rt = Array(null/* TODO Change to default(_) if this is not a reference type */);
            else
                rt = Array(ifNone);
        }

        if (IsObject(rt(0)))
            userChoiceFromDc = rt(0);
        else
            userChoiceFromDc = rt(0);
    }

    public Scripting.Dictionary dcGnsPrpPtDvl_2021_1112(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary dc01;
        Scripting.Dictionary dcVlGn;
        ADODB.Recordset rs;
        // '
        string aiPartNum;
        string aiFamily;
        string aiSubType;
        // '
        // Dim aiPropsUser As Inventor.PropertySet
        Inventor.PropertySet aiPropsDesign;
        // '
        Inventor.Property prPartNum;
        Inventor.Property prFamily;
        // '
        // Dim aiPartNum   As String 'will be same as gnPartNum
        // Dim aiPartFam   As String
        // Dim aiMatlNum   As String
        // Dim aiMatlFam   As String
        // Dim aiMatlQty   As Double
        // Dim aiQtyUnit   As String
        Inventor.BOMStructureEnum aiBomType;
        // '
        // '
        rt = new Scripting.Dictionary();
        dc01 = new Scripting.Dictionary();

        {
            var withBlock = invDoc;
            // Property Sets
            {
                var withBlock1 = withBlock.PropertySets;
                // aiPropsUser = .Item(gnCustom)
                aiPropsDesign = withBlock1.Item(gnDesign);
            }
            aiBomType = withBlock.ComponentDefinition.BOMStructure;
            aiSubType = withBlock.SubType;
        }

        // Part Number and Family
        // Properties from Design set
        {
            var withBlock = aiPropsDesign;
            prPartNum = withBlock.Item(pnPartNum);
            prFamily = withBlock.Item(pnFamily);
        }

        // Values of Part Number
        // and Family Properties
        aiPartNum = prPartNum.Value;
        aiFamily = prFamily.Value;
        /// NOTE[2021.11.12]
        /// The preceding three sections
        /// can PROBABLY be consolidated
        /// into one, using fewer variables
        /// and probably just one With block

        // dc01
        {
            var withBlock = cnGnsDoyle();
            Information.Err.Clear();
            rs = withBlock.Execute(sqlOf_ASDF(aiPartNum)); // 
            if (Information.Err.Number == 0)
            {
                dcVlGn = dcFromAdoRSrow(rs, "");
                {
                    var withBlock1 = dcVlGn;
                }

                {
                    var withBlock1 = rs;
                    if (withBlock1.BOF & withBlock1.EOF)
                    {
                    }
                    else
                    {
                        {
                            var withBlock2 = withBlock1.Fields;
                        }

                        withBlock1.MoveNext();
                        if (!withBlock1.EOF)
                        {
                            System.Diagnostics.Debugger.Break(); // to handle multiple raw materials
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }
                    }

                    withBlock1.Close();
                }
            }
            else
            {
                Debug.Print(Information.Err.Number);
                Debug.Print(Information.Err.Description);
                System.Diagnostics.Debugger.Break();
            }

            withBlock.Close();
        }

        if (aiBomType == kNormalBOMStructure)
        {
            if (aiSubType == guidSheetMetal)
            {
            }
        }
        else if (aiBomType == kPurchasedBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kPhantomBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kInseparableBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kReferenceBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kNormalBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kPhantomBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kDefaultBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kVariesBOMStructure)
            System.Diagnostics.Debugger.Break();
        else if (aiBomType == kDefaultBOMStructure)
            System.Diagnostics.Debugger.Break();
        else
            System.Diagnostics.Debugger.Break();

        // '
        dcGnsPrpPtDvl_2021_1112 = rt;
    }

    public Scripting.Dictionary dcGeniusPropsPartRev20180530_broken2(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        // Dim dcPr As Scripting.Dictionary
        Scripting.Dictionary dcVlGn;
        Scripting.Dictionary dcProp;
        Scripting.Dictionary dcVlPr;
        Scripting.Dictionary dcVlAi;
        Scripting.Dictionary dcVlFP;
        Inventor.Property pr;
        Variant ky;
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
        /// UPDATE[2021.11.08] MAJOR CHANGE
        /// Overhaul variable names to better
        /// reflect TWO distinct value sets
        /// one from Genius
        /// another from Inventor
        /// in order to better compare
        /// and synchronize them.
        /// 
        /// First set are the Genius variables:
        /// the original set, renamed en masse:
        /// 
        string gnPartNum; // was pnModel
        string gnPartFam; // was ptFamily
        string gnMatlNum; // was pnStock
        string gnMatlFam; // was mtFamily
        double gnMatlQty; // was qtRawMatl
        string gnQtyUnit; // was qtUnit
        Inventor.BOMStructureEnum gnBomType;
        /// Second set are the new Inventor variables.
        /// These should replace the Genius instances
        /// anywhere their values are taken
        /// from the model.
        /// 
        string aiPartNum; // will be same as gnPartNum
        string aiPartFam;
        string aiMatlNum;
        string aiMatlFam;
        double aiMatlQty;
        string aiQtyUnit;
        Inventor.BOMStructureEnum aiBomType;
        /// 
        // '
        VbMsgBoxResult ck;
        aiBoxData bd;
        /// UPDATE[2021.11.03]:
        /// 
        /// 
        ADODB.Recordset rs;
        ADODB.Field fdItem;
        ADODB.Field fdFamly;
        ADODB.Field fdOrder;
        ADODB.Field fdMatrl;
        ADODB.Field fdMtFam;
        ADODB.Field fdQty;
        ADODB.Field fdUnit;

        if (dc == null)
            dcGeniusPropsPartRev20180530_broken2 = dcGeniusPropsPartRev20180530_broken2(invDoc, new Scripting.Dictionary());
        else
        {
            aiBomType = invDoc.ComponentDefinition.BOMStructure;
            /// UPDATE[2021.11.11]
            /// Moved Property  collection
            /// to top of program to permit
            /// collection of Design Properties
            /// in second step. Also pulled up
            /// BOM Structure capture (above)
            /// along with the Values of
            /// of Design Properties.

            {
                var withBlock = invDoc;
                {
                    var withBlock1 = withBlock.PropertySets;
                    aiPropsDesign = withBlock1.Item(gnDesign);
                    aiPropsUser = withBlock1.Item(gnCustom);
                }

                aiBomType = withBlock.ComponentDefinition.BOMStructure;

                if (aiBomType == kNormalBOMStructure)
                {
                    if (withBlock.SubType == guidSheetMetal)
                    {
                    }
                }
            }

            // Part Number and Family properties
            // are from Design, NOT Custom set
            {
                var withBlock = aiPropsDesign // we know they're present
       ;
                // so we can grab them directly
                prPartNum = withBlock.Item(pnPartNum);
                prFamily = withBlock.Item(pnFamily);
            }
            aiPartNum = prPartNum.Value;
            aiPartFam = prFamily.Value;

            dcProp = dcGnsPropsPart(invDoc, null/* Conversion error: Set to default value for this argument */, 0); // dcAiPropsInSet
            dcVlPr = new Scripting.Dictionary();
            {
                var withBlock = dcProp;
                withBlock.Add(pnPartNum, prPartNum);
                withBlock.Add(pnFamily, prFamily);
                foreach (var ky in withBlock.Keys)
                {
                    pr = aiProperty(withBlock.Item(ky));
                    if (pr == null)
                        System.Diagnostics.Debugger.Break();
                    else
                        dcVlPr.Add(ky, pr.Value);
                }

                if (withBlock.Exists(pnRawMaterial))
                {
                    prRawMatl = withBlock.Item(pnRawMaterial);
                    aiMatlNum = prRawMatl.Value;
                }
                else
                    aiMatlNum = "";

                if (withBlock.Exists(pnRmUnit))
                {
                    prRmUnit = withBlock.Item(pnRmUnit);
                    aiQtyUnit = prRmUnit.Value;
                }
                else
                    aiQtyUnit = "";

                if (withBlock.Exists(pnRmUnit))
                {
                    prRmQty = withBlock.Item(pnRmQty);
                    aiMatlQty = prRmQty.Value;
                }
                else
                    aiMatlQty = 0;
            }
            Debug.Print("=== Check Existing Model Genius Properties ===");
            Debug.Print(dumpLsKeyVal(dcVlPr, "="));
            Debug.Print();
            System.Diagnostics.Debugger.Break();

            /// NOTE[2021.11.11]
            /// Assignment of initial rt Dictionary
            /// now essentially duplicates the new
            /// process now preceding this section.
            /// The only difference is, that version
            /// does NOT apply Genius Property col-
            /// lection to the supplied Dictionary dc.
            rt = dcGnsPropsPart(invDoc, dc, 0); // dcAiPropsInSet
            dcVlAi = new Scripting.Dictionary();
            {
                var withBlock = rt;
                withBlock.Add(pnPartNum, prPartNum);
                withBlock.Add(pnFamily, prFamily);
                foreach (var ky in withBlock.Keys)
                {
                    pr = aiProperty(withBlock.Item(ky));
                    if (pr == null)
                        System.Diagnostics.Debugger.Break();
                    else
                        dcVlAi.Add(ky, pr.Value);
                }
                pr = null/* TODO Change to default(_) if this is not a reference type */;
            }
            /// Ultimately, processes which populate
            /// returned Dictionary rt, and set the
            /// Properties it should receive, should
            /// be moved toward the end of the function.

            {
                var withBlock = cnGnsDoyle();
                // Pre-clear all relevant variables
                // to be set from query results,
                // if available.

                // gnPartNum = aiPartFam
                // gnPartFam = ""
                // fdOrder = .Item("Ord")
                // gnMatlNum = ""
                // gnMatlFam = ""
                // gnMatlQty = 0
                // gnQtyUnit = ""
                // gnBomType = kDefaultBOMStructure
                // use this to indicate no BOM type
                // or structure returned from Genius


                Information.Err.Clear();
                rs = withBlock.Execute(sqlOf_ASDF(aiPartNum)); // 
                if (Information.Err.Number == 0)
                {
                    dcVlGn = dcFromAdoRSrow(rs, "");
                    {
                        var withBlock1 = dcVlGn;
                        gnPartNum = withBlock1.Item("Item");
                        gnPartFam = withBlock1.Item("Family");
                        gnBomType = withBlock1.Item("bomStr");
                        // fdOrder = .Item("Ord")
                        gnMatlNum = withBlock1.Item("Material");
                        gnMatlFam = withBlock1.Item("MtFamily");
                        gnMatlQty = withBlock1.Item("Qty");
                        gnQtyUnit = withBlock1.Item("Unit");
                    }

                    {
                        var withBlock1 = rs;
                        if (withBlock1.BOF & withBlock1.EOF)
                        {
                        }
                        else
                        {
                            {
                                var withBlock2 = withBlock1.Fields;
                            }

                            withBlock1.MoveNext();
                            if (!withBlock1.EOF)
                            {
                                System.Diagnostics.Debugger.Break(); // to handle multiple raw materials
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }
                        }

                        withBlock1.Close();
                    }
                }
                else
                {
                    Debug.Print(Information.Err.Number);
                    Debug.Print(Information.Err.Description);
                    System.Diagnostics.Debugger.Break();
                }

                withBlock.Close();
            }

            Debug.Print("== Prop Check ==");
            Debug.Print("---- Genius ----");
            Debug.Print(dumpLsKeyVal(dcVlGn, "="));
            Debug.Print("--- Inventor ---");
            Debug.Print(dumpLsKeyVal(dcVlPr, "=")); Debug.Print("================");
            System.Diagnostics.Debugger.Break();

            {
                var withBlock = invDoc;
                /// UPDATE[2021.11.11]
                /// Moved Property  collection
                /// to top of program, along with
                /// collection of Design Properties
                /// and their values. BOM Structure
                /// as well.

                /// We should check HERE for possibly misidentified purchased parts
                /// UPDATE[2021.11.08]
                /// Another MAJOR overhaul, here:
                /// change Purchased Parts identification
                /// to defer to Genius. Only attempt to guess
                /// when no value comes back from Genius.
                // Stop 'BKPT-2021-1108-1608
                /// CHANGE NEEDED[2021.11.08]:
                /// indeterminate -- stopping work @endOfDay
                /// effort here is to separate collection
                /// and potential reassignment of
                /// based on Part's family, file location,
                /// and whatever other criteria, if any.
                /// 
                /// Likely need a counterpart variable
                /// which takes its value from the Model.
                /// The most likely Genius equivalent is
                /// probably the ItemType field in view
                /// table vgMfiItems, which will need
                /// translation.
                /// 
                if (gnBomType == kDefaultBOMStructure)
                {
                    // Genius didn't return an Item type
                    // or BOM structure. We need to get it here.

                    /// BKPT-2021-1109-1042
                    /// Checkpoint here. Verify desired
                    /// behavior here prior to removal.
                    System.Diagnostics.Debugger.Break();

                    /// BOM Structure type, correcting if appropriate,
                    /// and prepare Family value for part, if purchased.
                    /// 
                    /// UPDATE[2018.02.06]
                    /// Using new UserForm; see below
                    /// UPDATE[2018.05.31]
                    /// Combined both InStr checks by addition
                    /// to generate a single test for > 0
                    /// If EITHER string match succeeds, the total
                    /// SHOULD exceed zero, so this SHOULD work.
                    /// UPDATE[2021.11.08]
                    /// Removed extraneoous code previously
                    /// disabled under preceding update[2018.05.31]
                    /// Also reseparated InStr checks previously combined
                    if (aiBomType == kPurchasedBOMStructure)
                        // Just assume that's what it's supposed to be.
                        gnBomType = aiBomType;
                    else if (InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + aiPartFam + "|") > 0)
                        // it needs to be SET purchased.
                        gnBomType = kPurchasedBOMStructure;
                    else
                    {
                        // might need to ask User
                        if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
                            // Double check with User.
                            ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Is this a Purchased Part?");
                        else
                            ck = Constants.vbNo;

                        // Stop 'BKPT-2021-1105-0942
                        /// CHANGE NEEDED[2021.11.05]:
                        /// ONLY COLLECT desired BOMStructure here
                        /// while keeping track of current value.
                        /// Reassignment should take place along
                        /// with collective property changes
                        /// UPDATE[2021.11.09]
                        /// This section now reduced to setting
                        /// gnBomType from User response, if any.
                        /// Code to assign Model BOM structure
                        /// moved toward bottom for further work.
                        /// 
                        /// Check process below replaces duplicate check/responses above.
                        if (ck == Constants.vbYes)
                            gnBomType = kPurchasedBOMStructure;
                        else
                        {
                            gnBomType = aiBomType;
                            dcVlGn.Item("bomStr") = gnBomType;
                        }

                        // Request #2: Change Cost Center iProperty.
                        // If BOMStructure = Purchased and not content center,
                        // then Family = D-PTS, else Family = D-HDWR.
                        // 
                        /// UPDATE[2018.05.30]: Value produced here
                        /// will now be held for later processing,
                        /// more toward the end of this function.
                        /// UPDATE[2021.11.09]
                        /// Changed to set target (Genius)
                        /// Family, and ONLY if not set already.
                        /// '
                        /// MIGHT want to set up a more robust check
                        /// system, but see how this holds up, first.
                        /// '
                        if (Strings.Len(gnPartFam) == 0)
                        {
                            if (gnBomType == kPurchasedBOMStructure)
                            {
                                if (withBlock.IsContentMember)
                                {
                                    System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-0946
                                    gnPartFam = "D-HDWR";
                                }
                                else
                                {
                                    System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-0947
                                    gnPartFam = "D-PTS";
                                }
                            }
                            else
                            {
                            }
                        }
                    }
                }

                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// Request #1:  the Mass in Pounds
                    /// and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        // Stop 'BKPT-2021-1110-1551
                        /// CHANGE NEEDED[2021.11.10]
                        /// '
                        /// '
                        /// '
                        if (dcVlPr.Exists(pnMass))
                        {
                            if (Round(cvMassKg2LbM * withBlock2.Mass, 4) - System.Convert.ToDouble(dcVlPr.Item(pnMass)) == 0)
                            {
                            }
                            else
                                // Stop
                                dcVlPr.Item(pnMass) = Round(cvMassKg2LbM * withBlock2.Mass, 4);
                        }
                        else
                            dcVlPr.Add(pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4));
                    }
                }
                /// At this point, gnPartFam SHOULD be set
                /// to a non-blank value if Item is purchased.
                /// We should be able to check this later on,
                /// if Item BOMStructure is NOT Normal

                // Stop 'BKPT-2021-1109-1053
                /// HERE is where it starts to get interesting
                /// Actually, just a little further down, where
                /// Part SubType is checked for Sheet Metal.
                /// At that point, the function divides into two
                /// LONG, and possibly nearly identical branches.
                /// Ideally, these should be refactored, with as
                /// much of their processes as possible combined
                /// into a single path.

                // Request #4: Change Cost Center iProperty.
                // If BOMStructure = Normal, then Family = D-MTO,
                // else if BOMStructure = Purchased then Family = D-PTS.
                if (aiBomType == kNormalBOMStructure)
                {
                    // Custom Properties
                    // Stop 'BKPT-2021-1105-1144
                    /// CHANGE NEEDED[2021.11.05]:
                    /// these properties should NOT
                    /// be added immediately, but only
                    /// when it's time to set them,
                    /// towards the END of this function.
                    /// UPDATE[2021.11.09]
                    /// Custom Property collection/generation
                    /// moved into Normal BOM Part handling, as
                    /// no earlier usage appears to take place.
                    /// '
                    /// If possible, may wish to move even further.
                    /// Plan to review later, as time permits.
                    /// UPDATE[2021.11.10]
                    /// Disabled Genius Property collection here
                    /// since a Dictionary of ALL Genius Properties
                    /// is generated towards the beginning.
                    /// '
                    /// '
                    // With rt
                    // If .Exists(pnRawMaterial) Then  prRawMatl = .Item(pnRawMaterial)
                    // prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                    // If .Exists(pnRmUnit) Then  prRmUnit = .Item(pnRmUnit)
                    // prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                    // If .Exists(pnRmQty) Then  prRmQty = .Item(pnRmQty)
                    // prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                    // End With
                    /// Collecting them at this point
                    /// might still be appropriate,
                    /// or it may be more desirable
                    /// to hold off until later.
                    /// '
                    /// UPDATE[2021.11.08]
                    /// Design Properties have been moved
                    /// toward the top, as proposed.
                    /// Commentary recommending this
                    /// has been removed as extraneous.
                    /// '
                    /// UPDATE[2021.11.09]
                    /// BOM Structure collection has also been
                    /// moved up, alongside Design Properties,
                    /// using renamed variable aiBomType
                    /// (formerly bomStruct)
                    /// '

                    // ----------------------------------------------------'
                    if (withBlock.SubType == guidSheetMetal)
                    {
                        // ----------------------------------------------------'
                        // Request #3:
                        // sheet metal extent area
                        // and add to custom property "RMQTY"
                        /// UPDATE[2021.11.10]
                        /// Now collecting Flat Pattern Values
                        /// instead of the Properties for them.
                        /// If necessary, Properties should be
                        /// assigned in a separate function
                        /// CHANGE NEEDED[2021.11.05]:
                        /// not quite sure on this one yet,
                        /// but dcFlatPatProps might need its
                        /// own set of revisions to generate
                        /// assignment recommendations WITHOUT
                        /// performing them itself
                        /// UPDATE[2021.11.10]
                        /// Embedded Flat Pattern Property collection
                        /// in bypassed If branch. Preceding Stop,
                        /// when enabled, offers user/developer
                        /// an opportunity to run it, if desired.
                        System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-1105
                        if (true)
                        {
                            dcVlFP = dcFlatPatVals(withBlock.ComponentDefinition); // dcVlAi
                            {
                                var withBlock1 = dcVlFP;
                                aiMatlFam = withBlock1.Item("mtFamily");
                                withBlock1.Remove("mtFamily");
                                foreach (var ky in withBlock1.Keys)
                                {
                                    if (dcVlPr.Exists(ky))
                                    {
                                        if (System.Convert.ToHexString(dcVlPr.Item(ky)) == System.Convert.ToHexString(withBlock1.Item(ky)))
                                        {
                                        }
                                        else
                                            System.Diagnostics.Debugger.Break();
                                    }
                                    else
                                        System.Diagnostics.Debugger.Break();
                                }
                            }
                            System.Diagnostics.Debugger.Break();
                        }
                        else
                            rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
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
                                /// UPDATE[2018.05.30]:
                                /// Restoring original key check
                                /// and adding code for debug
                                /// Previously changed to "~OFFTHK"
                                /// to avoid this block and its issues.
                                /// (Might re-revert if not prepped to fix now)
                                Debug.Print(aiProperty(rt.Item("OFFTHK")).Value);
                                System.Diagnostics.Debugger.Break(); // because we're going to need to do something with this.

                                gnMatlNum = ""; // Originally the ONLY line in this block.
                                                // A more substantial response is required here.

                                if (0)
                                    System.Diagnostics.Debugger.Break(); // (just a skipover)
                            }
                            else if (Strings.Len(gnMatlNum) == 0)
                            {
                                System.Diagnostics.Debugger.Break(); // because we don't know IF this is sheet metal yet
                                gnMatlNum = ptNumShtMetal(withBlock.ComponentDefinition);
                            }
                        }
                        else
                        {
                            // '  ACTION ADVISED[2018.09.14]:
                            // '  gnMatlNum can probably be set
                            // '  to prRawMatl.Value and THEN
                            // '  checked for length to see
                            // '  if lookup needed.
                            // '  This might also allow us to check
                            // '  for machined or other non-sheet
                            // '  metal parts.

                            System.Diagnostics.Debugger.Break();
                            /// !!!WARNING!!![2021.11.04]:
                            /// Following section has been shuffled
                            /// and should be considered HIGHLY
                            /// UNSTABLE until verified functional
                            /// and SAFE! TWO Stop commands are
                            /// placed to emphasize the need for
                            /// EXTREME CAUTION at this point
                            System.Diagnostics.Debugger.Break();
                            /// UPDATE[2021.11.04]:
                            /// This section is being adjusted
                            /// in an attempt to improve the raw
                            /// material determination process.
                            /// 
                            /// This particular segment should
                            /// ONLY be invoked if gnMatlNum is not
                            /// successfully retrieved from Genius
                            /// 
                            if (Strings.Len(gnMatlNum) == 0)
                            {
                                // no stock retrieved from Genius
                                // attempt to retrieve from Model
                                // gnMatlNum = aiMatlNum

                                if (Strings.Len(aiMatlNum) > 0)
                                {
                                    // need to verify it against Genius
                                    // by retrieving its Family there
                                    /// This With block copied and modified [2021.03.11]
                                    /// from elsewhere in this function as a temporary measure
                                    /// to address a stopping situation later in the function.
                                    /// See comment below for details.
                                    /// 
                                    /// UPDATE[2021.11.04]:
                                    /// This section MIGHT be removed in future,
                                    /// 
                                    {
                                        var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + gnMatlNum + "';");
                                        if (withBlock1.BOF | withBlock1.EOF)
                                        {
                                            // Stop 'because Material value likely invalid
                                            System.Diagnostics.Debugger.Break(); // because we do NOT want to set gnMatlNum!
                                            /// want to assign it to a separate RETURN variable
                                            /// or most likely, the return Dictionary.
                                            gnMatlNum = ptNumShtMetal(invDoc.ComponentDefinition);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                        }
                                        else
                                        {
                                            // '  This section retained from source,
                                            // '  but disabled to avoid potential issues
                                            // '  with subsequent operations, just in case
                                            // '  anything depends on gnMatlFam remaining
                                            // '  uninitialized up to that point.
                                            /// UPDATE[2021.11.09]
                                            /// Re-enabling Genius Material
                                            /// Family assignment, as it SHOULD
                                            /// be set to match what Genius
                                            /// returns from this query.
                                            /// '
                                            /// Might not be the best place
                                            /// to do this, though. If ptNumShtMetal
                                            /// returns a valid Material Item above,
                                            /// a Family is still needed.
                                            /// '
                                            /// NOTE: Fix disabled With block between runs
                                            // '  With .Fields
                                            System.Diagnostics.Debugger.Break(); // because we do not want to set gnMatlFam
                                            /// for same reasons as above
                                            gnMatlFam = withBlock1.Fields.Item("Family").Value;
                                        }
                                    }
                                }
                            }

                            if (Strings.Len(gnMatlNum) == 0)
                            {
                                /// UPDATE[2018.05.30]:
                                /// Pulling ALL code/text from this section
                                /// to get rid of excessive cruft.
                                /// 
                                /// In fact, reversing logic to go directly
                                /// to User Prompt if no stock identified
                                /// 
                                /// IN DOUBLE FACT, hauling this WHOLE MESS
                                /// RIGHT UP after initial gnMatlNum assignment
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
                                                gnPartFam = withBlock2.Item(pnFamily);
                                                Debug.Print(pnFamily + "=" + gnPartFam);
                                            }

                                            if (withBlock2.Exists(pnRawMaterial))
                                            {
                                                gnMatlNum = withBlock2.Item(pnRawMaterial);
                                                Debug.Print(pnRawMaterial + "=" + gnMatlNum);
                                            }
                                        }
                                        if (0)
                                            System.Diagnostics.Debugger.Break(); // Use this for a debugging shim
                                    }
                                }
                            }
                            else if (Left(gnMatlNum, 2) == "LG")
                            {
                                Debug.Print(aiPartNum + ": PROBABLE LAGGING");
                                Debug.Print("  TRY TO IDENTIFY, AND FILL IN BELOW.");
                                Debug.Print("  PRESS ENTER ON gnMatlNum LINE WHEN");
                                Debug.Print("  COMPLETED, THEN F5 TO CONTINUE.");
                                Debug.Print("  gnMatlNum = \"" + gnMatlNum + "\"");
                                System.Diagnostics.Debugger.Break();
                            }

                            if (Strings.Len(gnMatlNum) > 0)
                            {
                                // do we look for a Raw Material Family!

                                /// NOTE[2021.11.10]
                                /// This query is probably WAY more than needed here.
                                /// Spec fields are probably not needed at all,
                                /// and it's not clear which of the others might be.
                                /// '
                                /// It might also be possible to REMOVE this query
                                /// based on the earlier one, which would return
                                /// Material Family along with Part Family,
                                /// providing the Part/Item were found in Genius.
                                /// '
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + gnMatlNum + "';") // ", Description1, Unit, " &"Specification1, Specification2, Specification3, " &"Specification4, Specification5, Specification6, " &"Specification7, Specification8, Specification9, " &"Specification15, Specification16 " &'''
       ;
                                    /// UPDATE[2021.11.10]
                                    /// Removed (likely) unneeded fields from query text.
                                    /// Will keep a lookout for any resulting errors.
                                    if (withBlock1.BOF | withBlock1.EOF)
                                        System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                    else
                                    {
                                        {
                                            var withBlock2 = withBlock1.Fields;
                                            if (Strings.Len(gnMatlFam) > 0)
                                            {
                                                if (gnMatlFam == withBlock2.Item("Family").Value)
                                                {
                                                }
                                                else
                                                    System.Diagnostics.Debugger.Break();
                                            }
                                            else
                                                gnMatlFam = withBlock2.Item("Family").Value;
                                        }
                                        /// NOTE[2021.11.10]
                                        /// Else branch should PROBABLY end here
                                        /// to permit Recordset to be closed,
                                        /// and probably a new If/Then block
                                        /// proceed based on results.
                                        /// '

                                        /// UPDATE[2021.06.18]:
                                        /// New pre-check for Material Item
                                        /// in Purchased Parts Family.
                                        /// VERY basic handler simply
                                        /// maps Material Family to D-BAR
                                        /// to force extra processing below.
                                        /// Further refinement VERY much needed!
                                        if (gnMatlFam) /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                        {
                                            // Debug.Print aiPartNum & " [" & aiMatlNum & "]: " & aiPropsDesign(pnDesc).Value
                                            Debug.Print(aiPartNum + "[" + prRmQty.Value + gnQtyUnit + "*" + gnMatlNum + ": " + aiPropsDesign(pnDesc).Value + "]"); System.Diagnostics.Debugger.Break(); // FULL Stop!
                                        }
                                        else if (gnMatlFam == "D-PTS")
                                        {
                                            gnPartFam = "D-RMT";
                                            System.Diagnostics.Debugger.Break(); // NOT SO FAST!
                                            gnMatlFam = "D-BAR";
                                        }
                                        else if (gnMatlFam == "R-PTS")
                                        {
                                            gnPartFam = "R-RMT";
                                            System.Diagnostics.Debugger.Break(); // NOT SO FAST!
                                            gnMatlFam = "D-BAR";
                                        }

                                        if (gnMatlFam == "DSHEET")
                                        {
                                            // We should be okay. This is sheet metal stock

                                            /// UPDATE[2021.11.04]:
                                            /// Expanding gnPartFam and gnQtyUnit
                                            /// assignments to check for pre-
                                            /// existing values, and validate
                                            /// them if found.
                                            if (Strings.Len(gnPartFam) == 0)
                                                gnPartFam = "D-RMT";
                                            else if (gnPartFam == "D-RMT")
                                            {
                                            }
                                            else
                                                System.Diagnostics.Debugger.Break();// because we have

                                            if (Strings.Len(gnQtyUnit) == 0)
                                                gnQtyUnit = "FT2";
                                            else if (gnQtyUnit == "FT2")
                                            {
                                            }
                                            else
                                                System.Diagnostics.Debugger.Break();// because we have

                                            /// UPDATE[2018.05.30]:
                                            /// Moving part family assignment
                                            /// to this section for better mapping
                                            /// and updating to new Family names
                                            /// as well as pulling up gnQtyUnit assignment
                                            System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-1120
                                        }
                                        else
                                        {
                                            if (gnMatlFam == "D-BAR")
                                            {
                                                /// UPDATE[2021.06.18]:
                                                /// Added check for Part Family already set
                                                /// to more properly handle new situation (above)
                                                if (Strings.Len(gnPartFam) == 0)
                                                    gnPartFam = "R-RMT";
                                                else
                                                {
                                                    if (gnPartFam == "R-RMT")
                                                    {
                                                    }
                                                    else
                                                        System.Diagnostics.Debugger.Break();
                                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                                }

                                                if (Strings.Len(gnQtyUnit) == 0)
                                                    gnQtyUnit = "IN"; // prRmUnit.Value '
                                                else
                                                {
                                                    if (gnQtyUnit == "IN")
                                                    {
                                                    }
                                                    else
                                                        System.Diagnostics.Debugger.Break();
                                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                                }
                                                // 'may want function here
                                                /// UPDATE[2018.05.30]: As noted above
                                                /// Will keep Stop for now
                                                /// pending further review,
                                                /// hopefully soon
                                                Debug.Print(aiPartNum + " [" + gnMatlNum + "]: " + aiPropsDesign(pnDesc).Value);                                             /// UPDATE[2021.03.11]: Replaced
                                                                                                                                                                             /// aiPropsDesign.Item(pnPartNum)
                                                                                                                                                                             /// with prPartNum (and now aiPartNum)
                                                                                                                                                                             /// since it's used in several places
                                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                                Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                                // Stop 'BKPT-2021-1105-1137
                                                /// CHANGE NEEDED[2021.11.05]:
                                                /// indent the following With,
                                                /// when possible to do so
                                                /// without resetting project
                                                {
                                                    var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                    Debug.Print(withBlock2.MaxPoint.X - withBlock2.MinPoint.X); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                                }
                                                // Debug.Print "CURRENT RAW MATERIAL QUANTITY (";
                                                // Debug.Print CStr(prRmQty.Value); ") IS SHOWN BELOW."
                                                // Debug.Print "IF NOT CORRECT, YOU MAY TYPE A NEW VALUE"
                                                // Debug.Print "IN ITS PLACE, AND PRESS ENTER TO CHANGE IT."
                                                // Debug.Print "SOME SUGGESTED VALUES INCLUDE X, Y, AND Z"
                                                // Debug.Print "EXTENTS (ABOVE) OR YOU MAY SUPPLY YOUR OWN."
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                // Debug.Print "YOU MAY ALSO CHANGE THE UNIT OF MEASURE BELOW,"
                                                // Debug.Print "IF DESIRED. BE SURE TO PRESS ENTER/RETURN"
                                                // Debug.Print "AFTER CHANGING EITHER LINE. WHEN FINISHED, "
                                                // Debug.Print "PRESS [F5] TO CONTINUE."
                                                Debug.Print("");
                                                Debug.Print("prRmQty.Value = ");
                                                // Debug.Print "gnQtyUnit = """; gnQtyUnit; """"
                                                Debug.Print("gnQtyUnit = \"IN\"");
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                                            }
                                            else
                                            {
                                                Debug.Print("NON-STANDARD MATERIAL FAMILY (" + gnMatlFam + ")");
                                                Debug.Print("PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW");
                                                Debug.Print("PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED");
                                                Debug.Print("PRESS [F5] WHEN READY TO CONTINUE");
                                                Debug.Print("");
                                                Debug.Print("gnPartFam = \"" + gnPartFam + "\" 'PART FAMILY");
                                                Debug.Print("gnQtyUnit = \"" + gnQtyUnit + "\" 'UNIT OF MEASURE");
                                                System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                                            }

                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                            System.Diagnostics.Debugger.Break();

                                            System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-1117
                                            /// CHANGE NEEDED[2021.11.05]:
                                            /// Property assignment needs moved
                                            /// to collective assignment sequence
                                            rt = dcAddProp(prRmQty, rt);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                                        }
                                    }
                                }
                            }
                            else if (0)
                                System.Diagnostics.Debugger.Break();// and regroup
                        }
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
                            if (invDoc.ComponentDefinition.Document == invDoc)
                            {
                                // following needs indented if not already

                                /// [2018.07.31 by AT]
                                /// Added the following to try to
                                /// preselect non-sheet metal stock
                                // .dbFamily.Value = "D-BAR"
                                // .lbxFamily.Value = "D-BAR"
                                /// Doesn't quite do it.
                                // With New aiBoxData
                                // bd = nuAiBoxData().UsingInches.UsingBox(invDoc.ComponentDefinition.RangeBox)
                                System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-0955
                                /// CHANGE NEEDED[2021.11.05]:
                                /// Probably want to move this
                                /// outside of this With block,
                                /// and closer to the beginning
                                /// of this function, as it could
                                /// prove helpful at other points.
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
                                            if (Strings.Len(gnPartFam) == 0)
                                                gnPartFam = withBlock2.Item(pnFamily);
                                            else if (gnPartFam == withBlock2.Item(pnFamily))
                                            {
                                            }
                                            else
                                            {
                                                Debug.Print("=====");
                                                Debug.Print("Model Family differs from Genius");
                                                Debug.Print("Genius: " + pnFamily);
                                                Debug.Print("Model:  " + withBlock2.Item(pnFamily));
                                                Debug.Print("gnPartFam = .Item(pnFamily) 'Press [ENTER] on this line to fix, and/or [F5] to continue'");
                                                System.Diagnostics.Debugger.Break();
                                            }
                                            Debug.Print(pnFamily + "=" + gnPartFam);
                                        }

                                        if (withBlock2.Exists(pnRawMaterial))
                                        {
                                            if (Strings.Len(gnMatlNum) == 0)
                                                gnMatlNum = withBlock2.Item(pnRawMaterial);
                                            else if (gnMatlNum == withBlock2.Item(pnRawMaterial))
                                            {
                                            }
                                            else
                                            {
                                                Debug.Print("=====");
                                                Debug.Print("Model Raw Material differs from Genius");
                                                Debug.Print("Genius: " + gnMatlNum);
                                                Debug.Print("Model:  " + withBlock2.Item(pnRawMaterial));
                                                Debug.Print("gnMatlNum = .Item(pnRawMaterial) 'Press [ENTER] on this line to fix, and/or [F5] to continue'");
                                                System.Diagnostics.Debugger.Break();
                                            }
                                            Debug.Print(pnRawMaterial + "=" + gnMatlNum);
                                        }
                                    }
                                    if (0)
                                        System.Diagnostics.Debugger.Break(); // Use this for a debugging shim
                                }
                                else
                                    System.Diagnostics.Debugger.Break();// shouldn't actually hit this line
                            }
                            else
                                System.Diagnostics.Debugger.Break();// because we've got a serious mismatch
                        }
                        /// 
                        /// 

                        if (Strings.Len(gnMatlNum) > 0)
                        {
                            // do we look for a Raw Material Family!

                            /// This enclosing With block should NOT be necessary
                            /// since the newFmTest1 above takes care of collecting
                            /// the Stock Family along with the Stock itself
                            {
                                var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + gnMatlNum + "';");
                                if (withBlock1.BOF | withBlock1.EOF)
                                    System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                else
                                {
                                    var withBlock2 = withBlock1.Fields;
                                    if (Strings.Len(gnMatlFam) == 0)
                                        gnMatlFam = withBlock2.Item("Family").Value;
                                    else if (gnMatlFam == withBlock2.Item("Family").Value)
                                    {
                                    }
                                    else
                                        System.Diagnostics.Debugger.Break();
                                }
                            }

                            if (gnMatlFam == "DSHEET")
                            {
                                System.Diagnostics.Debugger.Break(); // because we should NOT be doing Sheet Metal in this section.
                                                                     // This might require further investigation and/or development, if encountered.
                                                                     // We should be okay. This is sheet metal stock
                                /// UPDATE[2021.11.04]:
                                /// Expanding gnPartFam and gnQtyUnit
                                /// assignments to check for pre-
                                /// existing values, and validate
                                /// them if found.
                                if (Strings.Len(gnPartFam) == 0)
                                    gnPartFam = "D-RMT";
                                else if (gnPartFam == "D-RMT")
                                {
                                }
                                else
                                    System.Diagnostics.Debugger.Break();// because we have

                                if (Strings.Len(gnQtyUnit) == 0)
                                    gnQtyUnit = "FT2";
                                else if (gnQtyUnit == "FT2")
                                {
                                }
                                else
                                    System.Diagnostics.Debugger.Break();// because we have
                            }
                            else if (gnMatlFam == "D-BAR")
                            {
                                gnPartFam = "R-RMT";
                                if (Strings.Len(gnQtyUnit) == 0)
                                    // this might have to change
                                    // to better handle case
                                    // of missing prRmUnit
                                    gnQtyUnit = prRmUnit.Value;// "IN"
                                                               // 'may want function here
                                /// UPDATE[2018.05.30]: As noted above
                                /// Will keep Stop for now
                                /// pending further review,
                                /// hopefully soon
                                Debug.Print(aiPartNum + " [" + gnMatlNum + "]: " + System.Convert.ToHexString(aiPropsDesign(pnDesc).Value));                             /// UPDATE[2021.03.11]: Replaced
                                                                                                                                                                         /// aiPropsDesign.Item(pnPartNum)
                                                                                                                                                                         /// as noted above
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                Debug.Print(invDoc.ComponentDefinition.RangeBox.MaxPoint.X - invDoc.ComponentDefinition.RangeBox.MinPoint.X); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                Debug.Print("");
                                Debug.Print("PLACE CURSOR ON gnQtyUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.");
                                Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.");
                                Debug.Print("");
                                Debug.Print("gnMatlQty = ");
                                Debug.Print("gnQtyUnit = \"IN\"");
                                Debug.Print("");
                                System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                                /// Actually, we might NOT need to stop here
                                /// if bar stock is already selected,
                                /// because quantities would presumably
                                /// have been established already.
                                /// Any D-BAR handler probably needs
                                /// to be implemented in prior section(s)
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                System.Diagnostics.Debugger.Break();
                                System.Diagnostics.Debugger.Break(); // BKPT-2021-1110-1647
                                /// CHANGE NEEDED[2021.11.10]
                                /// This Dictionary Property assignment
                                /// MUST be moved to the END of the function!
                                rt = dcWithProp(aiPropsUser, pnRmQty, gnMatlQty, rt); // dcAddProp(prRmQty, rt)
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                            }
                            else
                            {
                                Debug.Print("NON-STANDARD MATERIAL FAMILY (" + gnMatlFam + ")");
                                Debug.Print("PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW");
                                Debug.Print("PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED");
                                Debug.Print("PRESS [F5] WHEN READY TO CONTINUE");
                                Debug.Print("");
                                Debug.Print("gnPartFam = \"" + gnPartFam + "\" 'PART FAMILY");
                                Debug.Print("gnQtyUnit = \"" + gnQtyUnit + "\" 'UNIT OF MEASURE");
                                // gnPartFam = ""
                                // gnQtyUnit = "" 'may want function here
                                /// UPDATE[2018.05.30]: As noted above
                                /// However, might need more handling here.
                                System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                            }
                        }
                        else if (0)
                            System.Diagnostics.Debugger.Break();// and regroup
                    } // Sheetmetal vs Part

                    // Stop 'BKPT-2021-1105-1011
                    /// UPDATE[2021.11.10]
                    /// Disabled this prRawMatl assignment pending removal.
                    /// Counterpart moved below from sheet metal branch
                    /// should serve in place of both branch instances.
                    /// '
                    /// Extraneous commentary removed.
                    /// '
                    // With prRawMatl
                    // If Len(Trim$(.Value)) > 0 Then
                    // If gnMatlNum <> .Value Then
                    // 'Debug.Print "Raw Stock Selection"
                    // 'Debug.Print "  Current : " & prRawMatl.Value
                    // 'Debug.Print "  Proposed: " & gnMatlNum
                    // 'Stop 'because we might not want to change existing stock setting
                    // 'if
                    // ck = MsgBox('                Join(Array('                    "Raw Stock Change Suggested",'                    "  for Item " & aiPartNum,'                    "",'                    "  Current : " & prRawMatl.Value,'                    "  Proposed: " & gnMatlNum,'                    "", "Change It?", ""'                ), vbNewLine),'                vbYesNo, aiPartNum & " Stock"'            )
                    // '"Change Raw Material?"
                    // '"Suggested Sheet Metal"
                    // If ck = vbYes Then .Value = gnMatlNum
                    // End If
                    // Else
                    // .Value = gnMatlNum
                    // End If
                    // End With
                    // rt = dcAddProp(prRawMatl, rt)

                    // Stop 'BKPT-2021-1110-1130
                    /// UPDATE[2021.11.10]
                    /// Disabled this prRmUnit assignment pending removal.
                    /// Duplicate moved below from sheet metal branch
                    /// should serve in place of both branch instances.
                    /// '
                    /// Also moved End If AHEAD of this block to minimize
                    /// comment clutter WITHIN branches.
                    // With prRmUnit
                    // If Len(.Value) > 0 Then
                    // If Len(gnQtyUnit) > 0 Then
                    // If .Value <> gnQtyUnit Then
                    // Stop 'and check both so we DON'T
                    // 'automatically "fix" the RMUNIT value
                    // 
                    // .Value = gnQtyUnit
                    // 
                    // If 0 Then Stop 'Ctrl-9 here to skip changing
                    // End If
                    // End If
                    // Else 'we're setting a new quantity unit
                    // .Value = gnQtyUnit
                    // End If
                    // End With
                    // rt = dcAddProp(prRmUnit, rt)

                    System.Diagnostics.Debugger.Break(); // BKPT-2021-1109-1610
                    /// UPDATE[2021.11.10]
                    /// Transported this prRawMatl assignment
                    /// from sheet metal branch to consolidate
                    /// both instances of duplicated process
                    /// into one following both branches.
                    /// '
                    /// Extraneous commentary also removed.
                    /// '
                    if (prRawMatl == null)
                    {
                        rt = dcWithProp(aiPropsUser, pnRawMaterial, gnMatlNum, rt);
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                    else
                    {
                        {
                            var withBlock1 = prRawMatl;
                            if (Len(Trim(withBlock1.Value)) > 0)
                            {
                                if (gnMatlNum != withBlock1.Value)
                                {
                                    // Debug.Print "Raw Stock Selection"
                                    // Debug.Print "  Current : " & prRawMatl.Value
                                    // Debug.Print "  Proposed: " & gnMatlNum
                                    // Stop 'because we might not want to change existing stock setting
                                    // if
                                    ck = MsgBox(Join(Array("Raw Stock Change Suggested", "  for Item " + aiPartNum, "", "  Current : " + prRawMatl.Value, "  Proposed: " + gnMatlNum, "", "Change It?", ""), Constants.vbNewLine), Constants.vbYesNo, aiPartNum + " Stock");
                                    // "Change Raw Material?"
                                    // "Suggested Sheet Metal"
                                    if (ck == Constants.vbYes)
                                        withBlock1.Value = gnMatlNum;
                                }
                            }
                            else
                                withBlock1.Value = gnMatlNum;
                        }
                        rt = dcAddProp(prRawMatl, rt);
                    }

                    // Stop 'BKPT-2021-1110-1133
                    /// UPDATE[2021.11.10]
                    /// Transported this prRmUnit assignment
                    /// from sheet metal branch to consolidate
                    /// both instances of duplicated process
                    /// into one following both sheet metal
                    /// and structural branches
                    if (prRmUnit == null)
                    {
                        rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt);
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                    else
                    {
                        {
                            var withBlock1 = prRmUnit;
                            if (Len(withBlock1.Value) > 0)
                            {
                                if (Strings.Len(gnQtyUnit) > 0)
                                {
                                    if (withBlock1.Value != gnQtyUnit)
                                    {
                                        System.Diagnostics.Debugger.Break(); // and check both so we DON'T
                                                                             // automatically "fix" the RMUNIT value

                                        withBlock1.Value = gnQtyUnit;

                                        if (0)
                                            System.Diagnostics.Debugger.Break(); // Ctrl-9 here to skip changing
                                    }
                                }
                            }
                            else
                                withBlock1.Value = gnQtyUnit;
                        }
                        rt = dcAddProp(prRmUnit, rt);
                    }
                    // rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt) 'gnQtyUnit WAS "FT2"
                    /// Plan to remove commented line above,
                    /// superceded by the one above that
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing


                    // Stop 'BKPT-2021-1110-1133
                    /// UPDATE[2021.11.09]
                    /// This is a VERY crude implementation
                    /// of the closing BOM Structure assignment.
                    /// Plan on revision and cleanup in future.
                    if (gnBomType == aiBomType)
                    {
                        {
                            var withBlock1 = withBlock.ComponentDefinition;
                            if (withBlock1.BOMStructure != gnBomType)
                            {
                                withBlock1.BOMStructure = gnBomType;
                                if (Information.Err.Number == 0)
                                {
                                }
                                else
                                {
                                    System.Diagnostics.Debugger.Break();
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                }

                                System.Diagnostics.Debugger.Break();
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debugger.Break();
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                }
                else if (aiBomType == kPurchasedBOMStructure)
                {
                    /// As mentioned above, gnPartFam
                    /// SHOULD be set at this point
                    if (Strings.Len(gnPartFam) == 0)
                    {
                        if (1)
                            System.Diagnostics.Debugger.Break(); // because we might
                                                                 // need to check out the situation
                        gnPartFam = "D-PTS"; // by default
                    }
                }
                else
                    System.Diagnostics.Debugger.Break();// because we might need

                // Stop 'BKPT-2021-1105-1020
                /// CHANGE NEEDED[2021.11.05]:
                /// Family assignment should be
                /// ported up into collective
                /// Property assignment, although
                /// its position here assures
                /// one instance of the sequence
                /// catches ALL divergent cases
                /// leading up to this point.
                /// '
                /// Ultimately, those cases probably
                /// need to be consolidated HERE
                /// if, or WHEN possible.
                /// '
                // the design tracking property set,
                // and update the Cost Center Property
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
                else if (Strings.Len(gnPartFam) > 0)
                {
                    dcVlGn.Item("Family") = gnPartFam;
                    if (aiPartFam == gnPartFam)
                    {
                    }
                    else
                    {
                        prFamily.Value = gnPartFam;
                        if (Information.Err.Number)
                        {
                            Debug.Print("CHGFAIL[FAMILY]{'" + prFamily.Value + "' -> '" + gnPartFam + "'}: " + invDoc.DisplayName + " (" + invDoc.FullDocumentName + ")");
                            if (MsgBox("Couldn't Change Family" + vbNewLine, Constants.vbYesNo | Constants.vbDefaultButton2, invDoc.DisplayName) == Constants.vbYes)
                                System.Diagnostics.Debugger.Break();
                        }
                        else
                        {
                        }
                    }
                    rt = dcAddProp(prFamily, rt);
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                }
            }
            /// UPDATE[2021.11.09]
            /// Moved Part Mass Property assignment
            /// out of the main With block, modified
            /// to take its value from the new Values
            /// Dictionary.
            rt = dcWithProp(aiPropsUser, pnMass, dcVlPr.Item(pnMass), rt); // Round(cvMassKg2LbM * .Mass, 4)

            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            dcGeniusPropsPartRev20180530_broken2 = rt;
        }
    }

    public Scripting.Dictionary dcGeniusPropsPartRev20180530_broken(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary dcChg;
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

        string txFilePath;

        if (dc == null)
            dcGeniusPropsPartRev20180530_broken = dcGeniusPropsPartRev20180530_broken(invDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;
            dcChg = new Scripting.Dictionary();

            {
                var withBlock = invDoc;
                txFilePath = withBlock.FullFileName;

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
                prPartNum = aiGetProp(aiPropsDesign, pnPartNum); // ADDED 2021.03.11
                pnModel = prPartNum.Value;
                prFamily = aiGetProp(aiPropsDesign, pnFamily);

                /// UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// Request #1:  the Mass in Pounds
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                    }

                    bomStruct = withBlock1.BOMStructure; // kDefaultBOMStructure '''''''''''
                    dcChg = d2g1f1(prFamily, dcChg);
                }
                /// At this point, nmFamily SHOULD be set

                // Request #4: Change Cost Center iProperty.
                if (bomStruct == kNormalBOMStructure)
                {
                    // ----------------------------------------------------'
                    if (withBlock.SubType == guidSheetMetal)
                    {
                        // ----------------------------------------------------'
                        /// NOTE[2018-05-31]: At this point, we MAY wish
                        // Request #3:  sheet metal extent area
                        rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                        /// NOTE[2018-05-30]: Raw Material Quantity value

                        // NOTE: THIS call might best be combined somehow
                        if (prRawMatl == null)
                        {
                            if (rt.Exists("OFFTHK"))
                            {
                                /// UPDATE[2018.05.30]: Restoring original key check
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
                            // '  ACTION ADVISED[2018.09.14]: pnStock can probably be set
                            if (Len(prRawMatl.Value) > 0)
                            {
                                pnStock = prRawMatl.Value;
                                /// This With block copied and modified [2021.03.11]
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                                    if (withBlock1.BOF | withBlock1.EOF)
                                    {
                                        // Stop 'because Material value likely invalid
                                        pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                    }
                                    else
                                    {
                                    }
                                }
                            }
                            else
                                pnStock = ptNumShtMetal(withBlock.ComponentDefinition);

                            if (Strings.Len(pnStock) == 0)
                            {
                                /// UPDATE[2018.05.30]: Pulling ALL code/text from this section
                                {
                                    var withBlock1 = newFmTest1();
                                    if (!(invDoc.ComponentDefinition.Document == invDoc))
                                        System.Diagnostics.Debugger.Break();

                                    bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                                    ck = withBlock1.AskAbout(invDoc, "No Stock Found! Please Review" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                                    if (ck == Constants.vbYes)
                                    {
                                        /// UPDATE[2018.05.30]: Pulling some extraneous commented code
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
                                Debug.Print(pnModel + ": PROBABLE LAGGING");
                                Debug.Print("  TRY TO IDENTIFY, AND FILL IN BELOW.");
                                Debug.Print("  PRESS ENTER ON pnStock LINE WHEN");
                                Debug.Print("  COMPLETED, THEN F5 TO CONTINUE.");
                                Debug.Print("  pnStock = \"" + pnStock + "\"");
                                System.Diagnostics.Debugger.Break();
                            }

                            if (Strings.Len(pnStock) > 0)
                            {
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

                                        /// UPDATE[2021.06.18]: New pre-check for Material Item
                                        if (mtFamily) /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                        {
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
                                            /// UPDATE[2021.06.18]: Added check for Part Family already set
                                            if (Strings.Len(nmFamily) == 0)
                                                nmFamily = "R-RMT";
                                            else
                                                Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing

                                            qtUnit = prRmUnit.Value; // "IN"
                                                                     // 'may want function here
                                            /// UPDATE[2018.05.30]: As noted above
                                            Debug.Print(pnModel + " [" + prRawMatl.Value + "]: " + aiPropsDesign(pnDesc).Value);
                                            /// UPDATE[2021.03.11]: Replaced aiPropsDesign.Item(pnPartNum)
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                            Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                            Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                            {
                                                var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                Debug.Print(withBlock2.MaxPoint.X - withBlock2.MinPoint.X); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                            }
                                            // Debug.Print "CURRENT RAW MATERIAL QUANTITY (";
                                            // Debug.Print CStr(prRmQty.Value); ") IS SHOWN BELOW."
                                            Debug.Print("");
                                            Debug.Print("prRmQty.Value = ");
                                            // Debug.Print "qtUnit = """; qtUnit; """"
                                            Debug.Print("qtUnit = \"IN\"");
                                            // Debug.Print ""
                                            System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                                            /// Actually, we might NOT need to stop here
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                            System.Diagnostics.Debugger.Break();
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
                                    ck = MsgBox(Join(Array("Raw Stock Change Suggested", "  for Item " + pnModel, "", "  Current : " + prRawMatl.Value, "  Proposed: " + pnStock, "", "Change It?", ""), Constants.vbNewLine), Constants.vbYesNo, pnModel + " Stock");
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
                                        System.Diagnostics.Debugger.Break(); // and check both so we DON'T automatically "fix" the RMUNIT value

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
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Another landing line
                    }
                    else
                    {
                        // --------------------------------------------'
                        /// [2018.07.31 by AT] Duped following block from above
                        {
                            var withBlock1 = newFmTest1();
                            if (!(invDoc.ComponentDefinition.Document == invDoc))
                                System.Diagnostics.Debugger.Break();

                            /// [2018.07.31 by AT] Added the following to try to
                            bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);

                            ck = withBlock1.AskAbout(invDoc, "Please Select Stock for Machined Part" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                            if (ck == Constants.vbYes)
                            {
                                /// UPDATE[2018.05.30]: Pulling some extraneous commented code
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
                        /// The following If block is copied wholesale from sheet metal section above.
                        if (Strings.Len(pnStock) > 0)
                        {

                            /// This enclosing With block should NOT be necessary
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

                            if (mtFamily == "DSHEET")
                            {
                                System.Diagnostics.Debugger.Break();
                                // because we should NOT be doing Sheet Metal in this section.
                                nmFamily = "D-RMT";
                                qtUnit = "FT2";
                            }
                            else if (mtFamily == "D-BAR")
                            {
                                nmFamily = "R-RMT";
                                qtUnit = prRmUnit.Value; // "IN"
                                                         // 'may want function here
                                /// UPDATE[2018.05.30]: As noted above Will keep Stop for now
                                Debug.Print(pnModel);
                                /// UPDATE[2021.03.11]: Replaced aiPropsDesign.Item(pnPartNum)
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                Debug.Print(invDoc.ComponentDefinition.RangeBox.MaxPoint.X - invDoc.ComponentDefinition.RangeBox.MinPoint.X); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                Debug.Print("");
                                Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.");
                                Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.");
                                Debug.Print("");
                                Debug.Print("prRmQty.Value = ");
                                Debug.Print("qtUnit = \"IN\"");
                                Debug.Print("");
                                System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                                /// Actually, we might NOT need to stop here
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                System.Diagnostics.Debugger.Break();
                                rt = dcAddProp(prRmQty, rt);
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                            }
                            else
                            {
                                nmFamily = "";
                                qtUnit = ""; // may want function here
                                /// UPDATE[2018.05.30]: As noted above
                                System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                            }
                        }
                        else if (0)
                            System.Diagnostics.Debugger.Break();// and regroup

                        {
                            var withBlock1 = prRawMatl;
                            if (Len(Trim(withBlock1.Value)) > 0)
                            {
                                if (pnStock != withBlock1.Value)
                                {
                                    ck = MsgBox(Join(Array("Raw Stock Change Suggested", "  Current : " + prRawMatl.Value, "  Proposed: " + pnStock, "", "Change It?", ""), Constants.vbNewLine), Constants.vbYesNo, "Change Raw Material?");
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
                                        System.Diagnostics.Debugger.Break(); // and check both so we DON'T automatically "fix" the RMUNIT value

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
                    /// As mentioned above, nmFamily SHOULD be set at this point
                    if (Strings.Len(nmFamily) == 0)
                    {
                        if (1)
                            System.Diagnostics.Debugger.Break(); // because we might need to check out the situation
                        nmFamily = "D-PTS"; // by default
                    }
                }
                else
                    System.Diagnostics.Debugger.Break();// because we might need to do something else

                // the design tracking property set,
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
            dcGeniusPropsPartRev20180530_broken = rt;
        }
    }

    public Scripting.Dictionary dcGeniusPropsPartDvl20210929(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
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
            dcGeniusPropsPartDvl20210929 = dcGeniusPropsPartDvl20210929(invDoc, new Scripting.Dictionary());
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
                prPartNum = aiGetProp(aiPropsDesign, pnPartNum); // ADDED 2021.03.11
                pnModel = prPartNum.Value;
                prFamily = aiGetProp(aiPropsDesign, pnFamily);

                /// Request #1:  the Mass in Pounds
                {
                    var withBlock1 = withBlock.ComponentDefinition.MassProperties;
                    rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock1.Mass, 4), rt);
                }

                /// NOTE[2021.10.01]: This block is for Purchased Part Determination! (see below)
                /// UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// BOM Structure type, correcting if appropriate,
                    ck = Constants.vbNo;
                }
                /// NOTE[2021.10.01]: END OF BLOCK for Purchased Part Determination!

                // Request #4: Change Cost Center iProperty.
                if (bomStruct == kNormalBOMStructure)
                {
                }
                else if (bomStruct == kPurchasedBOMStructure)
                {
                }
                else
                {
                }

                // the design tracking property set,
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
                else
                {
                }
            }
        }
    }

    public Scripting.Dictionary d2g2f1(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */) // Inventor.BOMStructureEnum
    {
        /// d2g2f1 -- (to be determined)
        /// 
        /// code here extracted for development
        /// from Function dcGeniusPropsPartRev20180530
        /// in module modGPUpdateAT (start line 559)
        /// lines 1056 (497 down from start)
        /// to 1201 (146 lines copied)
        /// along with necessary declarations:
        string pnModel;
        string nmFamily;
        string mtFamily;
        string pnStock;
        string qtUnit;
        aiBoxData bd;
        VbMsgBoxResult ck;
        /// followed by new declarations:
        Scripting.Dictionary rt;

        rt = new Scripting.Dictionary();

        {
            var withBlock = invDoc.ComponentDefinition;
            if (withBlock.Document == invDoc)
            {
                bd = nuAiBoxData().UsingInches.SortingDims(withBlock.RangeBox);
                {
                    var withBlock1 = newFmTest1() // ''== Original Line 1056 ==
       ;
                    // If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                    // moved this check outside this form block (see above)

                    /// [2018.07.31 by AT]
                    /// Added the following to try to
                    /// preselect non-sheet metal stock
                    // .dbFamily.Value = "D-BAR"
                    // .lbxFamily.Value = "D-BAR"
                    /// Doesn't quite do it.

                    ck = withBlock1.AskAbout(invDoc, "Please Select Stock for Machined Part" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                    if (ck == Constants.vbYes)
                    {
                        /// UPDATE[2018.05.30]:
                        /// Pulling some extraneous commented code
                        /// from here and beginning of block
                        {
                            var withBlock2 = withBlock1.ItemData();
                            if (withBlock2.Exists(pnFamily))
                            {
                                nmFamily = withBlock2.Item(pnFamily);
                                rt.Add(pnFamily, nmFamily);
                                Debug.Print(pnFamily + "=" + nmFamily);
                            }

                            if (withBlock2.Exists(pnRawMaterial))
                            {
                                pnStock = withBlock2.Item(pnRawMaterial);
                                rt.Add(pnRawMaterial, pnStock);
                                Debug.Print(pnRawMaterial + "=" + pnStock);
                            }
                        }
                        if (0)
                            System.Diagnostics.Debugger.Break(); // Use this for a debugging shim
                    }
                }
            }
            else
                System.Diagnostics.Debugger.Break();
        }

        if (Strings.Len(pnStock) > 0)
        {
            // do we look for a Raw Material Family!

            /// This enclosing With block should NOT be necessary
            /// since the newFmTest1 above takes care of collecting
            /// the Stock Family along with the Stock itself
            {
                var withBlock = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                if (withBlock.BOF | withBlock.EOF)
                    System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                else
                {
                    var withBlock1 = withBlock.Fields;
                    mtFamily = withBlock1.Item("Family").Value;
                }
            }

            if (mtFamily == "DSHEET")
            {
                System.Diagnostics.Debugger.Break(); // because we should NOT be doing Sheet Metal in this section.
                                                     // might require further investigation and/or development, if encountered.
                                                     // We should be okay. This is sheet metal stock
                nmFamily = "D-RMT";
                qtUnit = "FT2";
            }
            else if (mtFamily == "D-BAR")
            {
                nmFamily = "R-RMT";
                System.Diagnostics.Debugger.Break(); // and note disabled qtUnit -- needs work here
                                                     // qtUnit = prRmUnit.Value '"IN"
                                                     // 'may want function here
                /// UPDATE[2018.05.30]: As noted above
                /// Will keep Stop for now
                /// pending further review,
                /// hopefully soon
                System.Diagnostics.Debugger.Break(); // and note disabled prRawMatl too
                                                     // Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                /// UPDATE[2021.03.11]: Replaced
                /// aiPropsDesign.Item(pnPartNum)
                /// as noted above
                System.Diagnostics.Debugger.Break(); // and note disabled prRmQty
                                                     // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                Debug.Print(invDoc.ComponentDefinition.RangeBox.MaxPoint.X - invDoc.ComponentDefinition.RangeBox.MinPoint.X); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                Debug.Print("");
                Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.");
                Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.");
                Debug.Print("");
                System.Diagnostics.Debugger.Break(); // and note disabled prRmQty again
                                                     // Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                Debug.Print("qtUnit = \"IN\"");
                Debug.Print("");
                System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                /// Actually, we might NOT need to stop here
                /// if bar stock is already selected,
                /// because quantities would presumably
                /// have been established already.
                /// Any D-BAR handler probably needs
                /// to be implemented in prior section(s)
                System.Diagnostics.Debugger.Break(); // and note one moredisabled prRmQty
                                                     // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                System.Diagnostics.Debugger.Break();
                System.Diagnostics.Debugger.Break(); // and note prRmQty once more disabled
                                                     // this one really DOES need removed from this function
                                                     // rt = dcAddProp(prRmQty, rt)
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
            }
            else
            {
                System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                nmFamily = "";
                qtUnit = ""; // may want function here
            }
        }
        else if (0)
            System.Diagnostics.Debugger.Break();// and regroup

        d2g2f1 = rt;
    }

    public Scripting.Dictionary d2g1f1(Inventor.Property prFamily, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */) // Inventor.BOMStructureEnum
    {
        Scripting.Dictionary rt;
        // Dim invDoc As Inventor.PartDocument
        VbMsgBoxResult ck;
        string txFilePath;
        Inventor.BOMStructureEnum isNow;
        Inventor.BOMStructureEnum shdBe;
        string gnFam;
        string aiFam;
        string ptNum;
        fmTest2 fm = new fmTest2();

        if (dc == null)
            rt = d2g1f1(prFamily, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = prFamily;
                /// Family from Model
                aiFam = withBlock.Value;

                {
                    var withBlock1 = withBlock.Parent // Property 
       ;
                    /// Part Number from Model
                    ptNum = withBlock1.Item(pnPartNum).Value;

                    /// Then try to get Family from Genius
                    {
                        var withBlock2 = cnGnsDoyle();
                        {
                            var withBlock3 = withBlock2.Execute(Join(Array("select ISNULL(i.Family, '') Family", "from vgMfiItems i right join", "(values ('" + ptNum + "')) ls(Item)", "on i.Item = ls.Item", ";"), Constants.vbNewLine));
                            if (withBlock3.BOF | withBlock3.EOF)
                            {
                                System.Diagnostics.Debugger.Break(); // because something went wrong
                                gnFam = "";
                            }
                            else
                                gnFam = withBlock3.GetRows()(0, 0);
                            withBlock3.Close();
                        }
                    }

                    {
                        var withBlock2 = withBlock1.Parent  /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
       ;
                        /// File Path to check for Purchased Part
                        txFilePath = aiDocument(withBlock2.Parent).FullFileName;

                        // Request #2: Change Cost Center iProperty.
                        if (ck4ContentMember(withBlock2.Parent))
                        {
                            if (Strings.Len(gnFam) == 0)
                                gnFam = "D-HDWR";
                            else if (gnFam == "D-HDWR")
                            {
                            }
                            else if (gnFam == "D-PTS")
                            {
                            }
                            else if (gnFam == "R-PTS")
                            {
                            }
                            else
                                System.Diagnostics.Debugger.Break();
                        }

                        isNow = bomStructOf(withBlock2.Parent);

                        // fm = newFmTest2()

                        /// Check Model Family against Genius Family,
                        /// if defined, and if different, ask whether
                        /// it should be changed.
                        if (Strings.Len(gnFam) > 0)
                        {
                            if (gnFam != aiFam)
                            {
                                ck = fm.AskAbout(withBlock2.Parent, Join(Array("Model Family " + aiFam + " does not", "match Genius Part Family " + gnFam), Constants.vbNewLine), Join(Array("Should Model be updated", "to match Genius?"), Constants.vbNewLine));
                                if (ck == Constants.vbCancel)
                                    System.Diagnostics.Debugger.Break();

                                if (ck == Constants.vbYes)
                                    rt.Add(prFamily.Name, gnFam);
                                else
                                    gnFam = aiFam;
                            }
                        }

                        /// BOM Structure type,
                        /// correcting if appropriate,
                        /// UPDATE[2018.05.31]: Combined both InStr checks
                        if (InStr(1, txFilePath, @"\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0)
                            shdBe = kPurchasedBOMStructure;
                        else
                            shdBe = isNow;

                        if (shdBe != isNow)
                        {
                            ck = fm.AskAbout(withBlock2.Parent, Join(Array("Model Family " + gnFam + " or File Path", "(" + txFilePath + ")", "indicates a Purchased Part, but BOM", "Structure is NOT set to match"), Constants.vbNewLine), Join(Array("Should BOM Structure", "be set to Purchased?"), Constants.vbNewLine));
                            if (ck == Constants.vbCancel)
                                System.Diagnostics.Debugger.Break();

                            if (ck == Constants.vbYes)
                                // 
                                // .BOMStructure = kPurchasedBOMStructure
                                // If Err.Number = 0 Then
                                // bomStruct = .BOMStructure
                                // Else
                                // bomStruct = kPurchasedBOMStructure
                                // ''' WARNING: NOT a good way to go about this
                                // '''     but will go with it for now
                                // End If
                                // 
                                rt.Add("BOMstructure", shdBe);
                            else
                                shdBe = isNow;
                        }
                    }
                }
                Debug.Print(); // Breakpoint Landing
            }
        }

        d2g1f1 = rt;
    }

    public Scripting.Dictionary dcCtOfEach(Variant ls)
    {
        Scripting.Dictionary rt;
        Variant ck;
        long mx;
        long dx;

        rt = new Scripting.Dictionary();
        if (IsArray(ls))
        {
            mx = UBound(ls);
            dx = LBound(ls);
            {
                var withBlock = rt;
                while (!dx > mx)
                {
                    ck = ls(dx);
                    if (withBlock.Exists(ck))
                        withBlock.Item(ck) = withBlock.Item(ck) + 1;
                    else
                        withBlock.Add(ck, 1);

                    dx = 1 + dx;
                }
            }

            dcCtOfEach = rt;
        }
        else
            dcCtOfEach = dcCtOfEach(Array(ls));
    }

    public Scripting.Dictionary dcGnsMatlOps(Scripting.Dictionary DimCt, string MtSpec = "") // defaulted to SS, but maybe not such a great idea
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary rw;
        ADODB.Recordset rs;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = cnGnsDoyle();
            Information.Err.Clear();
            rs = withBlock.Execute(sqlOf_GnsMatlOptions(MtSpec, DimCt.Keys));

            if (Information.Err.Number == 0)
            {
                {
                    var withBlock1 = dcFromAdoRS(rs, "");
                    foreach (var ky in withBlock1.Keys)
                    {
                        rw = dcOb(withBlock1.Item(ky));
                        if (rw == null)
                            System.Diagnostics.Debugger.Break();
                        else
                            rt.Add(rw.Item("Item"), rw);
                    }
                }

                rs.Close();
            }
            else
            {
                System.Diagnostics.Debugger.Break();
                Information.Err.Clear();
            }


            withBlock.Close();
        }
        dcGnsMatlOps = rt;
    }

    public bool ck4ContentMember(Inventor.Document AiDoc)
    {
        ck4ContentMember = ptIsContentMember(aiDocPart(AiDoc));
    }

    public bool ptIsContentMember(Inventor.PartDocument AiDoc)
    {
        if (AiDoc == null)
            ptIsContentMember = 0;
        else
            ptIsContentMember = AiDoc.ComponentDefinition.IsContentMember;
    }

    public Inventor.BOMStructureEnum bomStructOfPart(Inventor.PartDocument AiDoc)
    {
        if (AiDoc == null)
            bomStructOfPart = 0;
        else
            bomStructOfPart = AiDoc.ComponentDefinition.BOMStructure;
    }

    public Inventor.BOMStructureEnum bomStructOfAssy(Inventor.AssemblyDocument AiDoc)
    {
        if (AiDoc == null)
            bomStructOfAssy = 0;
        else
            bomStructOfAssy = AiDoc.ComponentDefinition.BOMStructure;
    }

    public Inventor.BOMStructureEnum bomStructOf(Inventor.Document AiDoc)
    {
        if (AiDoc == null)
            bomStructOf = 0;
        else if (AiDoc is Inventor.PartDocument)
            bomStructOf = bomStructOfPart(AiDoc);
        else if (AiDoc is Inventor.AssemblyDocument)
            bomStructOf = bomStructOfAssy(AiDoc);
        else
            bomStructOf = 0;
    }
}