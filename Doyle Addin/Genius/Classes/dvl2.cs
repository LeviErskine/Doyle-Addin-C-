using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class dvl2
{
    public Dictionary dcGnsMatlSpecPairings()
    {
        // dcGnsMatlSpecPairings -- Genius Raw Material Spec Relations
        // Return a Dictionary of Dictionaries
        // keyed to each Specification value
        // found in ANY Spec field of any
        // Raw Material Item, each listing
        // all OTHER Spec values found in
        // conjunction with each value.
        // 

        var rt = new Dictionary();

        var wk = dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_MatlSpecXref())));
        if (wk == null)
        {
        }
        else if (wk.Exists("val"))
        {
            Dictionary dcVl = wk.get_Item("val");
            {
                foreach (var kyVl in dcVl.Keys)
                    rt.Add(kyVl, new Dictionary());

                foreach (var kyVl in dcVl.Keys)
                {
                    Dictionary dcAl = rt.get_Item(kyVl);
                    {
                        var withBlock1 = dcOb(dcVl.get_Item(kyVl));
                        foreach (var dxAl in withBlock1.Keys)
                        {
                            string kyAl = dcOb(withBlock1.get_Item(dxAl)).get_Item("also");
                            {
                                if (rt.Exists(kyAl))
                                    dcAl.Add(kyAl, rt.get_Item(kyAl));
                                else
                                    Debugger.Break(); // because something went wrong
                            }
                        }
                    }
                }
            }
        }

        return rt;
    }

    public Dictionary dcOfDcWithXrefsDep1st(Dictionary dc, Dictionary wk = null, string pt = "#")
    {
        // dcOfDcWithXrefsDep1st
        // Replace rudundant / recursive Dictionary
        // Objects in hierarchical Dictionary structure
        // '
        // This is a depth-first implementation, which
        // might locate an initial Dictionary reference
        // deep inside an early branch before finding a
        // shallower instance that might be preferable.
        // '
        // A breadth-first implementation might be preferred.
        // 
        Dictionary rt;

        if (wk == null)
            rt = dcOfDcWithXrefsDep1st(dc, new Dictionary());
        else
        {
            rt = new Dictionary();
            {
                foreach (var ky in dc.Keys)
                {
                    var ar = new[] { dc.get_Item(ky) };
                    Dictionary ck = dcOb(obOf(ar(0)));

                    if (ck == null)
                    {
                    }
                    else if (wk.Exists(ck))
                        ar = [wk.get_Item(ck)];
                    else
                    {
                        // prep new $ref path
                        var sp = pt + "/" + Convert.ToString(ky as string);

                        // add new $ref to wk
                        {
                            var withBlock1 = nuDcPopulator().Setting("$ref", sp);
                            wk.Add(ck, withBlock1.Dictionary);
                        }

                        // go ahead and process
                        // subdictionary
                        ar = [dcOfDcWithXrefsDep1st(ck, wk, sp)];
                    }

                    rt.Add(ky, ar(0));
                }
            }

            return rt;
        }

        ;

        return rt;
    }

    public Dictionary dcOfDcWithXrefsBrd1st(Dictionary dc, Dictionary wk = null, string pt = "#")
    {
        // dcOfDcWithXrefsBrd1st
        // Replace redundant / recursive Dictionary
        // Objects in hierarchical Dictionary structure
        // '
        // This is a depth-first implementation, which
        // might locate an initial Dictionary reference
        // deep inside an early branch before finding a
        // shallower instance that might be preferable.
        // '
        // A breadth-first implementation might be preferred.
        // 
        Dictionary rt;
        var ar;

        if (wk == null)
            rt = dcOfDcWithXrefsBrd1st(dc, new Dictionary());
        else
        {
            // create returned Dictionary
            rt = new Dictionary();

            // create local working
            // Dictionary of Dictionaries
            var ls = new Dictionary();

            // being processing
            // supplied Dictionary
            {
                // first pass: collect and process
                // all sub Dictionary Objects
                foreach (var ky in dc.Keys)
                {
                    Dictionary ck = dcOb(obOf(dc.get_Item(ky)));

                    if (ck != null) continue;
                    if (wk.Exists(ck))
                        // add existing $ref Dictionary
                        // to Dictionary list. thinking
                        // recursion should NOT be an issue
                        ls.Add(ky, wk.get_Item(ck));
                    else
                    {
                        // add new Dictionary to list
                        // for subsequent recursion
                        ls.Add(ky, dc.get_Item(ky));

                        // prep new $ref path
                        var sp = pt + "/" + Convert.ToString(ky as string);

                        {
                            // add new $ref Dictionary
                            wk.Add(ck, new Dictionary());

                            // add path to Dictionary
                            dcOb(wk.get_Item(ck)).Add("$ref", sp);
                        }
                    }
                }

                foreach (var ky in dc.Keys)
                {
                    rt.Add(ky,
                        ls.Exists(ky)
                            ? dcOfDcWithXrefsBrd1st(ls.get_Item(ky), wk, pt + "/" + Convert.ToString(ky as string))
                            : dc.get_Item(ky));
                }
            }
        }

        return rt;
    }

    public Dictionary dcGnsMatlSpecPairings4json()
    {
        // dcGnsMatlSpecPairings4json -- check on dcGnsMatlSpecPairings
        // 
        var rt = dcGnsMatlSpecPairings();

        {
            foreach (var ky in rt.Keys())
            {
                // .get_Item(ky) = Join(dcOb(.get_Item(ky)).Keys)
                {
                    var withBlock1 = dcOb(rt.get_Item(ky));
                    foreach (var k2 in withBlock1.Keys())
                        withBlock1.get_Item(k2) = Join(dcOb(withBlock1.get_Item(k2)).Keys);
                }
            }
        }

        return rt;
    }

    public Dictionary dcSpecSubsetWith(string txSpec, Dictionary inDc)
    {
        Dictionary rt;

        return inDc.Exists(txSpec)
            ? (Dictionary)dcKeysInCommon(inDc, dcOb(inDc.get_Item(txSpec)), 1)
            : new Dictionary();
    }

    public Dictionary dcSpecSubsetWithAll(Dictionary dcSpec, Dictionary inDc)
    {
        var rt = inDc;
        foreach (var ky in dcSpec.Keys)
            rt = dcSpecSubsetWith(Convert.ToHexString(ky), rt);
        return rt;
    }

    public Dictionary dcSpecSetFromUser()
    {
        string nx;

        var rt = new Dictionary();
        var dc = dcGnsMatlSpecPairings();

        // Debug.Print Join(dc.Keys)
        do
        {
            var fm = nuSelFromDict(dc);
            nx = fm.GetReply(null, "");
            if (Strings.Len(nx) <= 0) continue;
            rt.Add(nx, nx);
            dc = dcSpecSubsetWith(nx, dc);
            if (dc.Count == 0)
                nx = "";
        } while (Strings.Len(nx) > 0);

        // Stop
        return rt;
    }

    public Dictionary d2g3f1(PartDocument Part, Dictionary dc = null)
    {
        // d2g3f1 -- Return Dictionary
        // of relevant Part Properties
        // and information for use in
        // Genius data extraction
        // 
        Dictionary rt;
        string txPartNum;

        if (dc == null)
            rt = d2g3f1(Part, new Dictionary());
        else
        {
            rt = dc;

            {
                {
                    var withBlock1 = Part.PropertySets.get_Item(gnDesign);
                    rt.Add(pnPartNum, withBlock1.get_Item(pnPartNum));
                    rt.Add(pnFamily, withBlock1.get_Item(pnFamily));
                }
            }

            rt = dcGnsInfoCompDef(aiCompDefOf(Part), rt);
        }

        return rt;
    }

    public Dictionary dcGnsInfoAiDocBase(Document AiDoc, Dictionary dc = null)
    {
        // dcGnsInfoAiDocBase (formerly d2g3f1a)
        // Return Dictionary of Document Properties
        // and information relevant to Genius
        // for data extraction
        // 
        Dictionary rt;
        string txPartNum;

        if (dc == null)
            rt = dcGnsInfoAiDocBase(AiDoc, new Dictionary());
        else
        {
            rt = dc;

            {
                {
                    var withBlock1 = AiDoc.PropertySets.get_Item(gnDesign);
                    rt.Add(pnPartNum, withBlock1.get_Item(pnPartNum));
                    rt.Add(pnFamily, withBlock1.get_Item(pnFamily));
                }

                if (false)
                {
                    rt.Add("subType", AiDoc.SubType);
                    rt.Add("docType", AiDoc.DocumentType);
                    rt.Add("dsbType", AiDoc.DocumentSubType.DocumentSubTypeID);
                }
            }

            rt = dcGnsInfoCompDef(aiCompDefOf(AiDoc), rt);
        }

        return rt;
    }

    public Dictionary dcGnsInfoCompDef(ComponentDefinition CpDef, Dictionary dcWkg = null)
    {
        // dcGnsInfoCompDef -- Generate and/or populate
        // Dictionary (new or supplied) with data for
        // Genius from supplied ComponentDefinition.
        // This is the "generic" dynamic, which dispatches
        // the supplied ComponentDefinition to a dynamic
        // more specific to its Class. (some Class
        // variants remain to be implemented)
        // Note that this function follows the convention
        // of a recursive call with a new Dictionary
        // dynamic when none is supplied. Duplication
        // of the basic function structure should ensure
        // this pattern is followed by all specialized
        // variants. While this should not usually be
        // necessary under normal usage (dispatch to
        // specialized variants from here), it should
        // help accommodate the possibility of direct
        // ' calls from other client functions.
        // 

        var rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDef(CpDef, new Dictionary());
        else if (CpDef == null)
        {
        }
        else
        {
            {
                rt.Add("bomStr", CpDef.BOMStructure);
                switch (CpDef.BOMStructure)
                {
                    case kNormalBOMStructure:
                        rt.Add("Type", "M");
                        break;
                    case kPurchasedBOMStructure:
                        rt.Add("Type", "R");
                        break;
                    case kDefaultBOMStructure:
                    case kPhantomBOMStructure:
                    case kReferenceBOMStructure:
                    case kInseparableBOMStructure:
                    case kVariesBOMStructure:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                {
                    var withBlock1 = nuAiBoxData().UsingBox(CpDef.RangeBox) // .SortingDims
                        ;
                    {
                        var withBlock2 = withBlock1.UsingInches() // WARNING[2021.12.15]
                            ;
                        // Forcing inch conversion MAY lead
                        // to issues in future development.
                        // It is absolutely ESSENTIAL that
                        // unit measurement be tracked and
                        // kept consistent throughout the
                        // entire management process.
                        rt.Add(pnLength, Round(withBlock2.SpanX, 6));
                        rt.Add(pnWidth, Round(withBlock2.SpanY, 6));
                        rt.Add("Height", Round(withBlock2.SpanZ, 6));
                    }
                }
            }

            switch (CpDef)
            {
                case SheetMetalComponentDefinition:
                    rt = dcGnsInfoCompDefShtMtl(CpDef, rt);
                    break;
                case WeldmentComponentDefinition:
                case WeldsComponentDefinition:
                    Debugger.Break(); // using general Assembly handler
                    rt = dcGnsInfoCompDefAssy(CpDef, rt);
                    break;
                case PartComponentDefinition:
                    rt = dcGnsInfoCompDefPart(CpDef, rt);
                    break;
                case AssemblyComponentDefinition:
                    rt = dcGnsInfoCompDefAssy(CpDef, rt);
                    break;
                default:
                    break;
            }
        }

        return rt;
    }

    public Dictionary dcGnsInfoCompDefShtMtl(SheetMetalComponentDefinition CpDef, Dictionary dcWkg = null)
    {
        // dcGnsInfoCompDefShtMtl -- Generate and/or populate Dictionary
        // (new or supplied) with data for Genius
        // from supplied ComponentDefinition.
        // This is the Assembly dynamic.
        // '
        // 

        var rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefShtMtl(CpDef, new Dictionary());
        else
        {
            rt = dcGnsInfoCompDefPart(CpDef, rt);
            string s6;
            Dictionary rm;
            {
                if (!rt.Exists("SPEC06"))
                {
                    Debugger.Break();
                    rt.Add("SPEC06", steelSpec6("", 1));
                }

                s6 = rt.get_Item("SPEC06");

                if (!rt.Exists("RMLIST"))
                    rt.Add("RMLIST", new Dictionary());
                rm = dcOb(rt.get_Item("RMLIST"));
            }

            {
                double tk = CpDef.Thickness.Value / (double)cvLenIn2cm; // thickness
                // NOTE conversion to Inches from Centimeters.
                // keep in mind we're grabbing Thickness HERE
                // and will use Height (below) in an effort
                // to validate the Flat Pattern, and determine
                // this Part is MEANT to be Sheet Metal.

                double wd; // width
                double ck; // check height vs thickness
                double lg; // length
                double ht; // height
                double ar; // area
                if (CpDef.HasFlatPattern)
                {
                    {
                        var withBlock1 = CpDef.FlatPattern;
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

                    ck = Round(double.Abs(ht - tk), 6);
                    if (ck > 0.002)
                    {
                        {
                            var withBlock1 = dcFlatPatSpansByVertices(CpDef.FlatPattern);
                            if (ht > withBlock1.get_Item("Z"))
                            {
                                ht = withBlock1.get_Item("Z");

                                Debug.Print.get_Item("X");
                                Debug.Print.get_Item("Y");
                                Debug.Print(""); // Breakpoint Landing
                            }
                        }
                    }

                    if (Round(double.Abs(ht - tk), 6) > 0)
                    {
                    }

                    ar = lg * wd; // .SpanX * .SpanY
                }
                else
                {
                    {
                        if (rt.Exists(pnLength))
                            lg = rt.get_Item(pnLength);

                        if (rt.Exists(pnWidth))
                            wd = rt.get_Item(pnWidth);

                        if (rt.Exists("Height"))
                            ht = rt.get_Item("Height");
                    }

                    ar = 0; // STOPGAP[2021.12.08]
                }

                // At this point, we should have either
                // likely dimensions of the flat pattern, OR
                // the original dimensions of the part itself.
                // 
                // The next step is to determine whether they
                // are consistent with a valid sheet metal part.
                // If not, it's likely a structural one.
                // 
                // The key criterion is how closely the height
                // dimension matches the given thickness.
                // 
                ck = Round(double.Abs(ht - tk), 6);
                if (ck > 0.002)
                {
                }

                // REV[2021.12.15]:
                // add material option collection
                // specific to sheet metal
                {
                    var withBlock1 = dcGnsMatlOps(dcCtOfEach(new[] { tk, lg, wd, ht }), s6);
                    foreach (var ky in withBlock1.Keys)
                    {
                        if (!rm.Exists(ky))
                            rm.Add(ky, withBlock1.get_Item(ky));
                    }
                }

                // 
                // 
                {
                    // first, remove any previous
                    // dimensional values
                    if (rt.Exists(pnLength))
                        rt.Remove(pnLength);
                    if (rt.Exists(pnWidth))
                        rt.Remove(pnWidth);
                    if (rt.Exists("Height"))
                        rt.Remove("Height");
                    // (not sure if this is the best way
                    // but going to try it for now)

                    rt.Add(pnThickness, tk);
                    rt.Add(pnLength, lg);
                    rt.Add(pnWidth, wd);
                    rt.Add(pnArea, ar);
                    rt.Add("Height", ht);
                }
            }
        }

        return rt;
    }

    public Dictionary dcFlatPatSpansByVertices(FlatPattern smFlat)
    {
        // dcFlatPatSpansByVertices -- get extents of
        // Sheet Metal Flat Pattern
        // from a scan of its Vertices.
        // this is a last resort,
        // in case an erroneous Z span
        // reported from the Range Box
        // fails to match Thickness.
        // 
        double xmn;
        double xmx;
        double ymn;
        double ymx;
        double zmn;
        double zmx;

        var rt = new Dictionary();

        if (!smFlat.Body == null)
        {
            {
                var withBlock = smFlat.Body // .Vertices'.RangeBox
                    ;
                foreach (Vertex vx in withBlock.Vertices)
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
            rt.Add("X", xmx - xmn);
            rt.Add("Y", ymx - ymn);
            rt.Add("Z", zmx - zmn);
        }

        return rt;
    }

    public Dictionary dcGnsInfoCompDefPart(PartComponentDefinition CpDef, Dictionary dcWkg = null)
    {
        // dcGnsInfoCompDefPart -- Generate and/or populate Dictionary
        // (new or supplied) with data for Genius
        // from supplied ComponentDefinition.
        // This is the general Part dynamic.
        // It's grown somewhat in complexity
        // since development's begun.
        // Here is rough flow map:
        // (inside component definition)
        // - stage supplied Dictionary for return
        // (start a new one, if none supplied
        // one USUALLY should be)
        // - get Mass -- don't add to Dictionary yet
        // (no data should be added to Dictionary
        // until all data are collected and verified)
        // - get Active Material, and its name
        // - use this to set target Spec 6
        // (inside returning Dictionary)
        // - collect length, width, and height
        // dimensions from Dictionary
        // (this is why one should be supplied)
        // - collect raw material candidate items
        // from Genius into a Recordset, using
        // an SQL query generated from collected
        // dimensions, and material Spec 6
        // - generate Dictionary of candidates
        // from the Recordset, keyed on item names
        // - add data to Dictionary:
        // - mass
        // - material name
        // - spec 6
        // - Dictionary of raw material
        // item candidates
        // 
        Dictionary d2; // may be temporary
        double ck;

        var rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefPart(CpDef, new Dictionary());
        else
        {
            string mtName;
            string s6;
            double ms; // mass
            {
                {
                    var withBlock1 = CpDef.MassProperties;
                    ms = Round(withBlock1.Mass * cvMassKg2LbM, 4);
                }

                // ptNumShtMetal
                MaterialAsset mt = aiDocPart(CpDef.Document).ActiveMaterial;
                mtName = mt == null ? "" : mt.DisplayName;
                s6 = steelSpec6(mtName);
            }

            {
                var wk = new Dictionary(); // may be temporary
                foreach (var ky in new[] { pnLength, pnWidth, "Height" })
                {
                    if (rt.Exists(ky))
                        wk.Add(ky, Round(Convert.ToDouble(rt.get_Item(ky)), 6));
                }

                wk = dcCtOfEach(wk.Items);
                if (wk.Count == 0)
                    wk.Add(0.075, 1);
                // another kludge to trap an error
                // which should NOT occur as long as
                // a prepared Dictionary is supplied.

                // Here is where we'll attempt to collect
                // raw material Item candidates from Genius
                // 

                // present setup -- plan to change
                // "select d.v "
                // Debug.Print "from (values (" & txDumpLs(wk.Keys, "), (") & "))"
                // " as d(v)"

                // future proposal -- counts occurrences
                // "select d.v, d.c "
                // Debug.Print "from (values (" & dumpLsKeyVal(wk, ", ", "), (") & "))"
                // " as d(v, c)"

                wk = dcGnsMatlOps(wk, s6);
                // REV[2021.12.15]:
                // preceding line replaces With block below,
                // moving Genius material options request
                // to function dcGnsMatlOps, so it can be
                // called from other functions, like,
                // again, dcGnsInfoCompDefShtMtl
                // With cnGnsDoyle()
                // Dim rs As ADODB.Recordset
                // 
                // 'wk.RemoveAll
                // 
                // 
                // Err().Clear
                // rs = .Execute(' sqlOf_GnsMatlOptions(' s6, wk.Keys' ))
                // 
                // If Err().Number = 0 Then
                // With dcFromAdoRS(rs, "") wk =
                // For Each ky In .Keys
                // d2 = dcOb(.get_Item(ky))
                // If d2 Is Nothing Then
                // Stop
                // Else
                // wk.Add d2.get_Item("Item"), d2
                // End If
                // 'Stop
                // ''' ENDOFDAY[2021.12.08]:
                // ''' Need to setup process of remapping
                // ''' raw material Items from Genius
                // ''' to their Item names
                // Next: End With
                // 
                // rs.Close
                // Else
                // Stop
                // Err().Clear
                // End If
                // 
                // 
                // .Close
                // End With

                // .Add pnRawMaterial, wk

                rt.Add(pnMass, ms);
                rt.Add(pnMaterial, mtName);
                rt.Add("SPEC06", s6);

                // If False Then
                // not quite ready for this one yet
                rt.Add("RMLIST", wk);
                // End If
                Debug.Print(""); // Breakpoint Landing
            }
        }

        return rt;
    }

    public Dictionary dcGnsInfoCompDefAssy(AssemblyComponentDefinition CpDef, Dictionary dcWkg = null)
    {
        // dcGnsInfoCompDefAssy -- Generate and/or populate Dictionary
        // (new or supplied) with data for Genius
        // from supplied ComponentDefinition.
        // This is the general Assembly dynamic.
        // '
        // 

        var rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefAssy(CpDef, new Dictionary());
        else
        {
            {
                var withBlock1 = CpDef.MassProperties;
                rt.Add(pnMass, Round(withBlock1.Mass * cvMassKg2LbM, 4));
            }
        }

        return rt;
    }

    public Dictionary dcGnsInfoCompDefTBD(ComponentDefinition CpDef, Dictionary dcWkg = null)
    {
        // dcGnsInfoCompDefTBD (formerly d2g4f2zz)
        // Generate and/or populate Dictionary
        // (new or supplied) with data for Genius
        // from supplied ComponentDefinition.
        // This is the <TBD> dynamic. (formerly <zz>)
        // Use it as a template for others.
        // (be sure to modify comments accordingly)
        // 

        var rt = dcWkg;
        if (rt == null)
            rt = dcGnsInfoCompDefTBD(CpDef, new Dictionary());
        else
        {
            var withBlock = CpDef;
        }

        return rt;
    }

    public Dictionary dcGnsInfoSQLitem(string Item)
    {
        // dcGnsInfoSQLitem -- Return a Dictionary
        // of Part data from Genius
        // for the indicated Item
        // 
        Dictionary rt;

        {
            var withBlock = cnGnsDoyle();
            Information.Err().Clear();
            var rs = withBlock.Execute(sqlOf_GnsPartInfo(Item));
            if (Information.Err().Number == 0)
                rt = dcFromAdoRSrow(rs, "");
            else
            {
                Debug.Print(Information.Err().Number);
                Debug.Print(Information.Err().Description);
                Debugger.Break();
            }

            rs = withBlock.Execute(sqlOf_GnsPartMatl(Item)); // sqlOf_ASDF
            if (Information.Err().Number == 0)
            {
                var mt = dcFromAdoRS(rs, "");
                {
                    if (mt.Count > 0)
                    {
                        if (mt.Count > 1)
                        {
                            Debugger.Break(); // to handle multiple raw materials
                            Debug.Print(""); // Breakpoint Landing
                        }
                        else
                        {
                            var withBlock2 = dcOb(mt.get_Item(mt.Keys(0)));
                            foreach (var ky in withBlock2.Keys)
                            {
                                if (rt.Exists(ky))
                                    Debugger.Break(); // to deal with collision
                                else
                                    rt.Add(ky, withBlock2.get_Item(ky));
                            }
                        }
                    }
                }
            }
            else
            {
                Debug.Print(Information.Err().Number);
                Debug.Print(Information.Err().Description);
                Debugger.Break();
            }

            withBlock.Close();
        }

        return rt;
    }

    public Dictionary d2g3f4(PartDocument Part, Dictionary dc = null)
    {
        // d2g3f4 -- Return a Dictionary
        // of Properties and info from
        // Inventor Part Document for
        // Genius Interface.
        // 

        var rt =
            // Dim rs As ADODB.Recordset
            dcProps4genius(Part, d2g3f1(Part, dc), 0);

        return rt;
    }

    public Dictionary d2g3f5(Document AiDoc)
    {
        // d2g3f5 -- Gather Dictionaries of Inventor
        // Properties and Genius info from supplied
        // Document for correlation and potential
        // revision.
        // REV[2021.12.15]:
        // Parameter Part renamed to AiDoc, with Class
        // changed from PartDocument to the more general
        // Document, as it would appear all supporting
        // functions will accept and work with it.
        // 
        // base Genius info and inherent Properties
        // base + custom Genius Properties
        // values of all collected Properties
        // values of all collected Properties
        // information from Genius database

        var dcPt = dcGnsInfoAiDocBase(AiDoc);
        Dictionary dcVlAi = dcMapAiProps2vals(dcPt);

        // REV[2021.12.16]:
        // additional value Dictionary
        // collects values ONLY from
        // the inherent document data
        // and Properties, which gets
        // overridden at the next step:
        var dcPr = dcProps4genius(AiDoc, dcCopy(dcPt), 2);

        // REV[2021.12.15] argument 2 replaces 0 in order
        // to generate references to missing Properties
        // (see dcGnsPropsListed) without trying to
        // create them. That way, client functions may
        // be made aware of Properties that need created.
        // 
        // Modifications to those functions might be needed
        // to be prepared for missing Properties, whose
        // names will map to Nothing (void references)
        // UPDATE[2021.12.16]:
        // just happened today. created new function
        // blankIfNoValElseSelf to address this issue.
        // see dcMapAiProps2vals for application
        // 
        // = dcProps4genius(AiDoc, d2g3f1(AiDoc, dc), 0)
        // = d2g3f4(AiDoc)
        Dictionary dcVlPr = dcMapAiProps2vals(dcPr); // dcPt
        {
            foreach (var ky in new[] { pnThickness, pnWidth, pnLength, pnArea, pnRmQty }) // '' THIS IS A KLUDGE
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
                // REV[2021.12.16]:
                // added pnRmQty to list as quick and
                // dirty method to force a blank value
                // to zero, and prevent an error in
                // the correction code after the loop.
                // 
                // This must be the sort of cruft Joel
                // Spolsky was talking about in that
                // essay of his. Still not a justified
                // defense for crud programming, which
                // this is, let's face it! Right there
                // with you, Brando!
                // 
                if (dcVlPr.Exists(ky))
                    dcVlPr.get_Item(ky) = Val(Split("0" + Convert.ToHexString(dcVlPr.get_Item(ky)), " ")(0));
            }

            if (dcVlPr.Exists(pnRmQty))
                dcVlPr.get_Item(pnRmQty) = Round(dcVlPr.get_Item(pnRmQty), 8);
        }

        Dictionary dcGn = dcGnsInfoSQLitem(dcVlPr.get_Item(pnPartNum));

        var rt = new Dictionary();
        {
            rt.Add("aiVal", dcVlAi);
            rt.Add("inv", dcVlPr);
            rt.Add("gns", dcGn);
            rt.Add("prp", dcPr);
        }

        return rt;
    }

    public Dictionary d2g3f5as(AssemblyDocument Assy, long ThisToo = 0)
    {
        // d2g3f5as -- Assembly counterpart to d2g3f5
        // not sure what's actually to be done with it yet.
        // probably just remove it; d2g3f5 can handle both.
        // 

        var dc = dcRemapByPtNum(dcAiDocComponents(Assy,
            null, ThisToo));
    }

    public dynamic dcMapAiProps2vals(Dictionary dc, long Flags = 0)
    {
        // dcMapAiProps2vals --
        // Return a Dictionary
        // containing the Values of
        // any Inventor Properties
        // in supplied Dictionary,
        // with all other members
        // returned as they are.
        // 
        // related functions:
        // dcOfDcAiPropVals
        // dcAiPropValsFromDc
        // dcOfPropsInAiDoc
        // 
        Property pr;

        var rt = new Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, blankIfNoValElseSelf(valIfAiPropElseSelf(withBlock.get_Item(ky))));
        }
        return rt;
    }

    public dynamic valIfAiPropElseSelf(dynamic vl)
    {
        // valIfAiPropElseSelf --
        // Return the Value of any
        // supplied Inventor Property.
        // Any other type of argument
        // should be returned directly.
        // 

        if (!IsObject(vl)) return vl;
        Property pr = aiProperty(obOf(vl));
        return pr == null ? vl : pr.Value;
    }

    public dynamic blankIfNoValElseSelf(dynamic vl)
    {
        // blankIfNoValElseSelf --
        // Return the Value of any
        // supplied Inventor Property.
        // Any other type of argument
        // should be returned directly.
        // 
        Property pr;

        return vl switch
        {
            not null => obOf(vl) == null ? "" : vl,
            null => "",
            _ => IsEmpty(vl) ? "" : vl
        };
    }

    public Dictionary d2g3f7(Document AiDoc)
    {
        // d2g3f7 --
        // 
        Dictionary rt;

        // rt = New Scripting.Dictionary
        {
            var withBlock = d2g3f5(AiDoc);
            Debug.Print(""); // Breakpoint Landing
            rt = dcTreeReKeyedInPlc("src1", "gns",
                dcTreeReKeyedInPlc("src0", "inv",
                    dcWBQbyCmpResult(dcCmpTextOf2dc(withBlock.get_Item("inv"), withBlock.get_Item("gns")))));

            {
                var withBlock1 = rt;
            }
            rt.Add("prp", withBlock.get_Item("prp"));
            rt.Add("doc", AiDoc);
        }
        return rt;
    }

    public Dictionary d2g3f8(Document AiDoc = null)
    {
        // d2g3f8 --
        // 
        Dictionary rt;
        // Dim ck As Inventor.Document

        if (AiDoc == null)
        {
            {
                var withBlock = ThisApplication;
                if (withBlock.ActiveDocument == null)
                    Debugger.Break();
                else
                    rt = d2g3f8(withBlock.ActiveDocument);
            }
        }
        else
        {
            rt = new Dictionary();

            {
                var withBlock = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(dcAiDocComponents(AiDoc));
                // 
                {
                    var withBlock1 = withBlock.dcIn() // Parts
                        ;
                    foreach (var ky in withBlock1.Keys)
                        rt.Add(ky, d2g3f7(aiDocPart(obOf(withBlock1.get_Item(ky)))));
                }

                {
                    var withBlock1 = withBlock.dcOut() // Assemblies
                        ;
                }
            }
        }

        return rt;
    }

    public Dictionary dcTreeMembersWithKey(dynamic tg, Dictionary dc, Dictionary wk = null)
    {
        // dcTreeMembersWithKey (formerly d2g5f1)
        // Given a Dictionary that might contain
        // other Dictionaries, check it and any
        // sub Dictionaries for target key (tg)
        // and return a Dictionary of those
        // Dictionaries containing it, each
        // keyed to the number already found.
        // This should ensure a unique key
        // for each match found, with no
        // need to track any other keys.
        // 
        // The ultimate goal of this function is to
        // support a Key Find/Replace operation
        // across a hierarchy of Dictionaries.
        // 
        // This is initially and specifically to map
        // comparison keys "src0" and "src1" to
        // the names of sources they represent.
        // 
        // This is of course the 'Find' component
        // of the ultimate product
        // 
        Dictionary rt;

        if (wk == null)
            rt = dcTreeMembersWithKey(tg, dc, new Dictionary());
        else
        {
            rt = wk;
            if (!dc != null) return rt;
            {
                if (dc.Exists(tg))
                {
                    {
                        rt.Add.Count(null, dc);
                    }
                }

                foreach (var ky in dc.Keys)
                    rt = dcTreeMembersWithKey(tg, dcOb(obOf(dc.get_Item(ky))), rt);
            }
        }

        return rt;
    }

    public Dictionary dcTreeMemWithReplcmt(dynamic rp, Dictionary dc)
    {
        // dcTreeMemWithReplcmt (formerly d2g5f2)
        // Given a Dictionary of Dictionaries,
        // check for any Dictionary containing target
        // replacement Key rp, and return a Dictionary
        // containing any results.
        // 
        // This is a check for potential Key collisions.
        // The Dictionary returned should be null.
        // 
        // This is presently accomplished by first calling
        // dcTreeMembersWithKey against the supplied Dictionary,
        // which is normally expected to be the result
        // of a PRIOR call to dcTreeMembersWithKey using the target
        // key to be replaced.
        // 
        // It is therefore possible that the supplied
        // Dictionary might contain replacement key rp,
        // and thus be included in the local result.
        // That Dictionary should NOT be included
        // in the FINAL result returned.
        // 
        // It is therefore necessary to scan the result
        // of the local dcTreeMembersWithKey call, and remove it,
        // if found.
        // 

        var rt = dcTreeMembersWithKey(rp, dc);
        {
            foreach (var ky in rt.Keys)
            {
                if (dcOb(obOf(rt.get_Item(ky))) != dc) continue;
                Debugger.Break();
                rt.Remove(ky);
            }
        }
        return rt;
    }

    public Dictionary dcTreeReKeyedInPlc(dynamic tg, dynamic rp, Dictionary dc)
    {
        // dcTreeReKeyedInPlc (formerly d2g5f3)
        // Given a target Key tg (to be replaced),
        // a replacement Key rp, and a Dictionary
        // that includes other Dictionaries,
        // attempt to replace all instances of the
        // target Key with the replacement Key in
        // all Dictionaries within the hierarchy.
        // 
        // Note that this is a DESTRUCTIVE replacement
        // operation. A preferable option might be
        // to generate a NEW hierarchical Dictionary
        // replicating the original, with the desired
        // key substitution. Will consider that for
        // a later implementation.
        // 
        // Note also that error checking/handling
        // in this implementation is presently minimal.
        // A more robust process should also be considered.
        // 

        var wk = dcTreeMembersWithKey(tg, dc);
        var ck = dcTreeMemWithReplcmt(rp, wk);

        if (ck.Count > 0)
            Debugger.Break();
        else
        {
            foreach (var ky in wk.Keys)
            {
                {
                    var withBlock1 = dcOb(obOf(wk.get_Item(ky)));
                    // A Dictionary dynamic is assumed, here.
                    // Though typically risky in a With block,
                    // it SHOULD be guaranteed here,
                    // so no error should occur.
                    // Don't be surprised if it does, though.
                    if (withBlock1.Exists(rp))
                        Debugger.Break(); // because this
                    else
                    {
                        // a proper error handler might
                        // be desired here in future
                        // 
                        // for now, keep disabled

                        // note order of operations here
                        withBlock1.Add(rp, withBlock1.get_Item(tg));
                        withBlock1.Remove(tg);
                    }
                }
            }
        }

        return dc;
    }

    public dynamic userChoiceFromDc(dynamic dcAs = Dictionary == null, dynamic ifNone = null)
    {
        // userChoiceFromDc (formerly d2g3f2)
        // Request User Selection from
        // a Dictionary of options.
        // 
        // A list of Dictionary Keys is
        // presented to the user. After
        // User selects a Key, matching
        // Item is returned for use.
        // 

        // REV[2023.05.17.1304]
        // add ifNone processing to present
        // User with information on default
        // option(s), if supplied

        Information.Err().Clear();
        var msNoSel = Convert.ToHexString(ifNone);

        if (Information.Err().Number == 0)
        {
            if (Strings.Len(msNoSel) > 0)
                msNoSel = "Use default value (" + msNoSel + ")?";
        }
        else
        {
            msNoSel = "";
            Information.Err().Clear();

            if (ifNone is not null)
            {
                if (false)
                    msNoSel = Join(
                        new[] { "Use default " + TypeName(ifNone) + " dynamic?", "(dynamic details not available)" },
                        Constants.vbCrLf);
            }
            else
                Debugger.Break();
        }

        if (Strings.Len(msNoSel) > 0)
            msNoSel = Constants.vbCrLf + msNoSel;
        // msNoSel = Join(new [] {"User selection was requested", "with no available options!", msNoSel), vbCrLf)

        dynamic[] rt;
        if (dc == null)
            rt = [userChoiceFromDc(dcAiDocsVisible())];
        else if (dc.Count > 0)
        {
            var rp = nuSelFromDict(dc).GetReply();
            // , , , , ,, Join(new [] {"No option selected!",msNoSel), vbCrLf)
            if (dc.Exists(rp))
                rt = new[] { dc.get_Item(rp) };
            else
                rt = [ifNone];
        }
        else
        {
            MsgBoxResult ck =
                MessageBox.Show(
                    Join(new[] { "User selection was requested", "with no available options!" }, Constants.vbCrLf),
                    Constants.vbOKOnly, "No Options!");
            // IIf(Len(msNoSel) > 0,vbYesNo, vbOKOnly),msNoSel
            // 
            if (ck == Constants.vbNo)
                rt = [null];
            else
                rt = [ifNone];
        }

        if (IsObject(rt(0)))
            return rt(0);
        return rt(0);
    }

    public Dictionary dcGnsPrpPtDvl_2021_1112(PartDocument invDoc, Dictionary dc = null)
    {
        // '
        string aiSubType;
        // '
        // Dim aiPropsUser As Inventor.PropertySet
        PropertySet aiPropsDesign;
        // '
        Property prPartNum;
        Property prFamily;
        // '
        // Dim aiPartNum As String 'will be same as gnPartNum
        // Dim aiPartFam As String
        // Dim aiMatlNum As String
        // Dim aiMatlFam As String
        // Dim aiMatlQty As Double
        // Dim aiQtyUnit As String
        BOMStructureEnum aiBomType;
        // '
        // '
        var rt = new Dictionary();
        var dc01 = new Dictionary();

        {
            // Property Sets
            {
                var withBlock1 = invDoc.PropertySets;
                // aiPropsUser = .get_Item(gnCustom)
                aiPropsDesign = withBlock1.get_Item(gnDesign);
            }
            aiBomType = invDoc.ComponentDefinition.BOMStructure;
            aiSubType = invDoc.SubType;
        }

        // Part Number and Family
        // Properties from Design set
        {
            prPartNum = aiPropsDesign.get_Item(pnPartNum);
            prFamily = aiPropsDesign.get_Item(pnFamily);
        }

        // Values of Part Number
        // and Family Properties
        string aiPartNum = prPartNum.Value;
        string aiFamily = prFamily.Value;
        // NOTE[2021.11.12]
        // The preceding three sections
        // can PROBABLY be consolidated
        // into one, using fewer variables
        // and probably just one With block
        // dc01
        {
            var withBlock = cnGnsDoyle();
            Information.Err().Clear();
            var rs = withBlock.Execute(sqlOf_ASDF(aiPartNum));
            if (Information.Err().Number == 0)
            {
                var dcVlGn = dcFromAdoRSrow(rs, "");
                {
                    var withBlock1 = dcVlGn;
                }

                {
                    if (rs.BOF & rs.EOF)
                    {
                    }
                    else
                    {
                        {
                            var withBlock2 = rs.Fields;
                        }

                        rs.MoveNext();
                        if (!rs.EOF)
                        {
                            Debugger.Break(); // to handle multiple raw materials
                            Debug.Print(""); // Breakpoint Landing
                        }
                    }

                    rs.Close();
                }
            }
            else
            {
                Debug.Print(Information.Err().Number);
                Debug.Print(Information.Err().Description);
                Debugger.Break();
            }

            withBlock.Close();
        }

        switch (aiBomType)
        {
            case kNormalBOMStructure:
            {
                if (aiSubType == guidSheetMetal)
                {
                }

                break;
            }
            case kPurchasedBOMStructure:
            case kPhantomBOMStructure:
            case kInseparableBOMStructure:
            case kReferenceBOMStructure:
                Debugger.Break();
                break;
            case kDefaultBOMStructure:
            case kVariesBOMStructure:
            default:
            {
                if (aiBomType is kNormalBOMStructure or kPhantomBOMStructure or kDefaultBOMStructure
                    or kVariesBOMStructure or kDefaultBOMStructure)
                {
                }

                Debugger.Break();
                break;
            }
        }

        // '
        return rt;
    }

    public Dictionary dcGeniusPropsPartRev20180530_broken2(PartDocument invDoc, Dictionary dc = null)
    {
        while (true)
        {
            // Dim dcPr As Scripting.Dictionary
            // '
            PropertySet aiPropsUser;
            PropertySet aiPropsDesign;
            // '
            Property prPartNum; // pnPartNum
            // ADDED[2021.03.11] to simplify access
            // to Part Number of Model, since it's
            // requested several times in function
            Property prFamily;
            Property prRawMatl; // pnRawMaterial
            Property prRmUnit; // pnRmUnit
            Property prRmQty; // pnRmQty
            // '
            // UPDATE[2021.11.08] MAJOR CHANGE
            // Overhaul variable names to better
            // reflect TWO distinct value sets
            // one from Genius
            // another from Inventor
            // in order to better compare
            // and synchronize them.
            // 
            // First set are the Genius variables:
            // the original set, renamed en masse:
            // 
            string gnPartFam; // was ptFamily
            string gnMatlNum; // was pnStock
            string gnMatlFam; // was mtFamily
            double gnMatlQty; // was qtRawMatl
            string gnQtyUnit; // was qtUnit
            // Second set are the new Inventor variables.
            // These should replace the Genius instances
            // anywhere their values are taken
            // from the model.
            // 
            double aiMatlQty;
            string aiQtyUnit;
            // 
            // '
            aiBoxData bd;
            // UPDATE[2021.11.03]:
            // 
            // 
            ADODB.Field fdItem;
            ADODB.Field fdFamly;
            ADODB.Field fdOrder;
            ADODB.Field fdMatrl;
            ADODB.Field fdMtFam;
            ADODB.Field fdQty;
            ADODB.Field fdUnit;

            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            var aiBomType = invDoc.ComponentDefinition.BOMStructure;
            // UPDATE[2021.11.11]
            // Moved Property collection
            // to top of program to permit
            // collection of Design Properties
            // in second step. Also pulled up
            // BOM Structure capture (above)
            // along with the Values of
            // of Design Properties.

            {
                {
                    var withBlock1 = invDoc.PropertySets;
                    aiPropsDesign = withBlock1.get_Item(gnDesign);
                    aiPropsUser = withBlock1.get_Item(gnCustom);
                }

                aiBomType = invDoc.ComponentDefinition.BOMStructure;

                if (aiBomType == kNormalBOMStructure)
                {
                    if (invDoc.SubType == guidSheetMetal)
                    {
                    }
                }
            }

            // Part Number and Family properties
            // are from Design, NOT Custom set
            {
                // so we can grab them directly
                prPartNum = aiPropsDesign.get_Item(pnPartNum);
                prFamily = aiPropsDesign.get_Item(pnFamily);
            }
            string aiPartNum = prPartNum.Value; // will be same as gnPartNum
            string aiPartFam = prFamily.Value;

            var dcProp = dcGnsPropsPart(invDoc, null, 0); // dcAiPropsInSet
            var dcVlPr = new Dictionary();
            Property pr;
            string aiMatlNum;
            {
                dcProp.Add(pnPartNum, prPartNum);
                dcProp.Add(pnFamily, prFamily);
                foreach (var ky in dcProp.Keys)
                {
                    pr = aiProperty(dcProp.get_Item(ky));
                    if (pr == null)
                        Debugger.Break();
                    else
                        dcVlPr.Add(ky, pr.Value);
                }

                if (dcProp.Exists(pnRawMaterial))
                {
                    prRawMatl = dcProp.get_Item(pnRawMaterial);
                    aiMatlNum = prRawMatl.Value;
                }
                else
                    aiMatlNum = "";

                if (dcProp.Exists(pnRmUnit))
                {
                    prRmUnit = dcProp.get_Item(pnRmUnit);
                    aiQtyUnit = prRmUnit.Value;
                }
                else
                    aiQtyUnit = "";

                if (dcProp.Exists(pnRmUnit))
                {
                    prRmQty = dcProp.get_Item(pnRmQty);
                    aiMatlQty = prRmQty.Value;
                }
                else
                    aiMatlQty = 0;
            }
            Debug.Print("=== Check Existing Model Genius Properties ===");
            Debug.Print(dumpLsKeyVal(dcVlPr, "="));
            Debug.Print("");
            Debugger.Break();

            // NOTE[2021.11.11]
            // Assignment of initial rt Dictionary
            // now essentially duplicates the new
            // process now preceding this section.
            // The only difference is, that version
            // does NOT apply Genius Property col-
            // lection to the supplied Dictionary dc.
            var rt = dcGnsPropsPart(invDoc, dc, 0);
            var dcVlAi = new Dictionary();
            {
                rt.Add(pnPartNum, prPartNum);
                rt.Add(pnFamily, prFamily);
                foreach (var ky in rt.Keys)
                {
                    pr = aiProperty(rt.get_Item(ky));
                    if (pr == null)
                        Debugger.Break();
                    else
                        dcVlAi.Add(ky, pr.Value);
                }

                pr = null;
            }
            // Ultimately, processes which populate
            // returned Dictionary rt, and set the
            // Properties it should receive, should
            // be moved toward the end of the function.

            Dictionary dcVlGn;
            BOMStructureEnum gnBomType;
            {
                var withBlock = cnGnsDoyle();
                // Pre-clear all relevant variables
                // to be set from query results,
                // if available.

                // gnPartNum = aiPartFam
                // gnPartFam = ""
                // fdOrder = .get_Item("Ord")
                // gnMatlNum = ""
                // gnMatlFam = ""
                // gnMatlQty = 0
                // gnQtyUnit = ""
                // gnBomType = kDefaultBOMStructure
                // use this to indicate no BOM type
                // or structure returned from Genius

                Information.Err().Clear();
                var rs = withBlock.Execute(sqlOf_ASDF(aiPartNum));
                if (Information.Err().Number == 0)
                {
                    dcVlGn = dcFromAdoRSrow(rs, "");
                    {
                        string gnPartNum = dcVlGn.get_Item("Item"); // was pnModel
                        gnPartFam = dcVlGn.get_Item("Family");
                        gnBomType = dcVlGn.get_Item("bomStr");
                        // fdOrder = .get_Item("Ord")
                        gnMatlNum = dcVlGn.get_Item("Material");
                        gnMatlFam = dcVlGn.get_Item("MtFamily");
                        gnMatlQty = dcVlGn.get_Item("Qty");
                        gnQtyUnit = dcVlGn.get_Item("Unit");
                    }

                    {
                        if (rs.BOF & rs.EOF)
                        {
                        }
                        else
                        {
                            {
                                var withBlock2 = rs.Fields;
                            }

                            rs.MoveNext();
                            if (!rs.EOF)
                            {
                                Debugger.Break(); // to handle multiple raw materials
                                Debug.Print(""); // Breakpoint Landing
                            }
                        }

                        rs.Close();
                    }
                }
                else
                {
                    Debug.Print(Information.Err().Number);
                    Debug.Print(Information.Err().Description);
                    Debugger.Break();
                }

                withBlock.Close();
            }

            Debug.Print("== Prop Check ==");
            Debug.Print("---- Genius ----");
            Debug.Print(dumpLsKeyVal(dcVlGn, "="));
            Debug.Print("--- Inventor ---");
            Debug.Print(dumpLsKeyVal(dcVlPr, "="));
            Debug.Print("================");
            Debugger.Break();

            {
                // UPDATE[2021.11.11]
                // Moved Property collection
                // to top of program, along with
                // collection of Design Properties
                // and their values. BOM Structure
                // as well.

                // We should check HERE for possibly misidentified purchased parts
                // UPDATE[2021.11.08]
                // Another MAJOR overhaul, here:
                // change Purchased Parts identification
                // to defer to Genius. Only attempt to guess
                // when no value comes back from Genius.
                // Stop 'BKPT-2021-1108-1608
                // CHANGE NEEDED[2021.11.08]:
                // indeterminate -- stopping work @endOfDay
                // effort here is to separate collection
                // and potential reassignment of
                // based on Part's family, file location,
                // and whatever other criteria, if any.
                // 
                // Likely need a counterpart variable
                // which takes its value from the Model.
                // The most likely Genius equivalent is
                // probably the ItemType field in view
                // table vgMfiItems, which will need
                // translation.
                // 
                MsgBoxResult ck;
                if (gnBomType == kDefaultBOMStructure)
                {
                    // Genius didn't return an Item type
                    // or BOM structure. We need to get it here.

                    // BKPT-2021-1109-1042
                    // Checkpoint here. Verify desired
                    // behavior here prior to removal.
                    Debugger.Break();

                    // BOM Structure type, correcting if appropriate,
                    // and prepare Family value for part, if purchased.
                    // 
                    // UPDATE[2018.02.06]
                    // Using new UserForm; see below
                    // UPDATE[2018.05.31]
                    // Combined both InStr checks by addition
                    // to generate a single test for > 0
                    // If EITHER string match succeeds, the total
                    // SHOULD exceed zero, so this SHOULD work.
                    // UPDATE[2021.11.08]
                    // Removed extraneoous code previously
                    // disabled under preceding update[2018.05.31]
                    // Also reseparated InStr checks previously combined
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
                            ck = newFmTest2().AskAbout(invDoc, null, "Is this a Purchased Part?");
                        else
                            ck = Constants.vbNo;

                        // Stop 'BKPT-2021-1105-0942
                        // CHANGE NEEDED[2021.11.05]:
                        // ONLY COLLECT desired BOMStructure here
                        // while keeping track of current value.
                        // Reassignment should take place along
                        // with collective property changes
                        // UPDATE[2021.11.09]
                        // This section now reduced to setting
                        // gnBomType from User response, if any.
                        // Code to assign Model BOM structure
                        // moved toward bottom for further work.
                        // 
                        // Check process below replaces duplicate check/responses above.
                        if (ck == Constants.vbYes)
                            gnBomType = kPurchasedBOMStructure;
                        else
                        {
                            gnBomType = aiBomType;
                            dcVlGn.get_Item("bomStr") = gnBomType;
                        }

                        // Request #2: Change Cost Center iProperty.
                        // If BOMStructure = Purchased and not content center,
                        // then Family = D-PTS, else Family = D-HDWR.
                        // 
                        // UPDATE[2018.05.30]: Value produced here
                        // will now be held for later processing,
                        // more toward the end of this function.
                        // UPDATE[2021.11.09]
                        // Changed to set target (Genius)
                        // Family, and ONLY if not set already.
                        // '
                        // MIGHT want to set up a more robust check
                        // system, but see how this holds up, first.
                        // '
                        if (Strings.Len(gnPartFam) == 0)
                        {
                            if (gnBomType == kPurchasedBOMStructure)
                            {
                                if (invDoc.IsContentMember)
                                {
                                    Debugger.Break(); // BKPT-2021-1105-0946
                                    gnPartFam = "D-HDWR";
                                }
                                else
                                {
                                    Debugger.Break(); // BKPT-2021-1105-0947
                                    gnPartFam = "D-PTS";
                                }
                            }
                        }
                    }
                }

                {
                    var withBlock1 = invDoc.ComponentDefinition;
                    // Request #1: the Mass in Pounds
                    // and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        // Stop 'BKPT-2021-1110-1551
                        // CHANGE NEEDED[2021.11.10]
                        // '
                        // '
                        // '
                        if (dcVlPr.Exists(pnMass))
                        {
                            if (Round(cvMassKg2LbM * withBlock2.Mass, 4) - Convert.ToDouble(dcVlPr.get_Item(pnMass)) ==
                                0)
                            {
                            }
                            else
                                // Stop
                                dcVlPr.get_Item(pnMass) = Round(cvMassKg2LbM * withBlock2.Mass, 4);
                        }
                        else
                            dcVlPr.Add(pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4));
                    }
                }
                // At this point, gnPartFam SHOULD be set
                // to a non-blank value if Item is purchased.
                // We should be able to check this later on,
                // if Item BOMStructure is NOT Normal

                // Stop 'BKPT-2021-1109-1053
                // HERE is where it starts to get interesting
                // Actually, just a little further down, where
                // Part SubType is checked for Sheet Metal.
                // At that point, the function divides into two
                // LONG, and possibly nearly identical branches.
                // Ideally, these should be refactored, with as
                // much of their processes as possible combined
                // into a single path.

                switch (aiBomType)
                {
                    // Request #4: Change Cost Center iProperty.
                    // If BOMStructure = Normal, then Family = D-MTO,
                    // else if BOMStructure = Purchased then Family = D-PTS.
                    case kNormalBOMStructure:
                    {
                        // Custom Properties
                        // Stop 'BKPT-2021-1105-1144
                        // CHANGE NEEDED[2021.11.05]:
                        // these properties should NOT
                        // be added immediately, but only
                        // when it's time to set them,
                        // towards the END of this function.
                        // UPDATE[2021.11.09]
                        // Custom Property collection/generation
                        // moved into Normal BOM Part handling, as
                        // no earlier usage appears to take place.
                        // '
                        // If possible, may wish to move even further.
                        // Plan to review later, as time permits.
                        // UPDATE[2021.11.10]
                        // Disabled Genius Property collection here
                        // since a Dictionary of ALL Genius Properties
                        // is generated towards the beginning.
                        // '
                        // '
                        // With rt
                        // If .Exists(pnRawMaterial) Then prRawMatl = .get_Item(pnRawMaterial)
                        // prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1)
                        // If .Exists(pnRmUnit) Then prRmUnit = .get_Item(pnRmUnit)
                        // prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                        // If .Exists(pnRmQty) Then prRmQty = .get_Item(pnRmQty)
                        // prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1)
                        // End With
                        // Collecting them at this point
                        // might still be appropriate,
                        // or it may be more desirable
                        // to hold off until later.
                        // '
                        // UPDATE[2021.11.08]
                        // Design Properties have been moved
                        // toward the top, as proposed.
                        // Commentary recommending this
                        // has been removed as extraneous.
                        // '
                        // UPDATE[2021.11.09]
                        // BOM Structure collection has also been
                        // moved up, alongside Design Properties,
                        // using renamed variable aiBomType
                        // (formerly bomStruct)
                        // '

                        // ----------------------------------------------------'
                        if (invDoc.SubType == guidSheetMetal)
                        {
                            // ----------------------------------------------------'
                            // Request #3:
                            // sheet metal extent area
                            // and add to custom property "RMQTY"
                            // UPDATE[2021.11.10]
                            // Now collecting Flat Pattern Values
                            // instead of the Properties for them.
                            // If necessary, Properties should be
                            // assigned in a separate function
                            // CHANGE NEEDED[2021.11.05]:
                            // not quite sure on this one yet,
                            // but dcFlatPatProps might need its
                            // own set of revisions to generate
                            // assignment recommendations WITHOUT
                            // performing them itself
                            // UPDATE[2021.11.10]
                            // Embedded Flat Pattern Property collection
                            // in bypassed If branch. Preceding Stop,
                            // when enabled, offers user/developer
                            // an opportunity to run it, if desired.
                            Debugger.Break(); // BKPT-2021-1105-1105
                            if (true)
                            {
                                var dcVlFP = dcFlatPatVals(invDoc.ComponentDefinition); // dcVlAi
                                {
                                    string aiMatlFam = dcVlFP.get_Item("mtFamily");
                                    dcVlFP.Remove("mtFamily");
                                    foreach (var ky in dcVlFP.Keys)
                                    {
                                        if (dcVlPr.Exists(ky))
                                        {
                                            if (Convert.ToHexString(dcVlPr.get_Item(ky)) ==
                                                Convert.ToHexString(dcVlFP.get_Item(ky)))
                                            {
                                            }
                                            else
                                                Debugger.Break();
                                        }
                                        else
                                            Debugger.Break();
                                    }
                                }
                                Debugger.Break();
                            }
                            else
                                rt = dcFlatPatProps(invDoc.ComponentDefinition, rt);
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
                                    // UPDATE[2018.05.30]:
                                    // Restoring original key check
                                    // and adding code for debug
                                    // Previously changed to "~OFFTHK"
                                    // to avoid this block and its issues.
                                    // (Might re-revert if not prepped to fix now)
                                    Debug.Print(aiProperty(rt.get_Item("OFFTHK")).Value);
                                    Debugger.Break(); // because we're going to need to do something with this.

                                    gnMatlNum = ""; // Originally the ONLY line in this block.
                                    // A more substantial response is required here.

                                    if (false) Debugger.Break(); // (just a skipover)
                                }
                                else if (Strings.Len(gnMatlNum) == 0)
                                {
                                    Debugger.Break(); // because we don't know IF this is sheet metal yet
                                    gnMatlNum = ptNumShtMetal(invDoc.ComponentDefinition);
                                }
                            }
                            else
                            {
                                // ' ACTION ADVISED[2018.09.14]:
                                // ' gnMatlNum can probably be set
                                // ' to prRawMatl.Value and THEN
                                // ' checked for length to see
                                // ' if lookup needed.
                                // ' This might also allow us to check
                                // ' for machined or other non-sheet
                                // ' metal parts.

                                Debugger.Break();
                                // !!!WARNING!!![2021.11.04]:
                                // Following section has been shuffled
                                // and should be considered HIGHLY
                                // UNSTABLE until verified functional
                                // and SAFE! TWO Stop commands are
                                // placed to emphasize the need for
                                // EXTREME CAUTION at this point
                                Debugger.Break();
                                // UPDATE[2021.11.04]:
                                // This section is being adjusted
                                // in an attempt to improve the raw
                                // material determination process.
                                // 
                                // This particular segment should
                                // ONLY be invoked if gnMatlNum is not
                                // successfully retrieved from Genius
                                // 
                                if (Strings.Len(gnMatlNum) == 0)
                                {
                                    // no stock retrieved from Genius
                                    // attempt to retrieve from Model
                                    // gnMatlNum = aiMatlNum

                                    if (Strings.Len(aiMatlNum) > 0)
                                    {
                                        // need to verify it against Genius
                                        // by retrieving its Family there
                                        // This With block copied and modified [2021.03.11]
                                        // from elsewhere in this function as a temporary measure
                                        // to address a stopping situation later in the function.
                                        // See comment below for details.
                                        // 
                                        // UPDATE[2021.11.04]:
                                        // This section MIGHT be removed in future,
                                        // 
                                        {
                                            var withBlock1 = cnGnsDoyle().Execute("select Family " +
                                                "from vgMfiItems " + "where Item='" + gnMatlNum + "';");
                                            if (withBlock1.BOF | withBlock1.EOF)
                                            {
                                                // Stop 'because Material value likely invalid
                                                Debugger.Break(); // because we do NOT want to set gnMatlNum!
                                                // want to assign it to a separate RETURN variable
                                                // or most likely, the return Dictionary.
                                                gnMatlNum = ptNumShtMetal(invDoc.ComponentDefinition);
                                                Debug.Print(""); // Breakpoint Landing
                                            }
                                            else
                                            {
                                                // ' This section retained from source,
                                                // ' but disabled to avoid potential issues
                                                // ' with subsequent operations, just in case
                                                // ' anything depends on gnMatlFam remaining
                                                // ' uninitialized up to that point.
                                                // UPDATE[2021.11.09]
                                                // Re-enabling Genius Material
                                                // Family assignment, as it SHOULD
                                                // be set to match what Genius
                                                // returns from this query.
                                                // '
                                                // Might not be the best place
                                                // to do this, though. If ptNumShtMetal
                                                // returns a valid Material Item above,
                                                // a Family is still needed.
                                                // '
                                                // NOTE: Fix disabled With block between runs
                                                // ' With .Fields
                                                Debugger.Break(); // because we do not want to set gnMatlFam
                                                // for same reasons as above
                                                gnMatlFam = withBlock1.Fields.get_Item("Family").Value;
                                            }
                                        }
                                    }
                                }

                                if (Strings.Len(gnMatlNum) == 0)
                                {
                                    // UPDATE[2018.05.30]:
                                    // Pulling ALL code/text from this section
                                    // to get rid of excessive cruft.
                                    // 
                                    // In fact, reversing logic to go directly
                                    // to User Prompt if no stock identified
                                    // 
                                    // IN DOUBLE FACT, hauling this WHOLE MESS
                                    // RIGHT UP after initial gnMatlNum assignment
                                    // to prompt user IMMEDIATELY if no stock found
                                    {
                                        var withBlock1 = newFmTest1();
                                        if (!(invDoc.ComponentDefinition.Document == invDoc)) Debugger.Break();

                                        bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
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
                                                    gnPartFam = withBlock2.get_Item(pnFamily);
                                                    Debug.Print(pnFamily + "=" + gnPartFam);
                                                }

                                                if (withBlock2.Exists(pnRawMaterial))
                                                {
                                                    gnMatlNum = withBlock2.get_Item(pnRawMaterial);
                                                    Debug.Print(pnRawMaterial + "=" + gnMatlNum);
                                                }
                                            }
                                            if (false) Debugger.Break(); // Use this for a debugging shim
                                        }
                                    }
                                }
                                else if (Left(gnMatlNum, 2) == "LG")
                                {
                                    Debug.Print(aiPartNum + ": PROBABLE LAGGING");
                                    Debug.Print(" TRY TO IDENTIFY, AND FILL IN BELOW.");
                                    Debug.Print(" PRESS ENTER ON gnMatlNum LINE WHEN");
                                    Debug.Print(" COMPLETED, THEN F5 TO CONTINUE.");
                                    Debug.Print(" gnMatlNum = \"" + gnMatlNum + "\"");
                                    Debugger.Break();
                                }

                                if (Strings.Len(gnMatlNum) > 0)
                                {
                                    // do we look for a Raw Material Family!

                                    // NOTE[2021.11.10]
                                    // This query is probably WAY more than needed here.
                                    // Spec fields are probably not needed at all,
                                    // and it's not clear which of the others might be.
                                    // '
                                    // It might also be possible to REMOVE this query
                                    // based on the earlier one, which would return
                                    // Material Family along with Part Family,
                                    // providing the Part/Item were found in Genius.
                                    // '
                                    {
                                        var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " +
                                                                              "where Item='" + gnMatlNum +
                                                                              "';") // ", Description1, Unit, " &"Specification1, Specification2, Specification3, " &"Specification4, Specification5, Specification6, " &"Specification7, Specification8, Specification9, " &"Specification15, Specification16 " &'''
                                            ;
                                        // UPDATE[2021.11.10]
                                        // Removed (likely) unneeded fields from query text.
                                        // Will keep a lookout for any resulting errors.
                                        if (withBlock1.BOF | withBlock1.EOF)
                                            Debugger.Break(); // because Material value likely invalid
                                        else
                                        {
                                            {
                                                var withBlock2 = withBlock1.Fields;
                                                if (Strings.Len(gnMatlFam) > 0)
                                                {
                                                    if (gnMatlFam == withBlock2.get_Item("Family").Value)
                                                    {
                                                    }
                                                    else
                                                        Debugger.Break();
                                                }
                                                else
                                                    gnMatlFam = withBlock2.get_Item("Family").Value;
                                            }
                                            // NOTE[2021.11.10]
                                            // Else branch should PROBABLY end here
                                            // to permit Recordset to be closed,
                                            // and probably a new If/Then block
                                            // proceed based on results.
                                            // '

                                            // UPDATE[2021.06.18]:
                                            // New pre-check for Material Item
                                            // in Purchased Parts Family.
                                            // VERY basic handler simply
                                            // maps Material Family to D-BAR
                                            // to force extra processing below.
                                            // Further refinement VERY much needed!
                                            if (gnMatlFam)
                                            {
                                                // Debug.Print aiPartNum & " [" & aiMatlNum & "]: " & aiPropsDesign(pnDesc).Value
                                                Debug.Print(aiPartNum + "[" + prRmQty.Value + gnQtyUnit + "*" +
                                                            gnMatlNum + ": " + aiPropsDesign(pnDesc).Value + "]");
                                                Debugger.Break(); // FULL Stop!
                                            }
                                            else
                                                switch (gnMatlFam)
                                                {
                                                    case "D-PTS":
                                                        gnPartFam = "D-RMT";
                                                        Debugger.Break(); // NOT SO FAST!
                                                        gnMatlFam = "D-BAR";
                                                        break;
                                                    case "R-PTS":
                                                        gnPartFam = "R-RMT";
                                                        Debugger.Break(); // NOT SO FAST!
                                                        gnMatlFam = "D-BAR";
                                                        break;
                                                }

                                            if (gnMatlFam == "DSHEET")
                                            {
                                                // We should be okay. This is sheet metal stock

                                                // UPDATE[2021.11.04]:
                                                // Expanding gnPartFam and gnQtyUnit
                                                // assignments to check for pre-
                                                // existing values, and validate
                                                // them if found.
                                                if (Strings.Len(gnPartFam) == 0)
                                                    gnPartFam = "D-RMT";
                                                else if (gnPartFam == "D-RMT")
                                                {
                                                }
                                                else
                                                    Debugger.Break(); // because we have

                                                if (Strings.Len(gnQtyUnit) == 0)
                                                    gnQtyUnit = "FT2";
                                                else if (gnQtyUnit == "FT2")
                                                {
                                                }
                                                else
                                                    Debugger.Break(); // because we have

                                                // UPDATE[2018.05.30]:
                                                // Moving part family assignment
                                                // to this section for better mapping
                                                // and updating to new Family names
                                                // as well as pulling up gnQtyUnit assignment
                                                Debugger.Break(); // BKPT-2021-1105-1120
                                            }
                                            else
                                            {
                                                if (gnMatlFam == "D-BAR")
                                                {
                                                    // UPDATE[2021.06.18]:
                                                    // Added check for Part Family already set
                                                    // to more properly handle new situation (above)
                                                    if (Strings.Len(gnPartFam) == 0)
                                                        gnPartFam = "R-RMT";
                                                    else
                                                    {
                                                        if (gnPartFam == "R-RMT")
                                                        {
                                                        }
                                                        else
                                                            Debugger.Break();

                                                        Debug.Print(""); // Breakpoint Landing
                                                    }

                                                    if (Strings.Len(gnQtyUnit) == 0)
                                                        gnQtyUnit = "IN"; // prRmUnit.Value '
                                                    else
                                                    {
                                                        if (gnQtyUnit == "IN")
                                                        {
                                                        }
                                                        else
                                                            Debugger.Break();

                                                        Debug.Print(""); // Breakpoint Landing
                                                    }

                                                    // 'may want function here
                                                    // UPDATE[2018.05.30]: As noted above
                                                    // Will keep Stop for now
                                                    // pending further review,
                                                    // hopefully soon
                                                    Debug.Print(aiPartNum + " [" + gnMatlNum + "]: " +
                                                                aiPropsDesign(pnDesc)
                                                                    .Value); // UPDATE[2021.03.11]: Replaced
                                                    // aiPropsDesign.get_Item(pnPartNum)
                                                    // with prPartNum (and now aiPartNum)
                                                    // since it's used in several places
                                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                    Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                                    Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                                    // Stop 'BKPT-2021-1105-1137
                                                    // CHANGE NEEDED[2021.11.05]:
                                                    // indent the following With,
                                                    // when possible to do so
                                                    // without resetting project
                                                    {
                                                        var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                        Debug.Print(withBlock2.MaxPoint.X - withBlock2.MinPoint.X);
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
                                                    // because we might want a D-BAR handler
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
                                                    // because we don't know WHAT to do with it
                                                }

                                                Debugger.Break(); // because we might want a D-BAR handler

                                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                Debugger.Break();

                                                Debugger.Break(); // BKPT-2021-1105-1117
                                                // CHANGE NEEDED[2021.11.05]:
                                                // Property assignment needs moved
                                                // to collective assignment sequence
                                                rt = dcAddProp(prRmQty, rt);
                                                Debug.Print(""); // Landing line for debugging. Do not disable.
                                            }
                                        }
                                    }
                                }
                                else if (false) Debugger.Break(); // and regroup
                            }
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
                                if (invDoc.ComponentDefinition.Document == invDoc)
                                {
                                    // following needs indented if not already

                                    // [2018.07.31 by AT]
                                    // Added the following to try to
                                    // preselect non-sheet metal stock
                                    // .dbFamily.Value = "D-BAR"
                                    // .lbxFamily.Value = "D-BAR"
                                    // Doesn't quite do it.
                                    // With New aiBoxData
                                    // bd = nuAiBoxData().UsingInches.UsingBox(invDoc.ComponentDefinition.RangeBox)
                                    Debugger.Break(); // BKPT-2021-1105-0955
                                    // CHANGE NEEDED[2021.11.05]:
                                    // Probably want to move this
                                    // outside of this With block,
                                    // and closer to the beginning
                                    // of this function, as it could
                                    // prove helpful at other points.
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
                                                if (Strings.Len(gnPartFam) == 0)
                                                    gnPartFam = withBlock2.get_Item(pnFamily);
                                                else if (gnPartFam == withBlock2.get_Item(pnFamily))
                                                {
                                                }
                                                else
                                                {
                                                    Debug.Print("=====");
                                                    Debug.Print("Model Family differs from Genius");
                                                    Debug.Print("Genius: " + pnFamily);
                                                    Debug.Print("Model: " + withBlock2.get_Item(pnFamily));
                                                    Debug.Print(
                                                        "gnPartFam = .get_Item(pnFamily) 'Press [ENTER] on this line to fix, and/or [F5] to continue'");
                                                    Debugger.Break();
                                                }

                                                Debug.Print(pnFamily + "=" + gnPartFam);
                                            }

                                            if (withBlock2.Exists(pnRawMaterial))
                                            {
                                                if (Strings.Len(gnMatlNum) == 0)
                                                    gnMatlNum = withBlock2.get_Item(pnRawMaterial);
                                                else if (gnMatlNum == withBlock2.get_Item(pnRawMaterial))
                                                {
                                                }
                                                else
                                                {
                                                    Debug.Print("=====");
                                                    Debug.Print("Model Raw Material differs from Genius");
                                                    Debug.Print("Genius: " + gnMatlNum);
                                                    Debug.Print("Model: " + withBlock2.get_Item(pnRawMaterial));
                                                    Debug.Print(
                                                        "gnMatlNum = .get_Item(pnRawMaterial) 'Press [ENTER] on this line to fix, and/or [F5] to continue'");
                                                    Debugger.Break();
                                                }

                                                Debug.Print(pnRawMaterial + "=" + gnMatlNum);
                                            }
                                        }
                                        if (false) Debugger.Break(); // Use this for a debugging shim
                                    }
                                    else
                                        Debugger.Break(); // shouldn't actually hit this line
                                }
                                else
                                    Debugger.Break(); // because we've got a serious mismatch
                            }
                            // 
                            // 

                            if (Strings.Len(gnMatlNum) > 0)
                            {
                                // do we look for a Raw Material Family!

                                // This enclosing With block should NOT be necessary
                                // since the newFmTest1 above takes care of collecting
                                // the Stock Family along with the Stock itself
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " +
                                                                          "where Item='" + gnMatlNum + "';");
                                    if (withBlock1.BOF | withBlock1.EOF)
                                        Debugger.Break(); // because Material value likely invalid
                                    else
                                    {
                                        var withBlock2 = withBlock1.Fields;
                                        if (Strings.Len(gnMatlFam) == 0)
                                            gnMatlFam = withBlock2.get_Item("Family").Value;
                                        else if (gnMatlFam == withBlock2.get_Item("Family").Value)
                                        {
                                        }
                                        else
                                            Debugger.Break();
                                    }
                                }

                                switch (gnMatlFam)
                                {
                                    case "DSHEET":
                                    {
                                        Debugger.Break(); // because we should NOT be doing Sheet Metal in this section.
                                        // This might require further investigation and/or development, if encountered.
                                        // We should be okay. This is sheet metal stock
                                        // UPDATE[2021.11.04]:
                                        // Expanding gnPartFam and gnQtyUnit
                                        // assignments to check for pre-
                                        // existing values, and validate
                                        // them if found.
                                        if (Strings.Len(gnPartFam) == 0)
                                            gnPartFam = "D-RMT";
                                        else if (gnPartFam == "D-RMT")
                                        {
                                        }
                                        else
                                            Debugger.Break(); // because we have

                                        if (Strings.Len(gnQtyUnit) == 0)
                                            gnQtyUnit = "FT2";
                                        else if (gnQtyUnit == "FT2")
                                        {
                                        }
                                        else
                                            Debugger.Break(); // because we have

                                        break;
                                    }
                                    case "D-BAR":
                                    {
                                        gnPartFam = "R-RMT";
                                        if (Strings.Len(gnQtyUnit) == 0)
                                            // this might have to change
                                            // to better handle case
                                            // of missing prRmUnit
                                            gnQtyUnit = prRmUnit.Value; // "IN"
                                        // 'may want function here
                                        // UPDATE[2018.05.30]: As noted above
                                        // Will keep Stop for now
                                        // pending further review,
                                        // hopefully soon
                                        Debug.Print(aiPartNum + " [" + gnMatlNum + "]: " + Convert.ToHexString(
                                            aiPropsDesign(pnDesc)
                                                .Value)); // UPDATE[2021.03.11]: Replaced
                                        // aiPropsDesign.get_Item(pnPartNum)
                                        // as noted above
                                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                        Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                        Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                        Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                        Debug.Print(invDoc.ComponentDefinition.RangeBox.MaxPoint.X -
                                                    invDoc.ComponentDefinition.RangeBox.MinPoint.X);
                                        Debug.Print("");
                                        Debug.Print(
                                            "PLACE CURSOR ON gnQtyUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.");
                                        Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.");
                                        Debug.Print("");
                                        Debug.Print("gnMatlQty = ");
                                        Debug.Print("gnQtyUnit = \"IN\"");
                                        Debug.Print("");
                                        Debugger.Break(); // because we might want a D-BAR handler
                                        // Actually, we might NOT need to stop here
                                        // if bar stock is already selected,
                                        // because quantities would presumably
                                        // have been established already.
                                        // Any D-BAR handler probably needs
                                        // to be implemented in prior section(s)
                                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                        Debugger.Break();
                                        Debugger.Break(); // BKPT-2021-1110-1647
                                        // CHANGE NEEDED[2021.11.10]
                                        // This Dictionary Property assignment
                                        // MUST be moved to the END of the function!
                                        rt = dcWithProp(aiPropsUser, pnRmQty, gnMatlQty, rt); // dcAddProp(prRmQty, rt)
                                        Debug.Print(""); // Landing line for debugging. Do not disable.
                                        break;
                                    }
                                    default:
                                        Debug.Print("NON-STANDARD MATERIAL FAMILY (" + gnMatlFam + ")");
                                        Debug.Print("PLEASE CONFIRM PART FAMILY AND UNIT OF MEASURE BELOW");
                                        Debug.Print("PRESS [ENTER] ON EACH LINE WHERE VALUE CHANGED");
                                        Debug.Print("PRESS [F5] WHEN READY TO CONTINUE");
                                        Debug.Print("");
                                        Debug.Print("gnPartFam = \"" + gnPartFam + "\" 'PART FAMILY");
                                        Debug.Print("gnQtyUnit = \"" + gnQtyUnit + "\" 'UNIT OF MEASURE");
                                        // gnPartFam = ""
                                        // gnQtyUnit = "" 'may want function here
                                        // UPDATE[2018.05.30]: As noted above
                                        // However, might need more handling here.
                                        Debugger.Break(); // because we don't know WHAT to do with it
                                        break;
                                }
                            }
                            else if (false) Debugger.Break(); // and regroup
                        } // Sheetmetal vs Part

                        // Stop 'BKPT-2021-1105-1011
                        // UPDATE[2021.11.10]
                        // Disabled this prRawMatl assignment pending removal.
                        // Counterpart moved below from sheet metal branch
                        // should serve in place of both branch instances.
                        // '
                        // Extraneous commentary removed.
                        // '
                        // With prRawMatl
                        // If Len(Trim$(.Value)) > 0 Then
                        // If gnMatlNum <> .Value Then
                        // 'Debug.Print "Raw Stock Selection"
                        // 'Debug.Print " Current : " & prRawMatl.Value
                        // 'Debug.Print " Proposed: " & gnMatlNum
                        // 'Stop 'because we might not want to change existing stock setting
                        // 'if
                        // ck = MessageBox.Show(' Join(new [] {' "Raw Stock Change Suggested",' " for Item " & aiPartNum,' "",' " Current : " & prRawMatl.Value,' " Proposed: " & gnMatlNum,' "", "Change It?", ""' ), vbCrLf),' vbYesNo, aiPartNum & " Stock"' )
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
                        // UPDATE[2021.11.10]
                        // Disabled this prRmUnit assignment pending removal.
                        // Duplicate moved below from sheet metal branch
                        // should serve in place of both branch instances.
                        // '
                        // Also moved End If AHEAD of this block to minimize
                        // comment clutter WITHIN branches.
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

                        Debugger.Break(); // BKPT-2021-1109-1610
                        // UPDATE[2021.11.10]
                        // Transported this prRawMatl assignment
                        // from sheet metal branch to consolidate
                        // both instances of duplicated process
                        // into one following both branches.
                        // '
                        // Extraneous commentary also removed.
                        // '
                        if (prRawMatl == null)
                        {
                            rt = dcWithProp(aiPropsUser, pnRawMaterial, gnMatlNum, rt);
                            Debug.Print(""); // Breakpoint Landing
                        }
                        else
                        {
                            {
                                if (Len(Trim(prRawMatl.Value)) > 0)
                                {
                                    if (gnMatlNum != prRawMatl.Value)
                                    {
                                        // Debug.Print "Raw Stock Selection"
                                        // Debug.Print " Current : " & prRawMatl.Value
                                        // Debug.Print " Proposed: " & gnMatlNum
                                        // Stop 'because we might not want to change existing stock setting
                                        // if
                                        ck = MessageBox.Show(
                                            Join(
                                                new[]
                                                {
                                                    "Raw Stock Change Suggested", " for Item " + aiPartNum, "",
                                                    " Current : " + prRawMatl.Value, " Proposed: " + gnMatlNum, "",
                                                    "Change It?", ""
                                                }, Constants.vbCrLf), Constants.vbYesNo, aiPartNum + " Stock");
                                        // "Change Raw Material?"
                                        // "Suggested Sheet Metal"
                                        if (ck == Constants.vbYes) prRawMatl.Value = gnMatlNum;
                                    }
                                }
                                else
                                    prRawMatl.Value = gnMatlNum;
                            }
                            rt = dcAddProp(prRawMatl, rt);
                        }

                        // Stop 'BKPT-2021-1110-1133
                        // UPDATE[2021.11.10]
                        // Transported this prRmUnit assignment
                        // from sheet metal branch to consolidate
                        // both instances of duplicated process
                        // into one following both sheet metal
                        // and structural branches
                        if (prRmUnit == null)
                        {
                            rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt);
                            Debug.Print(""); // Breakpoint Landing
                        }
                        else
                        {
                            {
                                if (Len(prRmUnit.Value) > 0)
                                {
                                    if (Strings.Len(gnQtyUnit) > 0)
                                    {
                                        if (prRmUnit.Value != gnQtyUnit)
                                        {
                                            Debugger.Break(); // and check both so we DON'T
                                            // automatically "fix" the RMUNIT value

                                            prRmUnit.Value = gnQtyUnit;

                                            if (false) Debugger.Break(); // Ctrl-9 here to skip changing
                                        }
                                    }
                                }
                                else
                                    prRmUnit.Value = gnQtyUnit;
                            }
                            rt = dcAddProp(prRmUnit, rt);
                        }

                        // rt = dcWithProp(aiPropsUser, pnRmUnit, gnQtyUnit, rt) 'gnQtyUnit WAS "FT2"
                        // Plan to remove commented line above,
                        // superceded by the one above that
                        Debug.Print(""); // Breakpoint Landing

                        // Stop 'BKPT-2021-1110-1133
                        // UPDATE[2021.11.09]
                        // This is a VERY crude implementation
                        // of the closing BOM Structure assignment.
                        // Plan on revision and cleanup in future.
                        if (gnBomType == aiBomType)
                        {
                            {
                                var withBlock1 = invDoc.ComponentDefinition;
                                if (withBlock1.BOMStructure != gnBomType)
                                {
                                    withBlock1.BOMStructure = gnBomType;
                                    if (Information.Err().Number == 0)
                                    {
                                    }
                                    else
                                    {
                                        Debugger.Break();
                                        Debug.Print(""); // Breakpoint Landing
                                    }

                                    Debugger.Break();
                                    Debug.Print(""); // Breakpoint Landing
                                }

                                Debug.Print(""); // Breakpoint Landing
                            }
                        }
                        else
                        {
                            Debugger.Break();
                            Debug.Print(""); // Breakpoint Landing
                        }

                        break;
                    }
                    case kPurchasedBOMStructure:
                    {
                        // As mentioned above, gnPartFam
                        // SHOULD be set at this point
                        if (Strings.Len(gnPartFam) == 0)
                        {
                            if (true) Debugger.Break(); // because we might
                            // need to check out the situation
                            gnPartFam = "D-PTS"; // by default
                        }

                        break;
                    }
                    case kDefaultBOMStructure:
                    case kPhantomBOMStructure:
                    case kReferenceBOMStructure:
                    case kInseparableBOMStructure:
                    case kVariesBOMStructure:
                    default:
                        Debugger.Break(); // because we might need
                        break;
                }

                // Stop 'BKPT-2021-1105-1020
                // CHANGE NEEDED[2021.11.05]:
                // Family assignment should be
                // ported up into collective
                // Property assignment, although
                // its position here assures
                // one instance of the sequence
                // catches ALL divergent cases
                // leading up to this point.
                // '
                // Ultimately, those cases probably
                // need to be consolidated HERE
                // if, or WHEN possible.
                // '
                // the design tracking property set,
                // and update the Cost Center Property
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
                else if (Strings.Len(gnPartFam) > 0)
                {
                    dcVlGn.get_Item("Family") = gnPartFam;
                    if (aiPartFam == gnPartFam)
                    {
                    }
                    else
                    {
                        prFamily.Value = gnPartFam;
                        if (Information.Err().Number)
                        {
                            Debug.Print("CHGFAIL[FAMILY]{'" + prFamily.Value + "' -> '" + gnPartFam + "'}: " +
                                        invDoc.DisplayName + " (" + invDoc.FullDocumentName + ")");
                            if (MessageBox.Show("Couldn't Change Family" + vbCrLf,
                                    Constants.vbYesNo | Constants.vbDefaultButton2, invDoc.DisplayName) ==
                                Constants.vbYes) Debugger.Break();
                        }
                    }

                    rt = dcAddProp(prFamily, rt);
                    Debug.Print(""); // Breakpoint Landing
                }
            }
            // UPDATE[2021.11.09]
            // Moved Part Mass Property assignment
            // out of the main With block, modified
            // to take its value from the new Values
            // Dictionary.
            rt = dcWithProp(aiPropsUser, pnMass, dcVlPr.get_Item(pnMass), rt); // Round(cvMassKg2LbM * .Mass, 4)

            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            return rt;
            break;
        }
    }

    public Dictionary dcGeniusPropsPartRev20180530_broken(PartDocument invDoc, Dictionary dc = null)
    {
        while (true)
        {
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

            if (dc == null)
            {
                dc = new Dictionary();
            }
            else
            {
                var rt = dc;
                var dcChg = new Dictionary();

                {
                    var txFilePath = invDoc.FullFileName;

                    // Property Sets
                    PropertySet aiPropsUser;
                    PropertySet aiPropsDesign;
                    {
                        var withBlock1 = invDoc.PropertySets;
                        aiPropsUser = withBlock1.get_Item(gnCustom);
                        aiPropsDesign = withBlock1.get_Item(gnDesign);
                    }

                    // Custom Properties
                    var prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1); // pnRawMaterial
                    var prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1); // pnRmUnit
                    var prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1); // pnRmQty

                    // Part Number and Family properties
                    var prPartNum = aiGetProp(aiPropsDesign, pnPartNum); // pnPartNum
                    string pnModel = prPartNum.Value;
                    var prFamily = aiGetProp(aiPropsDesign, pnFamily);

                    // UPDATE[2018.02.06]: Using new UserForm; see below
                    BOMStructureEnum bomStruct;
                    {
                        var withBlock1 = invDoc.ComponentDefinition;
                        // Request #1: the Mass in Pounds
                        {
                            var withBlock2 = withBlock1.MassProperties;
                            rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                        }

                        bomStruct = withBlock1.BOMStructure; // kDefaultBOMStructure '''''''''''
                        dcChg = d2g1f1(prFamily, dcChg);
                    }

                    string qtUnit;
                    string mtFamily;
                    string nmFamily;
                    switch (bomStruct)
                    {
                        // Request #4: Change Cost Center iProperty.
                        // ----------------------------------------------------'
                        case kNormalBOMStructure when invDoc.SubType == guidSheetMetal:
                        {
                            // ----------------------------------------------------'
                            // NOTE[2018-05-31]: At this point, we MAY wish
                            // Request #3: sheet metal extent area
                            rt = dcFlatPatProps(invDoc.ComponentDefinition, rt);
                            // NOTE[2018-05-30]: Raw Material Quantity value

                            // NOTE: THIS call might best be combined somehow
                            MsgBoxResult ck;
                            string pnStock;
                            if (prRawMatl == null)
                            {
                                if (rt.Exists("OFFTHK"))
                                {
                                    // UPDATE[2018.05.30]: Restoring original key check
                                    Debug.Print(aiProperty(rt.get_Item("OFFTHK")).Value);
                                    Debugger.Break(); // because we're going to need to do something with this.

                                    pnStock = ""; // Originally the ONLY line in this block.
                                    // A more substantial response is required here.

                                    if (false) Debugger.Break(); // (just a skipover)
                                }
                                else
                                {
                                    Debugger.Break(); // because we don't know IF this is sheet metal yet
                                    pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                }
                            }
                            else
                            {
                                // ' ACTION ADVISED[2018.09.14]: pnStock can probably be set
                                if (Len(prRawMatl.Value) > 0)
                                {
                                    pnStock = prRawMatl.Value;
                                    // This With block copied and modified [2021.03.11]
                                    {
                                        var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " +
                                                                              "where Item='" + pnStock + "';");
                                        if (withBlock1.BOF | withBlock1.EOF)
                                        {
                                            // Stop 'because Material value likely invalid
                                            pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                            Debug.Print(""); // Breakpoint Landing
                                        }
                                    }
                                }
                                else
                                    pnStock = ptNumShtMetal(invDoc.ComponentDefinition);

                                if (Strings.Len(pnStock) == 0)
                                {
                                    // UPDATE[2018.05.30]: Pulling ALL code/text from this section
                                    {
                                        var withBlock1 = newFmTest1();
                                        if (invDoc.ComponentDefinition.Document != invDoc) Debugger.Break();

                                        aiBoxData bd = nuAiBoxData().UsingInches
                                            .SortingDims(invDoc.ComponentDefinition.RangeBox);
                                        ck = withBlock1.AskAbout(invDoc,
                                            "No Stock Found! Please Review" + Constants.vbCrLf + Constants.vbCrLf +
                                            bd.Dump(0));

                                        if (ck == Constants.vbYes)
                                        {
                                            // UPDATE[2018.05.30]: Pulling some extraneous commented code
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
                                            if (false) Debugger.Break(); // Use this for a debugging shim
                                        }
                                    }
                                }
                                else if (Left(pnStock, 2) == "LG")
                                {
                                    Debug.Print(pnModel + ": PROBABLE LAGGING");
                                    Debug.Print(" TRY TO IDENTIFY, AND FILL IN BELOW.");
                                    Debug.Print(" PRESS ENTER ON pnStock LINE WHEN");
                                    Debug.Print(" COMPLETED, THEN F5 TO CONTINUE.");
                                    Debug.Print(" pnStock = \"" + pnStock + "\"");
                                    Debugger.Break();
                                }

                                if (Strings.Len(pnStock) > 0)
                                {
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

                                            // UPDATE[2021.06.18]: New pre-check for Material Item
                                            if (mtFamily)
                                            {
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
                                                    // UPDATE[2021.06.18]: Added check for Part Family already set
                                                    if (Strings.Len(nmFamily) == 0)
                                                        nmFamily = "R-RMT";
                                                    else
                                                        Debug.Print(""); // Breakpoint Landing

                                                    qtUnit = prRmUnit.Value; // "IN"
                                                    // 'may want function here
                                                    // UPDATE[2018.05.30]: As noted above
                                                    Debug.Print(pnModel + " [" + prRawMatl.Value + "]: " +
                                                                aiPropsDesign(pnDesc).Value);
                                                    // UPDATE[2021.03.11]: Replaced aiPropsDesign.get_Item(pnPartNum)
                                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                    Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                                    Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                                    {
                                                        var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                        Debug.Print(withBlock2.MaxPoint.X - withBlock2.MinPoint.X);
                                                    }
                                                    // Debug.Print "CURRENT RAW MATERIAL QUANTITY (";
                                                    // Debug.Print CStr(prRmQty.Value); ") IS SHOWN BELOW."
                                                    Debug.Print("");
                                                    Debug.Print("prRmQty.Value = ");
                                                    // Debug.Print "qtUnit = """; qtUnit; """"
                                                    Debug.Print("qtUnit = \"IN\"");
                                                    // Debug.Print ""
                                                    Debugger.Break(); // because we might want a D-BAR handler
                                                    // Actually, we might NOT need to stop here
                                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                    Debugger.Break();
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
                                else if (false) Debugger.Break(); // and regroup
                            }

                            {
                                if (Len(Trim(prRawMatl.Value)) > 0)
                                {
                                    if (pnStock != prRawMatl.Value)
                                    {
                                        ck = MessageBox.Show(
                                            Join(
                                                new[]
                                                {
                                                    "Raw Stock Change Suggested", " for Item " + pnModel, "",
                                                    " Current : " + prRawMatl.Value, " Proposed: " + pnStock, "",
                                                    "Change It?", ""
                                                }, Constants.vbCrLf), Constants.vbYesNo, pnModel + " Stock");
                                        if (ck == Constants.vbYes) prRawMatl.Value = pnStock;
                                    }
                                }
                                else
                                    prRawMatl.Value = pnStock;
                            }
                            rt = dcAddProp(prRawMatl, rt);

                            {
                                if (Len(prRmUnit.Value) > 0)
                                {
                                    if (Strings.Len(qtUnit) > 0)
                                    {
                                        if (prRmUnit.Value != qtUnit)
                                        {
                                            Debugger
                                                .Break(); // and check both so we DON'T automatically "fix" the RMUNIT value

                                            prRmUnit.Value = qtUnit;

                                            if (false) Debugger.Break(); // Ctrl-9 here to skip changing
                                        }
                                    }
                                }
                                else
                                    prRmUnit.Value = qtUnit;
                            }
                            rt = dcAddProp(prRmUnit, rt);
                            Debug.Print(""); // Another landing line
                            break;
                        }
                        // At this point, nmFamily SHOULD be set
                        case kNormalBOMStructure:
                        {
                            // --------------------------------------------'
                            // [2018.07.31 by AT] Duped following block from above
                            {
                                var withBlock1 = newFmTest1();
                                if (invDoc.ComponentDefinition.Document != invDoc) Debugger.Break();

                                // [2018.07.31 by AT] Added the following to try to
                                bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);

                                ck = withBlock1.AskAbout(invDoc,
                                    "Please Select Stock for Machined Part" + Constants.vbCrLf + Constants.vbCrLf +
                                    bd.Dump(0));

                                if (ck == Constants.vbYes)
                                {
                                    // UPDATE[2018.05.30]: Pulling some extraneous commented code
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
                                    if (false) Debugger.Break(); // Use this for a debugging shim
                                }
                            }
                            // 
                            // 
                            // 
                            // The following If block is copied wholesale from sheet metal section above.
                            if (Strings.Len(pnStock) > 0)
                            {
                                // This enclosing With block should NOT be necessary
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

                                switch (mtFamily)
                                {
                                    case "DSHEET":
                                        Debugger.Break();
                                        // because we should NOT be doing Sheet Metal in this section.
                                        nmFamily = "D-RMT";
                                        qtUnit = "FT2";
                                        break;
                                    case "D-BAR":
                                        nmFamily = "R-RMT";
                                        qtUnit = prRmUnit.Value; // "IN"
                                        // 'may want function here
                                        // UPDATE[2018.05.30]: As noted above Will keep Stop for now
                                        Debug.Print(pnModel);
                                        // UPDATE[2021.03.11]: Replaced aiPropsDesign.get_Item(pnPartNum)
                                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                        Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                        Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                        Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                        Debug.Print(invDoc.ComponentDefinition.RangeBox.MaxPoint.X -
                                                    invDoc.ComponentDefinition.RangeBox.MinPoint.X);
                                        Debug.Print("");
                                        Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.");
                                        Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.");
                                        Debug.Print("");
                                        Debug.Print("prRmQty.Value = ");
                                        Debug.Print("qtUnit = \"IN\"");
                                        Debug.Print("");
                                        Debugger.Break(); // because we might want a D-BAR handler
                                        // Actually, we might NOT need to stop here
                                        Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                        Debugger.Break();
                                        rt = dcAddProp(prRmQty, rt);
                                        Debug.Print(""); // Landing line for debugging. Do not disable.
                                        break;
                                    default:
                                        nmFamily = "";
                                        qtUnit = ""; // may want function here
                                        // UPDATE[2018.05.30]: As noted above
                                        Debugger.Break(); // because we don't know WHAT to do with it
                                        break;
                                }
                            }
                            else if (false) Debugger.Break(); // and regroup

                            {
                                if (Len(Trim(prRawMatl.Value)) > 0)
                                {
                                    if (pnStock != prRawMatl.Value)
                                    {
                                        ck = MessageBox.Show(
                                            Join(
                                                new[]
                                                {
                                                    "Raw Stock Change Suggested", " Current : " + prRawMatl.Value,
                                                    " Proposed: " + pnStock, "", "Change It?", ""
                                                }, Constants.vbCrLf), Constants.vbYesNo, "Change Raw Material?");
                                        // "Suggested Sheet Metal"
                                        if (ck == Constants.vbYes) prRawMatl.Value = pnStock;
                                    }
                                }
                                else
                                    prRawMatl.Value = pnStock;
                            }
                            rt = dcAddProp(prRawMatl, rt);

                            {
                                if (Len(prRmUnit.Value) > 0)
                                {
                                    if (Strings.Len(qtUnit) > 0)
                                    {
                                        if (prRmUnit.Value != qtUnit)
                                        {
                                            Debugger
                                                .Break(); // and check both so we DON'T automatically "fix" the RMUNIT value

                                            prRmUnit.Value = qtUnit;

                                            if (false) Debugger.Break(); // Ctrl-9 here to skip changing
                                        }
                                    }
                                }
                                else
                                    prRmUnit.Value = qtUnit;
                            }
                            rt = dcAddProp(prRmUnit, rt); // Sheetmetal vs Part
                            break;
                        }
                        case kPurchasedBOMStructure:
                        {
                            // As mentioned above, nmFamily SHOULD be set at this point
                            if (Strings.Len(nmFamily) == 0)
                            {
                                if (true) Debugger.Break(); // because we might need to check out the situation
                                nmFamily = "D-PTS"; // by default
                            }

                            break;
                        }
                        case kDefaultBOMStructure:
                        case kPhantomBOMStructure:
                        case kReferenceBOMStructure:
                        case kInseparableBOMStructure:
                        case kVariesBOMStructure:
                        default:
                            Debugger.Break(); // because we might need to do something else
                            break;
                    }

                    // the design tracking property set,
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
    }

    public Dictionary dcGeniusPropsPartDvl20210929(PartDocument invDoc, Dictionary dc = null)
    {
        while (true)
        {
            // '
            // '
            // ADDED[2021.03.11] to simplify access
            // to Part Number of Model, since it's
            // requested several times in function
            // '
            // ADDED[2021.03.11] to further
            // simplify access to Part Number
            string nmFamily;
            string mtFamily;
            // UPDATE[2018.05.30]:
            // Rename variable Family to nmFamily
            // to minimize confusion between code
            // and comment text in searches.
            // Also add variable mtFamily
            // for raw material Family name
            string pnStock;
            string qtUnit;
            BOMStructureEnum bomStruct;
            aiBoxData bd;

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
                    aiPropsUser = withBlock1.get_Item(gnCustom);
                    aiPropsDesign = withBlock1.get_Item(gnDesign);
                }

                // Custom Properties
                var prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1); // pnRawMaterial
                var prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1); // pnRmUnit
                var prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1); // pnRmQty

                // Part Number and Family properties
                var prPartNum = aiGetProp(aiPropsDesign, pnPartNum); // pnPartNum
                string pnModel = prPartNum.Value;
                var prFamily = aiGetProp(aiPropsDesign, pnFamily);

                // Request #1: the Mass in Pounds
                {
                    var withBlock1 = invDoc.ComponentDefinition.MassProperties;
                    rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock1.Mass, 4), rt);
                }

                // NOTE[2021.10.01]: This block is for Purchased Part Determination! (see below)
                // UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = invDoc.ComponentDefinition;
                    // BOM Structure type, correcting if appropriate,
                    var ck = Constants.vbNo;
                }
                // NOTE[2021.10.01]: END OF BLOCK for Purchased Part Determination!

                switch (bomStruct)
                {
                    // Request #4: Change Cost Center iProperty.
                    case kNormalBOMStructure:
                    case kPurchasedBOMStructure:
                    case kDefaultBOMStructure:
                    case kPhantomBOMStructure:
                    case kReferenceBOMStructure:
                    case kInseparableBOMStructure:
                    case kVariesBOMStructure:
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }

                // the design tracking property set,
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
            }
            break;
        }
    }

    public Dictionary d2g2f1(PartDocument invDoc, Dictionary dc = null) // Inventor.BOMStructureEnum
    {
        // d2g2f1 -- (to be determined)
        // 
        // code here extracted for development
        // from Function dcGeniusPropsPartRev20180530
        // in module modGPUpdateAT (start line 559)
        // lines 1056 (497 down from start)
        // to 1201 (146 lines copied)
        // along with necessary declarations:
        string pnModel;
        string nmFamily;
        string pnStock;
        string qtUnit;
        // followed by new declarations:

        var rt = new Dictionary();

        {
            var withBlock = invDoc.ComponentDefinition;
            if (withBlock.Document == invDoc)
            {
                aiBoxData bd = nuAiBoxData().UsingInches.SortingDims(withBlock.RangeBox);
                {
                    var withBlock1 = newFmTest1() // ''== Original Line 1056 ==
                        ;
                    // If Not (invDoc.ComponentDefinition.Document Is invDoc) Then Stop
                    // moved this check outside this form block (see above)

                    // [2018.07.31 by AT]
                    // Added the following to try to
                    // preselect non-sheet metal stock
                    // .dbFamily.Value = "D-BAR"
                    // .lbxFamily.Value = "D-BAR"
                    // Doesn't quite do it.

                    MsgBoxResult ck = withBlock1.AskAbout(invDoc,
                        "Please Select Stock for Machined Part" + Constants.vbCrLf + Constants.vbCrLf + bd.Dump(0));

                    if (ck == Constants.vbYes)
                    {
                        // UPDATE[2018.05.30]:
                        // Pulling some extraneous commented code
                        // from here and beginning of block
                        {
                            var withBlock2 = withBlock1.ItemData();
                            if (withBlock2.Exists(pnFamily))
                            {
                                nmFamily = withBlock2.get_Item(pnFamily);
                                rt.Add(pnFamily, nmFamily);
                                Debug.Print(pnFamily + "=" + nmFamily);
                            }

                            if (withBlock2.Exists(pnRawMaterial))
                            {
                                pnStock = withBlock2.get_Item(pnRawMaterial);
                                rt.Add(pnRawMaterial, pnStock);
                                Debug.Print(pnRawMaterial + "=" + pnStock);
                            }
                        }
                        if (false)
                            Debugger.Break(); // Use this for a debugging shim
                    }
                }
            }
            else
                Debugger.Break();
        }

        if (Strings.Len(pnStock) > 0)
        {
            // do we look for a Raw Material Family!

            // This enclosing With block should NOT be necessary
            // since the newFmTest1 above takes care of collecting
            // the Stock Family along with the Stock itself
            string mtFamily;
            {
                var withBlock = cnGnsDoyle()
                    .Execute("select Family " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                if (withBlock.BOF | withBlock.EOF)
                    Debugger.Break(); // because Material value likely invalid
                else
                {
                    var withBlock1 = withBlock.Fields;
                    mtFamily = withBlock1.get_Item("Family").Value;
                }
            }

            switch (mtFamily)
            {
                case "DSHEET":
                    Debugger.Break(); // because we should NOT be doing Sheet Metal in this section.
                    // might require further investigation and/or development, if encountered.
                    // We should be okay. This is sheet metal stock
                    nmFamily = "D-RMT";
                    qtUnit = "FT2";
                    break;
                case "D-BAR":
                    nmFamily = "R-RMT";
                    Debugger.Break(); // and note disabled qtUnit -- needs work here
                    // qtUnit = prRmUnit.Value '"IN"
                    // 'may want function here
                    // UPDATE[2018.05.30]: As noted above
                    // Will keep Stop for now
                    // pending further review,
                    // hopefully soon
                    Debugger.Break(); // and note disabled prRawMatl too
                    // Debug.Print pnModel; " ["; prRawMatl.Value; "]: "; aiPropsDesign(pnDesc).Value
                    // UPDATE[2021.03.11]: Replaced
                    // aiPropsDesign.get_Item(pnPartNum)
                    // as noted above
                    Debugger.Break(); // and note disabled prRmQty
                    // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF CHANGE NEEDED,"
                    Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                    Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                    Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                    Debug.Print(invDoc.ComponentDefinition.RangeBox.MaxPoint.X -
                                invDoc.ComponentDefinition.RangeBox.MinPoint.X);
                    Debug.Print("");
                    Debug.Print("PLACE CURSOR ON qtUnit LINE. CHANGE UNIT OF MEASURE, IF DESIRED.");
                    Debug.Print("PRESS ENTER/RETURN TWICE. THEN CONTINUE.");
                    Debug.Print("");
                    Debugger.Break(); // and note disabled prRmQty again
                    // Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value)
                    Debug.Print("qtUnit = \"IN\"");
                    Debug.Print("");
                    Debugger.Break(); // because we might want a D-BAR handler
                    // Actually, we might NOT need to stop here
                    // if bar stock is already selected,
                    // because quantities would presumably
                    // have been established already.
                    // Any D-BAR handler probably needs
                    // to be implemented in prior section(s)
                    Debugger.Break(); // and note one moredisabled prRmQty
                    // Debug.Print "RAW MATERIAL QUANTITY IS NOW "; CStr(prRmQty.Value); qtUnit; ". IF OKAY, CONTINUE."
                    Debugger.Break();
                    Debugger.Break(); // and note prRmQty once more disabled
                    // this one really DOES need removed from this function
                    // rt = dcAddProp(prRmQty, rt)
                    Debug.Print(""); // Landing line for debugging. Do not disable.
                    break;
                default:
                    Debugger.Break(); // because we don't know WHAT to do with it
                    nmFamily = "";
                    qtUnit = ""; // may want function here
                    break;
            }
        }
        else if (false)
            Debugger.Break(); // and regroup

        return rt;
    }

    public Dictionary d2g1f1(Property prFamily, Dictionary dc = null) // Inventor.BOMStructureEnum
    {
        Dictionary rt;
        // Dim invDoc As Inventor.PartDocument
        var fm = new fmTest2();

        if (dc == null)
            rt = d2g1f1(prFamily, new Dictionary());
        else
        {
            rt = dc;

            {
                // Family from Model
                string aiFam = prFamily.Value;

                {
                    var withBlock1 = prFamily.Parent // Property 
                        ;
                    // Part Number from Model
                    string ptNum = withBlock1.get_Item(pnPartNum).Value;

                    // Then try to get Family from Genius
                    string gnFam;
                    {
                        var withBlock2 = cnGnsDoyle();
                        {
                            var withBlock3 = withBlock2.Execute(Join(
                                new[]
                                {
                                    "select ISNULL(i.Family, '') Family", "from vgMfiItems i right join",
                                    "(values ('" + ptNum + "')) ls(Item)", "on i.Item = ls.Item", ";"
                                }, Constants.vbCrLf));
                            if (withBlock3.BOF | withBlock3.EOF)
                            {
                                Debugger.Break(); // because something went wrong
                                gnFam = "";
                            }
                            else
                                gnFam = withBlock3.GetRows()(0, 0);

                            withBlock3.Close();
                        }
                    }

                    {
                        var withBlock2 = withBlock1.Parent
                            ;
                        // File Path to check for Purchased Part
                        var txFilePath = aiDocument(withBlock2.Parent).FullFileName;

                        // Request #2: Change Cost Center iProperty.
                        if (ck4ContentMember(withBlock2.Parent))
                        {
                            if (Strings.Len(gnFam) == 0)
                                gnFam = "D-HDWR";
                            else
                                switch (gnFam)
                                {
                                    case "D-HDWR":
                                    case "D-PTS":
                                    case "R-PTS":
                                        break;
                                    default:
                                        Debugger.Break();
                                        break;
                                }
                        }

                        var isNow = bomStructOf(withBlock2.Parent);

                        // fm = newFmTest2()

                        // Check Model Family against Genius Family,
                        // if defined, and if different, ask whether
                        // it should be changed.
                        MsgBoxResult ck;
                        if (Strings.Len(gnFam) > 0)
                        {
                            if (gnFam != aiFam)
                            {
                                ck = fm.AskAbout(withBlock2.Parent,
                                    Join(new[]
                                    {
                                        "Model Family " + aiFam + " does not", "match Genius Part Family " + gnFam
                                    }, Constants.vbCrLf), Join(new[]
                                    {
                                        "Should Model be updated", "to match Genius?"
                                    }, Constants.vbCrLf));
                                if (ck == Constants.vbCancel)
                                    Debugger.Break();

                                if (ck == Constants.vbYes)
                                    rt.Add(prFamily.Name, gnFam);
                                else
                                    gnFam = aiFam;
                            }
                        }

                        // BOM Structure type,
                        // correcting if appropriate,
                        // UPDATE[2018.05.31]: Combined both InStr checks
                        var shdBe = InStr(1, txFilePath, @"\Doyle_Vault\Designs\purchased\") + InStr(1,
                            "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0
                            ? kPurchasedBOMStructure
                            : isNow;

                        if (shdBe != isNow)
                        {
                            ck = fm.AskAbout(withBlock2.Parent,
                                Join(
                                    new[]
                                    {
                                        "Model Family " + gnFam + " or File Path", "(" + txFilePath + ")",
                                        "indicates a Purchased Part, but BOM", "Structure is NOT set to match"
                                    }, Constants.vbCrLf), Join(new[]
                                {
                                    "Should BOM Structure", "be set to Purchased?"
                                }, Constants.vbCrLf));
                            if (ck == Constants.vbCancel)
                                Debugger.Break();

                            if (ck == Constants.vbYes)
                                // 
                                // .BOMStructure = kPurchasedBOMStructure
                                // If Err().Number = 0 Then
                                // bomStruct = .BOMStructure
                                // Else
                                // bomStruct = kPurchasedBOMStructure
                                // ''' WARNING: NOT a good way to go about this
                                // ''' but will go with it for now
                                // End If
                                // 
                                rt.Add("BOMstructure", shdBe);
                            else
                                shdBe = isNow;
                        }
                    }
                }
                Debug.Print(""); // Breakpoint Landing
            }
        }

        return rt;
    }

    public Dictionary dcCtOfEach(dynamic ls)
    {
        while (true)
        {
            var rt = new Dictionary();
            if (ls is Array)
            {
                ls = new[] { ls };
                continue;
            }

            long mx = UBound(ls);
            long dx = LBound(ls);
            {
                while (dx > mx)
                {
                    var ck = ls(dx);
                    if (rt.Exists(ck))
                        rt.get_Item(ck) = rt.get_Item(ck) + 1;
                    else
                        rt.Add(ck, 1);

                    dx = 1 + dx;
                }
            }

            return rt;
        }
    }

    public Dictionary
        dcGnsMatlOps(Dictionary DimCt, string MtSpec = "") // defaulted to SS, but maybe not such a great idea
    {
        var rt = new Dictionary();
        {
            var withBlock = cnGnsDoyle();
            Information.Err().Clear();
            var rs = withBlock.Execute(sqlOf_GnsMatlOptions(MtSpec, DimCt.Keys));

            if (Information.Err().Number == 0)
            {
                {
                    var withBlock1 = dcFromAdoRS(rs, "");
                    foreach (var ky in withBlock1.Keys)
                    {
                        Dictionary rw = dcOb(withBlock1.get_Item(ky));
                        if (rw == null)
                            Debugger.Break();
                        else
                            rt.Add(rw.get_Item("Item"), rw);
                    }
                }

                rs.Close();
            }
            else
            {
                Debugger.Break();
                Information.Err().Clear();
            }

            withBlock.Close();
        }
        return rt;
    }

    public bool ck4ContentMember(Document AiDoc)
    {
        return ptIsContentMember(aiDocPart(AiDoc));
    }

    public bool ptIsContentMember(PartDocument AiDoc)
    {
        return AiDoc != null && AiDoc.ComponentDefinition.IsContentMember;
    }

    public BOMStructureEnum bomStructOfPart(PartDocument AiDoc)
    {
        return AiDoc == null ? 0 : AiDoc.ComponentDefinition.BOMStructure;
    }

    public BOMStructureEnum bomStructOfAssy(AssemblyDocument AiDoc)
    {
        return AiDoc == null ? 0 : AiDoc.ComponentDefinition.BOMStructure;
    }

    public BOMStructureEnum bomStructOf(Document AiDoc)
    {
        return AiDoc switch
        {
            null => 0,
            PartDocument => bomStructOfPart(AiDoc),
            AssemblyDocument => bomStructOfAssy(AiDoc),
            _ => 0
        };
    }
}