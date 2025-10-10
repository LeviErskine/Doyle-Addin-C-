class dvl0
{
    public Scripting.Dictionary d0g0f0(Inventor.SheetMetalComponentDefinition cd, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// New Sheet Metal Part processing function
        /// Reference function d0g0f0
        /// 
        Scripting.Dictionary rt;
        // '
        Inventor.PartDocument pt;
        Inventor.PropertySet ps;
        Inventor.Parameter prThk;
        // '
        VbMsgBoxResult ck;
        long ec;
        string ed;
        // Dim op As Boolean
        Inventor.View v1;
        Inventor.View v2;
        // '
        double dLength;
        double dWidth;
        double dArea;
        string strWidth;
        string strLength;
        string strArea;
        // '
        double dHeight;
        double dfHtThk;
        string strDVNs;
        // '

        if (dc == null)
            d0g0f0 = d0g0f0(cd, new Scripting.Dictionary());
        else
        {
            rt = dc;
            if (cd == null)
            {
            }
            else
            {
                pt = cd.Document;
                ps = pt.PropertySets.Item(gnCustom);
                v1 = d0g1f0(pt);
                // op = pt.Open


                Information.Err.Clear();
                prThk = cd.Thickness;
                ec = Information.Err.Number;
                ed = Information.Err.Description;


                if (ec == 0)
                {
                    if (!cd.HasFlatPattern)
                    {
                        ck = newFmTest2().AskAbout(pt, null/* Conversion error: Set to default value for this argument */, "NO FLAT PATTERN!" + Constants.vbNewLine + "Try to generate one?");
                        if (ck == Constants.vbYes)
                        {
                            // If Not op Then Stop
                            // Want to see if forcing an Unfold
                            // causes an unopened document to open.
                            // Also want to see how Open Property
                            // relates to a referenced Document
                            // not (yet) separately opened.

                            Information.Err.Clear();
                            cd.Unfold();
                            if (Information.Err.Number == 0)
                            {
                                if (cd.HasFlatPattern)
                                    cd.FlatPattern.ExitEdit();
                            }
                            else
                                System.Diagnostics.Debugger.Break();// Couldn't make Flat Pattern
                            Information.Err.Clear();

                            v2 = d0g1f0(pt);
                            if (!v2 == null)
                            {
                                if (v1 == null)
                                    v2.Close();
                            }
                        }
                        else
                        {
                        }
                    }

                    if (cd.HasFlatPattern)
                    {
                        {
                            var withBlock = cd.FlatPattern;
                            // First, make sure it's VALID
                            {
                                var withBlock1 = withBlock.Body.RangeBox;
                                // Check height against thickness
                                // Valid flat pattern should return
                                // zero or VERY minimal difference
                                dHeight = (withBlock1.MaxPoint.Z - withBlock1.MinPoint.Z);
                                dfHtThk = Abs(dHeight - prThk.Value);

                                // the extent of the face.
                                // Extract the width, length and area from the range.
                                dLength = (withBlock1.MaxPoint.X - withBlock1.MinPoint.X);
                                dWidth = (withBlock1.MaxPoint.Y - withBlock1.MinPoint.Y);
                                dArea = dLength * dWidth;
                            }
                            // Stop
                            /// At this point, we should have enough
                            /// to check at least a few things,
                            /// and possibly pick out stock.
                            /// 
                            if (dfHtThk > 0.01)
                            {
                                // Stop 'and prep for machined (non sheet metal) specs
                                /// Pretty sure dimension values
                                /// come through in centimeters
                                /// so try converting them here
                                // sort3dimsUp
                                d0g1f4(cd);
                                rt = d0g1f3(pt, sort3dimsUp(dHeight / cvLenIn2cm, dWidth / cvLenIn2cm, dLength / cvLenIn2cm), rt);
                            }
                            else
                            {
                            }

                            if (dArea > 0)
                            {
                                /// an invalid flat pattern SHOULD have no geometry,
                                /// which means it SHOULD have no area to speak of.
                                /// '
                                /// One would think this obvious, in retrospect,
                                /// but one would not be surprised to be proven wrong.
                                /// Again.

                                {
                                    var withBlock1 = pt;

                                    // Convert values into document units.
                                    // This will result in strings that are identical
                                    // to the strings shown in the Extent dialog.
                                    {
                                        var withBlock2 = withBlock1.UnitsOfMeasure;
                                        strWidth = withBlock2.GetStringFromValue(dWidth, withBlock2.GetStringFromType(withBlock2.LengthUnits));
                                        strLength = withBlock2.GetStringFromValue(dLength, withBlock2.GetStringFromType(withBlock2.LengthUnits));
                                        strArea = withBlock2.GetStringFromValue(dArea, withBlock2.GetStringFromType(withBlock2.LengthUnits) + "^2");

                                        if (dfHtThk > 0.01)
                                            strDVNs = withBlock2.GetStringFromValue(dfHtThk, withBlock2.GetStringFromType(withBlock2.LengthUnits));
                                        else
                                            strDVNs = "";
                                    }
                                }
                            }
                            else
                            {
                                if (MsgBox(Join(Array("The flat pattern for this", "part has no features,", "and is likely not valid.", "", "Pause here to review?", "(Click 'NO' to just keep going)"), Constants.vbNewLine), Constants.vbYesNo, "Invalid Flat Pattern") == Constants.vbYes)
                                    System.Diagnostics.Debugger.Break();// and let the user look into it
                                Debug.Print(aiDocument(withBlock.Document).FullDocumentName);
                            }
                        }

                        // Add area to custom property set
                        // rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt)

                        // Add Width to custom property set
                        // rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt)

                        // Add Length to custom property set
                        // rt = dcWithProp(aiPropSet, pnLength, strLength, rt)

                        // Add AreaDescription to custom property set
                        // rt = dcWithProp(aiPropSet, pnArea, strArea, rt)

                        if (Strings.Len(strDVNs) > 0)
                        {
                        }
                    }
                    else
                    {
                    }
                }
                else
                    System.Diagnostics.Debugger.Break();
            }
            d0g0f0 = rt;
        }
    }
    // For Each dc In aiDocAssy(aiDocActive).ComponentDefinition.Occurrences: Debug.Print aiDocument(aiCompOcc(obOf(dc)).Definition.Document).Open, aiDocument(aiCompOcc(obOf(dc)).Definition.Document).FullDocumentName: Next
    // Looks like Open property will NOT distinguish documents in tab list from those not
    // All entries came up True

    public Inventor.View d0g1f0(Inventor.Document rf)
    {
        Inventor.View rt;
        Inventor.View vw;

        rt = null/* TODO Change to default(_) if this is not a reference type */;
        foreach (var vw in ThisApplication.Views)
        {
            if (vw.Document == rf)
                rt = vw;
        }
        d0g1f0 = rt;
    }

    public Inventor.SheetMetalComponentDefinition d0g1f1(Inventor.PartDocument rf)
    {
        if (rf == null)
            d0g1f1 = null/* TODO Change to default(_) if this is not a reference type */;
        else
            d0g1f1 = aiCompDefShtMetal(rf.ComponentDefinition);
    }

    public Variant noVal(VbVarType vt = )
    {
        if (vt & Constants.vbArray)
            noVal = Array();
        else
            switch (vt)
            {
                case object _ when Constants.vbString:
                    {
                        noVal = "";
                        break;
                    }

                case object _ when Constants.vbLong:
                    {
                        noVal = System.Convert.ToInt64(0);
                        break;
                    }

                case object _ when Constants.vbVariant:
                    {
                        noVal = Empty;
                        break;
                    }

                case object _ when Constants.vbInteger:
                    {
                        noVal = System.Convert.ToInt32(0);
                        break;
                    }

                case object _ when Constants.vbSingle:
                    {
                        noVal = System.Convert.ToSingle(0);
                        break;
                    }

                case object _ when Constants.vbDouble:
                    {
                        noVal = System.Convert.ToDouble(0);
                        break;
                    }

                case object _ when Constants.vbDecimal:
                    {
                        noVal = System.Convert.ToDecimal(0);
                        break;
                    }

                case object _ when Constants.vbCurrency:
                    {
                        noVal = CCur(0);
                        break;
                    }

                case object _ when Constants.vbBoolean:
                    {
                        noVal = System.Convert.ToBoolean(0);
                        break;
                    }

                case object _ when Constants.vbByte:
                    {
                        noVal = System.Convert.ToByte(0);
                        break;
                    }

                case Constants.vbEmpty:
                    {
                        noVal = Empty;
                        break;
                    }

                case object _ when Constants.vbNull:
                    {
                        noVal = Null;
                        break;
                    }

                case object _ when Constants.vbObject:
                    {
                        noVal = null/* TODO Change to default(_) if this is not a reference type */;
                        break;
                    }

                case object _ when Constants.vbDate:
                    {
                        System.Diagnostics.Debugger.Break(); // noVal = Empty
                        break;
                    }

                case object _ when vbError:
                    {
                        System.Diagnostics.Debugger.Break(); // noVal = Empty
                        break;
                    }

                case object _ when vbDataObject:
                    {
                        System.Diagnostics.Debugger.Break(); // noVal = Empty
                        break;
                    }

                case object _ when Constants.vbUserDefinedType:
                    {
                        System.Diagnostics.Debugger.Break(); // noVal = Empty
                        break;
                    }
            }
    }

    public double[] pt3d(double d0 = 0#, double d1 = 0#, double d2 = 0#)
    {
        double[] rt = new double[3];

        rt[0] = d0;
        rt[1] = d1;
        rt[2] = d2;

        pt3d = rt;
    }

    public double[] sort3dimsUp(double d0, double d1, double d2)
    {
        double[] rt;

        if (d1 < d0)
            rt = sort3dimsUp(d1, d0, d2);
        else if (d2 < d1)
            rt = sort3dimsUp(d2, d0, d1);
        else
        {
            rt = new double[3];
            rt[0] = d0;
            rt[1] = d1;
            rt[2] = d2;
        }
        sort3dimsUp = rt;
    }

    public double[] sort3dimsDn(double d0, double d1, double d2)
    {
        double[] rt;

        if (d1 > d0)
            rt = sort3dimsDn(d1, d0, d2);
        else if (d2 > d1)
            rt = sort3dimsDn(d2, d0, d1);
        else
        {
            rt = new double[3];
            rt[0] = d0;
            rt[1] = d1;
            rt[2] = d2;
        }
        sort3dimsDn = rt;
    }

    public double[] aiBoxDims(Inventor.Box RefBox)
    {
        double[] rt;
        Inventor.Point mx;
        Inventor.Point mn;

        {
            var withBlock = RefBox;
            mx = withBlock.MaxPoint;
            mn = withBlock.MinPoint;
        }

        rt[0] = mx.X - mn.X;
        rt[1] = mx.Y - mn.Y;
        rt[2] = mx.Z - mn.Z;

        aiBoxDims = rt;
    }

    public Inventor.Box aiBoxSortDown(Inventor.Box RefBox)
    {
        Inventor.Box rt;
        Inventor.Point mx;
        Inventor.Point mn;

        {
            var withBlock = ThisApplication.TransientGeometry;
            rt = withBlock.CreateBox();
            {
                var withBlock1 = RefBox;
                mx = withBlock1.MaxPoint;
                mn = withBlock1.MinPoint;
                rt.PutBoxData(pt3d(), sort3dimsDn(mx.X - mn.X, mx.Y - mn.Y, mx.Z - mn.Z));
            }
        }

        aiBoxSortDown = rt;
    }

    public Scripting.Dictionary dcSteelType2Spec6()
    {
        Scripting.Dictionary rt;

        rt = new Scripting.Dictionary();
        {
            var withBlock = rt;
            withBlock.Add("Steel, Mild", "MS");
            withBlock.Add("Stainless Steel", "SS");
            withBlock.Add("Stainless Steel, Austenitic", "SS");
        }

        dcSteelType2Spec6 = rt;
    }

    public string steelSpec6(string stl, long Ask = 0)
    {
        string rt;
        // With dcSteelType2Spec6()
        // If .Exists(stl) Then
        // steelSpec6 = .Item(stl)
        // Else
        // steelSpec6 = ""
        // End If
        // End With

        switch (stl)
        {
            case "Stainless Steel":
                {
                    rt = "SS";
                    break;
                }

            case "Stainless Steel, Austenitic":
                {
                    rt = "SS";
                    break;
                }

            case "Stainless Steel 304":
                {
                    rt = "SS";
                    break;
                }

            case "Steel, Mild":
                {
                    rt = "MS";
                    break;
                }

            case "Rubber":
                {
                    rt = "";  // LG
                    break;
                }

            case "Rubber, Silicone":
                {
                    rt = "";  // LG
                    break;
                }

            case "UHMW, White":
                {
                    rt = "";  // LG
                    break;
                }

            default:
                {
                    if (Ask)
                    {
                        Debug.Print("=== UNKNOWN MATERIAL ===");
                        Debug.Print("   (" + stl + ")");
                        Debug.Print("Please supply a code for Specification 6,");
                        Debug.Print("if applicable, on the line below, and");
                        Debug.Print("press [ENTER] or [RETURN] to modify.");
                        Debug.Print("Press [F5] when ready to continue.");
                        Debug.Print("rt  = \"\" '<-( place code between double quotes )");
                        System.Diagnostics.Debugger.Break();
                    }
                    else
                        rt = "";
                    break;
                }
        }

        steelSpec6 = rt;
    }

    public Scripting.Dictionary d0g1f3(Inventor.PartDocument rf, double[] dm, Scripting.Dictionary dc = )
    {
        Scripting.Dictionary rt;
        // '
        double aspect;
        double offSqr;
        double length;
        // '
        string rmType;
        string rmSpc6;
        string rmItem;
        string rmUnit;
        string rmQty;

        if (dc == null)
            rt = d0g1f3(rf, dm, new Scripting.Dictionary());
        else
        {
            rt = dc;

            aspect = dm[1] / dm[0];
            offSqr = aspect - 1#;
            length = dm[2];
            rmType = rf.PropertySets(gnDesign).Item(pnMaterial).Value;
            rmSpc6 = steelSpec6(rmType);

            // Debug.Print ".Add """ & rmType & """, """""
            Debug.Print("Material: " + rmType + " (" + rmSpc6 + ")");
            /// 
            Debug.Print("Cross Section: " + Format(dm[1], "0.000") + " X " + Format(dm[0], "0.000"));
            Debug.Print("Length: " + Format(dm[2], "0.000"));
            if (offSqr < 0.01)
                Debug.Print("Likely Square or Round");
            else if (offSqr > 20)
                Debug.Print("Likely Sheet or Plate");
            else
                Debug.Print("Likely Rectangular, Uneven?");

            rmItem = rmType + "???"; // User has to help from info given
            rmQty = Format(dm[2], "0.000");
            rmUnit = "IN";

            dc.Add("RM", rmItem);
            dc.Add("RMQTY", rmQty);
            dc.Add("RMUNIT", rmUnit);
            System.Diagnostics.Debugger.Break();
        }
        d0g1f3 = rt;
    }

    public Scripting.Dictionary d0g1f4(Inventor.SheetMetalComponentDefinition cd)
    {
        Inventor.SurfaceBody sb;
        Inventor.Face fc;
        Inventor.Point pt;

        Inventor.ReferenceKeyManager rkMgr;
        long kyContx;
        byte[] kyBytes;
        string kyLabel;

        Scripting.Dictionary d0;
        Scripting.Dictionary d1;
        Variant ky;
        Variant k1;

        d0 = new Scripting.Dictionary();
        d1 = new Scripting.Dictionary();

        rkMgr = aiDocument(cd.Document).ReferenceKeyManager;

        kyContx = rkMgr.CreateKeyContext;

        foreach (var sb in cd.SurfaceBodies)
        {
            foreach (var fc in sb.Faces)
            {
                fc.GetReferenceKey(kyBytes, kyContx);
                kyLabel = rkMgr.KeyToString(kyBytes);
                d1.Add(kyLabel, fc);
                {
                    var withBlock = fc.Evaluator;
                    Debug.Print(TypeName(fc.Geometry) + "(" + System.Convert.ToHexString(0 + IIf(withBlock.IsExtrudedShape, 1, 0) + IIf(withBlock.IsRevolvedShape, 2, 0)) + ")");
                    {
                        var withBlock1 = withBlock.RangeBox;
                        d0.Add(d0.Count, withBlock1.MaxPoint);
                        d0.Add(d0.Count, withBlock1.MinPoint);
                    }
                }
            }

            foreach (var ky in d0.Keys)
            {
                pt = d0.Item(ky);
                foreach (var k1 in d1.Keys) // fc In sb.Faces
                {
                    fc = d1.Item(k1);
                    // If d1.Exists(k1) Then
                    if (!fc.Evaluator.RangeBox.Contains(pt))
                        d1.Remove(k1);
                }
                System.Diagnostics.Debugger.Break();
            }
        }
    }

    public Scripting.Dictionary d0g1f5(Inventor.SurfaceBody sb)
    {
        Inventor.Face fc;
        Inventor.Point pt;

        Scripting.Dictionary dFc;
        Scripting.Dictionary dPt;
        Variant kPt;

        dFc = new Scripting.Dictionary();
        dPt = new Scripting.Dictionary();

        foreach (var fc in sb.Faces)
        {
            dFc.Add(fc.InternalName, fc);
            {
                var withBlock = fc.Evaluator;
                Debug.Print(TypeName(fc.Geometry) + "(" + System.Convert.ToHexString(0 + IIf(withBlock.IsExtrudedShape, 1, 0) + IIf(withBlock.IsRevolvedShape, 2, 0)) + ")");
                {
                    var withBlock1 = withBlock.RangeBox;
                    dPt.Add(dPt.Count, withBlock1.MaxPoint);
                    dPt.Add(dPt.Count, withBlock1.MinPoint);
                }
            }
        }

        foreach (var kPt in dPt.Keys)
        {
            pt = dPt.Item(kPt);
            foreach (var fc in sb.Faces)
            {
                if (dFc.Exists(fc.InternalName))
                {
                    if (!fc.Evaluator.RangeBox.Contains(pt))
                    {
                        dFc.Remove(fc.InternalName);
                        if (dFc.Count == 0)
                            System.Diagnostics.Debugger.Break();
                    }
                }
            }
            System.Diagnostics.Debugger.Break();
        }

        d0g1f5 = dFc;
    }

    public string d0g1f6(Inventor.Face fc)
    {
        byte[] kyBytes;
        long kyContx;
        string rt;

        {
            var withBlock = aiDocument(fc.SurfaceBody.ComponentDefinition.Document).ReferenceKeyManager // .CreateKeyContext
       ;
            kyContx = withBlock.CreateKeyContext;
            fc.GetReferenceKey(kyBytes, kyContx);
            rt = withBlock.KeyToString(kyBytes);
        }
        d0g1f6 = rt;
    }

    public Inventor.Point aiPoint(object ob)
    {
        if (ob is Inventor.Point)
            aiPoint = ob;
        else
            aiPoint = null/* TODO Change to default(_) if this is not a reference type */;
    }

    /// d0g2: Testing

    /// 

    /// 

    public void d0g2f1()
    {
        /// Verify 3-way sorting function sort3dimsUp
        double[] ck;
        ck = sort3dimsUp(2, 3, 5); System.Diagnostics.Debugger.Break();
        ck = sort3dimsUp(2, 5, 3); System.Diagnostics.Debugger.Break();
        ck = sort3dimsUp(3, 2, 5); System.Diagnostics.Debugger.Break();
        ck = sort3dimsUp(3, 5, 2); System.Diagnostics.Debugger.Break();
        ck = sort3dimsUp(5, 2, 3); System.Diagnostics.Debugger.Break();
        ck = sort3dimsUp(5, 3, 2); System.Diagnostics.Debugger.Break();
    }

    public void d0g2f2()
    {
        /// Testing new spec pickup system
        Variant ky;

        {
            var withBlock = dcAiDocComponents(ThisApplication.ActiveDocument, null/* Conversion error: Set to default value for this argument */, 0);
            foreach (var ky in withBlock.Keys)
            {
                Debug.Print(ky);
                withBlock.Item(ky) = d0g0f0(aiCompDefShtMetal(aiCompDefOf(aiDocPart(withBlock.Item(ky)))));
                if (withBlock.Item(ky) == null)
                    withBlock.Remove(ky);
                else
                {
                }
            }
        }
    }

    public void d0g2f3()
    {
        /// Checking some behaviors
        /// on string arrays
        /// vs variants
        Variant ky;
        {
            var withBlock = new aiPropSetter();
            Debug.Print(Join(withBlock.PropList(), "|"));
            foreach (var ky in withBlock.PropList())
                Debug.Print(ky);
        }
    }

    public Scripting.Dictionary d0g2f4(Scripting.Dictionary dc)
    {
        /// Return Dictionary of ALLOCATED Property
        /// Values (True/False) attached to all components
        /// and subcomponents of the active Document.
        /// 
        /// Where the ALLOCATED Property is not present,
        /// represent it as "<default>"
        /// 
        Scripting.Dictionary rt;
        Inventor.Property pr;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                // Debug.Print ky
                {
                    var withBlock1 = aiDocument(withBlock.Item(ky)).PropertySets.Item(gnCustom);
                    Information.Err.Clear();
                    pr = withBlock1.Item("ALLOCATED");
                    if (Information.Err.Number == 0)
                        rt.Add(ky, System.Convert.ToHexString(pr.Value) + "|" + ky);
                    else
                        rt.Add(ky, "<default>" + "|" + ky);
                }
            }
        }
        d0g2f4 = rt;
    }
    // Debug.Print Join(d0g2f4(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbNewLine)

    public Scripting.Dictionary d0g2f6(Scripting.Dictionary dc, string pn, string gn = gnCustom, string df = "<NOPROP>")
    {
        /// Return Dictionary of named Property Values
        /// attached to all Inventor Documents
        /// in supplied Dictionary.
        /// 
        /// 
        /// Where the ALLOCATED Property is not present,
        /// represent it as "<default>"
        /// 
        Inventor.Document ad;
        Scripting.Dictionary rt;
        Inventor.Property pr;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocument(withBlock.Item(ky));
                if (ad == null)
                {
                }
                else
                {
                }
                {
                    var withBlock1 = ad.PropertySets.Item(gn);
                    Information.Err.Clear();
                    pr = withBlock1.Item("ALLOCATED");
                    if (Information.Err.Number == 0)
                        rt.Add(ky, System.Convert.ToHexString(pr.Value) + "|" + ky);
                    else
                        rt.Add(ky, "<default>" + "|" + ky);
                }
            }
        }
        d0g2f6 = rt;
    }
    // Debug.Print Join(d0g2f6(dcAiDocComponents(ThisApplication.ActiveDocument)).Items, vbNewLine)

    public Scripting.Dictionary d0g2f5(Scripting.Dictionary dc)
    {
        /// Attempt to "transpose" contents of Dictionary
        /// and return a dictionary of Items mapped
        /// to sub-Dictionaries containing all keys
        /// which mapped to each value
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant kv;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                kv = d0g2f5a(withBlock.Item(ky));
                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(kv))
                        withBlock1.Add(kv, new Scripting.Dictionary());

                    {
                        var withBlock2 = dcOb(withBlock1.Item(kv));
                        withBlock2.Add.Count(null/* Conversion error: Set to default value for this argument */, ky);
                    }
                }
            }
        }
    }

    public Variant d0g2f5a(Variant vr)
    {
        /// Return any Variant that is NOT an Object
        /// Object handling MAY be addressed later.
        /// 
        if (IsObject(vr))
            System.Diagnostics.Debugger.Break();
        else
            d0g2f5a = vr;
    }

    public Scripting.Dictionary dcMaterialUsage(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Inventor.PartDocument pt;
        Variant ky;
        string pn;
        string mt;

        rt = new Scripting.Dictionary();
        foreach (var ky in dc)
        {
            pt = aiDocPart(obOf(dc.Item(ky)));
            if (pt == null)
            {
            }
            else
            {
                pn = pt.PropertySets(gnDesign).Item(pnPartNum).Value;
                mt = pt.PropertySets(gnDesign).Item(pnMaterial).Value;
                {
                    var withBlock = rt;
                    if (withBlock.Exists(mt))
                        withBlock.Item(mt) = withBlock.Item(mt) + Constants.vbNewLine + Constants.vbTab + pn;
                    else
                        withBlock.Add(mt, mt + Constants.vbNewLine + Constants.vbTab + pn);
                }
                pt = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }
        dcMaterialUsage = rt;
    }
    // lsDump dcMaterialUsage(dcAiDocsOfType(kPartDocumentObject, dcAiDocComponents(aiDocActive()))).Items

    public Scripting.Dictionary d0g3f0()
    {
        Scripting.Dictionary rt;
        Inventor.Asset mt;

        rt = new Scripting.Dictionary();
        {
            var withBlock = ThisApplication.ActiveMaterialLibrary;
            foreach (var mt in withBlock.MaterialAssets)
                rt.Add(mt.DisplayName, mt);
        }
        d0g3f0 = rt;
    }
    // lsDump d0g3f0().Keys

    public Scripting.Dictionary dcGrpByPtNum(Scripting.Dictionary dc)
    {
        // '
        // '  Returns Dictionary of Dictionaries
        // '  grouping Inventor Documents in
        // '  supplied Dictionary by their Part
        // '  Numbers.
        // '
        // '  Ideally, each Document's Part Number
        // '  should be unique, and each sub Dictionary
        // '  should contain only one Document, however,
        // '  it is possible for more than one Document
        // '  to have the same Part Number.
        // '
        // '  By returning a Dictionary of Dictionaries,
        // '  this function provides a way for the client
        // '  to detect and respond to any conflicts.
        // '
        Scripting.Dictionary rt;
        Inventor.Document pt;
        Variant ky;
        string pn;
        string dn;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                dn = pt.FullDocumentName;
                pn = System.Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));
                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(pn))
                        withBlock1.Add(pn, new Scripting.Dictionary());

                    {
                        var withBlock2 = dcOb(withBlock1.Item(pn));
                        if (withBlock2.Exists(dn))
                            System.Diagnostics.Debugger.Break(); // because something went wrong
                        else
                            withBlock2.Add(dn, pt);
                    }
                }
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
            }
        }

        dcGrpByPtNum = rt;
    }

    public Scripting.Dictionary dcRemapByPtNum(Scripting.Dictionary dc)
    {
        // '  Returns Dictionary of Inventor
        // '  Documents keyed on Part Number
        Scripting.Dictionary rt;
        Scripting.Dictionary xt;
        Inventor.Document pt;
        Variant ky;
        string pn;

        rt = new Scripting.Dictionary();
        xt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                pn = System.Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign));
                /// 

                /// 
                // If pt.DocumentType = kPartDocumentObject Then
                // ''' UPDATE[2021.06.22]
                // ''' moving hardware component check outside of
                // ''' and before collision check in order to skip
                // ''' known hardware items entirely.
                // ''' '
                // ''' UPDATE[2021.06.21]
                // ''' implementing a set of checks for hardware components
                // ''' probably want to move outside Dictionary check
                // ''' to catch hardware elements before they're added.
                // ''' Could lead to trouble if this prevents new hardware
                // ''' Items from being added to Genius, but don't believe
                // ''' this is a high risk.
                // If aiDocPart(pt).ComponentDefinition.IsContentMember Then
                // 'it's commodity hardware
                // With cnGnsDoyle().Execute('            "select Family from vgMfiItems where Item = '"'            & pn & "';"'        )
                // If .BOF And .EOF Then
                // 'probably not in Genius
                // 'keep it -- may need added
                // ElseIf Split(.GetString('                adClipString, , "", vbVerticalTab'            ), vbVerticalTab)(0) = "D-HDWR" Then
                // pt = Nothing 'and move on
                // Else
                // Debug.Print ; 'Breakpoint Landing
                // 'Stop
                // End If
                // End If
                // End With
                // ElseIf pt.PropertySets.Item(gnDesign).Item(pnFamily).Value = "D-HDWR" Then
                // 'it's in commodity hardware family
                // pt = Nothing 'and move on
                // ElseIf InStr(1, "|D-HDWR|D-PTS|R-PTS|", "|" & pt.PropertySets.Item(gnDesign).Item(pnFamily).Value & "|") > 0 Then
                // 'it's PROBABLY hardware
                // 'but keep it, just in case
                // Debug.Print ; 'Breakpoint Landing
                // Else 'nothing special to worry about, probably
                // Debug.Print ; 'Breakpoint Landing
                // 'Stop
                // End If
                // Else 'we've got an Assembly
                // Debug.Print ; 'Breakpoint Landing
                // 'Stop
                // End If
                /// 

                /// 
                if (Strings.Len(pn) > 0)
                {
                    {
                        var withBlock1 = rt;
                        if (withBlock1.Exists(pn))
                        {
                            {
                                var withBlock2 = xt // report' it here
;
                                if (!withBlock2.Exists(pn))
                                    withBlock2.Add(pn, new Scripting.Dictionary());

                                {
                                    var withBlock3 = dcOb(withBlock2.Item(pn));
                                    withBlock3.Add(pt.FullDocumentName, pt);
                                }
                            }
                        }
                        else
                            withBlock1.Add(pn, pt);
                    }
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                }
                else
                {
                    Debug.Print(InputBox("This component has no part number:" + Constants.vbNewLine + pt.DisplayName + Constants.vbNewLine + System.Convert.ToHexString(aiDocPropVal(pt, pnDesc, gnDesign)) + Constants.vbNewLine + Constants.vbNewLine + "Copy file path from text box for later review.", pt.DisplayName, pt.FullDocumentName));
                    if (getFromClipBdWin10() == pt.FullDocumentName)
                    {
                    }
                    else if (MsgBox("Are you sure you want to continue" + Constants.vbNewLine + "without recording this file path?", Constants.vbExclamation + Constants.vbYesNo, "File Path not copied!") == Constants.vbNo)
                        System.Diagnostics.Debugger.Break();
                }
            }
        }

        if (xt.Count > 0)
            Debug.Print(MsgBox(Join(Array("The following Part Numbers are", "assigned to more than one Model:", "", Constants.vbTab + Join(xt.Keys, Constants.vbNewLine + Constants.vbTab), ""), Constants.vbNewLine), Constants.vbOKOnly | Constants.vbInformation, "Duplicate Part Numbers!"));

        dcRemapByPtNum = rt;
    }

    public Scripting.Dictionary dcRemapByFilePath(Scripting.Dictionary dc)
    {
        // '  Returns Dictionary of Inventor
        // '  Documents re-keyed to File Path.
        // '  Typically for a Dictionary
        // '  previously remapped to another
        // '  key (most likely Part Number)
        Scripting.Dictionary rt;
        Inventor.Document pt;
        Variant ky;
        string pn;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                pn = pt.FullDocumentName;
                rt.Add(pn, pt);
            }
        }
        dcRemapByFilePath = rt;
    }

    public Scripting.Dictionary dcRemapByPtNumFilePath(Scripting.Dictionary dc)
    {
        // '  Returns Dictionary of Inventor
        // '  Documents keyed on Part Number
        // '  combined with original key
        // '  (which SHOULD be full doc path)
        Scripting.Dictionary rt;
        Inventor.Document pt;
        Variant ky;
        string pn;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocument(withBlock.Item(ky));
                pn = System.Convert.ToHexString(aiDocPropVal(pt, pnPartNum, gnDesign)) + Constants.vbTab + pt.FullDocumentName; // ky
                rt.Add(pn, pt);
            }
        }
        dcRemapByPtNumFilePath = rt;
    }

    public Scripting.Dictionary dcGeniusItems()
    {
        // '  Generates Dictionary of Items in Genius
        // '  Formerly d0g3f2
        Scripting.Dictionary rt;
        ADODB.Field ky;
        ADODB.Field vl;

        rt = new Scripting.Dictionary();
        {
            var withBlock = cnGnsDoyle();
            {
                var withBlock1 = withBlock.Execute("select Item, ItemID from vgMfiItems");
                ky = withBlock1.Fields("Item");
                vl = withBlock1.Fields("ItemID");
                while (!withBlock1.EOF | withBlock1.BOF)
                {
                    rt.Add(ky.Value, vl.Value);
                    withBlock1.MoveNext();
                }
            }
        }
        dcGeniusItems = rt;
    }

    public Scripting.Dictionary[] d0g3f3(Scripting.Dictionary dc)
    {
        // '  Appears intended to separate
        // '  parts in Genius from those
        // '  not there yet. I don't think
        // '  this one was working quite
        // '  right yet. New kyPick system
        // '  should handle this properly,
        // '  now, in any case.
        Variant ky;
        Scripting.Dictionary[] rt = new Scripting.Dictionary[2];

        rt[0] = new Scripting.Dictionary();
        rt[1] = new Scripting.Dictionary();
        {
            var withBlock = dcGeniusItems();
            foreach (var ky in dc)
            {
                if (withBlock.Exists(ky))
                    rt[1].Add(ky, withBlock.Item(ky));
                else
                    rt[0].Add(ky, "");
            }
        }
        d0g3f3 = rt;
    }

    public long deConstrainAssyComponent(Inventor.ComponentOccurrence co)
    {
        // '  Deletes all constraints on an occurrence
        // '  !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
        Inventor.AssemblyConstraint cs;
        long ct;

        ct = 0;
        foreach (var cs in co.Constraints)
        {
            cs.Delete();
            ct = ct + 1;
        }

        deConstrainAssyComponent = ct;
    }

    public long deConstrainAssyDocument(Inventor.AssemblyDocument ad)
    {
        // '  Calls deConstrainAssyComponent over all occurrences
        // '  in an assembly to remove all their constraints
        // '  !!!DO NOT USE ON ANY PRODUCTION MODEL!!!
        // '  !!!That goes DOUBLE for THIS function!!!
        Inventor.ComponentOccurrence co;
        long ct;

        ct = 0;
        foreach (var co in ad.ComponentDefinition.Occurrences)
            ct = ct + deConstrainAssyComponent(co);

        deConstrainAssyDocument = ct;
    }

    public Scripting.Dictionary dcPartsInGeniusOrNot()
    {
        Scripting.Dictionary dcInGns;
        Scripting.Dictionary dcNotIn;
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        dcInGns = new Scripting.Dictionary();
        dcNotIn = dcRemapByPtNum(dcAiDocComponents(aiDocActive()));
        // dcNotIn = dcAiDocComponents(aiDocActive())
        // dcNotIn = dcAssyDocsByPtNum(aiDocActive())
        {
            var withBlock = dcFrom2Fields(cnGnsDoyle().Execute("select Item from vgMfiItems"), "Item", "Item");
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcNotIn;
                    if (withBlock1.Exists(ky))
                    {
                        dcInGns.Add(ky, withBlock1.Item(ky));
                        withBlock1.Remove(ky);
                    }
                }
            }
        }
        // Stop

        rt.Add("INGNS", dcInGns);
        rt.Add("NOTIN", dcNotIn);
        dcPartsInGeniusOrNot = rt;
    }

    public Scripting.Dictionary d0g4f1()
    {
        Scripting.Dictionary rt;
        ADODB.Connection cn;
        ADODB.Recordset rs;

        rt = new Scripting.Dictionary();
        cn = cnGnsDoyle();
        rs = cn.Execute("select Item from vgMfiItems");
        d0g4f1 = rt;
    }

    public Inventor.ComponentOccurrence compOccFromProxy(Inventor.ComponentOccurrence oc)
    {
        Inventor.ComponentOccurrenceProxy px;

        // If TypeOf oc Is Inventor.ComponentOccurrenceProxy Then
        // px = oc
        // Stop
        // compOccFromProxy = compOccFromProxy(px.ContainingOccurrence)
        // Else
        compOccFromProxy = oc;
    }

    public kyPick nuPicker(kyPick ob = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (ob == null)
            nuPicker = new kyPick();
        else
            nuPicker = ob;
    }

    public dcSplitter nuSplitter(dcSplitter ob = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (ob == null)
            nuSplitter = new dcSplitter();
        else
            nuSplitter = ob;
    }

    public Scripting.Dictionary dcAiPartDocsWithRMv0(Scripting.Dictionary dcIn, long WantOut = 0)
    {
        Variant ky;
        Scripting.Dictionary[] rt = new Scripting.Dictionary[2];

        if (WantOut < 0 | WantOut > 1)
            dcAiPartDocsWithRMv0 = dcAiPartDocsWithRMv0(dcIn, 1);
        else
        {
            // With nuSplitter().WithSel(New kyPickAiDocWithRM)
            {
                var withBlock = new kyPickAiDocWithRM();
                rt[0] = withBlock.dcIn();
                rt[1] = withBlock.dcOut();
                foreach (var ky in dcIn.Keys)
                {
                    {
                        var withBlock1 = withBlock.dcFor(dcIn.Item(ky));
                        if (withBlock1.Exists(ky))
                            System.Diagnostics.Debugger.Break();
                        else
                            withBlock1.Add(ky, dcIn.Item(ky));
                    }
                }
            }
            dcAiPartDocsWithRMv0 = rt[WantOut];
        }
    }
    // Debug.Print txDumpLs(dcAiPartDocsWithRMv0(dcAiDocComponents(aiDocActive()), 1).Keys)

    public kyPick kyScanned(Scripting.Dictionary dcIn, kyPick pkr = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (pkr == null)
            kyScanned = kyScanned(dcIn, new kyPick());
        else
            kyScanned = pkr.AfterScanning(dcIn);
    }
    // Debug.Print txDumpLs(kyScanned(dcAiDocComponents(aiDocActive()), New kyPickAiPartVsAssy).dcIn().Keys)

    public Scripting.Dictionary dcAiDocsPicked(Scripting.Dictionary dcIn, kyPick pkr = null/* TODO Change to default(_) if this is not a reference type */, long WantOut = 0)
    {
        Variant ky;
        Scripting.Dictionary[] rt = new Scripting.Dictionary[2];

        if (pkr == null)
            dcAiDocsPicked = dcAiDocsPicked(dcIn, new kyPick(), WantOut);
        else if (WantOut < 0 | WantOut > 1)
            dcAiDocsPicked = dcAiDocsPicked(dcIn, pkr, 1);
        else
        {
            {
                var withBlock = pkr;
                rt[0] = withBlock.dcIn();
                rt[1] = withBlock.dcOut();
                foreach (var ky in dcIn.Keys)
                {
                    {
                        var withBlock1 = withBlock.dcFor(dcIn.Item(ky));
                        if (withBlock1.Exists(ky))
                            System.Diagnostics.Debugger.Break();
                        else
                            withBlock1.Add(ky, dcIn.Item(ky));
                    }
                }
            }
            dcAiDocsPicked = rt[WantOut];
        }
    }
    // Debug.Print txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), 1).Keys)
    // Debug.Print txDumpLs(dcAiDocsPicked(dcAiDocComponents(aiDocActive()), New kyPickAiDocContentCtr, 0).Keys)

    public Scripting.Dictionary dcAiPartDocsWithRM(Scripting.Dictionary dcIn, long WantOut = 0)
    {
        dcAiPartDocsWithRM = dcAiDocsPicked(dcAiDocsPicked(dcIn, new kyPickAiPartVsAssy(), 0), new kyPickAiDocWithRM(), WantOut);
    }
    // Debug.Print txDumpLs(dcAiPartDocsWithRM(dcAiDocComponents(aiDocActive()), 1).Keys)

    public void d0g5f0()
    {
    }

    public Scripting.Dictionary dcAiDocGrpsByForm(Scripting.Dictionary dcIn)
    {
        /// dcAiDocGrpsByForm -- Separate a Dictionary
        /// of Inventor Documents into
        /// categorical sub-Dictionaries
        /// according to various criteria:
        /// -   PRCH    Purchased Items
        /// -   ASSY    Assemblies
        /// -   HDWR    Hardware (Content Center)
        /// -   DBAR    Structural Parts (was BSTK)
        /// Subtype NOT Sheet Metal
        /// (some "Sheet Metal" Parts
        /// might also technically
        /// belong here. see below)
        /// -   MAYB    Likely Structural Parts
        /// Sheet Metal subtype, but
        /// has either no flat pattern,
        /// or an invalid one.
        /// -   SHTM    Sheet Metal Parts
        /// Indicated both by Subtype
        /// and presence of a valid
        /// flat pattern.
        /// 
        /// Presence in Genius, a distinction
        /// originally intended to be made here,
        /// is now planned to be made to a separate
        /// Dictionary, possibly also subcategorized,
        /// to be processed in conjunction with
        /// the results of this function.
        /// 
        /// The notion of passing different subgroups
        /// of this Dictionary to separate handlers
        /// for more specialized processing, while
        /// still an option, is no longer considered
        /// its primary role. Instead, the set is
        /// expected to be used in a form application
        /// which will present the various groups to
        /// the user for review, and modification as
        /// desired or necessary.
        /// 
        /// REV[2022.03.08.1212] All new text
        /// in function description above.
        /// see notes_2022-0308_general-01.txt
        /// for prior description
        /// 
        Scripting.Dictionary rt;
        // Dim pkGns As kyPick
        kyPick pkBuy;
        kyPick pkPrt;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkMbe;

        rt = new Scripting.Dictionary();

        /// REV[2022.03.08.1112]
        /// Disabled split on presence
        /// in Genius. Believe better
        /// addressed separately
        // '  separate items already in Genius
        // '  from those not yet in
        // pkGns = nuPicker('    New kyPickInGenius').AfterScanning(dcIn)
        // '  NOTE: no further processing
        // '  implemented on this yet
        // '  MIGHT be better applied
        // '  at a different stage?

        /// REV[2022.03.08.1115]
        /// Add division on Purchased Parts
        /// with "out" Dictionary replacing
        /// main for Part/Assy separation.
        pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        {
            var withBlock = pkBuy;
            rt.Add("PRCH", withBlock.dcIn);

            // '  separate parts from assemblies
            pkPrt = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkPrt;
            rt.Add("ASSY", withBlock.dcOut);

            // '  isolate Content Center
            // '  parts from the rest
            pkCtC = nuPicker(new kyPickAiDocContentCtr()).AfterScanning(withBlock.dcIn);
        }

        {
            var withBlock = pkCtC;
            rt.Add("HDWR", withBlock.dcIn);

            // '  separate (potential) sheet
            // '  metal parts from non-sheet
            pkSht = nuPicker(new kyPickAiSheetMetal()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkSht;
            rt.Add("DBAR", withBlock.dcOut);
            pkMbe = nuPicker(new kyPickAiShMtl4sure()).AfterScanning(withBlock.dcIn);
        }

        {
            var withBlock = pkMbe;
            rt.Add("SHTM", withBlock.dcIn);
            rt.Add("MAYB", withBlock.dcOut);
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing

        dcAiDocGrpsByForm = rt;
    }

    public Scripting.Dictionary dcAiDocGrpsByFormAndIfac(Scripting.Dictionary dcIn)
    {
        /// dcAiDocGrpsByFormAndIfac -- Separate a Dictionary
        /// of Inventor Documents into
        /// categorical sub-Dictionaries
        /// according to various criteria:
        /// -   PRCH    Purchased Items
        /// -   ASSY    Assemblies
        /// -   IASM    iAssembly Factories
        /// -   IPRT    iPart Factories
        /// -   HDWR    Hardware (Content Center)
        /// -   DBAR    Structural Parts (was BSTK)
        /// Subtype NOT Sheet Metal
        /// (some "Sheet Metal" Parts
        /// might also technically
        /// belong here. see below)
        /// -   MAYB    Likely Structural Parts
        /// Sheet Metal subtype, but
        /// has either no flat pattern,
        /// or an invalid one.
        /// -   SHTM    Sheet Metal Parts
        /// Indicated both by Subtype
        /// and presence of a valid
        /// flat pattern.
        /// 
        /// REV[2023.01.24.0921]
        /// copied from dcAiDocGrpsByForm to produce
        /// new variant with additional groupings for
        /// iPart (IPRT) and iAssembly (IASM) members.
        /// additional groups for their corresponding
        /// Factories will likely also be added.
        /// 
        /// text of prior REV[2022.03.08.1212] removed.
        /// see dcAiDocGrpsByForm for that.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        /// REV[2023.01.24.1156]
        /// add working Dictionary
        /// to collect iAssembly
        /// and iPart Factories
        Variant ky;
        Inventor.Document mb;
        Inventor.Document md;
        string fp;

        // Dim pkGns As kyPick
        kyPick pkBuy;
        kyPick pkPrt;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkMbe;

        /// REV[2023.01.24.1009]
        /// add new pickers for
        /// iAssemblies and iParts
        kyPick pkIas;
        kyPick pkIpt;

        rt = new Scripting.Dictionary();

        /// REV[2022.03.08.1112]
        /// Disabled split on presence
        /// in Genius. Believe better
        /// addressed separately
        // '  separate items already in Genius
        // '  from those not yet in
        // pkGns = nuPicker('    New kyPickInGenius').AfterScanning(dcIn)
        // '  NOTE: no further processing
        // '  implemented on this yet
        // '  MIGHT be better applied
        // '  at a different stage?

        /// REV[2022.03.08.1115]
        /// Add division on Purchased Parts
        /// with "out" Dictionary replacing
        /// main for Part/Assy separation.
        pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        {
            var withBlock = pkBuy;
            if (dcIn.Count > 0)
                rt.Add("PRCH", withBlock.dcIn);

            // '  separate parts from assemblies
            pkPrt = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkPrt;
            // '  separate iAssembly members
            // '  from stand-alone assemblies
            pkIas = nuPicker(new kyPickAiAssyMember()).AfterScanning(withBlock.dcOut);

            // '  isolate Content Center
            // '  parts from the rest
            pkCtC = nuPicker(new kyPickAiDocContentCtr()).AfterScanning(withBlock.dcIn);
        }

        {
            var withBlock = pkIas;
            if (withBlock.dcOut.Count > 0)
                rt.Add("ASSY", withBlock.dcOut);
            {
                var withBlock1 = withBlock.dcIn;
                wk = new Scripting.Dictionary();

                foreach (var ky in withBlock1.Keys)
                {
                    {
                        var withBlock2 = aiDocAssy(withBlock1.Item(ky)).ComponentDefinition;
                        mb = withBlock2.Parent;
                        {
                            var withBlock3 = withBlock2.iAssemblyMember.ParentFactory.Parent;
                            md = withBlock3.Parent;
                            fp = md.FullDocumentName;

                            {
                                var withBlock4 = wk;
                                if (!withBlock4.Exists(fp))
                                {
                                    withBlock4.Add(fp, new Scripting.Dictionary());
                                    dcOb(withBlock4.Item(fp)).Add("", md);
                                }

                                {
                                    var withBlock5 = dcOb(withBlock4.Item(fp));
                                    if (withBlock5.Exists(ky))
                                        System.Diagnostics.Debugger.Break();
                                    else
                                        withBlock5.Add(ky, mb);
                                }
                            }
                        }
                    }
                }

                if (wk.Count > 0)
                    rt.Add("IASM", wk);
            }
        }

        {
            var withBlock = pkCtC;
            if (withBlock.dcIn.Count > 0)
                rt.Add("HDWR", withBlock.dcIn);

            // '  separate iPart members
            // '  from stand-alone parts
            pkIpt = nuPicker(new kyPickAiPartMember()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkIpt;
            {
                var withBlock1 = withBlock.dcIn;
                wk = new Scripting.Dictionary();

                foreach (var ky in withBlock1.Keys)
                {
                    {
                        var withBlock2 = aiDocPart(withBlock1.Item(ky)).ComponentDefinition;
                        mb = withBlock2.Document; // .Parent
                        md = aiDocPart(withBlock2.iPartMember.ParentFactory.Parent); // .PropertySets.Parent ' .Parent
                        {
                            var withBlock3 = md;
                            fp = md.FullDocumentName;

                            {
                                var withBlock4 = wk;
                                if (!withBlock4.Exists(fp))
                                {
                                    withBlock4.Add(fp, new Scripting.Dictionary());
                                    dcOb(withBlock4.Item(fp)).Add("", md);
                                }

                                {
                                    var withBlock5 = dcOb(withBlock4.Item(fp));
                                    if (withBlock5.Exists(ky))
                                        System.Diagnostics.Debugger.Break();
                                    else
                                        withBlock5.Add(ky, mb);
                                }
                            }
                        }
                    }
                }

                if (wk.Count > 0)
                    rt.Add("IPRT", wk);
            }

            // '  add iPart Factories to
            // '  Dictionary of non-Members
            {
                var withBlock1 = withBlock.dcOut;
                foreach (var ky in wk.Keys)
                {
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(ky, dcOb(wk.Item(ky)).Item(""));
                }
            }

            // '  separate (potential) sheet
            // '  metal parts from non-sheet
            pkSht = nuPicker(new kyPickAiSheetMetal()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkSht;
            if (withBlock.dcOut.Count > 0)
                rt.Add("DBAR", withBlock.dcOut);
            pkMbe = nuPicker(new kyPickAiShMtl4sure()).AfterScanning(withBlock.dcIn);
        }

        {
            var withBlock = pkMbe;
            if (withBlock.dcIn.Count > 0)
                rt.Add("SHTM", withBlock.dcIn);
            if (withBlock.dcOut.Count > 0)
                rt.Add("MAYB", withBlock.dcOut);
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing

        dcAiDocGrpsByFormAndIfac = rt;
    }

    public Scripting.Dictionary d0g5f2(Scripting.Dictionary dcIn)
    {
        /// function d0g5f2
        /// 
        /// INITIATED[2021.03.23]
        /// this variant on dcAiDocGrpsByForm
        /// is intended to separate items
        /// in Genius from those not yet in
        /// and purchased items from those
        /// to be made, cross-referencing
        /// the two to determine individual
        /// needs for processing.
        /// 
        /// presently in a nonfunctional state
        /// as End of Day approaches. will hope
        /// to continue development tomorrow
        /// 
        Scripting.Dictionary rt;
        kyPick pkGns;
        kyPick pkPvA;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkBuy;

        rt = new Scripting.Dictionary();

        // '  separate items already in Genius
        // '  from those not yet in
        pkGns = nuPicker(new kyPickInGenius()).AfterScanning(dcIn);

        // '  separate purchased items
        // '  from those to be made
        pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);
        // '  NOTE: no further processing
        // '  implemented on this yet
        // '  MIGHT be better applied
        // '  at a different stage?



        // '  separate parts from assemblies
        // pkPvA = nuPicker('    New kyPickAiPartVsAssy').AfterScanning(dcIn)
        // 'rt.Add "ASSY", dck pkPvA.dcOut

        // '  isolate Content Center
        // '  parts from the rest
        // pkCtC = nuPicker('    New kyPickAiDocContentCtr').AfterScanning(pkPvA.dcIn)
        // 'rt.Add "HDWR", pkCtC.dcIn

        // '  separate (potential)
        // '  sheet metal parts
        // '  from non-sheet
        // pkSht = nuPicker('    New kyPickAiSheetMetal').AfterScanning(pkCtC.dcOut)
        // rt.Add "SHTM", pkSht.dcIn
        // rt.Add "BSTK", pkSht.dcOut

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing

        d0g5f2 = rt;
    }

    public Scripting.Dictionary d0g5f3(Scripting.Dictionary dcIn)
    {
        /// function d0g5f3 -- essentially a recreation of dcAiDocGrpsByForm
        /// 
        Scripting.Dictionary rt;
        kyPick pkBuy;
        kyPick pkPrt;
        kyPick pkCtC;
        kyPick pkSht;
        kyPick pkMbe;
        // Dim pkGns As kyPick
        // Dim pk___ As kyPick

        rt = new Scripting.Dictionary();
        // pkGns = nuPicker(New kyPickInGenius).AfterScanning(dcIn)

        pkBuy = nuPicker(new kyPickAiDocPurchased()).AfterScanning(dcIn);

        {
            var withBlock = pkBuy;
            rt.Add("PRCH", withBlock.dcIn);
            pkPrt = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkPrt;
            rt.Add("ASSY", withBlock.dcOut);
            pkCtC = nuPicker(new kyPickAiDocContentCtr()).AfterScanning(withBlock.dcIn);
        }

        {
            var withBlock = pkCtC;
            rt.Add("HDWR", withBlock.dcIn);
            pkSht = nuPicker(new kyPickAiSheetMetal()).AfterScanning(withBlock.dcOut);
        }

        {
            var withBlock = pkSht;
            rt.Add("DBAR", withBlock.dcOut);
            pkMbe = nuPicker(new kyPickAiShMtl4sure()).AfterScanning(withBlock.dcIn);
        }

        {
            var withBlock = pkMbe;
            rt.Add("MAYB", withBlock.dcOut);
            rt.Add("SHTM", withBlock.dcIn);
        }

        d0g5f3 = rt;
    }

    public long dcDepthAiDocGrp(Scripting.Dictionary dc)
    {
        Variant vr;
        object ob;
        long mx;
        long dx;
        long ck;
        long rt;

        {
            var withBlock = dc;
            mx = withBlock.Count;

            if (mx == 0)
                dcDepthAiDocGrp = 0; // indeterminate
            else
            {
                dx = 0;

                do
                {
                    // vr = Array(.Item(.Keys(dx)))
                    ob = obOf(withBlock.Item(withBlock.Keys(dx)));
                    if (ob == null)
                        ck = -1;
                    else if (ob is Scripting.Dictionary)
                    {
                        ck = dcDepthAiDocGrp(ob);
                        if (ck > 0)
                            ck = 1 + ck;
                    }
                    else if (ob is Inventor.Document)
                        ck = 1;
                    else
                        ck = -1;
                    dx = dx + 1;
                    if (dx > mx)
                        ck = -1;
                }
                while (ck == 0)// vr(0)// invalid// invalid // invalid
    ; // indeterminate

                dcDepthAiDocGrp = ck;
            }
        }
    }

    public fmIfcTest05A nu_fmIfcTest05A(Scripting.Dictionary dcIn = null/* TODO Change to default(_) if this is not a reference type */)
    {
        {
            var withBlock = new fmIfcTest05A();
            nu_fmIfcTest05A = withBlock.Using(dcIn);
        }
    }

    public fmTest05 nu_fmTest05A(Scripting.Dictionary dcIn = null/* TODO Change to default(_) if this is not a reference type */)
    {
        System.Diagnostics.Debugger.Break(); // DO NOT USE THIS FUNCTION!
                                             // instead, use the Interface
                                             // generator nu_fmIfcTest05A

        {
            var withBlock = new fmTest05();
            nu_fmTest05A = withBlock.Holding(dcIn); // .Using(dcIn)
        }
    }

    public string lsAssyMembers(Inventor.AssemblyDocument aiAssy)
    {
        Scripting.Dictionary dc;
        string pn;
        string rt;
        Variant ky;

        dc = dcAiDocsByPtNum(dcAssyComponentsImmediate(aiAssy)); // dcAiDocPartNumbers
        pn = Constants.vbNewLine + aiAssy.PropertySets.Item(gnDesign).Item(pnPartNum).Value + Constants.vbTab;
        rt = pn + Join(dc.Keys, pn);

        {
            var withBlock = nuPicker(new kyPickAiPartVsAssy()).AfterScanning(dc);
            {
                var withBlock1 = withBlock.dcOut;
                foreach (var ky in withBlock1.Keys)
                    rt = rt + lsAssyMembers(aiDocument(obOf(withBlock1.Item(ky))));
            }
        }

        lsAssyMembers = rt;
    }

    public string d0g6f0(Inventor.Document AiDoc)
    {
        // '  Try to pick a distinct listing name
        // '  for a supplied Inventor Document
        string rt;
        string ds;

        {
            var withBlock = AiDoc;
            {
                var withBlock1 = withBlock.PropertySets(gnDesign);
                rt = Trim(withBlock1.Item(pnPartNum).Value);
                ds = Trim(withBlock1.Item(pnDesc).Value);
            }

            if (Strings.Len(rt) > 0)
            {
                if (Strings.Len(ds) > 0)
                    rt = rt + ": " + ds;
            }
            else if (Strings.Len(ds) > 0)
                rt = ds;

            if (Strings.Len(rt) == 0)
            {
                ds = withBlock.FullFileName;
                if (Strings.Len(ds) > 0)
                {
                    {
                        var withBlock1 = nuFso().GetFile(ds);
                        rt = withBlock1.Name + " (" + withBlock1.ParentFolder.Path + ")";
                    }
                }
                else
                    rt = withBlock.DisplayName;
            }

            d0g6f0 = rt;
        }
    }

    public void d0g6f1()
    {
        // '
        // '  testing form class fmTest0
        // '
        {
            var withBlock = new fmTest0();
            withBlock.imTNail.Visible = false;
            Debug.Print.Controls.Count();
            System.Diagnostics.Debugger.Break();
        }
    }

    public Scripting.Dictionary d0g6f2(Scripting.Dictionary dc)
    {
        /// Call this one from inside dcAiDocGrpsByForm (above)
        /// Try: debug.Print txDumpLs(d0g6f2(pkPvA.dcIn).Keys)
        /// 
        Scripting.Dictionary rt;
        Inventor.Document ad;
        Inventor.Property pr;
        Variant ky;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ad = aiDocument(withBlock.Item(ky));
                if (ad == null)
                {
                }
                else
                {
                    pr = ad.PropertySets(gnDesign).Item(pnFamily);
                    rt.Add(ky, pr);
                    {
                        var withBlock1 = pr;
                        withBlock1.Value = "R-PTS";
                    }
                }
            }
        }

        d0g6f2 = rt;
    }

    public void d0g6f3()
    {
        // '
        // '  testing new empty form class fmEmpty
        // '
        {
            var withBlock = new fmEmpty();
            // .imTNail.Visible = False
            {
                var withBlock1 = withBlock.Controls.Add("Forms.ComboBox.1", "test", true);
                Debug.Print.Name();
                withBlock1.Left = 10;
                withBlock1.Top = 10;
            }
            Debug.Print.Controls.Count();
            withBlock.Show(1);
            System.Diagnostics.Debugger.Break();
        }
    }

    public string d0g7f0()
    {
        /// This function used to transfer Property Values
        /// from blank model files GR12 ~ GR20
        /// to new versions generated from Intraflo's
        /// supplied STEP files. Save for reference,
        /// but this version should not likely be used
        /// as is for other tasks without review.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary dcPr;
        Inventor.Document sd;
        Inventor.Document td;
        Inventor.PropertySet psSc;
        Inventor.PropertySet psTg;
        Inventor.Property prSc;
        Inventor.Property prTg;
        Variant ky;
        Variant pn;
        string sn;

        rt = dcAiDocsByPtNum(dcAssyComponentsImmediate(aiDocActive()));
        {
            var withBlock = rt;
            foreach (var ky in withBlock.Keys)
            {
                Debug.Print(ky); sd = aiDocument(obOf(withBlock.Item(ky)));
                if (UCase(Left(ky, 2)) == "GR")
                {
                    sn = sd.PropertySets(gnDesign).Item(pnStockNum).Value;
                    if (withBlock.Exists(sn))
                    {
                        td = aiDocument(obOf(withBlock.Item(sn)));
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Stop
                        psTg = td.PropertySets(gnCustom);
                        dcPr = dcAiPropsInSet(psTg);
                        psSc = sd.PropertySets(gnCustom);
                        foreach (var prSc in psSc)
                        {
                            if (dcPr.Exists(prSc.Name))
                            {
                                {
                                    var withBlock1 = psTg.Item(prSc.Name);
                                    withBlock1.Value = prSc.Value;
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing Point -- Ctrl-F8 to here
                                }
                            }
                            else
                            {
                                var withBlock1 = prSc;
                                psTg.Add.Value(null/* Conversion error: Set to default value for this argument */, withBlock1.Name); Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing Point -- Ctrl-F8 to here
                            }
                        }

                        psSc = sd.PropertySets(gnDesign);
                        {
                            var withBlock1 = td.PropertySets(gnDesign);
                            foreach (var pn in Array(pnPartNum, pnStockNum, pnFamily, pnDesc, pnCatWebLink))
                            {
                                // .Item(pnStockNum).Value = psSc.Item(pnStockNum).Value
                                // .Item(pnFamily).Value = psSc.Item(pnFamily).Value
                                // .Item(pnCatWebLink).Value = psSc.Item(pnCatWebLink).Value
                                // .Item(pnDesc).Value = psSc.Item(pnDesc).Value
                                // .Item(pnPartNum).Value = psSc.Item(pnPartNum).Value
                                // .Item(pn).Value = psSc.Item(pn).Value
                                withBlock1.Item(System.Convert.ToHexString(pn)).Value = psSc.Item(System.Convert.ToHexString(pn)).Value;
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing Point -- Ctrl-F8 to here
                            }
                        }
                    }
                    else
                    {
                    }
                }
                else
                {
                }
            }
        }
    }

    public string d0g8f0()
    {
        long dx;
        string fn;

        for (dx = 1; dx <= 16; dx++)
        {
            fn = "Specification" + System.Convert.ToHexString(dx);
            {
                var withBlock = cnGnsDoyle().Execute(Join(Array("select distinct", fn, "from vgMfiItems", "where Family = 'D-BAR'", "and", fn, "is not null", "and", fn, "<> ''", "order by", fn, ";"), " "));
                if (withBlock.BOF | withBlock.EOF)
                {
                }
                else
                {
                    Debug.Print("[" + fn + "]");
                    Debug.Print.GetString();
                }
            }
        }
    }

    public string d0g9f0(Inventor.Document ad = null/* TODO Change to default(_) if this is not a reference type */, string pn = "")
    {
        Inventor.View vw;
        Inventor.Camera cm;
        string bp;

        if (ad == null)
            d0g9f0 = d0g9f0(ThisApplication.ActiveDocument, pn);
        else if (Strings.Len(pn) < 1)
            d0g9f0 = d0g9f0(ad, d0g9f3(ad));
        else
        {
            bp = @"C:\Doyle_Vault\Designs\Misc\andrewT\";
            vw = ad.Views.Item(1); // ThisApplication.ActiveView
            cm = vw.Camera;

            {
                var withBlock = vw;
                // Debug.Print .Left, .Top
                // Debug.Print .Width, .Height

                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // breakpoint anchor

                {
                    var withBlock1 = withBlock.Camera;
                    withBlock1.ViewOrientationType = kIsoTopRightViewOrientation;
                    withBlock1.Fit();
                    withBlock1.Apply();
                }
                withBlock.Fit();
                withBlock.Update();
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // breakpoint anchor
                                                                             // .SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                withBlock.SaveAsBitmap(bp + pn + "-I.png", withBlock.Width, withBlock.Height);
                {
                    var withBlock1 = withBlock.Camera;
                    withBlock1.ViewOrientationType = kFrontViewOrientation;
                    withBlock1.Fit();
                    withBlock1.Apply();
                }
                withBlock.Fit();
                withBlock.Update();
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // breakpoint anchor
                                                                             // .SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                withBlock.SaveAsBitmap(bp + pn + "-F.png", 0, 0);
                {
                    var withBlock1 = withBlock.Camera;
                    withBlock1.ViewOrientationType = kTopViewOrientation;
                    withBlock1.Fit();
                    withBlock1.Apply();
                }
                withBlock.Fit();
                withBlock.Update();
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // breakpoint anchor
                                                                             // .SaveAsBitmapWithOptions pn & "-I.png", 0, 0
                withBlock.SaveAsBitmap(bp + pn + "-T.png", 0, 0);
                withBlock.GoHome();
                withBlock.Update();
            }
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // breakpoint anchor

        d0g9f0 = "";
    }

    public string d0g9f1(Inventor.AssemblyDocument ad)
    {
        Inventor.iAssemblyTableRow rw;

        {
            var withBlock = ad.ComponentDefinition;
            if (withBlock.IsiAssemblyFactory)
            {
                {
                    var withBlock1 = withBlock.iAssemblyFactory;
                    foreach (var rw in withBlock1.TableRows)
                    {
                        {
                            var withBlock2 = rw;
                            Debug.Print.MemberName();
                        }
                    }
                }
            }
            else
            {
            }
        }

        d0g9f1 = "";
    }

    public string d0g9f2(Inventor.ComponentOccurrence oc)
    {
        {
            var withBlock = oc;
            if (withBlock.IsiAssemblyMember)
                System.Diagnostics.Debugger.Break();
            else if (withBlock.IsiPartMember)
                System.Diagnostics.Debugger.Break();
            else
                System.Diagnostics.Debugger.Break();
        }
    }

    public string d0g9f2as(Inventor.AssemblyComponentDefinition cd)
    {
        {
            var withBlock = cd;
            // .IsiAssemblyMember
            // .iAssemblyMember
            {
                var withBlock1 = withBlock.iAssemblyMember;
            }
        }
    }

    public string d0g9f3(Inventor.AssemblyDocument ad)
    {
        Inventor.AssemblyDocument cp;

        {
            var withBlock = ad.ComponentDefinition.Occurrences.Item(1);
            cp = aiDocAssy(withBlock.Definition.Document);
            if (cp == null)
                d0g9f3 = "NO-NUM-ASSY";
            else
                d0g9f3 = cp.PropertySets(gnDesign).Item(pnPartNum).Value;
        }
    }

    public void PlaceInAssembly()
    {
        Scripting.Dictionary dc;
        Inventor.CommandManager cm;
        Inventor.Document cd;
        Inventor.Document ad;
        VbMsgBoxResult rp;
        string nm;

        {
            var withBlock = ThisApplication;
            if (withBlock.ActiveDocumentType == kPartDocumentObjectOr.ActiveDocumentType == kAssemblyDocumentObjectThen)
            {
                cd = withBlock.ActiveDocument;
                dc = dcAiAssyDocs(dcAiDocsVisible());
                dc.Remove(cd.FullDocumentName);
                {
                    var withBlock1 = nuSelAiDoc().WithList(dc.Keys);
                    do
                    {
                        nm = withBlock1.GetReply();
                        if (dc.Exists(nm))
                        {
                            ad = dc.Item(nm);
                            rp = Constants.vbOK;
                        }
                        else
                        {
                            ad = null/* TODO Change to default(_) if this is not a reference type */;
                            rp = MsgBox("No Valid Assembly Selected.", Constants.vbRetryCancel, "No Assembly");
                        }
                    }
                    while (rp == Constants.vbRetry)// Try Again?
   ;
                }

                if (ad == null)
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
                else
                {
                    ad.Activate();
                    cm = withBlock.CommandManager;
                    {
                        var withBlock1 = cm;
                        withBlock1.PostPrivateEvent(kFileNameEvent, cd.FullDocumentName);
                        withBlock1.ControlDefinitions.Item("AssemblyPlaceComponentCmd").Execute();
                    }
                }
            }
        }
    }
}