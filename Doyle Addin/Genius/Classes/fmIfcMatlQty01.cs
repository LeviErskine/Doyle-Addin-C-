class fmIfcMatlQty01
{
    private void Class_Initialize()
    {
        // Dim ctl As MSForms.Control

        dcResult = new Scripting.Dictionary();
        {
            var withBlock = dcResult;
            withBlock.Add(pnRmQty, 0);
            withBlock.Add(pnRmUnit, "");
        }

        fm = new fmMatlQty();
        {
            var withBlock = fm;
            // For Each ctl In .Controls
            // Debug.Print ctl.Name
            // Next

            cbxUnitQty = withBlock.cbxUnitQty;
            cbxUnitQty.List = Split("IN FT FT2 IN2 EA");

            lbxMatlQty = withBlock.lbxMatlQty;
            txbMatlQty = withBlock.txbMatlQty;

            imgThmNail = withBlock.imThmNail;
            lblPartNumber = withBlock.lblPartNumber;
            lblPartInfo = withBlock.lblPartInfo;
            lblMatlNumber = withBlock.lblMatlNumber;
            lblMatlInfo = withBlock.lblMatlInfo;
        }
    }

    public Scripting.Dictionary Result()
    {
        Result = dcCopy(dcResult);
    }

    private Scripting.Dictionary Changes(Scripting.Dictionary wkg)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = wkg;
            foreach (var ky in dcResult.Keys)
            {
                if (withBlock.Exists(ky))
                {
                    if (withBlock.Item(ky) == dcResult.Item(ky))
                    {
                    }
                    else
                        rt.Add(ky, withBlock.Item(ky));
                }
                else
                {
                }
            }
        }

        Changes = dcCopy(dcResult);
    }

    private Scripting.Dictionary Commit(Scripting.Dictionary src)
    {
        Variant ky;

        {
            var withBlock = dcResult;
            foreach (var ky in withBlock.Keys)
            {
                if (src.Exists(ky))
                    withBlock.Item(ky) = src.Item(ky);
            }
        }

        Commit = dcCopy(dcResult);
    }

    public Scripting.Dictionary SeeUser(object About = null) // fmIfcMatlQty01
    {
        string ky;
        string ck;

        if (About == null)
            SeeUser = SeeUser(nuDcPopulator().Setting(pnRmQty + "()", nuDcPopulator().Setting(4, 1).Setting(2, 1).Setting(24, 1).Dictionary()).Setting(pnRmQty, 24).Setting(pnRmUnit, "IN").Setting(pnPartNum, "NO-ITM-GIVEN").Setting(pnRawMaterial, "NO-MTL-GIVEN").Dictionary());
        else if (About is Scripting.Dictionary)
            SeeUser = SeeUserWithDict(About);
        else if (About is Inventor.PartDocument)
            SeeUser = SeeUserWithPart(About);
        else if (About is Inventor.Property)
            SeeUser = SeeUserWithQtyProp(About);
        else
            SeeUser = SeeUser();
    }

    /// make this one Public later

    /// once Part version is working
    private fmIfcMatlQty01 SeeUserWithModel(Inventor.Property About)
    {
    }

    public Scripting.Dictionary SeeUserWithPart(Inventor.PartDocument About) // fmIfcMatlQty01
    {
        if (About == null)
            SeeUserWithPart = SeeUser(About);
        else
        {
            Scripting.Dictionary dcPr;
            Inventor.Property obPr;
            Variant kyPr;
            long op;

            dcPr = new Scripting.Dictionary();
            {
                var withBlock = About.PropertySets;
                {
                    var withBlock1 = withBlock.Item(gnDesign);
                    dcPr.Add(pnPartNum, withBlock1.Item(pnPartNum).Value);
                    dcPr.Add(pnDesc, withBlock1.Item(pnDesc).Value);
                }


                {
                    var withBlock1 = withBlock.Item(gnCustom);
                    foreach (var kyPr in Array(pnRawMaterial, pnRmQty, pnRmUnit))
                    {
                        Information.Err.Clear();
                        obPr = withBlock1.Item(System.Convert.ToHexString(kyPr));
                        if (Information.Err.Number == 0)
                            dcPr.Add(kyPr, obPr.Value);
                        else
                        {
                            Debug.Print(Information.Err.Description);
                            System.Diagnostics.Debugger.Break();
                            Information.Err.Clear();
                        }
                    }
                }
            }

            /// prepare Dictionary of Dimensions
            /// with Count of Each
            /// 
            Scripting.Dictionary dcDm;
            Variant vlDm;
            long ctDm;

            dcDm = new Scripting.Dictionary();

            {
                var withBlock = nuAiBoxData().UsingInches(1);
                for (var op = 0; op <= 1; op++)
                {
                    {
                        var withBlock1 = withBlock.UsingModel(About, op);
                        foreach (var vlDm in Array(Round(withBlock1.SpanX, 4), Round(withBlock1.SpanY, 4), Round(withBlock1.SpanZ, 4), 0))
                        {
                            {
                                var withBlock2 = dcDm;
                                if (vlDm > 0)
                                {
                                    if (withBlock2.Exists(vlDm))
                                    {
                                        ctDm = withBlock2.Item(vlDm) + 1;
                                        withBlock2.Item(vlDm) = ctDm;
                                    }
                                    else
                                        withBlock2.Add(vlDm, ctDm);
                                }
                            }
                        }
                    }
                }
            }

            {
                var withBlock = dcPr;
                withBlock.Add(pnRmQty + "()", dcDm);
                withBlock.Add("img", About.Thumbnail);
            }

            SeeUserWithPart = SeeUserWithDict(dcPr);
        }
    }

    private Scripting.Dictionary SeeUserWithQtyProp(Inventor.Property About) // fmIfcMatlQty01
    {
        /// this one will have to be heavily modified
        /// likely dumping a bunch of code now implemented
        /// in SeeUserWithPart, which can simply be
        /// called with the Document containing
        /// the supplied Property
        /// 

        if (About == null)
            System.Diagnostics.Debugger.Break();
        else
        {
            /// these variables are for use
            /// in separating quantity from
            /// unit of measure in Value of
            /// supplied Property
            string vlIn;
            Variant arIn;
            double qtIn;
            string unIn;
            /// split incoming Property Value into
            /// Quantity and Unit of Measurement
            vlIn = System.Convert.ToHexString(About.Value) + " ";
            /// note: concatenated space at end
            /// of Value text should ensure two
            /// members of arIn, as follows
            arIn = Split(vlIn, " ", 2);

            qtIn = Round(Val(arIn(0)), 4);
            if (UBound(arIn) > 0)
                unIn = Trim(arIn(1));
            /// this section and its associated variables
            /// will likely be exported to a separate function

            /// force blank Unit of
            /// Measure to default inches
            if (Strings.Len(unIn) == 0)
                unIn = "IN";

            /// the following section SHOULD be
            /// implemented now in SeeUserWithPart
            /// it should be possible to simply
            /// call that function, completely
            /// ignoring the supplied Property
            /// 
            /// prepare Dictionary of Dimensions
            /// with Count of Each
            /// 
            Scripting.Dictionary dcDm;
            Variant vlDm;
            long ctDm;

            dcDm = new Scripting.Dictionary();
            if (qtIn > 0)
                dcDm.Add(qtIn, 1);

            /// get all necessary information
            /// from Inventor Model
            /// 
            Inventor.Document md;
            Inventor.Property mdPt;
            Inventor.Property mdMt;

            md = aiDocument(About.Parent.Parent.Parent);
            {
                var withBlock = md.PropertySets;
                mdPt = withBlock.Item(gnDesign).Item(pnPartNum);

                Information.Err.Clear();
                mdMt = withBlock.Item(gnCustom).Item(pnRawMaterial);
                if (Information.Err.Number == 0)
                {
                }
                else
                    System.Diagnostics.Debugger.Break();
            }

            {
                var withBlock = nuAiBoxData().UsingInches(1).UsingModel(About);
                foreach (var vlDm in Array(Round(withBlock.SpanX, 4), Round(withBlock.SpanY, 4), Round(withBlock.SpanZ, 4), 0))
                {
                    {
                        var withBlock1 = dcDm;
                        if (vlDm > 0)
                        {
                            if (withBlock1.Exists(vlDm))
                            {
                                ctDm = withBlock1.Item(vlDm) + 1;
                                withBlock1.Item(vlDm) = ctDm;
                            }
                            else
                                withBlock1.Add(vlDm, ctDm);
                        }
                    }
                }
            }

            {
                var withBlock = nuDcPopulator().Setting(pnRmQty + "()", dcDm).Setting(pnRmQty, qtIn).Setting(pnRmUnit, unIn).Setting(pnPartNum, mdPt.Value).Setting(pnRawMaterial, mdMt.Value);
                SeeUserWithQtyProp = SeeUserWithDict(withBlock.Dictionary());
            }
        }
    }

    public Scripting.Dictionary SeeUserWithDict(Scripting.Dictionary About) // fmIfcMatlQty01
    {
        string ky;
        string ck;

        if (About == null)
            SeeUserWithDict = SeeUserWithDict(nuDcPopulator().Setting(pnRmQty + "()", nuDcPopulator().Setting(4, 1).Setting(2, 1).Setting(24, 1).Dictionary()).Setting(pnRmQty, 24).Setting(pnRmUnit, "IN").Setting(pnPartNum, "NO-ITM-GIVEN").Setting(pnRawMaterial, "NO-MTL-GIVEN").Dictionary());
        else
        {
            {
                var withBlock = About;
                // .Add "img", About.Thumbnail
                if (withBlock.Exists("img"))
                    imgThmNail.Picture = withBlock.Item("img");

                if (withBlock.Exists(pnDesc))
                    txbMatlQty.Value = Val(System.Convert.ToHexString(withBlock.Item(pnDesc)));

                ky = pnRmQty + "()";
                if (withBlock.Exists(ky))
                    lbxMatlQty.List = dcOb(withBlock.Item(ky)).Keys;

                if (withBlock.Exists(pnRmQty))
                    txbMatlQty.Value = Val(System.Convert.ToHexString(withBlock.Item(pnRmQty)));

                if (withBlock.Exists(pnRmUnit))
                {
                    Information.Err.Clear();
                    cbxUnitQty.Value = withBlock.Item(pnRmUnit);
                    if (Information.Err.Number)
                    {
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        cbxUnitQty.Value = "IN";
                    }
                }

                /// Following are "boilerplate" elements
                /// for Part/Item and Raw Material numbers,
                /// along with their descriptions.
                /// 
                /// A thumbnail image of the Part is also
                /// expected to be supplied at some point,
                /// but will be held off for now, pending
                /// successful testing of the form's main
                /// functions.
                /// 
                /// Part/Item Number
                if (withBlock.Exists(pnPartNum))
                    lblPartNumber.Caption = System.Convert.ToHexString(withBlock.Item(pnPartNum));

                /// Material Number
                if (withBlock.Exists(pnRawMaterial))
                    lblMatlNumber.Caption = System.Convert.ToHexString(withBlock.Item(pnRawMaterial));

                /// Item Description
                if (withBlock.Exists(pnDesc))
                    lblPartInfo.Caption = System.Convert.ToHexString(withBlock.Item(pnDesc));

                /// Material Description
                /// (not expected at this time)
                ky = pnRawMaterial + ":";
                if (withBlock.Exists(ky))
                    lblMatlInfo.Caption = System.Convert.ToHexString(withBlock.Item(ky));
            }

            {
                var withBlock = Commit(About);
            }

            fm.Show(1);
            // Stop

            {
                var withBlock = nuDcPopulator().Setting(pnRmQty, Round(Val(txbMatlQty.Value), 4)).Setting(pnRmUnit, cbxUnitQty.Value) // Mapping...
       ;
                // txbMatlQty -> pnRmQty
                // cbxUnitQty -> pnRmUnit

                SeeUserWithDict = Commit(withBlock.Dictionary);
            }
        }
    }

    public void Version()
    {
        Version = fmVersion;
    }

    private void Class_Terminate()
    {
        // dcWorkg = Nothing
        // dcGiven = Nothing

        imgThmNail.Picture = null;
        imgThmNail = null;

        cbxUnitQty = null;

        lbxMatlQty = null;
        txbMatlQty = null;

        lblPartNumber = null;
        lblPartInfo = null;
        lblMatlNumber = null;
        lblMatlInfo = null;

        fm = null;
    }

    private void fm_Sent(VbMsgBoxResult Signal)
    {
        VbMsgBoxResult ck;

        if (Signal == Constants.vbCancel)
        {
            ck = MsgBox(Join(Array("Material Quantity", "and Units will", "remain unchanged."), Constants.vbNewLine), Constants.vbYesNo, "Cancel Update?");
            // Stop

            if (ck == Constants.vbYes)
            {
                {
                    var withBlock = dcResult;
                    txbMatlQty.Value = System.Convert.ToHexString(withBlock.Item(pnRmQty));
                    cbxUnitQty.Value = withBlock.Item(pnRmUnit);
                }
                fm.Hide();
            }
            else if (ck == Constants.vbCancel)
                System.Diagnostics.Debugger.Break();// drop and debug
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
        else if (Signal == Constants.vbOK)
        {
            ck = MsgBox(Join(Array("Update Material", "Quantity to " + System.Convert.ToHexString(Round(Val(txbMatlQty.Value), 4)) + cbxUnitQty.Value + "?"), Constants.vbNewLine), Constants.vbYesNo, "Update Quantity?");
            // Stop

            if (ck == Constants.vbYes)
            {
                fm.Hide();
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }
            else if (ck == Constants.vbCancel)
                System.Diagnostics.Debugger.Break();// drop and debug
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
        else
            System.Diagnostics.Debugger.Break();
    }

    private void lbxMatlQty_DblClick(MSForms.ReturnBoolean Cancel)
    {
        txbMatlQty.Value = lbxMatlQty.Value;
    }

    private void lbxMatlQty_MouseMove(int Button, int Shift, float X, float Y)
    {
        MSForms.DataObject dt;
        int ef;

        if (Button == 1)
        {
            dt = new MSForms.DataObject();
            dt.SetText(lbxMatlQty.Value);
            ef = dt.StartDrag();
        }
    }

    private void txbMatlQty_Change()
    {
        double ck;
        string tx;
        Variant gp;
        long mx;
        long dx;

        {
            var withBlock = txbMatlQty;
            tx = withBlock.Value;
            gp = Split(tx, ".");
            mx = UBound(gp);

            for (dx = LBound(gp); dx <= mx; dx++)
            {
                ck = Val(gp(dx));

                if (ck > 0)
                    gp(dx) = System.Convert.ToHexString(ck);
                else if (dx > 0)
                    gp(dx) = "";
                else
                    gp(dx) = "0";
            }
            tx = Join(gp, ".");

            if (tx != withBlock.Value)
            {
                DoEvents();
                withBlock.Value = tx;
            }
        }
    }
}