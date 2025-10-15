using Doyle_Addin.Genius.Forms;
using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class fmIfcMatlQty01 : Form
{
    private void Class_Initialize()
    {
        // Dim ctl As MSForms.Control

        dcResult = new Dictionary();
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

    public Dictionary Result()
    {
        return dcCopy(dcResult);
    }

    private Dictionary Changes(Dictionary wkg)
    {

        var rt = new Dictionary();
        {
            foreach (var ky in dcResult.Keys)
            {
                if (!wkg.Exists(ky)) continue;
                if (wkg.get_Item(ky) == dcResult.get_Item(ky))
                {
                }
                else
                    rt.Add(ky, wkg.get_Item(ky));
            }
        }

        return dcCopy(dcResult);
    }

    private Dictionary Commit(Dictionary src)
    {

        {
            var withBlock = dcResult;
            foreach (var ky in withBlock.Keys)
            {
                if (src.Exists(ky))
                    withBlock.get_Item(ky) = src.get_Item(ky);
            }
        }

        return dcCopy(dcResult);
    }

    public Dictionary SeeUser(dynamic About = null) // fmIfcMatlQty01
    {
        while (true)
        {
            string ky;
            string ck;

            if (About != null)
                return About switch
                {
                    Dictionary => SeeUserWithDict(About),
                    PartDocument => SeeUserWithPart(About),
                    Property => SeeUserWithQtyProp(About),
                    _ => SeeUser()
                };
            About = nuDcPopulator()
                .Setting(pnRmQty + "()", nuDcPopulator().Setting(4, 1).Setting(2, 1).Setting(24, 1).Dictionary())
                .Setting(pnRmQty, 24)
                .Setting(pnRmUnit, "IN")
                .Setting(pnPartNum, "NO-ITM-GIVEN")
                .Setting(pnRawMaterial, "NO-MTL-GIVEN")
                .Dictionary();
            continue;

            break;
        }
    }

    // make this one Public later

    // once Part version is working
    private fmIfcMatlQty01 SeeUserWithModel(Property About)
    {
    }

    public Dictionary SeeUserWithPart(PartDocument About) // fmIfcMatlQty01
    {
        if (About == null)
            return SeeUser(About);

        var dcPr = new Dictionary();
        {
            var withBlock = About.PropertySets;
            {
                var withBlock1 = withBlock.get_Item(gnDesign);
                dcPr.Add(pnPartNum, withBlock1.get_Item(pnPartNum).Value);
                dcPr.Add(pnDesc, withBlock1.get_Item(pnDesc).Value);
            }

            {
                var withBlock1 = withBlock.get_Item(gnCustom);
                foreach (var kyPr in new[]
                         {
                             pnRawMaterial, pnRmQty, pnRmUnit
                         })
                {
                    Information.Err().Clear();
                    var obPr = withBlock1.get_Item(Convert.ToHexString(kyPr));
                    if (Information.Err().Number == 0)
                        dcPr.Add(kyPr, obPr.Value);
                    else
                    {
                        Debug.Print(Information.Err().Description);
                        Debugger.Break();
                        Information.Err().Clear();
                    }
                }
            }
        }

        // prepare Dictionary of Dimensions
        // with Count of Each
        // 

        var dcDm = new Dictionary();

        {
            var withBlock = nuAiBoxData().UsingInches(1);
            for (long op = 0; op <= 1; op++)
            {
                {
                    var withBlock1 = withBlock.UsingModel(About, op);
                    foreach (var vlDm in new[]
                             {
                                 Round(withBlock1.SpanX, 4), Round(withBlock1.SpanY, 4),
                                 Round(withBlock1.SpanZ, 4), 0
                             })
                    {
                        {
                            var withBlock2 = dcDm;
                            if (vlDm <= 0) continue;
                            long ctDm;
                            if (withBlock2.Exists(vlDm))
                            {
                                ctDm = withBlock2.get_Item(vlDm) + 1;
                                withBlock2.get_Item(vlDm) = ctDm;
                            }
                            else
                                withBlock2.Add(vlDm, ctDm);
                        }
                    }
                }
            }
        }

        {
            dcPr.Add(pnRmQty + "()", dcDm);
            dcPr.Add("img", About.Thumbnail);
        }

        return SeeUserWithDict(dcPr);
    }

            private Dictionary SeeUserWithQtyProp(Property About) // fmIfcMatlQty01
            {
                // this one will have to be heavily modified
                // likely dumping a bunch of code now implemented
                // in SeeUserWithPart, which can simply be
                // called with the Document containing
                // the supplied Property
                // 

                if (About == null)
                    Debugger.Break();
                else
                {
                    // these variables are for use
                    // in separating quantity from
                    // unit of measure in Value of
                    // supplied Property
                    string unIn;
                    // split incoming Property Value into
                    // Quantity and Unit of Measurement
                    var vlIn = Convert.ToHexString(About.Value) + " ";
                    // note: concatenated space at end
                    // of Value text should ensure two
                    // members of arIn, as follows
                    dynamic arIn = Split(vlIn, " ", 2);

                    double qtIn = Round(Val(arIn(0)), 4);
                    if (UBound(arIn) > 0)
                        unIn = Trim(arIn(1));
                    // this section and its associated variables
                    // will likely be exported to a separate function

                    // force blank Unit of
                    // Measure to default inches
                    if (Strings.Len(unIn) == 0)
                        unIn = "IN";

                    // the following section SHOULD be
                    // implemented now in SeeUserWithPart
                    // it should be possible to simply
                    // call that function, completely
                    // ignoring the supplied Property
                    // 
                    // prepare Dictionary of Dimensions
                    // with Count of Each
                    // 
                    long ctDm;

                    var dcDm = new Dictionary();
                    if (qtIn > 0)
                        dcDm.Add(qtIn, 1);

                    // get all necessary information
                    // from Inventor Model
                    // 
                    Property mdPt;
                    Property mdMt;

                    var md = aiDocument(About.Parent.Parent.Parent);
                    {
                        var withBlock = md.PropertySets;
                        mdPt = withBlock.get_Item(gnDesign).get_Item(pnPartNum);

                        Information.Err().Clear();
                        mdMt = withBlock.get_Item(gnCustom).get_Item(pnRawMaterial);
                        if (Information.Err().Number == 0)
                        {
                        }
                        else
                            Debugger.Break();
                    }

                    {
                        var withBlock = nuAiBoxData().UsingInches(1).UsingModel(About);
                        foreach (var vlDm in new[]
                                 {
                                     Round(withBlock.SpanX, 4), Round(withBlock.SpanY, 4), Round(withBlock.SpanZ, 4),
                                     0
                                 })
                        {
                            {
                                if (vlDm <= 0) continue;
                                if (dcDm.Exists(vlDm))
                                {
                                    ctDm = dcDm.get_Item(vlDm) + 1;
                                    dcDm.get_Item(vlDm) = ctDm;
                                }
                                else
                                    dcDm.Add(vlDm, ctDm);
                            }
                        }
                    }

                    {
                        var withBlock = nuDcPopulator().Setting(pnRmQty + "()", dcDm).Setting(pnRmQty, qtIn)
                            .Setting(pnRmUnit, unIn).Setting(pnPartNum, mdPt.Value)
                            .Setting(pnRawMaterial, mdMt.Value);
                        return SeeUserWithDict(withBlock.Dictionary());
                    }
                }
            }

            public Dictionary SeeUserWithDict(Dictionary About) // fmIfcMatlQty01
            {
                while (true)
                {
                    string ck;

                    if (About == null)
                    {
                        About = nuDcPopulator()
                            .Setting(pnRmQty + "()", nuDcPopulator().Setting(4, 1).Setting(2, 1).Setting(24, 1).Dictionary())
                            .Setting(pnRmQty, 24)
                            .Setting(pnRmUnit, "IN")
                            .Setting(pnPartNum, "NO-ITM-GIVEN")
                            .Setting(pnRawMaterial, "NO-MTL-GIVEN")
                            .Dictionary();
                        continue;
                    }

                    {
                        // .Add "img", About.Thumbnail
                        if (About.Exists("img")) imgThmNail.Picture = About.get_Item("img");

                        if (About.Exists(pnDesc))
                            txbMatlQty.Value = Val(Convert.ToHexString(About.get_Item(pnDesc)));

                        var ky = pnRmQty + "()";
                        if (About.Exists(ky)) lbxMatlQty.List = dcOb(About.get_Item(ky)).Keys;

                        if (About.Exists(pnRmQty))
                            txbMatlQty.Value = Val(Convert.ToHexString(About.get_Item(pnRmQty)));

                        if (About.Exists(pnRmUnit))
                        {
                            Information.Err().Clear();
                            cbxUnitQty.Value = About.get_Item(pnRmUnit);
                            if (Information.Err().Number)
                            {
                                Debug.Print(""); // Breakpoint Landing
                                cbxUnitQty.Value = "IN";
                            }
                        }

                        // Following are "boilerplate" elements
                        // for Part/Item and Raw Material numbers,
                        // along with their descriptions.
                        // 
                        // A thumbnail image of the Part is also
                        // expected to be supplied at some point,
                        // but will be held off for now, pending
                        // successful testing of the form's main
                        // functions.
                        // 
                        // Part/Item Number
                        if (About.Exists(pnPartNum))
                            lblPartNumber.Caption = Convert.ToHexString(About.get_Item(pnPartNum));

                        // Material Number
                        if (About.Exists(pnRawMaterial))
                            lblMatlNumber.Caption = Convert.ToHexString(About.get_Item(pnRawMaterial));

                        // Item Description
                        if (About.Exists(pnDesc)) lblPartInfo.Caption = Convert.ToHexString(About.get_Item(pnDesc));

                        // Material Description
                        // (not expected at this time)
                        ky = pnRawMaterial + ":";
                        if (About.Exists(ky)) lblMatlInfo.Caption = Convert.ToHexString(About.get_Item(ky));
                    }

                    {
                        var withBlock = Commit(About);
                    }

                    fm.Show(1);
                    // Stop

                    {
                        var withBlock = nuDcPopulator()
                                .Setting(pnRmQty, Round(Val(txbMatlQty.Value), 4))
                                .Setting(pnRmUnit, cbxUnitQty.Value) // Mapping...
                            ;
                        // txbMatlQty -> pnRmQty
                        // cbxUnitQty -> pnRmUnit

                        return Commit(withBlock.Dictionary);
                    }

                    break;
                }
            }

            public void Version()
            {
                return fmVersion;
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
                if (Signal == Constants.vbCancel)
                {
                    VbMsgBoxResult ck = MessageBox.Show(Join(new[]
                        {
                            "Material Quantity", "and Units will", "remain unchanged."
                        }, Constants.vbCrLf),
                        Constants.vbYesNo, "Cancel Update?");
                    // Stop
                    if (ck == Constants.vbYes)
                    {
                        {
                            var withBlock = dcResult;
                            txbMatlQty.Value = Convert.ToHexString(withBlock.get_Item(pnRmQty));
                            cbxUnitQty.Value = withBlock.get_Item(pnRmUnit);
                        }
                        fm.Hide();
                    }
                    else if (ck == Constants.vbCancel)
                        Debugger.Break(); // drop and debug

                    Debug.Print(""); // Breakpoint Landing
                }

                else if (Signal == Constants.vbOK)
                {
                    ck = MessageBox.Show(Join(new[]
                    {
                        "Update Material",
                        "Quantity to " + Convert.ToHexString(Round(Val(txbMatlQty.Value), 4)) +
                        cbxUnitQty.Value + "?"
                    }, Constants.vbCrLf), Constants.vbYesNo, "Update Quantity?");
                    // Stop
                    if (ck == Constants.vbYes)
                    {
                        fm.Hide();
                        Debug.Print(""); // Breakpoint Landing
                    }
                    else if (ck == Constants.vbCancel)
                        Debugger.Break(); // drop and debug

                    Debug.Print(""); // Breakpoint Landing
                }
                else
                    Debugger.Break();
            }

            private void lbxMatlQty_DblClick(MSForms.ReturnBoolean Cancel)
            {
                txbMatlQty.Value = lbxMatlQty.Value;
            }

            private void lbxMatlQty_MouseMove(int Button, int Shift, float X, float Y)
            {
                if (Button != 1) return;
                var dt = new DataObject();
                dt.SetText(lbxMatlQty.Value);
                int ef = dt.StartDrag();
            }

            private void txbMatlQty_Change()
            {
                {
                    var withBlock = txbMatlQty;
                    string tx = withBlock.Value;
                    dynamic gp = Split(tx, ".");
                    long mx = UBound(gp);

                    for (long dx = LBound(gp); dx <= mx; dx++)
                    {
                        double ck = Val(gp(dx));

                        if (ck > 0)
                            gp(dx) = Convert.ToHexString(ck);
                        else if (dx > 0)
                            gp(dx) = "";
                        else
                            gp(dx) = "0";
                    }

                    tx = Join(gp, ".");

                    if (tx == withBlock.Value) return;
                    DoEvents();
                    withBlock.Value = tx;
                }
            }

            /// <summary>
            /// Required method for Designer support - do not modify
            /// the contents of this method with the code editor.
            /// </summary>
            private void InitializeComponent()
            {
                SuspendLayout();
                // 
                // fmIfcMatlQty01
                // 
                ClientSize = new System.Drawing.Size(284, 261);
                ResumeLayout(false);
            }
}