using Microsoft.VisualBasic;

class fmTest1 : Form
{
    private var VB_Name = "fmTest1";
    private var VB_GlobalNameSpace = false;
    private var VB_Creatable = false;
    private var VB_PredeclaredId = true;
    private var VB_Exposed = false;

    private ADODB.Connection cn;
    private ADODB.Recordset rsFam;
    private ADODB.Recordset rsPrt;
    private ADODB.Recordset rsItm;

    private Dictionary dc;
    private Document ad;

    private PropertySet psDsn;
    private PropertySet psUsr;

    private Dictionary dcDsn;
    private Dictionary dcUsr;

    private Property prFam;
    private Property prStk;
    private Property prThk;

    public VbMsgBoxResult AskAbout(Document AiDoc, string txMsg = "")
    {
        stdole.IPictureDisp pc;
        VbMsgBoxResult ck;
        string pn; // part number
        string sn; // material (stock) number
        string pd; // part description

        ad = AiDoc;
        {
            var withBlock = ad;

            Information.Err().Clear();
            pc = withBlock.Thumbnail;
            if (Information.Err().Number == 0)
            {
            }

            psDsn = withBlock.PropertySets(gnDesign);
            psUsr = withBlock.PropertySets(gnCustom);

            dcDsn = dcAiPropsInSet(psDsn);
            dcUsr = dcAiPropsInSet(psUsr);

            prFam = psDsn.get_Item(pnFamily);

            // ' Get Sheet Metal Thickness Property
            prThk = aiPropShtMetalThickness(ad);
            // ' NOTE: Function returns Nothing
            // ' if Part is NOT Sheet Metal!

            {
                var withBlock1 = dcUsr;
                if (withBlock1.Exists(pnRawMaterial))
                    prStk = psUsr.get_Item(pnRawMaterial);
                else
                {
                    Information.Err().Clear();
                    prStk = psUsr.Add("", pnRawMaterial);
                    if (Information.Err().Number)
                        Debugger.Break();
                    else
                        withBlock1.Add(pnRawMaterial, prStk);
                }
            }

            // REV[2022.04.28.1615]
            // added initializtion of Dictionary dc
            // with initial raw material setting.
            // sn now assigned from the Dictionary.
            // NOTE: probably want to initial
            // values in a separate "recovery"
            // Dictionary to be restored if
            // the User chooses to cancel.
            // Also, see function/method dcUpd.
            // looks like it gets called when
            // something changes. Easy to miss!
            dc.get_Item(pnRawMaterial) = prStk.Value;
            sn = dc.get_Item(pnRawMaterial);
            pn = psDsn.get_Item(pnPartNum).Value;
            pd = psDsn.get_Item(pnDesc).Value;
        }

        {
            var withBlock = this;
            withBlock.Caption = "Please Review Part Number: " + pn;

            if (pc == null)
            {
            }
            else
                withBlock.imThmNail.Picture = pc;

            {
                var withBlock1 = withBlock.lbMsg;
                withBlock1.Caption = pn + ": " + pd
                                     + Constants.vbCrLf + txMsg + Interaction.IIf(Strings.Len(txMsg) > 0,
                                         Constants.vbCrLf, "")
                                     + ft1g0f0(pnCatWebLink, psDsn.get_Item(pnCatWebLink)) + Constants.vbCrLf
                                     + ft1g0f0(pnMaterial, psDsn.get_Item(pnMaterial)) + Constants.vbCrLf
                                     + ft1g0f0(pnThickness, prThk) + Constants.vbCrLf
                                     + "";

                //& vbCrLf                 & vbCrLf & pnThickness & ": " & psUsr.get_Item(pnThickness).Value 
            }
            float df = mdl1g1f2(withBlock.lbMsg);
            if (df > 0)
            {
                mdl1g1f3.lbMtFamily(null /* Conversion error: Set to default value for this argument */, 0, df);
                mdl1g1f3.lbxFamily(null /* Conversion error: Set to default value for this argument */, 0, df);
            }

            withBlock.dbFamily.Value = prFam.Value;

            if (Strings.Len(sn) > 0)
            {
                string sf; // material (stock) family
                {
                    var withBlock1 = cn.Execute("select Family from vgMfiItems where Item = '"
                                                + Replace(sn, "'", "''") + "'"
                    );
                    // REV[2022.08.19.1359]
                    // temporarily replacing direct use of sn
                    // with call to Replace single quotes
                    // in string with doubled single quotes
                    // (NOT double quotes!) to "escape" the
                    // character in a string value.
                    // '
                    // will ultimately want to produce some
                    // sort of 'handler' to preprocess values
                    // for use in SQL commands to avoid errors
                    // that arise from this sort of thing.
                    if (withBlock1.BOF | withBlock1.EOF)
                        sf = "";
                    else
                        sf = withBlock1.Fields(0).Value;
                }

                if (Strings.Len(sf) == 0)
                    // EITHER doesn't have a Family,
                    // OR is not (yet) in Genius.
                    // SO, let's just ...
                    sf = "DSHEET"; // as a default!

                Information.Err().Clear();
                withBlock.lbxFamily.Value = sf;
                if (Information.Err().Number)
                {
                    Debug.Print("FAILED TO MATERIAL FAMILY " + sf);
                    ck = MessageBox.Show(Join(new[]
                        {
                            "Part Number " + pn, "uses Material " + sn, "which is a" + IIf(
                                InStr(1, "AEIOU", UCase(Left(sf, 1))), "n ", " "
                            ) + sf + " Item.",
                            "", "This interface does not presently", "support Materials from this Family.", "",
                            "You might not be able to find the correct", "Material for this Part, and might wish",
                            "to avoid changing it here.", "", "Do you wish to proceed anyway?"
                        }, Constants.vbCrLf),
                        Constants.vbYesNoCancel + Constants.vbExclamation + Constants.vbDefaultButton2,
                        "Material Family not Supported"
                    );
                    if (ck == Constants.vbCancel)
                        Debugger.Break();
                }
                else
                {
                    Information.Err().Clear();
                    withBlock.lbxItem.Value = sn;

                    // REV[2022.05.06.1329]
                    // added intermediate error handler
                    // to capture failure in Material
                    // Family selector to adopt new Value.
                    // it re-implements process of Event
                    // handler Sub lbxFamily_Change
                    // against variable 'sf' directly
                    // in an effort to force population
                    // of Material list.
                    if (Information.Err().Number)
                    {
                        Debug.Print(""); // Breakpoint Landing
                        Information.Err().Clear();
                        rsItm.Filter = "Family = '" + sf + "'";
                        withBlock.lbxItem.List = m0g3f1(rsItm);
                        withBlock.lbxItem.Value = sn;
                    }
                    // something MIGHT have happened
                    // to prevent normal Value update
                    // when lbxFamily is above.
                    // further investigation may be
                    // warranted.

                    if (Information.Err().Number)
                    {
                        Debug.Print("FAILED TO MATERIAL " + sn);
                        ck = MessageBox.Show(Join(new[]
                            {
                                "!!WARNING!!", "", "Active Material " + sn, "for Part Number " + pn,
                                "could NOT be selected,", "and might be unavailable.", "",
                                "You might wish to avoid",
                                "making Material changes", "to this Part here.", "",
                                "Do you wish to proceed anyway?"
                            }, Constants.vbCrLf), Constants.vbYesNoCancel + Constants.vbExclamation,
                            "Active Material Not Found!"
                        );
                        if (ck == Constants.vbCancel)
                            Debugger.Break();
                    }
                    else
                    {
                        ck = Constants.vbYes;
                        lbxItem_Change();
                        // lbxFamily_Change
                        rsItm.Filter = "Family = '" + sf + "'";
                        withBlock.lbxItem.List = m0g3f1(rsItm);
                    }
                }
            }
            else
                ck = Constants.vbYes;

            if (ck == Constants.vbYes)
                withBlock.Show(1);
        }

        return ck; // vbYes ' = 1
    }

    private string ft1g0f0(string pn, Property pr
    )
    {
        if (pr == null)
            return "";
        return Constants.vbCrLf + pn + ": " + pr.Value;
    }

    private void dbFamily_Change()
    {
        Debug.Print(dcUpd(pnFamily, dbFamily.Value));
    }
    // Me.lbxItem.ColumnWidths = "84 pt;6 pt;180 pt"
    // Me.lbxItem.ColumnWidths = "84 pt;48 pt;216 pt"

    private void lbMsg_DblClick(MSForms.ReturnBoolean Cancel)
    {
        Debugger.Break();
    }

    private void lbxFamily_Change()
    {
        {
            var withBlock = this;
            rsItm.Filter = "Family = '" + withBlock.lbxFamily.Value + "'";
            withBlock.lbxItem.List = m0g3f1(rsItm);
        }
    }

    public Dictionary ItemData()
    {
        var rt = new Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, withBlock.get_Item(ky));
        }
        return rt;
    }

    public Dictionary Synch()
    {
        {
            var withBlock = dc;
            if (withBlock.Exists(pnFamily))
                prFam.Value = dc.get_Item(pnFamily);
            if (withBlock.Exists(pnRawMaterial))
                prStk.Value = dc.get_Item(pnRawMaterial);
        }

        return ItemData();
    }

    private string dcUpd(string ky, dynamic vl)
    {
        if (IsNull(vl))
            return dcUpd(ky, "");
        var withBlock = dc;
        if (withBlock.Exists(ky))
        {
            string rt = Convert.ToHexString(withBlock.get_Item(ky));
            withBlock.get_Item(ky) = vl;
            return "CHANGE[" + ky + "] FROM '" + rt
                   + "' TO '" + Convert.ToHexString(withBlock.get_Item(ky)) + "'";
        }

        withBlock.Add(ky, vl);
        return "[" + ky + "] TO '"
               + Convert.ToHexString(withBlock.get_Item(ky)) + "'";
    }

    private void lbxItem_Change()
    {
        Debug.Print(dcUpd(pnRawMaterial, lbxItem.Value));
    }

    private void UserForm_Initialize()
    {
        dc = new Dictionary();
        cn = cnGnsDoyle();

        {
            var withBlock = cn;

            // rsFam = .Execute(Join(new [] { "select Family, Description1", "from vgMfiFamilies", "order by Family" ), " ")) ', "where FamilyGroup = 'RAW'"rsPrt = withBlock.Execute(Join(new [] {"select Family, FamilyGroup, Description1", "from vgMfiFamilies", "order by Family"), Constants.vbCrLf)); // , ;/* Cannot convert EmptyStatementSyntax, CONVERSION ERROR: Conversion for EmptyStatement not implemented, please report this issue in '' at character 11306

            "where FamilyGroup = 'PARTS'"
            rsItm = withBlock.Execute(Join(new[]
            {
                "Select I.Item, I.Family, I.Description1, I.Specification1",
                "From vgMfiItems as I", "Inner Join vgMfiFamilies as F", "On I.Family = F.Family",
                "Where F.FamilyGroup = 'RAW'", "order by Family, Item"
            }, " "));
        }

        {
            var withBlock = this;
            rsPrt.Filter = "FamilyGroup = 'RAW'";
            withBlock.lbxFamily.List = m0g3f1(rsPrt); // rsFam

            rsPrt.Filter = "FamilyGroup = 'PARTS'";
            withBlock.dbFamily.List = m0g3f1(rsPrt);
        }
    }

    private void UserForm_QueryClose(int Cancel, int CloseMode
    )
    {
        Cancel = 1;
        this.Hide();
    }

    private void UserForm_Terminate()
    {
        cn.Close();
        cn = null;
    }
}