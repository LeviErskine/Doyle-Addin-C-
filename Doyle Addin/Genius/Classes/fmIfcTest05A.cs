using Doyle_Addin.Genius.Forms;
using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class fmIfcTest05A : Form
{
    public fmIfcTest05A Itself()
    {
        // returns this fmIfcTest04A class instance "Itself"
        // should be HIGHLY useful inside a With context
        return this;
    }

    public fmIfcTest05A Using(Dictionary Dict = null) // fmTest05
    {
        if (Dict == null)
        {
            using ( == Using(new Dictionary()))
            {
                ;
                allGroups = null;
                var dp = dcDepthAiDocGrp(Dict);

                switch (dp)
                {
                    case 1:
                    {
                        using ( == withDcFlat(Dict))
                        {
                            ;
                            Case 2:

                            using ( == withDcGrpd(Dict))
                            {
                                Else itemsFlat = new Dictionary();
                                {
                                    var withBlock = allGroups;
                                    foreach (var ky in withBlock.Keys)
                                        itemsFlat = dcKeysCombined(dcOb(withBlock.get_Item(ky)), itemsFlat, 1);
                                }
                                ;
                                Case Else:

                                using ( == this)
                                {
                                }
                            }
                        }

                        break;
                    }
                }
            }
        }
    }

    public string GroupNow()
    {
        return fm.GroupNow();
    }

    public fmIfcTest05A InGroup(string GrpId)
    {
        if (fm.InGroup(GrpId).GroupNow() == GrpId)
        {
        }

        return this;
    }

    public string ItemNow()
    {
        return fm.ItemNow();
    }

    public fmIfcTest05A OnItem(string ItemId)
    {
        if (fm.OnItem(ItemId).ItemNow() == ItemId)
        {
        }
        else
            // couldn't change!
            Debugger.Break();

        return this;
    }

    public fmIfcTest05A Show(dynamic Modal)
    {
        fm.Show(Modal);
        return this;
    }

    public fmIfcTest05A Hide()
    {
        fm.Hide();
        return this;
    }

    public Dictionary SaveAll()
    {
        // debug.Print aiDocument(itemsFlat.get_Item(itemsFlat.Keys(0))).Dirty
        // NOTE[2022.04.13.1224] (copied from ...)
        // want to initiate 'save all' operation here
        // or somewhere nearby. note immediate mode
        // command in comment above
        // Dim rtBd As Scripting.Dictionary

        // Dim mx As Long
        // Dim dx As Long
        var rtGd = new Dictionary();

        // rtBd = New Scripting.Dictionary
        {
            var withBlock = itemsFlat;
            foreach (var ky in withBlock.Keys)
            {
                Document wk = withBlock.get_Item(ky);
                {
                    if (wk.Dirty)
                    {
                        Information.Err().Clear();
                        wk.Save2();
                        if (Information.Err().Number == 0)
                            rtGd.Add(ky, wk);
                    }
                    else
                        rtGd.Add(ky, wk);
                }
            }
        }

        return rtGd;
    }

    private fmIfcTest05A withDcFlat(Dictionary Dict) // fmTest05
    {
        itemsFlat = dcCopy(Dict);
        return withDcGrpd(dcAiDocGrpsByForm(itemsFlat));
    }

    private fmIfcTest05A withDcGrpd(Dictionary Dict) // fmTest05
    {
        itmPicked = new Dictionary();
        docActive = null;

        MSForms.Tabs ls = tbsItemGrps.Tabs;
        ls.Clear();

        allGroups = Dict;

        {
            var withBlock = allGroups;
            foreach (var ky in Split("MAYB DBAR SHTM ASSY PRCH HDWR")) // instead of .Keys, to
            {
                // ensure preferred order
                if (withBlock.Exists(ky))
                {
                    {
                        var withBlock1 = dcOb(withBlock.get_Item(ky));
                        // ' check for group members
                        if (withBlock1.Count > 0)
                        {
                            itmPicked.Add(ky, withBlock1.Keys(0));
                            ls.Add(ky);
                        }
                    }
                }
            }
        }
        return this;
    }

    private string gpActive()
    {
        MSForms.Tab tb;

        {
            var withBlock = tbsItemGrps;
            tb = withBlock.Tabs.get_Item(withBlock.Value);
        }
        return tb.Name;
    }

    private void cmdOpenItem_Click()
    {
        if (docActive == null)
        {
        }
        else
        {
            var withBlock = docActive;
            string pn = withBlock.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;

            // REV[2022.05.06.1142]
            // added check for Part Document to avoid error
            // trying to edit Material for Assembly Documents.
            if (docActive is PartDocument)
                VbMsgBoxResult ck = MessageBox.Show(Join(new[]
                    {
                        "Would you rather just edit", "material for " + pn + "?", "", "(No to go ahead and open)"
                    },
                    Constants.vbCrLf), Constants.vbYesNoCancel + Constants.vbQuestion, "Edit Material?");
            else
                ck = Constants.vbNo;
            if (ck == Constants.vbCancel)
                Debugger.Break();
            else if (ck == Constants.vbYes)
                // NOTE[2022.05.06.1143]
                // this section throws an error
                // if the Document is an assembly.
                // REV[2022.05.06.1142] above adds
                // a check to prevent this branch
                // from being taken in that case.
                Debug.Print(ConvertToJson(askUserForPartMatlUpdate(itemsFlat.get_Item(pn)), Constants.vbTab));
            else
            {
                if (withBlock.Open)
                    ck = Constants.vbYes;
                else
                    ck = MessageBox.Show(Join(new[]
                    {
                        "Document " + pn, "is not presently open.",
                        "Go ahead and open it?"
                    }, Constants.vbCrLf), Constants.vbYesNo, "Open " + pn + "?");
                if (ck == Constants.vbYes)
                {
                    Information.Err().Clear();
                    withBlock.Activate();
                    if (Information.Err().Number == 0)
                    {
                    }
                    else
                    {
                        if (ThisApplication.Documents.Open(withBlock.FullDocumentName, true) == docActive)
                            Debug.Print(""); // Breakpoint Landing
                        else
                        {
                            Debugger.Break();
                            Debug.Print(""); // Breakpoint Landing
                        }

                        Information.Err().Clear();
                        withBlock.Activate();
                        if (Information.Err().Number)
                            Debugger.Break();
                        Debug.Print(""); // Breakpoint Landing
                    }
                }
            }
        }
    }

    private void fm_GroupIs(string Now)
    {
        // 
        GroupIs?.Invoke(Now);
    }

    private void fm_ItemIs(string Now)
    {
        // 
        ItemIs?.Invoke(Now);
    }

    private void fm_Sent(VbMsgBoxResult Signal)
    {
        // Public Event Sent(Signal As VbMsgBoxResult)
        // 

        if (obClient != null) return;
        VbMsgBoxResult ck = Constants.vbRetry;

        switch (Signal)
        {
            case dynamic _ when Constants.vbOK:
            {
                // ck = MessageBox.Show(Join(new [] {' "Save and Close",' "Operation Selected"'), vbCrLf),' vbYesNoCancel,' "Save Documents?"')

                // debug.Print aiDocument(itemsFlat.get_Item(itemsFlat.Keys(0))).Dirty
                // NOTE[2022.04.13.1224]
                // want to initiate 'save all' operation here
                // or somewhere nearby. note immediate mode
                // command in comment above
                {
                    var withBlock = dcKeysMissing(itemsFlat, SaveAll());
                    if (withBlock.Count > 0)
                        ck = MessageBox.Show(Join(new[]
                            {
                                "Errors encountered trying to", "save the following Documents:",
                                Constants.vbTab + txDumpLs(withBlock.Keys, Constants.vbCrLf + Constants.vbTab),
                                "", "Close anyway?"
                            }, Constants.vbCrLf), Constants.vbYesNoCancel,
                            "Errors on Save!");
                    else
                        ck = Constants.vbYes;
                }

                break;
            }

            case dynamic _ when Constants.vbAbort:
            case dynamic _ when Constants.vbCancel:
            {
                ck = MessageBox.Show(Join(new[]
                    {
                        "Cancel", "Operation", "Selected"
                    }, Constants.vbCrLf), Constants.vbYesNoCancel,
                    "Finished?");
                break;
            }

            default:
            {
                break;
            }
        }

        if (ck == Constants.vbCancel)
            Debugger.Break();
        else if (ck == Constants.vbYes)
            Hide();
    }
    else

    {
        Debugger.Break();
        Sent?.Invoke(Signal);
    }

    private void lbxItems_Change()
    {
        // Stop
        string pn = lbxItems.Value;
        {
            var withBlock = gdcActive;
            if (withBlock.Exists(pn))
            {
                docActive = aiDocument(withBlock.get_Item(pn));
                {
                    var withBlock1 = docActive;
                    {
                        var withBlock2 = withBlock1.PropertySets.get_Item(gnDesign);
                        lblPartNum.Caption = withBlock2.get_Item(pnPartNum).Value;
                        lblDesc.Caption = withBlock2.get_Item(pnDesc).Value;
                    }

                    stdole.StdPicture pc = withBlock1.Thumbnail;

                    if (Information.Err().Number == 0)
                    {
                    }
                    else
                        pc = null;

                    imgOfItem.Picture = pc;
                }
            }
            else
            {
                docActive = null;
                lblPartNum.Caption = "(select part)";
                lblDesc.Caption = "";
            }
        }

        // REV[2022.03.17.1348]
        // add DoEvents for rapid visual feedback
        // (see tbsItemGrps_Change for details)
        DoEvents();
    }

    private void tbsItemGrps_Change()
    {
        MSForms.Tab tb;

        var nm =
            // With tbsItemGrps
            // tb = .Tabs.get_Item(.Value)
            // End With
            gpActive(); // tb.Name

        gdcActive = dcOb(allGroups.get_Item(nm));
        {
            var withBlock = lbxItems;
            withBlock.List = gdcActive.Keys; // dcOb(allGroups.get_Item(nm)).Keys
            // If gdcActive.Count > 0 Then
            withBlock.Value = itmPicked.get_Item(nm); // gdcActive.get_Item()
        }

        // REV[2022.03.17.1348]
        // adding DoEvents steps to various
        // Change Event handlers to try to
        // ensure timely visual feedback
        // to the User in-process
        DoEvents();
    }

    private void Class_Initialize()
    {
        // fm =
        {
            var withBlock = new fmTest05();
            // End With
            // With fm
            lbxItems = withBlock.lbxItems;
            tbsItemGrps = withBlock.tbsItemGrps;
            lblPartNum = withBlock.lblPartNum;
            lblDesc = withBlock.lblDesc;
            imgOfItem = withBlock.imgOfItem;
            cmdOpenItem = withBlock.cmdOpenItem;
            // cmdEndCancel = .cmdEndCancel
            // cmdEndSave = .cmdEndSave

            fm = withBlock.Holding(this);
        }

        // dcActvSet = New Scripting.Dictionary
        allGroups = new Dictionary();
    }

    private void Class_Terminate()
    {
        allGroups = null;

        fm = fm.Dropping(this);

        // cmdEndSave = Nothing
        // cmdEndCancel = Nothing
        cmdOpenItem = null;
        imgOfItem = null;
        lblDesc = null;
        lblPartNum = null;
        tbsItemGrps = null;

        lbxItems = null;
        lbxItems = null;

        fm = null;
    }
}