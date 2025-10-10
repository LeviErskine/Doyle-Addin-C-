class fmIfcTest05A
{
    public fmIfcTest05A Itself()
    {
        /// returns this fmIfcTest04A class instance "Itself"
        /// should be HIGHLY useful inside a With context
        Itself = this;
    }

    public fmIfcTest05A Using(Scripting.Dictionary Dict = null/* TODO Change to default(_) if this is not a reference type */) // fmTest05
    {
        Variant ky;
        long dp;

        if (Dict == null)
        {
            using ( == this.Using(new Scripting.Dictionary()))
            {
                ;
    Else

                itemsFlat = null;
                allGroups = null;
                dp = dcDepthAiDocGrp(Dict);

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
                                    itemsFlat = new Scripting.Dictionary();
                                    {
                                        var withBlock = allGroups;
                                        foreach (var ky in withBlock.Keys)
                                            itemsFlat = dcKeysCombined(dcOb(withBlock.Item(ky)), itemsFlat, 1);
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
        GroupNow = fm.GroupNow();
    }

    public fmIfcTest05A InGroup(string GrpId)
    {
        if (fm.InGroup(GrpId).GroupNow() == GrpId)
        {
        }
        else
        {
        }

        InGroup = this;
    }

    public string ItemNow()
    {
        ItemNow = fm.ItemNow();
    }

    public fmIfcTest05A OnItem(string ItemId)
    {
        if (fm.OnItem(ItemId).ItemNow() == ItemId)
        {
        }
        else
            /// couldn't change!
            System.Diagnostics.Debugger.Break();

        OnItem = this;
    }

    public fmIfcTest05A Show(Variant Modal)
    {
        fm.Show(Modal);
        Show = this;
    }

    public fmIfcTest05A Hide()
    {
        fm.Hide();
        Hide = this;
    }

    public Scripting.Dictionary SaveAll()
    {
        // debug.Print aiDocument(itemsFlat.Item(itemsFlat.Keys(0))).Dirty
        /// NOTE[2022.04.13.1224] (copied from ...)
        /// want to initiate 'save all' operation here
        /// or somewhere nearby. note immediate mode
        /// command in comment above
        Scripting.Dictionary rtGd;
        // Dim rtBd As Scripting.Dictionary
        Inventor.Document wk;
        Variant ky;
        // Dim mx As Long
        // Dim dx As Long

        rtGd = new Scripting.Dictionary();
        // rtBd = New Scripting.Dictionary

        {
            var withBlock = itemsFlat;
            foreach (var ky in withBlock.Keys)
            {
                wk = withBlock.Item(ky);
                {
                    var withBlock1 = wk;
                    if (withBlock1.Dirty)
                    {
                        Information.Err.Clear();
                        withBlock1.Save2();
                        if (Information.Err.Number == 0)
                            rtGd.Add(ky, wk);
                        else
                        {
                        }
                    }
                    else
                        rtGd.Add(ky, wk);
                }
            }
        }

        SaveAll = rtGd;
    }

    private fmIfcTest05A withDcFlat(Scripting.Dictionary Dict) // fmTest05
    {
        itemsFlat = dcCopy(Dict);
        withDcFlat = withDcGrpd(dcAiDocGrpsByForm(itemsFlat));
    }

    private fmIfcTest05A withDcGrpd(Scripting.Dictionary Dict) // fmTest05
    {
        Variant ky;
        MSForms.Tabs ls;

        itmPicked = new Scripting.Dictionary();
        docActive = null;

        ls = tbsItemGrps.Tabs;
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
                        var withBlock1 = dcOb(withBlock.Item(ky));
                        // '  check for group members
                        if (withBlock1.Count > 0)
                        {
                            itmPicked.Add(ky, withBlock1.Keys(0));
                            ls.Add(ky);
                        }
                        else
                        {
                        }
                    }
                }
            }
        }
        withDcGrpd = this;
    }

    private string gpActive()
    {
        MSForms.Tab tb;

        {
            var withBlock = tbsItemGrps;
            tb = withBlock.Tabs.Item(withBlock.Value);
        }
        gpActive = tb.Name;
    }

    private void cmdOpenItem_Click()
    {
        VbMsgBoxResult ck;
        string pn;

        if (docActive == null)
        {
        }
        else
        {
            var withBlock = docActive;
            pn = withBlock.PropertySets.Item(gnDesign).Item(pnPartNum).Value;

            /// REV[2022.05.06.1142]
            /// added check for Part Document to avoid error
            /// trying to edit Material for Assembly Documents.
            if (docActive is Inventor.PartDocument)
                ck = MsgBox(Join(Array("Would you rather just edit", "material for " + pn + "?", "", "(No to go ahead and open)"), Constants.vbNewLine), Constants.vbYesNoCancel + Constants.vbQuestion, "Edit Material?");
            else
                ck = Constants.vbNo;

            if (ck == Constants.vbCancel)
                System.Diagnostics.Debugger.Break();
            else if (ck == Constants.vbYes)
                /// NOTE[2022.05.06.1143]
                /// this section throws an error
                /// if the Document is an assembly.
                /// REV[2022.05.06.1142] above adds
                /// a check to prevent this branch
                /// from being taken in that case.
                Debug.Print(ConvertToJson(askUserForPartMatlUpdate(itemsFlat.Item(pn)), Constants.vbTab));
            else
            {
                if (withBlock.Open)
                    ck = Constants.vbYes;
                else
                    ck = MsgBox(Join(Array("Document " + pn, "is not presently open.", "Go ahead and open it?"), Constants.vbNewLine), Constants.vbYesNo, "Open " + pn + "?");

                if (ck == Constants.vbYes)
                {
                    Information.Err.Clear();
                    withBlock.Activate();

                    if (Information.Err.Number == 0)
                    {
                    }
                    else
                    {
                        if (ThisApplication.Documents.Open(withBlock.FullDocumentName, true) == docActive)
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        else
                        {
                            System.Diagnostics.Debugger.Break();
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }

                        Information.Err.Clear();
                        withBlock.Activate();
                        if (Information.Err.Number)
                            System.Diagnostics.Debugger.Break();

                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                }
            }
        }
    }

    private void fm_GroupIs(string Now)
    {
        /// 
        GroupIs?.Invoke(Now);
    }

    private void fm_ItemIs(string Now)
    {
        /// 
        ItemIs?.Invoke(Now);
    }

    private void fm_Sent(VbMsgBoxResult Signal)
    {
        // Public Event Sent(Signal As VbMsgBoxResult)
        /// 
        VbMsgBoxResult ck;

        if (obClient == null)
        {
            ck = Constants.vbRetry;

            switch (Signal)
            {
                case object _ when Constants.vbOK:
                    {
                        // ck = MsgBox(Join(Array('    "Save and Close",'    "Operation Selected"'), vbNewLine),'    vbYesNoCancel,'    "Save Documents?"')

                        // debug.Print aiDocument(itemsFlat.Item(itemsFlat.Keys(0))).Dirty
                        /// NOTE[2022.04.13.1224]
                        /// want to initiate 'save all' operation here
                        /// or somewhere nearby. note immediate mode
                        /// command in comment above
                        {
                            var withBlock = dcKeysMissing(itemsFlat, SaveAll());
                            if (withBlock.Count > 0)
                                ck = MsgBox(Join(Array("Errors encountered trying to", "save the following Documents:", Constants.vbTab + txDumpLs(withBlock.Keys, Constants.vbNewLine + Constants.vbTab), "", "Close anyway?"), Constants.vbNewLine), Constants.vbYesNoCancel, "Errors on Save!");
                            else
                                ck = Constants.vbYes;
                        }

                        break;
                    }

                case object _ when Constants.vbAbort:
                case object _ when Constants.vbCancel:
                    {
                        ck = MsgBox(Join(Array("Cancel", "Operation", "Selected"), Constants.vbNewLine), Constants.vbYesNoCancel, "Finished?");
                        break;
                    }

                default:
                    {
                        break;
                    }
            }

            if (ck == Constants.vbCancel)
                System.Diagnostics.Debugger.Break();
            else if (ck == Constants.vbYes)
                Hide();
        }
        else
        {
            System.Diagnostics.Debugger.Break();
            Sent?.Invoke(Signal);
        }
    }

    private void lbxItems_Change()
    {
        string pn;
        stdole.StdPicture pc;

        // Stop
        pn = lbxItems.Value;
        {
            var withBlock = gdcActive;
            if (withBlock.Exists(pn))
            {
                docActive = aiDocument(withBlock.Item(pn));
                {
                    var withBlock1 = docActive;
                    {
                        var withBlock2 = withBlock1.PropertySets.Item(gnDesign);
                        lblPartNum.Caption = withBlock2.Item(pnPartNum).Value;
                        lblDesc.Caption = withBlock2.Item(pnDesc).Value;
                    }


                    pc = withBlock1.Thumbnail;

                    if (Information.Err.Number == 0)
                    {
                    }
                    else
                        pc = null/* TODO Change to default(_) if this is not a reference type */;

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

        /// REV[2022.03.17.1348]
        /// add DoEvents for rapid visual feedback
        /// (see tbsItemGrps_Change for details)
        DoEvents();
    }

    private void tbsItemGrps_Change()
    {
        MSForms.Tab tb;
        string nm;

        // With tbsItemGrps
        // tb = .Tabs.Item(.Value)
        // End With
        nm = gpActive(); // tb.Name

        gdcActive = dcOb(allGroups.Item(nm));
        {
            var withBlock = lbxItems;
            withBlock.List = gdcActive.Keys; // dcOb(allGroups.Item(nm)).Keys
                                             // If gdcActive.Count > 0 Then
            withBlock.Value = itmPicked.Item(nm); // gdcActive.Item()
        }

        /// REV[2022.03.17.1348]
        /// adding DoEvents steps to various
        /// Change Event handlers to try to
        /// ensure timely visual feedback
        /// to the User in-process
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
        allGroups = new Scripting.Dictionary();
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