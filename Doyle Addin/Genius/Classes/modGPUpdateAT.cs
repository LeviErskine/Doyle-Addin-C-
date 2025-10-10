class SurroundingClass
{
    public Scripting.Dictionary dcGeniusPropsPartRev20180530(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// WAYPOINTS (search on phrase)
        /// (NOT Sheet Metal)
        /// 
        // (removed lines 11-41)
        Scripting.Dictionary rt;
        /// REV[2022.01.21.1351] (lines 43-44)
        Scripting.Dictionary dcIn;
        /// to collect settings already in Genius
        Scripting.Dictionary dcFP;
        // (removed lines 48-51)

        // '
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;
        // '
        /// ADDED[2021.03.11] (lines 58-60)
        Inventor.Property prPartNum; // pnPartNum
        Inventor.Property prFamily;
        /// REV[2023.05.23.1134]
        /// introduce break between built-in
        /// and user-defined properties to
        /// better distinguish the two groups
        /// 
        /// also move ADDED note (above)
        /// before prPartNum to pull
        /// the two built-ins together
        /// 
        Inventor.Property prRawMatl; // pnRawMaterial
        Inventor.Property prRmUnit; // pnRmUnit
        Inventor.Property prRmQty; // pnRmQty
                                   // '
        string pnModel;
        /// ADDED[2021.03.11] (lines 67-68)
        string nmFamily;
        string mtFamily;
        /// REV[2022.05.05.1110] (lines 71-85)
        double qtRawMatl;
        string pnStock;
        string qtUnit;
        Inventor.BOMStructureEnum bomStruct;
        VbMsgBoxResult ck;
        aiBoxData bd;
        /// REV[2022.09.29.1448]
        /// added String variable txTmp
        /// as temporary text holder
        /// initially for lagging assignment (see below)
        /// but potentially useful in other places
        string txTmp;

        if (dc == null)
            dcGeniusPropsPartRev20180530 = dcGeniusPropsPartRev20180530(invDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = invDoc;
                /// REV[2022.05.06.1113] (lines 102-112)
                if (withBlock.ComponentDefinition.IsContentMember)
                {
                }

                // Property Sets
                {
                    var withBlock1 = withBlock.PropertySets;
                    aiPropsUser = withBlock1.Item(gnCustom);
                    aiPropsDesign = withBlock1.Item(gnDesign);
                }

                // Custom Properties...
                /// REV[2022.05.06.1124] (lines 124-130)
                if (withBlock.ComponentDefinition.IsContentMember)
                {
                    pnStock = "";
                    qtRawMatl = 0#;
                    qtUnit = "";
                }
                else
                {
                }

                // Part Number and Family properties
                // are from Design, NOT Custom set
                prPartNum = aiGetProp(aiPropsDesign, pnPartNum);
                // ADDED 2021.03.11
                pnModel = prPartNum.Value;
                prFamily = aiGetProp(aiPropsDesign, pnFamily);
                /// REV[2022.05.05.1551] (lines 179-185)
                nmFamily = famVsGenius(pnModel, prFamily.Value);

                /// REV[2022.06.29.1351] (lines 188-225)

                /// We should check HERE for possibly misidentified purchased parts
                /// UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// Request #1:  the Mass in Pounds
                    /// and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        /// REV[2021.11.12] (lines 233-241)

                        rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                        if (Information.Err.Number)
                            // (removed lines 247-260)
                            System.Diagnostics.Debugger.Break();
                    }

                    /// BOM Structure type, correcting if appropriate,
                    /// and prepare Family value for part, if purchased.
                    /// 
                    ck = Constants.vbNo;
                    /// REV[2022.05.06.1118] (lines 271-273)
                    if (withBlock1.IsContentMember)
                        ck = Constants.vbYes;
                    else if (InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|D-BAR|DSHEET|", "|" + nmFamily + "|") > 0)
                        /// REV[2022.06.29.1416] (lines 277-281)
                        ck = Constants.vbYes;
                    else if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|D-BAR|DSHEET|", "|" + nmFamily + "|") > 0)
                        /// REV[2022.05.06.1118] (lines 288-299)
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
                    // REV[2018.05.30] (lines 331-334)
                    /// REV[2023.01.16.1618]
                    /// embedded default nmFamily assignment
                    /// in length check on nmFamily to avoid
                    /// overwriting nonblank value,
                    /// such as from Genius
                    if (Strings.Len(nmFamily) == 0)
                    {
                        if (bomStruct == kPurchasedBOMStructure)
                        {
                            if (withBlock1.IsContentMember)
                                /// NOTE[2022.05.06.1130] (lines 337-340)
                                nmFamily = "D-HDWR";
                            else
                                nmFamily = "D-PTS";
                        }
                        else
                        {
                        }
                    }
                }
                // (remove lines 358-361)

                // Request #4: Change Cost Center iProperty.
                // If BOMStructure = Normal, then Family = D-MTO,
                // else if BOMStructure = Purchased then Family = D-PTS.
                if (bomStruct == kNormalBOMStructure)
                {
                    /// REV[2023.05.23.1148]
                    /// 
                    /// move collection of user-defined properties into
                    /// start of If block for bomStruct = kNormalBOMStructure
                    /// to avoid unecessary creation of these properties
                    /// where not needed; specifically in purchased parts.
                    /// 
                    /// search on REV tag above to find original location
                    /// source there remains in commented form, pending removal
                    /// 
                    prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1);
                    prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1);
                    prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1);

                    // ...and their initial Values
                    if (prRawMatl == null)
                        /// REV[2022.05.10.1427] (lines 142-144)
                        pnStock = "";
                    else
                        pnStock = prRawMatl.Value;

                    /// REV[2022.05.05.1517] (lines 149-152)
                    if (prRmQty == null)
                        qtRawMatl = 0#;
                    else if (IsNumeric(prRmQty.Value))
                        qtRawMatl = Round(prRmQty.Value, 4);
                    else
                        qtRawMatl = 0#;

                    if (prRmUnit == null)
                        qtUnit = "";
                    else
                        qtUnit = prRmUnit.Value;
                    /// REV[2022.05.05.1128]
                    /// added initial Value collection
                    /// for custom Raw Material Properties
                    // '''''
                    // '''''
                    /// END of REV[2023.05.23.1148]
                    // '''''
                    // '''''

                    /// REV[2022.01.28.1014] (lines 368-373)
                    pnStock = prRawMatl.Value;
                    /// REV[2022.02.08.1304] (lines 375-396)
                    dcIn = dcFromAdoRS(cnGnsDoyle().Execute(sqlOf_GnsPartMatl(pnModel)));
                    // (removed lines 400-402)
                    if (dcIn.Count > 0)
                    {
                        {
                            var withBlock1 = dcOb(dcDxFromRecSetDc(dcIn).Item(pnRawMaterial));
                            /// REV[2022.01.28.1336] (lines 405-413)
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
                                    // (removed lines 427-429)
                                    pnStock = withBlock1.Keys(0);

                                // and use it for the default...
                                if (withBlock1.Count > 1)
                                {
                                    /// REV[2022.04.14.1131] (lines 444-447)
                                    Debug.Print(pnModel + Constants.vbNewLine + Constants.vbTab + Join(withBlock1.Keys, Constants.vbNewLine + Constants.vbTab));                                 // (removed lines 449-452)

                                    pnStock = nuSelector().GetReply(withBlock1.Keys, pnStock);

                                    Debug.Print("Selected " + Interaction.IIf(Strings.Len(pnStock) > 0, pnStock, "(nothing)"));
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
                                    // it'll be taken care of further down
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                else if (pnStock == prRawMatl.Value)
                                    // (removed lines 477-483)
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                else
                                {
                                    // (removed lines 486-506)
                                    ck = Constants.vbOK;
                                    if (ck == Constants.vbCancel)
                                        System.Diagnostics.Debugger.Break(); // to check things out
                                    else if (ck == Constants.vbNo)
                                        /// NOTE[2022.02.08.1359]
                                        /// DO NOT DISABLE this instance
                                        /// of the pnStock assignment!
                                        pnStock = prRawMatl.Value;
                                }
                                /// REV[2022.01.28.1448] (lines 536-554)
                                if (withBlock1.Exists(pnStock))
                                {
                                    dcIn = dcOb(dcIn.Item(dcOb(withBlock1.Item(pnStock)).Keys(0)));
                                    // (removed lines 557-567)
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
                        /// NOTE[2018-05-31] (602-608)
                        /// REV[2022.01.28.0903] (609-614)
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
                            ck = Constants.vbRetry;// vbNo

                        /// REV[2022.01.31.1335] (lines 631-633)
                        if (ck == Constants.vbNo)
                            dcFP = new Scripting.Dictionary();
                        else
                        {
                            /// NOTE[2022.04.12.1157] (lines 639-644)
                            dcFP = dcFlatPatVals(withBlock.ComponentDefinition);
                            // (removed lines 646-649)

                            /// REV[2023.02.15.1503]
                            /// added duplication of pnStock to txTmp
                            /// in order to check for changes.
                            /// user prompt added to this segment
                            /// immediately disabled as redundant.
                            /// most changes to this segment can
                            /// PROBABLY be reverted
                            /// REV[2023.04.21.1503]
                            /// re-enabling user prompt (see below)
                            /// REV[2023.04.24.0928]
                            /// overhauling -- see below
                            if (dcFP.Exists(pnThickness))
                            {
                                /// REV[2023.04.24.0929]
                                /// switched assignment of projected
                                /// sheet metal item directly to txTmp
                                /// instead of using it as a placeholder
                                /// for pnStock. that way, only one
                                /// assignment is required at this stage
                                txTmp = ptNumShtMetal(invDoc.ComponentDefinition); // pnStock
                                                                                   // pnStock = ptNumShtMetal(invDoc.ComponentDefinition)

                                /// NOTE[2022.05.31.1158] (lines 653-662)
                                if (Strings.Len(pnStock) == 0)
                                    /// REV[2023.04.24.0949]
                                    /// added check for blank pnStock
                                    /// with automatic assignment from
                                    /// txTmp in that case
                                    /// 
                                    pnStock = txTmp;
                                else if (Strings.Len(txTmp) == 0)
                                    System.Diagnostics.Debugger.Break();
                                else if (pnStock != txTmp)
                                {
                                    // change confirmation code
                                    // duplicated with modification
                                    // from NOTE[2022.01.03] (below)
                                    // 
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                    /// REV[2023.04.21.1503]
                                    /// re-enabling user prompt to fix
                                    /// issue with potential changes not
                                    /// getting picked up
                                    /// also switching pnStock with txTmp
                                    /// based on assignments above, they
                                    /// were in the wrong order in the prompt
                                    /// REV[2023.04.21.1526] disabled AGAIN
                                    /// and changed check [REV:1527] below
                                    ck = newFmTest2().AskAbout(invDoc, "Suggest Material change from" + Constants.vbNewLine + pnStock + " to" + Constants.vbNewLine + txTmp + " for", "Change it?"); // vbYes '
                                    if (ck == Constants.vbCancel)
                                        // Debug.Print ConvertToJson(nuDcPopulator.Setting(pnModel, nuDcPopulator.Setting("from", prRawMatl.Value).Setting("into", pnStock).Dictionary).Dictionary, vbTab)
                                        System.Diagnostics.Debugger.Break(); // to check things out
                                    else if (ck == Constants.vbYes)
                                        /// REV[2023.04.21.1527] switch check
                                        /// from YES to NO to force new stock
                                        /// number. it SHOULD get picked up
                                        /// and prompted toward the end.
                                        pnStock = txTmp;
                                }
                                txTmp = "";
                                dcFP.Add(pnRawMaterial, pnStock);
                            }
                            else
                                // so clear it for now
                                pnStock = "";
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
                                        // Stop
                                        ck = newFmTest2().AskAbout(invDoc, "This Part: ", "might not be sheet metal. " + Constants.vbNewLine + vbNewLine);
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
                            Debug.Print(ConvertToJson(Array(dcIn, dcFP, pnModel), Constants.vbTab)); System.Diagnostics.Debugger.Break(); // so we can figure out what to do next.
                        }

                        // Request #3:
                        // sheet metal extent area
                        // and add to custom property "RMQTY"

                        /// REV[2022.01.28.1556] (lines 724-726)
                        if (ck == Constants.vbYes)
                            rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                        else if (ck == Constants.vbRetry)
                            rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                        else if (ck == Constants.vbNo)
                        {
                        }
                        else
                            // (removed lines 743-745)
                            System.Diagnostics.Debugger.Break();// and check it out

                        /// NOTE[2018-05-30] (lines 749-762)
                        if (prRawMatl == null)
                        {
                            if (rt.Exists("OFFTHK"))
                            {
                                /// NOTE[2021.12.10] (lines 765-769)
                                /// UPDATE[2018.05.30] (lines 770-775)
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
                            // '  ACTION ADVISED[2018.09.14] (lines 788-796)
                            /// REV[2021.12.17] (lines 797-800)
                            if (Len(prRawMatl.Value) > 0)
                            {
                                // we need to check it

                                if (Strings.Len(pnStock) == 0)
                                    /// REV[2022.01.28.1445] (lines 805-808)
                                    pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                                /// NOTE[2021.12.17@15:32] (lines 823-827)
                                /// NOTE[2021.12.17] (lines 828-835)
                                /// NOTE[2022.01.05] (lines 836-841)
                                if (Strings.Len(pnStock) > 0)
                                {
                                    if (pnStock != prRawMatl.Value)
                                    {
                                        // Stop
                                        /// REV[2022.04.19.0945] (lines 845-861)
                                        /// REV[2022.04.19.0944] (lines 862-865)
                                        if (UCase(prRawMatl.Value) == pnStock)
                                            ck = Constants.vbYes;
                                        else
                                            /// NOTE[2022.01.03] (lines 869-871)
                                            ck = newFmTest2().AskAbout(invDoc, "Suggest Sheet Metal change" + Constants.vbNewLine + "from " + prRawMatl.Value, "Change it?");

                                        if (ck == Constants.vbCancel)
                                        {
                                            Debug.Print(ConvertToJson(nuDcPopulator.Setting(pnModel, nuDcPopulator.Setting("from", prRawMatl.Value).Setting("into", pnStock).Dictionary).Dictionary, Constants.vbTab));
                                            System.Diagnostics.Debugger.Break(); // to check things out
                                        }
                                        else if (ck == Constants.vbYes)
                                        {
                                            Information.Err.Clear();
                                            prRawMatl.Value = pnStock;
                                            if (Information.Err.Number)
                                                System.Diagnostics.Debugger.Break();// and check for Member not Found
                                        }
                                    }
                                }
                            }
                            else if (Strings.Len(pnStock) > 0)
                            {
                                // go ahead and assign material

                                Information.Err.Clear();
                                prRawMatl.Value = pnStock;
                                if (Information.Err.Number)
                                    System.Diagnostics.Debugger.Break();
                            }

                            if (Len(prRawMatl.Value) > 0)
                            {
                                if (rt.Exists("OFFTHK"))
                                {
                                    // Stop 'and verify raw material item
                                    /// NOTE[2021.12.13] (lines 902-905)
                                    ck = newFmTest2().AskAbout(invDoc, "Assigned Raw Material " + prRawMatl.Value, "Clear it?");
                                    if (ck == Constants.vbCancel)
                                        System.Diagnostics.Debugger.Break(); // to check things out
                                    else if (ck == Constants.vbYes)
                                    {
                                        prRawMatl.Value = "";
                                        pnStock = prRawMatl.Value;
                                    }
                                }


                                if (pnStock == prRawMatl.Value)
                                    // no need to assign it again
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                else
                                {
                                    Debug.Print(ConvertToJson(Array(pnStock, prRawMatl.Value)));                                 // Stop 'before we do something stupid!
                                    pnStock = prRawMatl.Value;
                                }

                                /// The following With block copied and modified [2021.03.11]
                                /// from elsewhere in this function as a temporary measure
                                /// to address a stopping situation later in the function.
                                /// See comment below for details.
                                /// 
                                {
                                    var withBlock1 = cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", pnStock));
                                    // With cnGnsDoyle().Execute("select Family from vgMfiItems where Item='"& Replace(pnStock, "'", "''") & "';")
                                    /// REV[2022.08.26.1055] (lines 947-950)
                                    if (withBlock1.BOF | withBlock1.EOF)
                                    {
                                        if (pnStock != "0")
                                        {
                                            /// REV[2022.03.01.1553] (lines 953-960)
                                            if (Strings.Len(pnStock) > 0)
                                                /// REV[2022.07.07.1340]
                                                /// added secondary check for string length.
                                                /// an empty string requires no user attention.
                                                System.Diagnostics.Debugger.Break();// because Material value likely invalid
                                        }
                                        /// REV[2022.02.08.1413] (lines 968-976)
                                        /// UPDATE[2021.12.10] (lines 977-982)
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
                                /// UPDATE[2021.12.10] (lines 1021-1023)
                                pnStock = "";
                            else
                                pnStock = ptNumShtMetal(withBlock.ComponentDefinition);

                            if (Strings.Len(pnStock) == 0)
                            {
                                /// UPDATE[2018.05.30] (lines 1053-1062)
                                {
                                    var withBlock1 = newFmTest1();
                                    if (!(invDoc.ComponentDefinition.Document == invDoc))
                                        System.Diagnostics.Debugger.Break();

                                    bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                                    ck = withBlock1.AskAbout(invDoc, "No Stock Found! Please Review" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                                    if (ck == Constants.vbYes)
                                    {
                                        /// UPDATE[2018.05.30] (lines 1075-1077)
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
                                /// NOTE[2022.05.10.1559] (lines 1093-1097)
                                Debug.Print(pnModel + ": PROBABLE LAGGING [" + pnStock + "]");
                                Debug.Print("  TRY TO VERIFY. IF CHANGE REQUIRED,");
                                Debug.Print("  FILL IN NEW VALUE FOR pnStock BELOW, ");
                                Debug.Print("  AND PRESS ENTER ON THE LINE. WHEN ");
                                Debug.Print("  READY, PRESS [F5] TO CONTINUE.");
                                // Debug.Print "  pnStock = """ & pnStock & """"

                                ck = Constants.vbNo;
                                do
                                {
                                    txTmp = Trim(InputBox(Join(Array("Item " + pnModel + " appears", "to be lagging, likely " + pnStock + ".", "Try to verify, and if not correct,", "fill in correct material item below.", "", "(WARNING! update NOT working yet!", " Program will stop when entry complete", " to permit manual update)"), Constants.vbNewLine), "Verify Lagging " + pnStock + " for " + pnModel, pnStock));
                                    Debug.Print("  pnStock = \"" + txTmp + "\"");
                                    if (Strings.Len(txTmp) > 0)
                                    {
                                        if (txTmp == pnStock)
                                        {
                                            ck = MsgBox("Go ahead with " + txTmp + "?", Constants.vbYesNoCancel, "Confirm Same Material");
                                            if (ck == Constants.vbNo)
                                                ck = Constants.vbRetry;
                                        }
                                        else
                                            ck = MsgBox(Join(Array("Change Lagging Material ", pnStock + " to " + txTmp + "?"), Constants.vbNewLine), Constants.vbYesNoCancel, "Confirm Material Change");
                                    }
                                    else
                                    {
                                        // ck = MsgBox(Join(Array("Input appears to have been cleared.", "Are you sure you want to remove the", "current material," & pnStock & "?", "", "([Cancel] to debug)"), vbNewLine), vbYesNoCancel, "Remove Material?")
                                        // "Select [No] to keep it.",
                                        // 
                                        ck = MsgBox(Join(Array("No material entered.", "(perhaps entry was canceled?)", "", "Do you wish to remove the current", "material, " + pnStock + ", without replacement?"), Constants.vbNewLine), Constants.vbYesNoCancel, "No Material!");
                                        /// WARNING!!![2022.10.17.1434]
                                        /// there's something screwy going on here.
                                        /// Pressing [F8] in debug mode on the line above
                                        /// SHOULD stop on the next If statement, below.
                                        /// Instead, execution continues straight through
                                        /// to the Stop statement further down (which
                                        /// is SUPPOSED to go away, eventually!)
                                        /// 
                                        /// It's not clear what might be causing this,
                                        /// or what it might take to regain expected behavior.
                                        /// For now, have added a Breakpoint Landing
                                        /// in a crude attempt to address the matter.
                                        /// 
                                        Debug.Print();
                                        /// and THAT doesn't seem to be helping.
                                        /// will have to look into this more later.

                                        if (ck == Constants.vbNo)
                                        {
                                            ck = MsgBox("Do you want to keep " + pnStock + "?", Constants.vbYesNoCancel, "Keep Current?");
                                            if (ck == Constants.vbNo)
                                                ck = Constants.vbRetry;
                                            else if (ck == Constants.vbYes)
                                                ck = Constants.vbNo;
                                        }
                                    }
                                    // Stop

                                    if (ck == Constants.vbCancel)
                                        System.Diagnostics.Debugger.Break();
                                    else if (ck == Constants.vbRetry)
                                        ck = Constants.vbCancel;
                                    else if (ck == Constants.vbYes)
                                        pnStock = txTmp;
                                }
                                while (ck == Constants.vbCancel)/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing// to bypass debug below// to force retry
    ;
                            }

                            if (Strings.Len(pnStock) > 0)
                            {
                                // do we look for a Raw Material Family!
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing'Stop 'WAYPOINT to check block WITH@1764
                                /// REV[2022.08.26.1001]
                                /// placing temporary Stops at start
                                /// and end of following With block
                                /// to check use of fields normally
                                /// requested in SQL select statement.
                                /// 
                                {
                                    var withBlock1 = cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", pnStock));
                                    // preceding (disabled) With statement
                                    // to replace the following, assuming
                                    // tests prove successful. if so, it
                                    // might permit further streamlining
                                    // With cnGnsDoyle().Execute("select Family from vgMfiItems where Item='" &Replace(pnStock, "'", "''") & "';") ', Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16
                                    /// REV[2022.08.26.1059]
                                    /// (duping REV[2022.08.26.1055] above)
                                    /// replaced direct ref to pnStock
                                    /// with Replace operation to "escape"
                                    /// it, re REV[2022.08.19.1416] (below)
                                    /// REV[2022.08.26.1001] NOTE
                                    /// it is known that field Family
                                    /// is used directly below, however,
                                    /// usage of other fields is unclear.
                                    /// '
                                    /// to check their necessity, they
                                    /// have been removed from the SQL
                                    /// source string to a commend after
                                    /// the SQL call statement, to be
                                    /// recovered as needed.
                                    /// '
                                    /// Stops have been placed just before
                                    /// this With block (above), and just
                                    /// before its End (below) to mark both
                                    /// entry and exit from this block.
                                    /// in this way, it is hoped the critical
                                    /// period of execution may be delineated.
                                    /// '
                                    /// assuming no errors are encountered
                                    /// between entry and exit from this block,
                                    /// it may be assumed that no other fields
                                    /// but Family are required here, and they
                                    /// can likely be removed without harm.
                                    /// '
                                    /// this should permit replacement of the
                                    /// "hard-coded" SQL statement with a call
                                    /// to the new Function sqlOf_simpleSelWhere
                                    /// 
                                    if (withBlock1.BOF | withBlock1.EOF)
                                        System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                    else
                                    {
                                        {
                                            var withBlock2 = withBlock1.Fields;
                                            mtFamily = withBlock2.Item("Family").Value;
                                        }
                                        ;/* Cannot convert MultiLineIfBlockSyntax, System.NotSupportedException: LikeExpression not supported!
   at ICSharpCode.CodeConverter.CSharp.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/SyntaxKindExtensions.cs:line 278
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1415
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitMultiLineIfBlock(MultiLineIfBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 353
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                                    
                                    ''' UPDATE[2021.06.18] (lines 1172-1178)
                                    ''' REV[2022.04.15.1035] (lines 1179-1185)
                                    If mtFamily Like "?-MT*" Then
                                        ' (removed lines 1187-1188)
                                        Debug.Print pnModel & "["& qtRawMatl & qtUnit& " of " & pnStock & ": "& aiPropsDesign(pnDesc).Value& "]" 'prRmQty.Value prRawMatl.Value
                                        Stop 'FULL Stop!
                                    ''' NOTE[2022.05.05.1603]
                                    ''' new ElseIf branch called for here
                                    ''' see corresponding block under
                                    ''' Standard Part branch.
                                    ElseIf mtFamily = "D-PTS" Then
                                        'Stop 'NOT SO FAST!
                                        mtFamily = "D-BAR"
                                        'nmFamily = "D-RMT"
                                    ElseIf mtFamily = "R-PTS" Then
                                        Stop 'NOT SO FAST!
                                        mtFamily = "D-BAR"
                                        'nmFamily = "R-RMT"
                                    End If

 */
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

                                            /// UPDATE[2022.01.11] (lines 1229-1236)
                                            qtUnit = prRmUnit.Value; // "IN"
                                            ck = Constants.vbCancel;
                                            do
                                            {

                                                // 'may want function here
                                                /// UPDATE[2018.05.30]: (lines 1242-1272)
                                                /// REV[2022.02.09.0923] (lines 1273-1277)
                                                if (false)
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

                                                /// REV[2022.02.08.1446] (lines 1295-1299)
                                                Debug.Print("qtRawMatl = ");
                                                if (dcIn.Exists(pnRmQty))
                                                    Debug.Print("In Genius: ");
                                                Debug.Print();
                                                Debug.Print("qtUnit = \"");
                                                if (dcIn.Exists(pnRmUnit))
                                                    Debug.Print("In Genius: ");
                                                if (dcIn.Item(pnRmUnit) != "IN")
                                                    Debug.Print(" ( or try IN )");
                                                Debug.Print();
                                                // (removed lines 1307-1314)
                                                {
                                                    var withBlock2 = nu_fmIfcMatlQty01().SeeUser(invDoc);
                                                    if (withBlock2.Exists(pnRmQty))
                                                    {
                                                        /// REV[2022.04.04.1404] (lines 1317-1323)
                                                        if (System.Convert.ToDouble("0" + System.Convert.ToHexString(qtRawMatl)) == System.Convert.ToDouble(withBlock2.Item(pnRmQty)))
                                                        {
                                                        }
                                                        else
                                                        {
                                                            // Debug.Print "prRmQty.Value FROM " & prRmQty.Value & " TO " & .Item(pnRmQty)
                                                            Debug.Print("qtRawMatl FROM " + qtRawMatl + " TO " + withBlock2.Item(pnRmQty));

                                                            // Stop 'and double-check
                                                            // might still be equivalent
                                                            qtRawMatl = withBlock2.Item(pnRmQty);
                                                        }
                                                    }
                                                    else
                                                        System.Diagnostics.Debugger.Break();
                                                    if (withBlock2.Exists(pnRmUnit))
                                                    {
                                                        if (qtUnit == withBlock2.Item(pnRmUnit))
                                                        {
                                                        }
                                                        else
                                                        {
                                                            Debug.Print("qtUnit FROM " + qtUnit + " TO " + withBlock2.Item(pnRmUnit));
                                                            // Stop 'and double-check
                                                            // might still be equivalent
                                                            qtUnit = withBlock2.Item(pnRmUnit);
                                                        }
                                                    }
                                                    else
                                                        System.Diagnostics.Debugger.Break();
                                                }
                                                // (removed lines 1352-1358)
                                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                                ck = newFmTest2().AskAbout(invDoc, "Raw Material Quantity is now " + System.Convert.ToHexString(qtRawMatl) + qtUnit + " for", "If this is okay, click [YES]." + Constants.vbNewLine + "Otherwise, click [NO] to review." + Constants.vbNewLine + "" + Constants.vbNewLine + "( for debug, click [CANCEL] )");
                                                if (ck == Constants.vbCancel)
                                                    System.Diagnostics.Debugger.Break();
                                            }
                                            while (!ck == Constants.vbYes)/* TODO ERROR: Skipped SkippedTokensTrivia */// .Result()// prRmQty.Value// prRmQty.Value// to debug
    ;
                                            /// UPDATE[2022.01.11]:
                                            /// This is the terminal end of the
                                            /// Do..Loop Until block noted above

                                            /// REV[2023.02.22.1325]
                                            /// add error trap code
                                            /// and precheck for equality
                                            /// to reduce error potential here

                                            Information.Err.Clear();
                                            /// REV[2023.04.21.1517] on RMQTY
                                            /// added rounding of raw material
                                            /// quantity to four digits
                                            /// immediately ahead of assignment
                                            /// to RMQTY property to ensure
                                            /// the assigned value IS rounded
                                            qtRawMatl = Round(qtRawMatl, 4);
                                            if (prRmQty.Value != qtRawMatl)
                                                prRmQty.Value = qtRawMatl;

                                            if (Information.Err.Number)
                                            {
                                                System.Diagnostics.Debugger.Break();
                                                if (false)
                                                    /// REV[2023.05.10.1650]
                                                    /// attempt to work around problems
                                                    /// with iPart members by going to
                                                    /// the factory table cell itself.
                                                    /// '
                                                    /// doomed to fail if that column
                                                    /// isn't actually in the table...
                                                    /// '
                                                    /// not that it's likely
                                                    /// to succeed anyway.
                                                    /// '
                                                    aiDocPart(prRmQty.Parent.Parent.Parent).ComponentDefinition.iPartMember.Row.Item(dcIPartTbCols(aiDocPart(prRmQty.Parent.Parent.Parent).ComponentDefinition.iPartMember.ParentFactory.TableColumns).Item(pnRmQty)).Value = qtRawMatl;
                                            }


                                            rt = dcAddProp(prRmQty, rt);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                                        }
                                        else
                                        {
                                            /// REV[2022.09.20.1038]
                                            /// added step to notify user of situation,
                                            /// and offer opportunity to collect part
                                            /// and material numbers for later review.
                                            // Debug.Print "== DONTKNOW =="
                                            // Debug.Print "Item: " & pnModel
                                            // Debug.Print "Matl: " & pnStock
                                            Debug.Print(nu_FmGetList().AskUser(Join(Array("Unable to process", "the current Item.", "", "Copy the following ", "for later reference:", "", "Item: " + pnModel, "Matl: " + pnStock, ""), Constants.vbNewLine)));

                                            /// REV[2022.09.20.1042]
                                            /// in conjunction with REV[2022.09.20.1038]
                                            /// (above), disabled the following breakpoint,
                                            /// as the new User notification effectively
                                            /// supplants it
                                            // Stop 'because we don't know WHAT to do with it
                                            // and we do NOT want to clear anything
                                            // until we know what's going on!

                                            nmFamily = "";
                                            qtUnit = ""; // may want function here
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
                        /// REV[2022.05.04.1501] (lines 1400-1406)
                        if (withBlock.DocumentInterests.HasInterest(guidPipingSgmt))
                        {
                            // Stop
                            ck = newFmTest2().AskAbout(invDoc, "", Join(Array("", "appears to be Hose or Tubing,", "presently " + Interaction.IIf(Strings.Len(pnStock) > 0, pnStock, "unset") + ".", "", "Would you like to " + Interaction.IIf(Strings.Len(pnStock) > 0, "change", "set") + " it?"), Constants.vbNewLine));
                            // (removed lines 1425-1431)
                            if (ck == Constants.vbCancel)
                                System.Diagnostics.Debugger.Break();
                            else if (ck == Constants.vbYes)
                            {
                                // (removed lines 1435-1440)
                                pnStock = userChoiceFromDc(dcFrom2Fields(cnGnsDoyle().Execute(sqlOf_GnsTubeHose(withBlock.ComponentDefinition.Parameters.Item("Size_Designation").Value)), "Description", "Item"), pnStock);
                                qtUnit = Trim(UCase(aiPropsUser.Item("ROPL").Value));
                                qtRawMatl = Round(Val(Split(qtUnit + " ", " ")(0)), 4);
                                qtUnit = Split(qtUnit + " ", " ")(1);

                                ck = newFmTest2().AskAbout(invDoc, Join(Array("Stock Quantity of ", qtRawMatl + qtUnit, "selected for Item "), Constants.vbNewLine), Join(Array("If this is okay, click [YES]", "(CANCEL to debug)"), Constants.vbNewLine));
                                if (ck == Constants.vbCancel)
                                {
                                }
                                else if (ck == Constants.vbYes)
                                {
                                    prRawMatl.Value = pnStock;
                                    prRmQty.Value = qtRawMatl;
                                    prRmUnit.Value = qtUnit;
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                }
                                else
                                    System.Diagnostics.Debugger.Break();
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }
                        }
                        /// REV[2022.05.04.1501] ENDS HERE
                        /// NOTE[2022.05.04.1638] (lines 1477-1487)
                        /// [2018.07.31 by AT] (lines 1488-1498)
                        {
                            var withBlock1 = newFmTest1();
                            if (!(invDoc.ComponentDefinition.Document == invDoc))
                                System.Diagnostics.Debugger.Break();

                            /// [2018.07.31 by AT] (lines 1502-1511)
                            bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                            // End With

                            ck = withBlock1.AskAbout(invDoc, "Please Select Stock for Machined Part" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                            if (ck == Constants.vbYes)
                            {
                                /// UPDATE[2018.05.30]: (lines 1523-1525)
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
                            else
                                Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
                        }
                        // (removed lines 1547-1557) WAYPOINT ALERT!
                        if (Strings.Len(pnStock) > 0)
                        {
                            // do we look for a Raw Material Family!
                            // (removed lines 1560-1563)
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            // With cnGnsDoyle().Execute("select Family from vgMfiItems where Item='" &Replace(pnStock, "'", "''") & "';")
                            // Replace(pnStock, "'", "''")
                            /// REV[2022.08.19.1416]
                            /// temporarily replacing direct use
                            /// of pnStock with Replace operation
                            /// on single quotes in string
                            /// '
                            /// have already noted need for a 'handler',
                            /// or 'preprocessor' to prepare values
                            /// for SQL to avoid errors.
                            /// see REV[2022.08.19.1359]
                            /// '
                            {
                                var withBlock1 = cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", pnStock));
                                /// REV[2022.08.26.1104]
                                /// re 'handler' per REVS[
                                /// 2022.08.19.1359
                                /// 2022.08.19.1416
                                /// ]
                                /// new calls to sqlOf_simpleSelWhere
                                /// added in disabled (commented) form
                                /// to ultimately replace use of "hard
                                /// coded" SQL statements nearby.
                                /// '
                                /// search this Function for sqlOf_simpleSelWhere
                                /// to locate other instances of REV
                                /// '
                                /// new function sqlOf_simpleSelWhere
                                /// automatically escapes single quotes
                                /// in any String values supplied for
                                /// matching, eliminating the need for
                                /// this in the calling procedure.
                                /// '
                                if (withBlock1.BOF | withBlock1.EOF)
                                    System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                else
                                {
                                    {
                                        var withBlock2 = withBlock1.Fields;
                                        mtFamily = withBlock2.Item("Family").Value;
                                    }
                                    ;/* Cannot convert MultiLineIfBlockSyntax, System.NotSupportedException: LikeExpression not supported!
   at ICSharpCode.CodeConverter.CSharp.SyntaxKindExtensions.ConvertToken(SyntaxKind t, TokenContext context) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/SyntaxKindExtensions.cs:line 278
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.NodesVisitor.VisitBinaryExpression(BinaryExpressionSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/NodesVisitor.cs:line 1415
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingNodesVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingNodesVisitor.cs:line 28
   at ICSharpCode.CodeConverter.CSharp.VisualBasicConverter.MethodBodyVisitor.VisitMultiLineIfBlock(MultiLineIfBlockSyntax node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/MethodBodyVisitor.cs:line 353
   at Microsoft.CodeAnalysis.VisualBasic.VisualBasicSyntaxVisitor`1.Visit(SyntaxNode node)
   at ICSharpCode.CodeConverter.CSharp.CommentConvertingMethodBodyVisitor.DefaultVisit(SyntaxNode node) in /home/runner/work/CodeConverter/CodeConverter/.temp/codeconverter/ICSharpCode.CodeConverter/CSharp/CommentConvertingMethodBodyVisitor.cs:line 27

Input: 
                                    
                                    ''' UPDATE[2022.04.29.0852]
                                    ''' replicating code from UPDATE[2021.06.18]
                                    ''' above, noting also REV[2022.04.15.1035]
                                    If mtFamily Like "?-MT*" Then
                                        'Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        'Debug.Print pnModel & "[" & prRmQty.Value & qtUnit & "*" & pnStock & ": " & aiPropsDesign(pnDesc).Value & "]" ' prRawMatl.Value
                                        Debug.Print pnModel & "["& qtRawMatl & qtUnit& " of " & pnStock & ": "& aiPropsDesign(pnDesc).Value& "]" 'prRmQty.Value prRawMatl.Value
                                        Stop 'FULL Stop!
                                    ElseIf mtFamily Like "?-PT*" Then
                                    ''' REV[2022.05.05.1343] (lines 1626-1639)
                                        If nmFamily Like "?-RM*" Then
                                            'ck = vbNo
                                            Debug.Print ; 'Breakpoint Landing
                                        Else
                                            ck = MsgBox(Join(Array("Part " & pnModel & " uses " & pnStock,"which is not sheet metal.","","These parts are usually assigned","to the Riverview family, R-RMT.","","Do you want to use this Family?","Click [NO] to see other options.","(CANCEL to debug)"), vbNewLine),vbYesNoCancel + vbQuestion,"Select Part Family?")
                                            ' (removed lines 1658-1659)
                                            Debug.Print ; 'Breakpoint Landing
                                            If ck = vbCancel Then
                                                Stop 'to debug. (developers only!)
                                            ElseIf ck = vbYes Then
                                                nmFamily = "R-RMT"
                                            Else
                                                If Len(nmFamily) = 0 Then
                                                    nmFamily = "R-RMT"
                                                End If
                                                
                                                With nuDcPopulator().Setting("D-RMT", "Doyle (typ. sheet metal)").Setting("R-RMT", "Riverview (most others)")
                                                    If Not .Exists(nmFamily) Then
                                                        .Setting nmFamily, "Current (" & nmFamily & ")"
                                                    End If
                                                    
                                                    nmFamily = userChoiceFromDc(dcTransposed(.Dictionary()),nmFamily)
                                                End With
                                            End If
                                        End If
                                        
                                        mtFamily = "D-BAR"
                                    ElseIf mtFamily = "D-PTS" Then
                                        mtFamily = "D-BAR"
                                        Stop 'NOT SO FAST!
                                        'nmFamily = "D-RMT"
                                    ElseIf mtFamily = "R-PTS" Then
                                        mtFamily = "D-BAR"
                                        Stop 'NOT SO FAST!
                                        'nmFamily = "R-RMT"
                                    End If

 */
                                }
                            }
                            // (removed lines 1722-1725) WAYPOINT ALERT!
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
                                /// UPDATE[2022.01.11]: (lines 1738-1745)
                                nmFamily = "R-RMT";
                                qtUnit = prRmUnit.Value; // "IN"
                                ck = Constants.vbCancel;
                                do
                                {
                                    /// UPDATE[2021.03.11] (lines 1750-1776)
                                    if (true)
                                    {
                                        Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                        /// REV[2022.02.09.0904] (lines 1779-1783)
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

                                    /// REV[2022.02.08.1446] (lines 1799-1803)
                                    Debug.Print("qtRawMatl = ");
                                    if (dcIn.Exists(pnRmQty))
                                        Debug.Print("In Genius: ");
                                    Debug.Print();
                                    Debug.Print("qtUnit = \"");
                                    if (dcIn.Exists(pnRmUnit))
                                        Debug.Print("In Genius: ");
                                    Debug.Print(" ( or try IN )");

                                    /// REV[2022.02.08.1525] (lines 1811-1825)
                                    Debug.Print("");
                                    /// REV[2022.03.11.1112] (lines 1827-1830)
                                    {
                                        var withBlock1 = nu_fmIfcMatlQty01().SeeUser(invDoc);
                                        if (withBlock1.Exists(pnRmQty))
                                        {
                                            /// REV[2022.04.04.1404] (lines 1833-1839)
                                            if (System.Convert.ToDouble("0" + System.Convert.ToHexString(qtRawMatl)) == System.Convert.ToDouble(withBlock1.Item(pnRmQty)))
                                            {
                                            }
                                            else
                                            {
                                                // Debug.Print "prRmQty.Value FROM " & prRmQty.Value & " TO " & .Item(pnRmQty)
                                                Debug.Print("qtRawMatl FROM " + qtRawMatl + " TO " + withBlock1.Item(pnRmQty));

                                                // Stop 'and double-check
                                                // might still be equivalent
                                                qtRawMatl = withBlock1.Item(pnRmQty);
                                            }
                                        }
                                        else
                                            System.Diagnostics.Debugger.Break();
                                        if (withBlock1.Exists(pnRmUnit))
                                        {
                                            if (qtUnit == withBlock1.Item(pnRmUnit))
                                            {
                                            }
                                            else
                                            {
                                                Debug.Print("qtUnit FROM " + qtUnit + " TO " + withBlock1.Item(pnRmUnit));
                                                // Stop 'and double-check
                                                // might still be equivalent
                                                qtUnit = withBlock1.Item(pnRmUnit);
                                            }
                                        }
                                        else
                                            System.Diagnostics.Debugger.Break();
                                    }
                                    // (removed lines 1868-1874)
                                    /// REV[2022.10.17.1504] CANCELED
                                    /// disabled following confirmation
                                    /// UserForm prompt as redundant
                                    /// REV[2022.10.17.1511] undid 1504
                                    /// this prompt might not be quite
                                    /// so redundant as presumed
                                    Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                    ck = newFmTest2().AskAbout(invDoc, "Raw Material Quantity is now " + System.Convert.ToHexString(qtRawMatl) + qtUnit + " for", "If this is okay, click [YES]." + Constants.vbNewLine + "Otherwise, click [NO] to review." + Constants.vbNewLine + "" + Constants.vbNewLine + "( for debug, click [CANCEL] )");
                                    if (ck == Constants.vbCancel)
                                        System.Diagnostics.Debugger.Break();
                                }
                                while (!ck == Constants.vbYes)/* TODO ERROR: Skipped SkippedTokensTrivia */// .Result()// prRmQty.Value// prRmQty.Value// to debug.
;
                                /// UPDATE[2022.01.11]:
                                /// This is the terminal end of the
                                /// Do..Loop Until block noted above

                                prRmQty.Value = qtRawMatl;
                                rt = dcAddProp(prRmQty, rt);
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                            }
                            else
                            {
                                System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                                /// REV[2022.04.29.0755]
                                /// moved Stop AHEAD of the following assignments to
                                /// avoid clearing any potentially essential values.
                                nmFamily = "";
                                qtUnit = ""; // may want function here
                            }
                        }
                        else if (0)
                            System.Diagnostics.Debugger.Break();// and regroup
                    } // Sheetmetal vs Part
                      // --------------------------------------------'
                    /// REV[2022.05.05.1257] (lines 1914-1922)
                    if (Strings.Len(pnStock) > 0)
                    {
                        {
                            var withBlock1 = prRawMatl;
                            if (Len(Trim(withBlock1.Value)) > 0)
                            {
                                if (pnStock != withBlock1.Value)
                                {
                                    // (removed comment lines 1927-1931)
                                    ck = MsgBox(Join(Array("Raw Stock Change Suggested", "  for Item " + pnModel, "", "  Current : " + prRawMatl.Value, "  Proposed: " + pnStock, "", "Change It?", ""), Constants.vbNewLine), Constants.vbYesNo, pnModel + " Stock");
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
                    // rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
                    /// Plan to remove commented line above,
                    /// superceded by the one above that
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Another landing line
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
                    /// REV[2022.01.17.1135] (lines 1999-2004)
                    ck = newFmTest2().AskAbout(invDoc, "For some reason, THIS Item is marked Phantom:", "Is this okay? (Click [NO] OR [CANCEL] if not)");
                    if (ck == Constants.vbYes)
                    {
                    }
                    else
                        System.Diagnostics.Debugger.Break();
                }
                else if (bomStruct == kInseparableBOMStructure)
                {
                    // How the HECK does a PART get marked Inseparable?!?
                    ck = newFmTest2().AskAbout(invDoc, "This Item is marked Inseperable:", Join(Array("This is likely not correct,", "and should be fixed ASAP.", "Would you like to copy the Part", "Number for later review?", "", Constants.vbNewLine + Constants.vbNewLine + "([CANCEL] to debug)"), " "));
                    if (ck == Constants.vbYes)
                        // InputBox Join(Array("Copy this Part Number and paste it into another document or memo for review later."), vbNewLine), "Copy Part Number " & pnModel, pnModel
                        InputBox(Join(Array("Copy this Part Number, and paste", "it into another document or memo", "for later review."), Constants.vbNewLine), "Copy Part Number " + pnModel, pnModel);
                    else if (ck == Constants.vbCancel)
                        System.Diagnostics.Debugger.Break();// to debug. (developers only)

                    System.Diagnostics.Debugger.Break(); // really, just STOP!
                }
                else
                {
                    /// REV[2022.01.17.1138] (lines 2027-2032)
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
                    /// REV[2022.04.15.1044]
                    /// add check against current value.
                    /// why try to fix what ain't broken?
                    if (prFamily.Value != nmFamily)
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
                    }
                    rt = dcAddProp(prFamily, rt);
                }
            }

            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            dcGeniusPropsPartRev20180530 = rt;
        }
    }

    public Scripting.Dictionary dcAiDocComponents(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long incTop = 0, long inclPhantom = 0)
    {
        Scripting.Dictionary rt;

        if (dc == null)
            rt = dcAiDocComponents(AiDoc, new Scripting.Dictionary(), incTop, inclPhantom);
        else
            rt = dc;

        if (AiDoc == null)
        {
        }
        else if (AiDoc.DocumentType == kAssemblyDocumentObject)
        {
            if (incTop)
            {
                {
                    var withBlock = rt;
                    if (withBlock.Exists(AiDoc.FullFileName))
                    {
                        if (withBlock.Item(AiDoc.FullFileName) == AiDoc)
                        {
                        }
                        else
                            System.Diagnostics.Debugger.Break();// because somethin' ain't right.
                    }
                    else
                        withBlock.Add(AiDoc.FullFileName, AiDoc);
                }
            }
            /// NOTE[2023.01.27.1207]
            /// not sure following call is correct.
            /// it's a self-referential call passing
            /// the received 'include phantom' flag
            /// as the 'include top' argument, with
            /// no clear reason why.
            /// 
            /// no, wait; this is NOT a self-referential call
            /// 
            rt = dcAssyDocComponents(AiDoc, rt, inclPhantom);
        }
        else if (AiDoc.DocumentType == kPartDocumentObject)
        {
            /// REV[2022.04.12.1130]
            /// add guard code to catch key collision
            /// and check if matching Item is already
            /// filed under that key. If not, manual
            /// intervention may be required
            if (rt.Exists(AiDoc.FullFileName))
            {
                if (rt.Item(AiDoc.FullFileName) == AiDoc)
                {
                }
                else
                    System.Diagnostics.Debugger.Break();// because somethin' just ain't right.
            }
            else
                rt.Add(AiDoc.FullFileName, AiDoc);
        }
        else
        {
        }

        dcAiDocComponents = rt;
    }

    public Scripting.Dictionary dcAssyDocComponents(Inventor.AssemblyDocument Assy, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long inclPhantom = 0)
    {
        dcAssyDocComponents = dcAssyCompAndSub(Assy.ComponentDefinition.Occurrences, dc, inclPhantom);
    }

    public Scripting.Dictionary dcAssyCompAndSub(Inventor.ComponentOccurrences Occurences, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */, long inclPhantom = 0)
    {
        /// Traverse the assembly,
        /// including any/all subassemblies,
        /// and collect all parts to be processed.
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence invOcc;
        Inventor.ObjectTypeEnum tp;
        Inventor.ComponentDefinition ocDef;

        if (dc == null)
            rt = dcAssyCompAndSub(Occurences, new Scripting.Dictionary(), inclPhantom);
        else
        {
            rt = dc;
            foreach (var invOcc in Occurences)
            {
                {
                    var withBlock = compOccFromProxy(invOcc) // (instead of just invOcc) [AT:2018.09.28]
      ;
                    if (withBlock._IsSimulationOccurrence)
                        System.Diagnostics.Debugger.Break();
                    else
                   /// !!!WARNING!!!'''
                   // '  The latest modification above
                   // '  attempts to get around an issue with
                   // '  ComponentOccurrenceProxy Occurences.
                   // '  These seem to fail on attempts
                   // '  to retrieve their Definition,
                   // '  and its associated Document.
                   // '
                   // '  It is hoped the ContainingOccurrence
                   // '  will supply the correct objects.
                   // '  However, we DO NOT KNOW if this
                   // '  is what we actually get.
                   // '
                   // '  Function compOccFromProxy includes a Stop
                   // '  that occurs whenever a ComponentOccurrenceProxy
                   // '  is discovered. In these instances, the process
                   // '  should be carefully stepped through and traced
                   // '  for any indication of the actual relationship
                   // '  between a ComponentOccurrenceProxy
                   // '  and its ContainingOccurrence.

                   // Remove suppressed and excluded parts from the process
                   // Moved out here from inner checks
                   if (withBlock.Visible & !withBlock.Suppressed & !withBlock.Excluded)
                    {
                        /// UPDATE[2018.08.20,AT]
                        /// Error encountered on line noted.
                        /// Adding Error trap with code to try alternative

                        /// when stopped under REV[2023.01.27.1329] below
                        /// set next statement (Ctrl-F9) at
                        /// and continue (F5)

                        ocDef = withBlock.Definition;
                        if (Information.Err.Number != 0)
                        {
                            System.Diagnostics.Debugger.Break();
                            if (withBlock._IsSimulationOccurrence)
                                // Err.Clear
                                // ocDef = .ContainingOccurrence.Definition
                                // If Err.Number <> 0 Then
                                // Stop
                                // End If
                                System.Diagnostics.Debugger.Break();
                            else
                                System.Diagnostics.Debugger.Break();
                        }


                        if (ocDef == null)
                            System.Diagnostics.Debugger.Break();
                        else
                        {
                            // ''''
                            // tp = .ContextDefinition.Type
                            tp = ocDef.Type;

                            if (tp != kAssemblyComponentDefinitionObjectAnd) /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                            {
                                // (moved suppression/exclusion check OUTSIDE)
                                if (tp != kWeldsComponentDefinitionObject)

                                    // rt = dcAddAiDoc(aiDocument(ocDef.Document), rt)
                                    /// Recasting by aiDocument not likely necessary here.
                                    /// Revised to following:
                                    rt = dcAddAiDoc(ocDef.Document, rt); // inVisible, suppressed, excluded or Welds
                            }
                            else if (withBlock.BOMStructure == kPurchasedBOMStructure)
                                // Just add it to the Dictionary
                                rt = dcAddAiDoc(ocDef.Document, rt);
                            else if (withBlock.BOMStructure == kNormalBOMStructure)
                                // Gather its components
                                rt = dcAssyCompAndSub(withBlock.SubOccurrences, dcAddAiDoc(ocDef.Document, rt), inclPhantom); // NOT forgetting to add THIS document!
                            else if (withBlock.BOMStructure == kInseparableBOMStructure)
                            {
                                if (tp == kWeldmentComponentDefinitionObject)
                                    // Treat it like an assembly
                                    rt = dcAssyCompAndSub(withBlock.SubOccurrences, dcAddAiDoc(ocDef.Document, rt), inclPhantom);
                                else if (tp == kAssemblyComponentDefinitionObject)
                                    // just an ordinary Assembly.
                                    // Same handling as above,
                                    // but use own branch, just in case.
                                    rt = dcAssyCompAndSub(withBlock.SubOccurrences, dcAddAiDoc(ocDef.Document, rt), inclPhantom);
                                else
                                    System.Diagnostics.Debugger.Break();// and see if we can figure out what its type is
                            }
                            else if (withBlock.BOMStructure == kPhantomBOMStructure)
                            {
                                if ((inclPhantom & 1) == 1)
                                    // the Document as well as its components
                                    // (this is mainly for debugging/development)
                                    rt = dcAssyCompAndSub(withBlock.SubOccurrences, dcAddAiDoc(ocDef.Document, rt), inclPhantom);
                                else
                                    // Gather its components, but NOT the document itself
                                    rt = dcAssyCompAndSub(withBlock.SubOccurrences, rt, inclPhantom);
                            }
                            else if (withBlock.BOMStructure == kReferenceBOMStructure)
                                Debug.Print(newFmTest2.AskAbout(ocDef.Document, "Reference Component will not be processed.", "Click any button to acknowledge and continue."));
                            else
                            {
                                Debug.Print(newFmTest2.AskAbout(ocDef.Document, "Unhandled Condition on this component.", "Going to Debug -- Click any button."));
                                System.Diagnostics.Debugger.Break(); // and have a look at it
                            } // part or assembly
                        }
                    }
                    else
                        /// REV[2023.01.27.1329]
                        /// 
                        /// add Else branch to deal with missed items
                        /// under certain circumstances, for example,
                        /// when instances are brought from an iAssembly
                        /// (or iPart) factory in a Suppresed or Excluded state
                        /// 
                        /// USAGE: set breakpoint on following line when needed
                        Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// and then, on Break here, search within // (SimulationOccurrence)
                }
                ocDef = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }
        dcAssyCompAndSub = rt;
    }

    public Scripting.Dictionary dcAddAiDoc(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        string fn;

        if (dc == null)
            dcAddAiDoc = dcAddAiDoc(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;
            fn = AiDoc.FullFileName;
            {
                var withBlock = rt;
                if (!withBlock.Exists(fn))
                    withBlock.Add(fn, AiDoc);
            }
            dcAddAiDoc = rt;
        }
    }

    public Scripting.Dictionary dcPropVals(Scripting.Dictionary dc, long nonProp = 0)
    {
        /// dcPropVals -- Extract Values from
        /// Properties in supplied Dictionary
        /// non Property Items are processed
        /// according to nonProp:
        /// 0 - Key/Item NOT added
        /// 1 - Key/Item added as is
        /// >1 - Key/"blank" added
        /// 
        /// NOTE: similar functions may be
        /// due for deprecation:
        /// dcAiPropValsFromDc
        /// dcMapAiProps2vals
        /// '
        /// 
        Scripting.Dictionary rt;
        Inventor.Property pr;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcNewIfNone(dc);
            foreach (var ky in withBlock.Keys)
            {
                pr = aiProperty(obOf(withBlock.Item(ky)));
                if (pr == null)
                {
                    if (nonProp > 0)
                    {
                        if (nonProp > 1)
                            // Generate "blank" Item
                            rt.Add(ky, noVal(VarType(withBlock.Item(ky))));
                        else
                            // Keep non-Property Item
                            rt.Add(ky, withBlock.Item(ky));
                    }
                }
                else
                    rt.Add(ky, pr.Value);
            }
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        dcPropVals = rt;
    }

    public Scripting.Dictionary dcGeniusProps(Inventor.Document invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        {
            var withBlock = invDoc;
            /// 
            Debug.Print("== Item " + withBlock.PropertySets(gnDesign).Item(pnPartNum).Value + " <" + withBlock.FullDocumentName + ">");
            if (withBlock.DocumentType == kPartDocumentObject)
                dcGeniusProps = dcGeniusPropsPart(invDoc, dc);
            else if (withBlock.DocumentType == kAssemblyDocumentObject)
                dcGeniusProps = dcGeniusPropsAssy(invDoc, dc);
            else
                System.Diagnostics.Debugger.Break();// cuz we don't know WHAT to do with it
        }
    }

    public Scripting.Dictionary dcGeniusPropsAssy(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;

        Inventor.Property prPartNum; // pnPartNum
        Inventor.Property prFamily;
        string pnModel;
        string nmFamily;

        Scripting.Dictionary rt;
        VbMsgBoxResult ck;
        // Dim fm As String

        if (dc == null)
            rt = dcGeniusPropsAssy(AiDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            if (AiDoc == null)
            {
            }
            else
            {
                var withBlock = AiDoc;
                // the custom property set.
                {
                    var withBlock1 = withBlock.PropertySets;
                    aiPropsUser = withBlock1.Item(gnCustom);
                    aiPropsDesign = withBlock1.Item(gnDesign);
                }

                /// REV[2022.06.30.1546]
                /// duplicated prPartNum and prFamily
                /// from dcGeniusPropsPartRev20180530
                /// along with some related processes.
                {
                    var withBlock1 = aiPropsDesign;
                    prPartNum = withBlock1.Item(pnPartNum);
                    // aiGetProp( aiPropsDesign, pnPartNum)
                    prFamily = withBlock1.Item(pnFamily);
                }

                // For now, we'll just assume all assemblies are made here.
                // fm = "D-MTO"

                /// REV[2022.06.30.1550]
                /// replaced above assigment with the two below
                /// for more robust Family assignment for assemblies.
                /// 
                /// by checking model and Genius for an existing Family,
                /// one hopes to avoid indiscriminately overriding
                /// established Families, particularly in Genius.
                /// 
                pnModel = prPartNum.Value;

                /// REV[2023.0106.1623]
                /// due to replacement of Level of Detail
                /// with Model States, which use a different
                /// name for the default level/state, original
                /// check is demoted to a secondary check
                /// for the new default name.
                if (withBlock.LevelOfDetailName != "[Primary]")
                {
                    if (withBlock.LevelOfDetailName != "Master")
                        // If .ModelStateName <> "[Primary]" Then
                        System.Diagnostics.Debugger.Break();
                }

                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// REV[2023.0113.1624]
                    /// Family assignment moved into With statement
                    /// and modified to collect existing value, if set,
                    /// or otherwise generate one based on BOM structure,
                    /// so purchased assemblies might be identified

                    nmFamily = prFamily.Value;

                    if (Strings.Len(nmFamily) == 0)
                    {
                        if (withBlock1.BOMStructure == kPurchasedBOMStructure)
                            nmFamily = "D-PTS";
                        else
                            nmFamily = "D-MTO";
                    }

                    /// REV[2023.0113.1625]
                    /// simplified Family check against Genius
                    /// to simply use current value of nmFamily,
                    /// now it SHOULD be set to a non-blank value
                    nmFamily = famVsGenius(pnModel, nmFamily);

                    rt = dcWithProp(aiPropsDesign, pnFamily, nmFamily, rt);

                    {
                        var withBlock2 = withBlock1.MassProperties;
                        Information.Err.Clear();
                        rt = dcWithProp(aiPropsUser, pnMass, Round(withBlock2.Mass * cvMassKg2LbM, 4), rt);
                        if (Information.Err.Number)
                        {
                            Debug.Print(Join(Array("NOMASS", AiDoc.FullFileName), ":"));
                            ck = MsgBox(Join(Array("" + "An Error occurred while collecting", "or updating Mass Property information", "for " + AiDoc.DisplayName + ".", "", "Click [Cancel] to enter debug mode", "and attempt to review and correct.", "", "Otherwise click [OK] to continue.", "(Mass will probably be incorrect)"), Constants.vbNewLine), Constants.vbOKCancel, "ERROR(" + AiDoc.DisplayName + ")!");
                            if (ck == Constants.vbCancel)
                                System.Diagnostics.Debugger.Break();
                        }
                    }
                }
            }

            iSyncAssyFactory(AiDoc); // Backport Properties to iAssembly Factory
        }

        dcGeniusPropsAssy = rt;
    }

    public Scripting.Dictionary dcGeniusPropsPart(Inventor.PartDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (dc == null)
            dcGeniusPropsPart = dcGeniusPropsPart(AiDoc, new Scripting.Dictionary());
        else if (AiDoc == null)
            dcGeniusPropsPart = dc;
        else
            dcGeniusPropsPart = dcGeniusPropsPartRev20180530(AiDoc, dc);
    }

    public Scripting.Dictionary dcFlatPatVals(Inventor.SheetMetalComponentDefinition invSheetMetalComp, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        // '
        // Dim aiPropSet As Inventor.PropertySet
        Inventor.Parameter prThickness;
        // '
        double dLength;
        double dWidth;
        double dArea;
        string strWidth;
        string strLength;
        string strArea;
        string mtFamily;
        string mtType;

        double dHeight;
        double dfHtThk;
        string strDVNs;

        double ck;
        VbMsgBoxResult gn;

        // If dc Is Nothing Then
        // dcFlatPatVals = dcFlatPatVals('        invSheetMetalComp,'        New Scripting.Dictionary'    )
        // Else
        rt = dcNewIfNone(dc);

        // Request #3:  sheet metal extent area and add to custom property "RMQTY"
        // Check to see if flat exists
        if (invSheetMetalComp == null)
        {
        }
        else
        {
            var withBlock = invSheetMetalComp;
            Information.Err.Clear();
            prThickness = withBlock.Thickness;
            if (Information.Err.Number)
            {
                if (InStr(1, aiDocument(withBlock.Document).FullFileName, @"\Doyle_Vault\Designs\purchased\"))
                    System.Diagnostics.Debugger.Break(); // we probably got a purchased part, here
                else
                    System.Diagnostics.Debugger.Break();// and look into it
            }
            else
            {
                /// anything that fails to yield a valid Thickness object
                /// probably can't be processed as a Sheet Metal component.
                /// Therefore, the REST of the process should only proceed
                /// if the retrieval succeeded.
                if (withBlock.HasFlatPattern)
                {
                    mtFamily = "DSHEET";
                    {
                        var withBlock1 = withBlock.FlatPattern;
                        if (withBlock1.Body == null)
                        {
                            /// UPDATE[2021.06.11] Implementing check for Body
                            /// to try to avoid raising an Error
                            /// by diving blind into the With block
                            /// and handle the missing Body situation
                            /// in a more appropriate fashion.
                            /// This comment supercedes [2019.12.13],
                            /// now removed to notes_2021-0611_general-01.txt
                            {
                                var withBlock2 = newFmTest2();
                                if (withBlock2.AskAbout(invSheetMetalComp.Document, Join(Array("ISSUE WITH FLAT PATTERN:", "   NO BODY FOUND"), Constants.vbNewLine), Join(Array("Please consider reviewing model,", "and rebuilding its flat pattern.", "", "Pause for review? (not necessary)"), Constants.vbNewLine)) == Constants.vbYes)
                                    // (please check part for outdated flat pattern)
                                    System.Diagnostics.Debugger.Break();
                                else
                                {
                                }
                            }

                            System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-1256
                            /// CHANGE NEEDED[2021.11.05]:
                            /// Not sure what, exactly
                            /// NOTE[2021.12.13]
                            /// These values are converted to inches
                            /// from centimeters below in ...
                            /// They should NOT be converted HERE!
                            /// (don't think so, anyway)
                            /// disabling conversion operations
                            /// pending review on debug
                            dfHtThk = withBlock1.Parameters.Item("Thickness").Value; // / cvLenIn2cm
                            dLength = withBlock1.length; // / cvLenIn2cm
                            dWidth = withBlock1.Width; // / cvLenIn2cm
                            dArea = dLength * dWidth;
                        }
                        else
                        {
                            // Stop 'BKPT-2021-1105-1250
                            /// CHANGE NEEDED[2021.11.05]:
                            /// Actually have a function
                            /// to collect X, Y, and Z spans
                            /// which is used in the main
                            /// function. Might be usable here.
                            /// '
                            /// (one might think the note would
                            /// IDENTIFY the aforementioned
                            /// function here, wouldn't one?
                            /// Well... one would be WRONG!!!)
                            /// '
                            /// Here we go: nuAiBoxData()
                            {
                                var withBlock2 = nuAiBoxData().UsingInches(0).UsingBox(withBlock1.Body.RangeBox) // .SortingDims
;
                                /// UPDATE[2021.12.13]
                                /// Changed UsingInches argument
                                /// to zero (DON'T use) because
                                /// conversion is performed below.
                                // Check height against thickness
                                // Valid flat pattern should return
                                // zero or VERY minimal difference
                                dHeight = Round(withBlock2.SpanZ, 6); // (.MaxPoint.Z - .MinPoint.Z)

                                // the extent of the face.
                                // Extract the width, length and area from the range.
                                dLength = Round(withBlock2.SpanX, 6); // (.MaxPoint.X - .MinPoint.X)
                                dWidth = Round(withBlock2.SpanY, 6); // (.MaxPoint.Y - .MinPoint.Y)
                            }

                            {
                                var withBlock2 = withBlock1.Body.RangeBox;
                                /// CHECKPOINT[2021.12.07]:
                                /// not actually stopping here
                                /// but running a quick check
                                /// to make sure revised code
                                /// above works correctly.
                                /// If so, this section SHOULD
                                /// be good to remove or disable
                                ck = 0;
                                ck = ck + Abs(withBlock2.MaxPoint.Z - withBlock2.MinPoint.Z - dHeight); // * cvLenIn2cm
                                ck = ck + Abs(withBlock2.MaxPoint.X - withBlock2.MinPoint.X - dLength); // * cvLenIn2cm
                                ck = ck + Abs(withBlock2.MaxPoint.Y - withBlock2.MinPoint.Y - dWidth); // * cvLenIn2cm
                                /// UPDATE[2021.12.13]
                                /// conversion operations removed
                                /// since no longer converting
                                /// before end stage
                                if (Round(ck, 5) > 0)
                                    System.Diagnostics.Debugger.Break(); // BKPT-2021-1207-1158
                            }

                            /// UPDATE[2021.11.05]:
                            /// Moved derived calculations
                            /// outside of With block above.
                            /// Might or might not prove useful.
                            dfHtThk = Round(dHeight - prThickness.Value, 6); // / cvLenIn2cm
                            /// UPDATE[2021.12.13]
                            /// conversion operation removed since
                            /// no longer converting before end stage
                            /// UPDATE[2022.01.28.1512]
                            /// moved rounding operation to
                            /// end stage, AFTER conversion
                            dArea = dLength * dWidth; // Round(, 6)

                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }

                        if (dArea == 0)
                            // [2021.06.11] Moved alternate calculation
                            // sequence to new If-Then block preceding
                            // this one, and removed previous comment
                            // as no longer relevant.
                            System.Diagnostics.Debugger.Break();// and note when this branch taken

                        if (dArea > 0)
                            /// an invalid flat pattern SHOULD have no geometry,
                            /// which means it SHOULD have no area to speak of.
                            /// '
                            /// One would think this obvious, in retrospect,
                            /// but one would not be surprised to be proven wrong.
                            /// Again.
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        else
                        {
                            if (MsgBox(Join(Array("The flat pattern for this", "part has no features,", "and is likely not valid.", "", "Pause here to review?", "(Click 'NO' to just keep going)"), Constants.vbNewLine), Constants.vbYesNo, "Invalid Flat Pattern") == Constants.vbYes)
                                System.Diagnostics.Debugger.Break();// and let the user look into it
                            Debug.Print(aiDocument(withBlock1.Document).FullDocumentName);
                            System.Diagnostics.Debugger.Break();
                            mtFamily = "D-BAR";
                        }
                    }

                    // '''''
                    // '''''  The following section should be moved OUTSIDE this branch!
                    // '''''
                    {
                        var withBlock1 = aiDocPart(withBlock.Document);
                        // aiPropSet = .PropertySets.Item(gnCustom)
                        // prOffThknss

                        // Convert values into document units.
                        // This will result in strings that are identical
                        // to the strings shown in the Extent dialog.
                        /// 
                        /// NOTE[2021.11.09]
                        /// If UsingInches is set as shown above,
                        /// this section might not work properly.
                        /// Might be better to NOT use inches,
                        /// and simply let things take care of
                        /// themselves, here.
                        /// UPDATE[2021.12.13]
                        /// Changed UsingInches argument to zero
                        /// in two calls above.
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

                    // Stop 'BKPT-2021-1105-1304
                    /// CHANGE NEEDED[2021.11.05]:
                    /// This is where Properties are set.
                    /// Want to change this to simply collect
                    /// the generated values, and pass them
                    /// back to the client for processing.
                    /// '
                    /// A separate process can then
                    /// assign them to Properties.
                    /// '
                    // Add area to custom property set
                    // rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt)
                    rt.Add(pnRmQty, Round(dArea * cvArSqCm2SqFt, 4));
                    // 
                    // 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
                    // 
                    // /  1ft | 1in    \2     2                2
                    // ( ------+-------- ) * cm  = 0.00107639 ft
                    // \ 12in | 2.54cm /
                    // 
                    // That value really needs to go into a constant
                    // and so it HAS: cvArSqCm2SqFt (noted 2022.01.28)
                    /// REV[2022.01.28.1516]
                    /// add Raw Material Unit Quantity to output
                    rt.Add(pnRmUnit, "FT2");
                    /// Thickness, too:

                    // Add Thickness to returned values
                    rt.Add(pnThickness, withBlock.Thickness.Value / (double)cvLenIn2cm);

                    // Add Width to custom property set
                    // rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt)
                    rt.Add(pnWidth, strWidth);

                    // Add Length to custom property set
                    // rt = dcWithProp(aiPropSet, pnLength, strLength, rt)
                    rt.Add(pnLength, strLength);

                    // Add AreaDescription to custom property set
                    // rt = dcWithProp(aiPropSet, pnArea, strArea, rt)
                    rt.Add(pnArea, strArea);
                    if (Strings.Len(strDVNs) > 0)
                        rt.Add("OFFTHK", strDVNs);

                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                }
                else
                    mtFamily = "D-BAR";

                rt.Add("mtFamily", mtFamily);
            }
        }
        dcFlatPatVals = rt;
    }

    public Scripting.Dictionary dcFlatPatProps(Inventor.SheetMetalComponentDefinition invSheetMetalComp, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        // '
        Inventor.PropertySet aiPropSet;
        Inventor.Parameter prThickness;
        Inventor.Parameter prOffThknss;
        // '
        double dLength;
        double dWidth;
        double dArea;
        string strWidth;
        string strLength;
        string strArea;

        double dHeight;
        double dfHtThk;
        string strDVNs;

        double ck;
        VbMsgBoxResult gn;

        if (dc == null)
            dcFlatPatProps = dcFlatPatProps(invSheetMetalComp, new Scripting.Dictionary());
        else
        {
            rt = dc;

            // Request #3:  sheet metal extent area and add to custom property "RMQTY"
            // Check to see if flat exists
            if (invSheetMetalComp == null)
                /// 
                System.Diagnostics.Debugger.Break();
            else
            {
                var withBlock = invSheetMetalComp;
                Information.Err.Clear();
                prThickness = withBlock.Thickness;
                if (Information.Err.Number)
                {
                    if (InStr(1, aiDocument(withBlock.Document).FullFileName, @"\Doyle_Vault\Designs\purchased\"))
                    {
                        System.Diagnostics.Debugger.Break(); // we probably got a purchased part, here
                        /// NOTE[2018-05-31]: Don't recall hitting this stop recently,
                        /// likely because parts matching this file path
                        /// are getting caught in the calling function now.
                        /// May be cable to remove this section,
                        /// but retain for now pending further notice.
                        /// '
                        /// CHANGE NEEDED[2021.11.05]:
                        /// do NOT want to go changing ANYTHING
                        /// in the model inside THIS function,
                        /// unless it's unavoidable; for example,
                        /// if a flat pattern is not available.
                        withBlock.BOMStructure = kPurchasedBOMStructure;
                    }
                    else
                        System.Diagnostics.Debugger.Break();// and look into it
                }
                else
                {
                    /// anything that fails to yield a valid Thickness object
                    /// probably can't be processed as a Sheet Metal component.
                    /// Therefore, the REST of the process should only proceed
                    /// if the retrieval succeeded.
                    if (!withBlock.HasFlatPattern)
                    {
                        System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-1213
                        /// CHANGE NEEDED[2021.11.05]:
                        /// A new Flat Pattern should NOT
                        /// be generated in this section!
                        /// This should be handled in a new
                        /// check function to determine
                        /// whether identified Sheet Metal
                        /// Part is in fact sheet metal
                        /// UPDATE[2018.02.06]: New UserForm Available!
                        /// Modify code here to use UserForm fmTest2
                        /// to prompt user with part image and data
                        /// while asking about Flat Pattern
                        /// UPDATE[2021.12.13]: Bug Fix (hopefully)
                        /// Bug encountered, wherein a "No"
                        /// response is mistaken for a "Yes".
                        /// Exact cause unclear, but suspect
                        /// issue with terminated UserForm
                        /// leaving behind undefined result.
                        /// 
                        /// This fix is intended to keep
                        /// UserForm fmTest2 active while
                        /// retrieving result, and thereby
                        /// ensure correct result received.
                        /// 
                        /// Note that it depends on a new method
                        /// function added to UserForm fmTest2:
                        /// Using takes a supplied Inventor
                        /// Document and locks it in for use
                        /// on the next call to AskAbout, for
                        /// which the Document is now Optional.
                        /// 
                        {
                            var withBlock1 = newFmTest2().Using(withBlock.Document);
                            gn = withBlock1.AskAbout(null/* Conversion error: Set to default value for this argument */, null/* Conversion error: Set to default value for this argument */, "NO FLAT PATTERN!" + Constants.vbNewLine + "Try to generate one?");
                        }
                        if (gn == Constants.vbYes)
                        {
                            /// UPDATE[2018-05-31]: Removing comment-disabled
                            /// legacy code from switch to new UserForm above.
                            /// Successful use since update noted above
                            /// indicates no further need.

                            Information.Err.Clear();
                            withBlock.Unfold();
                            if (Information.Err.Number)
                                Debug.Print(Join(Array("FPFAIL", aiDocument(withBlock.Document).FullFileName), ":"));
                            else if (withBlock.HasFlatPattern)
                                withBlock.FlatPattern.ExitEdit();


                            // We'll want to do something else instead of the following
                            // to make sure any document openened by the Unfold operation
                            // gets closed before we move on.
                            {
                                var withBlock1 = ThisApplication.Documents.VisibleDocuments;
                                if (withBlock1.Item(withBlock1.Count) == invSheetMetalComp.Document)
                                    withBlock1.Item(withBlock1.Count).Close(true);
                            }
                        }
                        else
                        {
                        }
                    }

                    if (withBlock.HasFlatPattern)
                    {
                        {
                            var withBlock1 = withBlock.FlatPattern;
                            // First, make sure it's VALID
                            // If .Features.Count > 0 Then 'should be good? NOPE!!!
                            /// Turns out, most flat patterns don't HAVE features.
                            /// Not sure how they work, but they're not typical elements.

                            // If .BaseFace Is Nothing Then 'This is an alternate test
                            // Keep on hand in case primary doesn't work out.
                            // Changeover will require switching Then and Else blocks.

                            if (withBlock1.Body == null)
                            {
                                // [2021.06.11] Implementing check for Body
                                // to try to avoid raising an Error
                                // by diving blind into the With block
                                // and handle the missing Body situation
                                // in a more appropriate fashion.
                                // This comment supercedes [2019.12.13],
                                // now removed to notes_2021-0611_general-01.txt
                                {
                                    var withBlock2 = newFmTest2();
                                    // If newFmTest2().AskAbout(.Document, ,"NO FLAT PATTERN!" & vbNewLine &"Try to generate one?") = vbYes Then
                                    if (withBlock2.AskAbout(invSheetMetalComp.Document, Join(Array("ISSUE WITH FLAT PATTERN:", "   NO BODY FOUND"), Constants.vbNewLine), Join(Array("Please consider reviewing model,", "and rebuilding its flat pattern.", "", "Pause for review? (not necessary)"), Constants.vbNewLine)) == Constants.vbYes)
                                        // (please check part for outdated flat pattern)
                                        System.Diagnostics.Debugger.Break();
                                    else
                                    {
                                    }
                                }

                                System.Diagnostics.Debugger.Break(); // BKPT-2021-1105-1256
                                /// CHANGE NEEDED[2021.11.05]:
                                /// Not sure what, exactly
                                /// NOTE[2021.12.13]
                                /// These values are converted to inches
                                /// from centimeters below in ...
                                /// They should NOT be converted HERE!
                                /// (don't think so, anyway)
                                /// disabling conversion operations
                                /// pending review on debug
                                dfHtThk = withBlock1.Parameters.Item("Thickness").Value; // / cvLenIn2cm
                                dLength = withBlock1.length; // / cvLenIn2cm
                                dWidth = withBlock1.Width; // / cvLenIn2cm
                                dArea = dLength * dWidth;
                            }
                            else
                            {
                                // Stop 'BKPT-2021-1105-1250
                                /// CHANGE NEEDED[2021.11.05]:
                                /// Actually have a function
                                /// to collect X, Y, and Z spans
                                /// which is used in the main
                                /// function. Might be usable here.
                                /// '
                                /// (one might think the note would
                                /// IDENTIFY the aforementioned
                                /// function here, wouldn't one?
                                /// Well... one would be WRONG!!!)
                                /// '
                                /// Here we go: nuAiBoxData()
                                {
                                    var withBlock2 = nuAiBoxData().UsingInches(0).UsingBox(withBlock1.Body.RangeBox) // .SortingDims
   ;
                                    /// UPDATE[2021.12.13]
                                    /// Changed UsingInches argument
                                    /// to zero (DON'T use) because
                                    /// conversion is performed below.
                                    // Check height against thickness
                                    // Valid flat pattern should return
                                    // zero or VERY minimal difference
                                    dHeight = Round(withBlock2.SpanZ, 6); // (.MaxPoint.Z - .MinPoint.Z)

                                    // the extent of the face.
                                    // Extract the width, length and area from the range.
                                    dLength = Round(withBlock2.SpanX, 6); // (.MaxPoint.X - .MinPoint.X)
                                    dWidth = Round(withBlock2.SpanY, 6); // (.MaxPoint.Y - .MinPoint.Y)
                                }

                                {
                                    var withBlock2 = withBlock1.Body.RangeBox;
                                    /// CHECKPOINT[2021.12.07]:
                                    /// not actually stopping here
                                    /// but running a quick check
                                    /// to make sure revised code
                                    /// above works correctly.
                                    /// If so, this section SHOULD
                                    /// be good to remove or disable
                                    ck = 0;
                                    ck = ck + Abs(withBlock2.MaxPoint.Z - withBlock2.MinPoint.Z - dHeight); // * cvLenIn2cm
                                    ck = ck + Abs(withBlock2.MaxPoint.X - withBlock2.MinPoint.X - dLength); // * cvLenIn2cm
                                    ck = ck + Abs(withBlock2.MaxPoint.Y - withBlock2.MinPoint.Y - dWidth); // * cvLenIn2cm
                                    /// UPDATE[2021.12.13]
                                    /// conversion operations removed
                                    /// since no longer converting
                                    /// before end stage
                                    if (Round(ck, 5) > 0)
                                        System.Diagnostics.Debugger.Break(); // BKPT-2021-1207-1158
                                }

                                /// UPDATE[2021.11.05]:
                                /// Moved derived calculations
                                /// outside of With block above.
                                /// Might or might not prove useful.
                                dfHtThk = Round(dHeight - prThickness.Value, 6); // / cvLenIn2cm
                                /// UPDATE[2021.12.13]
                                /// conversion operation removed since
                                /// no longer converting before end stage
                                dArea = Round(dLength * dWidth, 6);

                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }

                            if (dArea == 0)
                                // [2021.06.11] Moved alternate calculation
                                // sequence to new If-Then block preceding
                                // this one, and removed previous comment
                                // as no longer relevant.
                                System.Diagnostics.Debugger.Break();// and note when this branch taken.

                            if (dArea > 0)
                            {
                            }
                            else
                            {
                                if (MsgBox(Join(Array("The flat pattern for this", "part has no features,", "and is likely not valid.", "", "Pause here to review?", "(Click 'NO' to just keep going)"), Constants.vbNewLine), Constants.vbYesNo, "Invalid Flat Pattern") == Constants.vbYes)
                                    System.Diagnostics.Debugger.Break();// and let the user look into it
                                Debug.Print(aiDocument(withBlock1.Document).FullDocumentName);
                            }
                        }

                        if (dfHtThk > 0)
                        {
                            {
                                var withBlock1 = dcFlatPatSpansByVertices(withBlock.FlatPattern);
                                ck = Round(withBlock1.Item("Z") - prThickness.Value, 6);
                                if (dfHtThk > ck)
                                {
                                    dHeight = withBlock1.Item("Z");
                                    dfHtThk = Round(dHeight - prThickness.Value, 6);
                                    Debug.Print(Round(withBlock1.Item("X") - dLength, 6));
                                    Debug.Print(Round(withBlock1.Item("Y") - dWidth, 6));
                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                }
                            }
                        }
                    }
                    else
                    {
                        // aiDocPart(.Document).
                        {
                            var withBlock1 = nuAiBoxData().UsingInches(0).UsingBox(withBlock.RangeBox) // .SortingDims
   ;
                            /// UPDATE[2021.12.13]
                            /// Changed UsingInches argument
                            /// to zero (DON'T use) because
                            /// conversion is performed below.
                            // Check height against thickness
                            // Valid flat pattern should return
                            // zero or VERY minimal difference
                            dHeight = Round(withBlock1.SpanZ, 6); // (.MaxPoint.Z - .MinPoint.Z)

                            // the extent of the face.
                            // Extract the width, length and area from the range.
                            dLength = Round(withBlock1.SpanX, 6); // (.MaxPoint.X - .MinPoint.X)
                            dWidth = Round(withBlock1.SpanY, 6); // (.MaxPoint.Y - .MinPoint.Y)

                            dArea = Round(dLength * dWidth, 6);
                            // not really valid, here
                            // but it's used below
                            dfHtThk = Round(dHeight - prThickness.Value, 6); // / cvLenIn2cm
                        }

                        Debug.Print(Join(Array("NOFLAT", aiDocument(withBlock.Document).FullFileName), ":"));
                    }
                    // '''''
                    // '''''  The following section should be moved OUTSIDE this branch!
                    // '''''
                    {
                        var withBlock1 = aiDocPart(withBlock.Document);
                        aiPropSet = withBlock1.PropertySets.Item(gnCustom);
                        // prOffThknss

                        // Convert values into document units.
                        // This will result in strings that are identical
                        // to the strings shown in the Extent dialog.
                        /// 
                        /// NOTE[2021.11.09]
                        /// If UsingInches is set as shown above,
                        /// this section might not work properly.
                        /// Might be better to NOT use inches,
                        /// and simply let things take care of
                        /// themselves, here.
                        /// UPDATE[2021.12.13]
                        /// Changed UsingInches argument to zero
                        /// in two calls above.
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

                    // Stop 'BKPT-2021-1105-1304
                    /// CHANGE NEEDED[2021.11.05]:
                    /// This is where Properties are set.
                    /// Want to change this to simply collect
                    /// the generated values, and pass them
                    /// back to the client for processing.
                    /// '
                    /// A separate process can then
                    /// assign them to Properties.
                    /// '
                    // Add area to custom property set
                    rt = dcWithProp(aiPropSet, pnRmQty, dArea * cvArSqCm2SqFt, rt);
                    // 
                    // 0.00107639 = (1ft / 12in/ft / 2.54 cm/in)^2
                    // 
                    // /  1ft | 1in    \2     2                2
                    // ( ------+-------- ) * cm  = 0.00107639 ft
                    // \ 12in | 2.54cm /
                    // 
                    // That value really needs to go into a constant
                    // 

                    // Add Width to custom property set
                    rt = dcWithProp(aiPropSet, pnWidth, strWidth, rt);

                    // Add Length to custom property set
                    rt = dcWithProp(aiPropSet, pnLength, strLength, rt);

                    // Add AreaDescription to custom property set
                    rt = dcWithProp(aiPropSet, pnArea, strArea, rt);

                    if (Strings.Len(strDVNs) > 0)
                        rt = dcWithProp(aiPropSet, "OFFTHK", strDVNs, rt);
                    else
                    {
                    }

                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                }
            }
            dcFlatPatProps = rt;
        }
    }

    public Scripting.Dictionary dcWithProp(PropertySet propSet, string Name, Variant Value, Scripting.Dictionary rt = null/* TODO Change to default(_) if this is not a reference type */)
    {
        dcWithProp = dcAddProp(aiSetProp(propSet, Name, Value), rt);
    }

    public Scripting.Dictionary dcAddProp(Inventor.Property aiProp, Scripting.Dictionary dcIn = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary dcOut;
        string nm;

        if (dcIn == null)
            dcAddProp = dcAddProp(aiProp, new Scripting.Dictionary());
        else
        {
            dcOut = dcIn;
            if (!aiProp == null)
            {
                nm = aiProp.Name;
                {
                    var withBlock = dcOut;
                    if (withBlock.Exists(nm))
                        withBlock.Remove(nm);
                    withBlock.Add(nm, aiProp);
                }
            }
            dcAddProp = dcOut;
        }
    }

    public Inventor.Property aiGetProp(PropertySet aiPropSet, string aiPropName, long Create = 0)
    {
        Inventor.Property aiProp;



        /// FORDEBUG[2021.08.09] -- disable when not debugging
        /// report names of Property  and desired Property
        // Debug.Print "PROPSET[" & aiPropSet.Name & "].ITEM[" & aiPropName & "]";

        // Attempt to get an existing custom property named aiPropName.
        aiProp = aiPropSet.Item(aiPropName);

        if (Information.Err.Number == 0)
        {
        }
        else if (Create == 0)
            /// FORDEBUG[2021.08.09] -- disable when not debugging
            // Debug.Print " NOTFOUND"

            aiProp = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            Information.Err.Clear();
            aiProp = aiPropSet.Add("", aiPropName);
            if (Information.Err.Number == 0)
            {
            }
            else
                /// FORDEBUG[2021.08.09] -- disable when not debugging
                // Debug.Print " CANTMAKE"

                aiProp = null/* TODO Change to default(_) if this is not a reference type */;
        }


        aiGetProp = aiProp;
    }

    public Inventor.Property aiSetProp(PropertySet aiPropSet, string aiPropName, Variant aiPropValue)
    {
        Inventor.Property aiProp;

        // Try to acquire Property Object
        aiProp = aiGetProp(aiPropSet, aiPropName, 1);

        if (aiProp == null)
            // Might have to handle that here
            Debug.Print("    NOPROP/CANTMAKE -- couldn't create Property " + aiPropName);
        else

                /// All Property Object acquisition code
                /// moved to aiGetProp (called above)
                /// remains in comment-disabled form below.
                /// Remove once functionality verified
                /// 

                // Attempt to get an existing custom property named aiPropName.
                // aiProp = aiPropSet.Item(aiPropName)

                // If Err.Number <> 0 Then 'failed to get property, which means it doesn't exist

                // Try to create it:
                // Err.Clear
                // aiProp = aiPropSet.Add((aiPropValue), aiPropName)
                // '  NOTE: aiPropValue apparently needs to be in parentheses
                // '      for some values. Specifically, number/unit strings
                // '      like "1.500 in" seem to trigger VBA Error 51: Internal Error
                // '      Embedding the variable in parentheses forces VBA
                // '      to resolve the Variant to a string, maybe?
                // '      In any case, that seems to fix the problem.

                // If Err.Number <> 0 Then 'creation also failed.
                // aiProp = Nothing
                // End If
                // Else 'Got the property so update the value:
                // 's check if it's different, first
                if (aiPropValue == aiProp.Value)
            // Stop 'to verify they're the same
            // Debug.Print "    SAMEVAL(" & aiPropValue & ")";
            Debug.Print("    SAMEVAL(" + aiProp.Name + "): " + aiProp.Value);
        else if (System.Convert.ToHexString(aiPropValue) == System.Convert.ToHexString(aiProp.Value))
            // Stop 'BKPT-2021-1105-1419
            /// CHANGE NEEDED[2021.11.05]:
            /// Need to make sure values really ARE different!
            /// RMQTY especially seems to have trouble with this.
            /// '
            /// Example: CHGVAL(RMQTY): 0.20833313172 ==> 0.20833313172
            /// 'VERY minor difference in values:
            /// '  Debug.Print aiPropValue - aiProp.Value
            /// '  -2.77555756156289E-17
            /// '
            /// Need to include a check between converted copies.
            /// Believe implemented before, but got lost in crash.
            /// '
            /// UPDATE[SAME_DAY]:
            /// This ElseIf clause adds confirmative test
            /// by checking String conversions of each
            /// against each other.
            /// '
            Debug.Print("    EQUIVAL(" + aiProp.Name + "): " + aiProp.Value);
        else
        {
            /// CHANGE NEEDED[2021.11.05]:
            /// Need to make sure values really ARE different!
            /// UPDATE[SAME_DAY]:
            /// Added confirmative test; see ElseIf above
            /// (hopefully no issue with failed CStr calls)
            Debug.Print("    CHGVAL(" + aiProp.Name + "): " + aiProp.Value + " => " + aiPropValue);
            // Stop 'and make sure it really IS different
            aiProp.Value = (aiPropValue);
            // '  See note above on setting Property.
            // '  Assuming parentheses also required here.

            if (Information.Err.Number == 0)
                // Debug.Print " ==> (" & aiProp.Value & ")";
                // Debug.Print " ==> " & aiProp.Value;
                Debug.Print("    OK!");
            else
                // Not much else we can do at this point
                // aiProp = Nothing
                Debug.Print("    FAILED:CANTCHG");
        }

        /// FORDEBUG[2021.08.09] -- disable when not debugging
        Debug.Print(); // forcing newline

        aiSetProp = aiProp;
    }

    public string ptNumShtMetal(Inventor.SheetMetalComponentDefinition aiSMdef)
    {
        // Request #2:  Genius SheetMetal
        // by matching Style Name and Material.
        // Add to Custom Property RM
        string invGeniusMaterial; // Return value
                                  // 
        Inventor.Parameter prThickness;
        ADODB.Recordset rs;
        Scripting.Dictionary dc;
        string invSheetMetalMaterial;
        double invSheetMetalThickness;
        string sqlText;
        string docName;

        {
            var withBlock = aiSMdef;
            Information.Err.Clear();
            prThickness = withBlock.Thickness;
            if (Information.Err.Number == 0)
            {
                /// For now, we must assume we can only proceed
                /// if a valid Thickness parameter is retrieved

                docName = aiDocPart(withBlock.Document).FullDocumentName;

                invSheetMetalMaterial = aiDocPart(withBlock.Document).PropertySets.Item(gnDesign).Item(pnMaterial).Value;
                invSheetMetalThickness = prThickness.Value / (double)cvLenIn2cm; // Internal Units in cm???
                                                                                 // invSheetMetalThickness = .Thickness.Value / cvLenIn2cm 'Internal Units in cm???
                sqlText = sqlSheetMetal(invSheetMetalMaterial, invSheetMetalThickness);

                {
                    var withBlock1 = cnGnsDoyle();
                    rs = withBlock1.Execute(sqlText);
                    {
                        var withBlock2 = rs;
                        if ((withBlock2.BOF & withBlock2.EOF))
                        {
                            withBlock2.Close();

                            // Here's where we resort to the HARD way.
                            // Debug.Print Val(aiSMdef.ActiveSheetMetalStyle.Thickness) - invSheetMetalThickness < 0.0001
                            if (Val(aiSMdef.ActiveSheetMetalStyle.Thickness) - invSheetMetalThickness < 0.0001)
                                /// UPDATE[2022.01.12.1314]:
                                /// Add check for matching Thickness
                                /// between Part Property and its
                                /// active Sheet Metal Style.
                                /// If they DON'T match, then it's
                                /// probably NOT a Sheet Metal Part.
                                /// Will probably need a better check
                                /// moving forward, but this SHOULD do
                                /// for now.
                                invGeniusMaterial = pnShtMetalHardCoded(invSheetMetalMaterial, aiSMdef.ActiveSheetMetalStyle.Name);
                            else
                            {
                                Debug.Print(docName);
                                Debug.Print(aiSMdef.ActiveSheetMetalStyle.Name);
                                Debug.Print(aiSMdef.ActiveSheetMetalStyle.Thickness);
                                System.Diagnostics.Debugger.Break();
                                invGeniusMaterial = "";
                            }

                            if (Strings.Len(invGeniusMaterial) > 0)
                            {
                                // something might be missing from Genius
                                if (Left(invGeniusMaterial, 2) == "LG")
                                    // might actually be lagging
                                    Debug.Print("POSSIBLE LAGGING ITEM");
                                else
                                {
                                }
                                System.Diagnostics.Debugger.Break(); // and review the situation
                            }
                            else
                                invGeniusMaterial = "";
                        }
                        else
                        {
                            // (or SOMETHING, anyway) '''REV[2023.05.17.1211]
                            /// added User Prompt to pick from multiple options
                            /// when more than one material option returned.
                            /// details noted below
                            // Stop
                            dc = dcFrom2Fields(rs, "Item", "Item"); // dcFromAdoRS(rs)
                            {
                                var withBlock3 = dc;
                                if (withBlock3.Count > 0)
                                {
                                    invGeniusMaterial = withBlock3.Keys(0); // .Fields(0).Value
                                    if (withBlock3.Count > 1)
                                        invGeniusMaterial = userChoiceFromDc(dc, invGeniusMaterial);// nuSelFromDict '.Fields(0).Value
                                }
                                else
                                {
                                    System.Diagnostics.Debugger.Break();
                                    invGeniusMaterial = "";
                                }
                            }

                            withBlock2.Close();
                        }
                    }

                    withBlock1.Close();
                }
            }
            else
                System.Diagnostics.Debugger.Break();// and review the situation
        }

        ptNumShtMetal = invGeniusMaterial;
    }

    public string sqlSheetMetal(string mtName = "", double thk = 0)
    {
        string hdr2match;
        string thk2match;
        string mtl2match;
        /// REV[2022.04.13.0939]
        /// modified to replace header match
        /// (FM-, FS-, etc.) with match against
        /// metal/material type (MS/SS) and
        /// thus be able to catch expanded
        /// metal options, assuming the sheet
        /// metal thickness is correctly set.
        /// NOTE: this is HIGHLY experimental.
        /// It SHOULD still work under most
        /// circumstances, but be aware of
        /// potential issues.

        if (mtName == "Stainless Steel")
        {
            hdr2match = "FS";
            mtl2match = "SS";
        }
        else if (mtName == "Stainless Steel, Austenitic")
        {
            hdr2match = "FS";
            mtl2match = "S4"; // for 409
        }
        else if (mtName == "Stainless Steel 304")
        {
            hdr2match = "FS";
            mtl2match = "SS";
        }
        else if (mtName == "304SS")
        {
            hdr2match = "FS";
            mtl2match = "SS";
        }
        else if (mtName == "Steel, Mild")
        {
            hdr2match = "FM";
            mtl2match = "MS";
        }
        else if (mtName == "Rubber")
        {
            hdr2match = "LG";
            mtl2match = ""; // not metal
        }
        else if (mtName == "Rubber, Silicone")
        {
            hdr2match = "LG";
            mtl2match = ""; // not metal
        }
        else if (mtName == "Lagging")
        {
            hdr2match = "LG";
            mtl2match = ""; // not metal
        }
        else if (mtName == "UHMW, White")
        {
            hdr2match = "UH";
            mtl2match = ""; // not metal
        }
        else
        {
            Debug.Print(mtName);
            // Stop
            hdr2match = "XX";
            mtl2match = ""; // not metal
        }

        thk2match = Format(thk, "0.000");
        // "Item Like '" & hdr2match & "%'"
        sqlSheetMetal = Trim(Join(Array("SELECT Item, Description1", "FROM vgMfiItems", "WHERE " + Interaction.IIf(Strings.Len(mtName) > 0, "Specification6 = '" + mtl2match + "'", "1=1"), "  AND Family = " + Interaction.IIf(hdr2match == "LG", "'D-PTS'", "'DSHEET'"), Interaction.IIf(hdr2match == "LG", "  AND Item LIKE 'LG%'", "  AND Specification1 = 'STANDARDSHEET'"), Interaction.IIf(thk > 0, "  AND Abs(Thickness - " + thk2match + ") < 0.007", ""), ";"), Constants.vbNewLine));
    }

    public string pnShtMetalHardCoded(string invSheetMetalMaterial, string invSheetMetalName)
    {
        string invGeniusMaterial;

        // Stop 'because this function should
        // not be getting called anymore

        // Map combination to corresponding Genius Part Number
        if (invSheetMetalMaterial == "Stainless Steel")
        {
            if (invSheetMetalName == "18 GA")
                invGeniusMaterial = "FS-48x96x0.048";
            else if (invSheetMetalName == "14 GA")
                invGeniusMaterial = "FS-60x120x0.075";
            else if (invSheetMetalName == "13 GA")
                invGeniusMaterial = "FS-60x97x0.09";
            else if (invSheetMetalName == "12 GA")
                invGeniusMaterial = "FS-60x120x0.105";
            else if (invSheetMetalName == "10 GA")
                invGeniusMaterial = "FS-60x144x0.135";
            else if (invSheetMetalName == "3/16\"")
                invGeniusMaterial = "FS-60x144x0.188";
            else if (invSheetMetalName == "1/4\"")
                invGeniusMaterial = "FS-60x144x0.25";
            else if (invSheetMetalName == "5/16\"")
                invGeniusMaterial = "FS-60x144x0.313";
            else if (invSheetMetalName == "3/8\"")
                invGeniusMaterial = "FS-60x144x0.375";
            else if (invSheetMetalName == "1/2\"")
                invGeniusMaterial = "FS-60x144x0.5";
            else
                invGeniusMaterial = "";
        }
        else if (invSheetMetalMaterial == "Steel, Mild")
        {
            if (invSheetMetalName == "14 GA")
                invGeniusMaterial = "FM-60x144x0.075";
            else if (invSheetMetalName == "12 GA")
                invGeniusMaterial = "FM-60x144x0.105";
            else if (invSheetMetalName == "10 GA")
                invGeniusMaterial = "FM-60x144x0.135";
            else if (invSheetMetalName == "3/16\"")
                invGeniusMaterial = "FM-60x144x0.188";
            else if (invSheetMetalName == "1/4\"")
                invGeniusMaterial = "FM-60x144x0.25";
            else if (invSheetMetalName == "5/16\"")
                invGeniusMaterial = "FM-60x144x0.313";
            else if (invSheetMetalName == "3/8\"")
                invGeniusMaterial = "FM-60x144x0.375";
            else if (invSheetMetalName == "1/2\"")
                invGeniusMaterial = "FM-60x144x0.5";
            else if (invSheetMetalName == "5/8\"")
                invGeniusMaterial = "FM-60x144x0.625";
            else if (invSheetMetalName == "3/4\"")
                invGeniusMaterial = "FM-60x120x0.75";
            else if (invSheetMetalName == "1\"")
                invGeniusMaterial = "FM-48x120x1";
            else
                invGeniusMaterial = "";
        }
        else if (invSheetMetalMaterial == "Rubber")
        {
            // Debug.Print "POSSIBLE LAGGING ITEM"
            // Debug.Print "invGeniusMaterial = ""LG"""
            System.Diagnostics.Debugger.Break();
            invGeniusMaterial = "LG";
        }
        else
            invGeniusMaterial = ""; // Mapping of material

        pnShtMetalHardCoded = invGeniusMaterial;
    }

    public Scripting.Dictionary dcAddDocPtNum(Scripting.Dictionary dcIn, Inventor.Document AiDoc)
    {
        Scripting.Dictionary rt;
        string pn;
        string fn;

        rt = dcIn;
        if (AiDoc == null)
        {
        }
        else
        {
            pn = AiDoc.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
            fn = AiDoc.FullFileName;
            {
                var withBlock = rt;
                if (withBlock.Exists(pn))
                    withBlock.Item(pn) = withBlock.Item(pn) + "|" + fn;
                else
                    withBlock.Add(pn, fn);
            }
        }
        dcAddDocPtNum = rt;
    }

    public Scripting.Dictionary dcAiDocPartNumbers(Scripting.Dictionary dcIn)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        foreach (var ky in dcIn)
            rt = dcAddDocPtNum(rt, aiDocument(obOf(dcIn.Item(ky))));
        dcAiDocPartNumbers = rt;
    }
}