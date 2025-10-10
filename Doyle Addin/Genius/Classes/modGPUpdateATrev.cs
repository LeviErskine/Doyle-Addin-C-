class SurroundingClass
{
    public Scripting.Dictionary dcGeniusPropsPartRev20180530_withComments(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// WAYPOINTS (search on phrase)
        /// (NOT Sheet Metal)
        /// 
        /// 
        /// '''
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
        /// dcGeniusPropsPartRev20180530_withComments_broken
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
        /// NOTE: predecessor variant function
        /// dcGeniusPropsPartPre20180530 moved
        /// to module modGPUpdateATrev to be
        /// retained for potential reference,
        /// prior to eventual deprecation
        /// and, presumably, removal.
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
        /// REV[2022.05.05.1110]
        /// add variable qtRawMatl to store
        /// material quantity BEFORE applying
        /// it to Property prRmQty
        /// Ultimate goal is to separate
        /// value changes from collection,
        /// moving the former as far down
        /// the process as possible, and
        /// ultimately, to the end.
        double qtRawMatl;
        string pnStock;
        string qtUnit;
        Inventor.BOMStructureEnum bomStruct;
        VbMsgBoxResult ck;
        aiBoxData bd;

        if (dc == null)
            dcGeniusPropsPartRev20180530_withComments = dcGeniusPropsPartRev20180530_withComments(invDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = invDoc;
                /// REV[2022.05.06.1113]
                /// add trap here for Content Center Items
                /// new ones likely won't, indeed CAN'T
                /// have custom properties, so attempts
                /// to read them will throw errors.
                /// '
                /// this is a stopgap to deal with a run
                /// in progress. a more thorough revision
                /// to properly address Content Center
                /// members (and other purchased parts)
                /// will be needed when possible.
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
                /// REV[2022.05.06.1124]
                /// embedded Custom Property collection
                /// process in Else branch of new check
                /// for Content Center Item.
                /// HOPEFULLY, this will help bypass
                /// error triggers when encountring
                /// Content Center member Items.
                if (withBlock.ComponentDefinition.IsContentMember)
                {
                    pnStock = "";
                    qtRawMatl = 0#;
                    qtUnit = "";
                }
                else
                {
                    prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1);
                    prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1);
                    prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1);

                    // ...and their initial Values
                    if (prRawMatl == null)
                        /// REV[2022.05.10.1427]
                        /// add check for successful collection
                        /// of custom properties
                        pnStock = "";
                    else
                        pnStock = prRawMatl.Value;
                    /// REV[2022.05.05.1517]
                    /// added trap to catch non-numeric values
                    /// in current Raw Material Quantity Property
                    /// and replace them with zero when encountered.
                    if (prRmQty == null)
                        qtRawMatl = 0#;
                    else if (IsNumeric(prRmQty.Value))
                        qtRawMatl = prRmQty.Value;
                    else
                        qtRawMatl = 0#;

                    if (prRmUnit == null)
                        qtUnit = "";
                    else
                        qtUnit = prRmUnit.Value;
                }

                // Part Number and Family properties
                // are from Design, NOT Custom set
                prPartNum = aiGetProp(aiPropsDesign, pnPartNum);
                // ADDED 2021.03.11
                pnModel = prPartNum.Value;
                prFamily = aiGetProp(aiPropsDesign, pnFamily);
                /// REV[2022.05.05.1551]
                /// added initial Value collection
                /// for Part Family as well as Number

                /// REV[2022.06.30.1531]
                /// exported Part Family Value collection
                /// and Genius check to function famVsGenius
                nmFamily = famVsGenius(pnModel, prFamily.Value);

                /// REV[2022.06.29.1351]
                /// changed Part Family Value collection
                /// to check against Genius
                /// UPDATE: superceded by
                /// REV[2022.06.30.1531]
                // nmFamily = Split(cnGnsDoyle().Execute("select Family from vgMfiItems where Item = '" & pnModel & "';").GetString(adClipString, , "", "", ""), vbCr)(0)
                // If Len(nmFamily) = 0 Then
                // nmFamily = prFamily.Value
                // ElseIf Len(prFamily.Value) > 0 Then
                // If nmFamily <> prFamily.Value Then
                // ck = MsgBox(Join(Array('            "Current Model Part Family " & prFamily.Value,'            "differs from Part Family " & nmFamily,'            "reported by Genius.", "",'            "Change to match Genius?", "",'            "(click [CANCEL] to debug)"'        ), vbNewLine),'            vbYesNoCancel + vbQuestion,'            "Match Genius Family?"'        )
                // 
                // If ck = vbCancel Then
                // Stop 'to debug
                // ElseIf ck = vbNo Then
                // nmFamily = prFamily.Value
                // 'to retain model value
                // Else 'do nothing, and Genius
                // 'Family should prevail
                // End If
                // End If
                // Else 'DO NOT SET IT HERE!
                // 'that's supposed to be done below
                // End If
                /// END of REV[2022.06.29.1351]
                /// (want to make sure the extent
                /// of this block is noted)

                /// We should check HERE for possibly misidentified purchased parts
                /// UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// Request #1:  the Mass in Pounds
                    /// and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        /// REV[2021.11.12]
                        /// Round mass to nearest ten-thousandth
                        /// to try to match expected Genius value.
                        /// This should reduce or minimize reported
                        /// discrepancies during ETM process.
                        /// REV[2022.05.06.1349]
                        /// adding (HOPEFULLY temporary) error trap here
                        /// to address issue with Application Error
                        /// when attempting to retrieve Mass.

                        rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                        if (Information.Err.Number)
                            /// suspect it's just an issue with a
                            /// particular Part Document (for Item SP344)
                            /// 
                            /// however, there may be some indication
                            /// of an issue relating to a protected
                            /// Excel worksheet
                            /// 
                            /// see https://docs.microsoft.com/en-us/office/troubleshoot/excel/run-time-error-2147467259-80004005
                            /// 
                            /// Error Number
                            /// -2147467259
                            /// (0x80004005)
                            /// Automation error
                            /// Unspecified error
                            System.Diagnostics.Debugger.Break();
                    }

                    /// BOM Structure type, correcting if appropriate,
                    /// and prepare Family value for part, if purchased.
                    /// 
                    ck = Constants.vbNo;
                    /// REV[2022.05.06.1118]
                    /// added separate check for Content Center Item.
                    /// (using code from REV[2022.05.06.1113] above)
                    if (withBlock1.IsContentMember)
                        ck = Constants.vbYes;
                    else if (InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + nmFamily + "|") > 0)
                        /// REV[2022.06.29.1416]
                        /// added this ElseIf to check against the
                        /// Family collected from Genius into nmFamily.
                        /// (see REV[2022.06.29.1351] above)
                        /// if Genius says it's purchased, it should be.
                        ck = Constants.vbYes;
                    else if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0)
                        // when possible, change prFamily.Value to nmFamily
                        /// REV[2022.05.06.1118]
                        /// changed If to ElseIf here to "chain" it
                        /// to Content Center check preceding. No need
                        /// to dig deeper if already have that, right?

                        /// REV[2018.05.31]: Combined both InStr checks
                        /// by addition to generate a single test for > 0
                        /// If EITHER string match succeeds, the total
                        /// SHOULD exceed zero, so this SHOULD work.
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
                    // REV[2018.05.30]: Value produced here
                    // will now be held for later processing,
                    // more toward the end of this function.
                    if (bomStruct == kPurchasedBOMStructure)
                    {
                        if (withBlock1.IsContentMember)
                            /// NOTE[2022.05.06.1130]
                            /// just noting this check has been here
                            /// for some time already. Probably since
                            /// 2018, noting REV[2018.05.30] above.
                            nmFamily = "D-HDWR";
                        else
                            nmFamily = "D-PTS";
                    }
                    else
                    {
                    }
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
                                    /// REV[2022.04.14.1131]
                                    /// added print statements to inform user
                                    /// of current part number and members
                                    /// of its BOM currently in Genius
                                    Debug.Print(prPartNum.Value + Constants.vbNewLine + Constants.vbTab + Join(withBlock1.Keys, Constants.vbNewLine + Constants.vbTab));
                                    // Stop 'because selection is going
                                    // to be a lot more complicated.
                                    // (just look at that pnStock
                                    // assignment up there!)

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
                                    // should only be minor quantity changes
                                    // Stop 'and make sure we want to do this.

                                    // dcIn = dcOb(dcIn.Item(dcOb(.Item(pnStock)).Keys(0)))
                                    // Deactivated, moved down and out of this If-Then nest.
                                    // Search below for active copy

                                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                                else
                                {
                                    Debug.Print("=== CURRENT GENIUS MATERIAL DATA ===");
                                    // Debug.Print dumpLsKeyVal(dcIn, ":" & vbTab)
                                    // ck = newFmTest2().AskAbout(invDoc,"Raw Material " & prRawMatl.Value& vbNewLine & " for Item","does not match " & pnStock& vbNewLine & "indicated in Genius."& vbNewLine & vbNewLine& "Change to match Genius?"& vbNewLine & "(Cancel to debug)")
                                    /// REV[2022.04.01.1443]
                                    /// short-circuiting this prompt
                                    /// and assuming automatic material
                                    /// change confirmation at this stage.
                                    /// '
                                    /// user gets another opportunity
                                    /// to confirm below. that should
                                    /// make this one redundant
                                    /// 
                                    ck = Constants.vbOK;
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
                            /// NOTE[2022.04.12.1157]
                            /// this section might want refinement
                            /// seems to be trying to determine
                            /// whether part is clearly sheet metal
                            /// might want to add something to further
                            /// determine NON sheet metal status
                            dcFP = dcFlatPatVals(withBlock.ComponentDefinition);
                            /// try to get flat pattern data
                            /// WITHOUT mucking up Properties!
                            /// Want to avoid dirtying file with
                            /// changes until absolutely necessary)

                            if (dcFP.Exists(pnThickness))
                            {
                                pnStock = ptNumShtMetal(invDoc.ComponentDefinition);
                                /// NOTE[2022.05.31.1158]
                                /// this attempt to capture sheet metal item
                                /// might NOT be appropriate! it appears to be
                                /// repeated below
                                /// NOTE[2022.05.31.1146]
                                /// need to better address failed capture
                                /// of sheet metal item number. material
                                /// selection dialog SHOULD be invoked
                                /// somewhere to address this!
                                /// see also NOTE[2022.05.31.1149] below
                                if (Strings.Len(pnStock) == 0)
                                    System.Diagnostics.Debugger.Break();
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
                            Debug.Print(ConvertToJson(Array(dcIn, dcFP, prPartNum.Value), Constants.vbTab));
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

                                        /// REV[2022.04.19.0945]
                                        /// made to following official:
                                        /// NOTE[2022.01.03]:
                                        /// Following text SHOULD no longer
                                        /// be needed. Verify function of
                                        /// fmTest2 following, and when good,
                                        /// disable and/or remove this block.
                                        // Debug.Print "!!! NOTICE !!!"
                                        // Debug.Print "Recommended Sheet Metal Stock (" & pnStock & ")"
                                        // Debug.Print "does not match current Stock (" & prRawMatl.Value & ")"
                                        // Debug.Print
                                        // Debug.Print "To continue with no change, just press [F5]. Otherwise,"
                                        // Debug.Print "press [ENTER] on the following line first to change:"
                                        // Debug.Print "prRawMatl.Value = """ & pnStock & """"
                                        // Debug.Print

                                        /// REV[2022.04.19.0944]
                                        /// added check for case mismatch.
                                        /// if that's the only difference,
                                        /// no need to bother the user.
                                        if (UCase(prRawMatl.Value) == pnStock)
                                            ck = Constants.vbYes;
                                        else
                                            /// NOTE[2022.01.03]:
                                            /// Now using fmTest2(?) to prompt
                                            /// user as in other checks (above?)
                                            ck = newFmTest2().AskAbout(invDoc, "Suggest Sheet Metal change from" + Constants.vbNewLine + prRawMatl.Value + " to" + Constants.vbNewLine + pnStock + " for", "Change it?");

                                        if (ck == Constants.vbCancel)
                                        {
                                            Debug.Print(ConvertToJson(nuDcPopulator.Setting(pnModel, nuDcPopulator.Setting("from", prRawMatl.Value).Setting("into", pnStock).Dictionary).Dictionary, Constants.vbTab));
                                            System.Diagnostics.Debugger.Break(); // to check things out
                                        }
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
                                // With cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", pnStock))
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family from vgMfiItems where Item='" + Replace(pnStock, "'", "''") + "';");
                                    /// REV[2022.08.26.1055]
                                    /// replaced direct ref to pnStock
                                    /// with Replace operation to "escape"
                                    /// it, re REV[2022.08.19.1416] (below)
                                    if (withBlock1.BOF | withBlock1.EOF)
                                    {
                                        if (pnStock != "0")
                                        {
                                            /// REV[2022.03.01.1553]
                                            /// embedded in check
                                            /// for string value "0"
                                            /// as this seems to come
                                            /// up as a legacy issue,
                                            /// and is readily remedied
                                            /// in this section. No stop
                                            /// is needed in that case.
                                            if (Strings.Len(pnStock) > 0)
                                                /// REV[2022.07.07.1340]
                                                /// added secondary check for string length.
                                                /// an empty string requires no user attention.
                                                System.Diagnostics.Debugger.Break();// because Material value likely invalid
                                        }
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
                                /// NOTE[2022.05.10.1559]
                                /// see NOTE[2022.05.10.1558]
                                /// on HOSE AND PIPING (search on this
                                /// to find in this function) for more
                                /// robust approach to PROBABLE LAGGING
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
                                System.Diagnostics.Debugger.Break(); // to check block WITH@1764
                                /// REV[2022.08.26.1001]
                                /// placing temporary Stops at start
                                /// and end of following With block
                                /// to check use of fields normally
                                /// requested in SQL select statement.
                                /// 
                                // With cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", pnStock))
                                // preceding (disabled) With statement
                                // to replace the following, assuming
                                // tests prove successful. if so, it
                                // might permit further streamlining
                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family from vgMfiItems where Item='" + Replace(pnStock, "'", "''") + "';") // , Description1, Unit, Specification1, Specification2, Specification3, Specification4, Specification5, Specification6, Specification7, Specification8, Specification9, Specification15, Specification16
       ;
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
                                    
                                    ''' UPDATE[2021.06.18]:
                                    '''     New pre-check for Material Item
                                    '''     in Purchased Parts Family.
                                    '''     VERY basic handler simply
                                    '''     maps Material Family to D-BAR
                                    '''     to force extra processing below.
                                    '''     Further refinement VERY much needed!
                                    ''' REV[2022.04.15.1035]
                                    '''     moved Stop statements to head
                                    '''     of their respective branches.
                                    '''     anticipate need to come up
                                    '''     with much better mechanism
                                    '''     to handle "special" raw stock
                                    '''     (read: D/R-PTS stock family)
                                    If mtFamily Like "?-MT*" Then
                                        'Debug.Print pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value
                                        'Debug.Print pnModel & "[" & prRmQty.Value & qtUnit & "*" & pnStock & ": " & aiPropsDesign(pnDesc).Value & "]" ' prRawMatl.Value
                                        Debug.Print pnModel & "["& qtRawMatl & qtUnit& " of " & pnStock & ": "& aiPropsDesign(pnDesc).Value& "]" 'prRmQty.Value prRawMatl.Value
                                        Stop 'FULL Stop!
                                    ''' NOTE[2022.05.05.1603]
                                    ''' new ElseIf branch called for here
                                    ''' see corresponding block under
                                    ''' Standard Part branch.
                                    ElseIf mtFamily = "D-PTS" Then
                                        Stop 'NOT SO FAST!
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

                                                /// REV[2022.04.04.1350]
                                                /// disabled immediate mode boilerplate text dump
                                                /// as user prompt appears to be functioning properly
                                                // Debug.Print "===== CHECK AND VERIFY RAW MATERIAL QUANTITY ====="
                                                // Debug.Print "  If change required, place new values at end"
                                                // Debug.Print "  of lines below for prRmQty.Value and qtUnit."
                                                // Debug.Print "  Press [ENTER] on each line to be changed."
                                                // Debug.Print "  Press [F5] when ready to continue."
                                                // Debug.Print "----- " & pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value & " -----"
                                                // Debug.Print ""

                                                /// REV[2022.02.09.0923]
                                                /// replication of REV[2022.02.09.0919]
                                                /// from section below: prep to replace
                                                /// old dimension dump operation with more
                                                /// compact call to aiBoxData's Dump method
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

                                                /// REV[2022.02.08.1446]
                                                /// removed block of Debug.Print lines
                                                /// disabled now for some time, as they
                                                /// do not seem to have been missed.
                                                // Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value); " 'in model. ";
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
                                                // Debug.Print "qtUnit = ""IN"""
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                // Debug.Print ""
                                                /// REV[2022.03.11.1125]
                                                /// now invoking new UserForm Interface
                                                /// for Material Quantity determination.
                                                /// as in REV[2022.03.11.1112] (above)
                                                {
                                                    var withBlock2 = nu_fmIfcMatlQty01().SeeUser(invDoc);
                                                    if (withBlock2.Exists(pnRmQty))
                                                    {
                                                        /// REV[2022.04.04.1404]
                                                        /// add checks for value difference
                                                        /// here and to units (below)
                                                        /// REV[2022.04.11.1007]
                                                        /// added additional "guard code"
                                                        /// to avoid error condition resulting
                                                        /// from blank value of property RMQTY
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
                                                // Stop 'because we might want a D-BAR handler
                                                /// Actually, we might NOT need to stop here
                                                /// if bar stock is already selected,
                                                /// because quantities would presumably
                                                /// have been established already.
                                                /// Any D-BAR handler probably needs
                                                /// to be implemented in prior section(s)
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

                                            prRmQty.Value = qtRawMatl;
                                            rt = dcAddProp(prRmQty, rt);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                                        }
                                        else
                                        {
                                            System.Diagnostics.Debugger.Break(); // because we don't know WHAT to do with it
                                                                                 // and we do NOT want to clear anything
                                                                                 // until we know what's going on!
                                            nmFamily = "";
                                            qtUnit = ""; // may want function here
                                        }
                                    }
                                    System.Diagnostics.Debugger.Break(); // at end of block WITH@1764
                                }
                            }
                            else if (0)
                                System.Diagnostics.Debugger.Break();// and regroup
                        }
                    }
                    else
                    {
                        // --------------------------------------------'
                        /// REV[2022.05.04.1501]
                        /// adding an option to try to handle
                        /// hose and piping elements
                        /// NOTE[2022.05.10.1558]
                        /// a similar process might be invoked
                        /// to address PROBABLE LAGGING (search
                        /// on that to find in this function)
                        if (withBlock.DocumentInterests.HasInterest(guidPipingSgmt))
                        {
                            // Stop
                            ck = newFmTest2().AskAbout(invDoc, "", Join(Array("", "appears to be Hose or Tubing,", "presently " + Interaction.IIf(Strings.Len(pnStock) > 0, pnStock, "unset") + ".", "", "Would you like to " + Interaction.IIf(Strings.Len(pnStock) > 0, "change", "set") + " it?"), Constants.vbNewLine));
                            /// & "Raw Material " & prRawMatl.Value& vbNewLine & "Unit of Measure currently "& .Value & vbNewLine & vbNewLine& "Change to " & qtUnit & "?"& vbNewLine & " "'''
                            if (ck == Constants.vbCancel)
                                System.Diagnostics.Debugger.Break();
                            else if (ck == Constants.vbYes)
                            {
                                // pnStock = userChoiceFromDc(nuDcPopulator().Setting("(" & pnStock & ")", pnStock).Setting("5/16"" OD HOSE (GR16)", "GR16").Dictionary(),pnStock)

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
                        /// NOTE[2022.05.04.1638]
                        /// originally an alternate branch to both
                        /// sheet metal and "standard" part handlers,
                        /// it was decided to move it to the start of
                        /// the "standard" handler to take advantage
                        /// of the property setting code there.
                        /// 
                        /// ultimately, things are going to have to be refactored
                        /// to better manage data gathering and assignment overall.


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
                            else
                                Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
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
                                var withBlock1 = cnGnsDoyle().Execute("select Family from vgMfiItems where Item='" + Replace(pnStock, "'", "''") + "';");
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
                                // With cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", pnStock))
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
                                    ''' REV[2022.05.05.1343]
                                    ''' inserted new check for purchased
                                    ''' (PTS) "material" items. This SHOULD
                                    ''' ultimately replace the following two
                                    ''' ElseIf statements, and consolidate
                                    ''' determination of Part Family.
                                        
                                        ''' REV[2022.05.05.1610]
                                        ''' added preliminary check for RMT
                                        ''' material family, bypassing User
                                        ''' prompt if encountered.
                                        ''' likely want to build on this
                                        ''' to confirm User wants to keep
                                        ''' existing Family setting.
                                        If nmFamily Like "?-RM*" Then
                                            'ck = vbNo
                                            Debug.Print ; 'Breakpoint Landing
                                        Else
                                            ck = MsgBox(Join(Array("Part " & pnModel & " uses " & pnStock,"which is not sheet metal.","","These parts are usually assigned","to the Riverview family, R-RMT.","","Do you want to use this Family?","Click [NO] to see other options.","(CANCEL to debug)"), vbNewLine),vbYesNoCancel + vbQuestion,"Select Part Family?")
                                            '"Part " & pnModel & " uses " & pnStock & vbNewLine & "which is " & mtFamily & " Material."
                                            '"Part " & pnModel & " uses " & pnStock & vbNewLine & "which is not sheet metal." & vbNewLine & "" & vbNewLine & "These parts are usually assigned" & vbNewLine & "to the Riverview family, R-RMT." & vbNewLine & "" & vbNewLine & "Do you want to use this Family?" & vbNewLine & "Click [NO] to see other options." & vbNewLine & "(CANCEL to debug)"
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

                                    /// REV[2022.04.04.1350]
                                    /// disabled immediate mode boilerplate text dump
                                    /// as user prompt appears to be functioning properly
                                    // Debug.Print "===== CHECK AND VERIFY RAW MATERIAL QUANTITY ====="
                                    // Debug.Print "  If change required, place new values at end"
                                    // Debug.Print "  of lines below for prRmQty.Value and qtUnit."
                                    // Debug.Print "  Press [ENTER] on each line to be changed."
                                    // Debug.Print "  Press [F5] when ready to continue."
                                    // Debug.Print "----- " & pnModel & " [" & prRawMatl.Value & "]: " & aiPropsDesign(pnDesc).Value & " -----"
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
                                    // Debug.Print "prRmQty.Value = "; CStr(prRmQty.Value); " 'in model. ";
                                    Debug.Print("qtRawMatl = ");
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
                                    /// REV[2022.03.11.1112]
                                    /// now invoking new UserForm Interface
                                    /// for Material Quantity determination.
                                    /// see also REV[2022.03.11.1125] (below)
                                    {
                                        var withBlock1 = nu_fmIfcMatlQty01().SeeUser(invDoc);
                                        if (withBlock1.Exists(pnRmQty))
                                        {
                                            /// REV[2022.04.04.1404]
                                            /// add checks for value difference
                                            /// here and to units (below)
                                            /// REV[2022.04.11.1007]
                                            /// added additional "guard code"
                                            /// to avoid error condition resulting
                                            /// from blank value of property RMQTY
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
                                    // Stop 'because we might want a D-BAR handler
                                    /// Actually, we might NOT need to stop here
                                    /// if bar stock is already selected,
                                    /// because quantities would presumably
                                    /// have been established already.
                                    /// Any D-BAR handler probably needs
                                    /// to be implemented in prior section(s)
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
                    /// REV[2022.05.05.1257]
                    /// begin consolidating common steps from end
                    /// of both Sheet Metal and Standard branches.

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
            dcGeniusPropsPartRev20180530_withComments = rt;
        }
    }

    public Scripting.Dictionary dcGeniusPropsPartRev20200409(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// dcGeniusPropsPartRev20200409
        /// [2020.04.09] begin new revision
        /// 
        Scripting.Dictionary rt;
        // '
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;
        // '
        Inventor.Property prFamily;
        Inventor.Property prRawMatl; // pnRawMaterial
        Inventor.Property prRmUnit; // pnRmUnit
        Inventor.Property prRmQty; // pnRmQty
                                   // '
        string nmFamily;
        string mtFamily;
        /// UPDATE[2018.05.30.01]
        string pnStock;
        string qtUnit;
        Inventor.BOMStructureEnum bomStruct;
        VbMsgBoxResult ck;
        aiBoxData bd;

        if (dc == null)
            dcGeniusPropsPartRev20200409 = dcGeniusPropsPartRev20200409(invDoc, new Scripting.Dictionary());
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

                // Family property is from Design, NOT Custom set
                prFamily = aiGetProp(aiPropsDesign, pnFamily);

                /// We should check HERE for possibly misidentified purchased parts
                /// UPDATE[2018.02.06.01]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    /// Request #1:  the Mass in Pounds
                    /// and add to Custom Property GeniusMass
                    {
                        var withBlock2 = withBlock1.MassProperties;
                        rt = dcWithProp(aiPropsUser, pnMass, Round(cvMassKg2LbM * withBlock2.Mass, 4), rt);
                    }

                    /// BOM Structure type, correcting if appropriate,
                    /// and prepare Family value for part, if purchased.
                    /// 
                    /// NOTE[2020.04.09.01]
                    ck = Constants.vbNo;
                    /// UPDATE[2018.05.31.01]
                    if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0)
                        /// UPDATE[2020.04.09.02]
                        ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Is this a Purchased Part?");

                    /// Check process below replaces duplicate check/responses above.
                    /// NOTE[2020.04.09.02]
                    if (ck == Constants.vbYes)
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
                    // UPDATE[2018.05.30.02]
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

                /// UPDATE[2020.04.09.03]
                if (bomStruct == kNormalBOMStructure)
                {
                    // ----------------------------------------------------'
                    if (withBlock.SubType == guidSheetMetal)
                    {
                        // ----------------------------------------------------'
                        /// NOTE[2018.05.31.01]
                        // Request #3:
                        // sheet metal extent area
                        // and add to custom property "RMQTY"
                        rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);
                        /// NOTE[2018.05.30.01]

                        // Moved to start of block to check for NON sheet metal

                        // NOTE: THIS call might best be combined somehow
                        // with the flat pattern prop pickup above.
                        // Note especially that if dcFlatPatProps
                        // FINDS NO .FlatPattern, then there should
                        // BE NO sheet metal part number!
                        if (prRawMatl == null)
                        {
                            System.Diagnostics.Debugger.Break(); // '' UPDATE[2020.04.09.04]
                            if (rt.Exists("OFFTHK"))
                            {
                                /// UPDATE[2018.05.30.05]
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
                            /// NOTE[2018.09.14.01]: ACTION ADVISED
                            if (Len(prRawMatl.Value) > 0)
                                pnStock = prRawMatl.Value;
                            else
                                pnStock = ptNumShtMetal(withBlock.ComponentDefinition);

                            if (Strings.Len(pnStock) == 0)
                            {
                                /// UPDATE[2018.05.30.03]
                                {
                                    var withBlock1 = newFmTest1();
                                    if (!(invDoc.ComponentDefinition.Document == invDoc))
                                        System.Diagnostics.Debugger.Break();

                                    bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);
                                    ck = withBlock1.AskAbout(invDoc, "No Stock Found! Please Review" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                                    if (ck == Constants.vbYes)
                                    {
                                        /// UPDATE[2018.05.30.04]
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

                            if (Strings.Len(pnStock) > 0)
                            {
                                // do we look for a Raw Material Family!

                                {
                                    var withBlock1 = cnGnsDoyle().Execute("select Family " + "from vgMfiItems " + "where Item='" + pnStock + "';");
                                    if (withBlock1.BOF | withBlock1.EOF)
                                        System.Diagnostics.Debugger.Break(); // because Material value likely invalid
                                    else
                                    {
                                        {
                                            var withBlock2 = withBlock1.Fields;
                                            mtFamily = withBlock2.Item("Family").Value;
                                        }

                                        if (mtFamily == "DSHEET")
                                        {
                                            // We should be okay. This is sheet metal stock
                                            nmFamily = "D-RMT";
                                            qtUnit = "FT2";
                                        }
                                        else if (mtFamily == "D-BAR")
                                        {
                                            nmFamily = "R-RMT";
                                            qtUnit = prRmUnit.Value; // "IN"
                                                                     // 'may want function here
                                            /// UPDATE[2018.05.30.07]
                                            Debug.Print(aiPropsDesign.Item(pnPartNum).Value);
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                            Debug.Print("THEN SELECT LENGTH FROM THE FOLLOWING SPANS,");
                                            Debug.Print("AND ENTER AT END OF prRmQty LINE BELOW.");
                                            Debug.Print("X SPAN", "Y SPAN", "Z SPAN");
                                            {
                                                var withBlock2 = invDoc.ComponentDefinition.RangeBox;
                                                Debug.Print(withBlock2.MaxPoint.X - withBlock2.MinPoint.X); /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
                                            }
                                            /// UPDATE[2020.04.09.05]
                                            Debug.Print("");
                                            Debug.Print("prRmQty.Value = ");
                                            /// UPDATE[2020.04.09.05]
                                            Debug.Print("qtUnit = \"IN\"");
                                            /// UPDATE[2020.04.09.05]
                                            System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                                            /// UPDATE[2020.04.09.05]
                                            Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                            System.Diagnostics.Debugger.Break();
                                            rt = dcAddProp(prRmQty, rt);
                                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                                        }
                                        else
                                        {
                                            nmFamily = "";
                                            qtUnit = ""; // may want function here
                                            /// UPDATE[2018.05.30.08]
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
                                    /// UPDATE[2020.04.09.06]
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
                                        System.Diagnostics.Debugger.Break(); // and check both so we DON'T
                                                                             // automatically "fix" the RMUNIT value

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
                        /// UPDATE[2020.04.09.07]
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Another landing line
                    }
                    else
                    {
                        // --------------------------------------------'
                        /// NOTE[2018.07.31.01]
                        {
                            var withBlock1 = newFmTest1();
                            if (!(invDoc.ComponentDefinition.Document == invDoc))
                                System.Diagnostics.Debugger.Break();

                            /// [2018.07.31.02][by AT]
                            bd = nuAiBoxData().UsingInches.SortingDims(invDoc.ComponentDefinition.RangeBox);

                            ck = withBlock1.AskAbout(invDoc, "Please Select Stock for Machined Part" + Constants.vbNewLine + Constants.vbNewLine + bd.Dump(0));

                            if (ck == Constants.vbYes)
                            {
                                /// UPDATE[2018.05.30.09]
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

                        /// NOTE[2020.04.09.05]
                        if (Strings.Len(pnStock) > 0)
                        {
                            // do we look for a Raw Material Family!

                            /// NOTE[2020.04.09.06]
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
                                nmFamily = "R-RMT";
                                qtUnit = prRmUnit.Value; // "IN"
                                                         // 'may want function here
                                /// UPDATE[2018.05.30.07]
                                Debug.Print(aiPropsDesign.Item(pnPartNum).Value);
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
                                /// UPDATE[2020.04.09.05]
                                Debug.Print("RAW MATERIAL QUANTITY IS NOW ");
                                System.Diagnostics.Debugger.Break();
                                rt = dcAddProp(prRmQty, rt);
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Landing line for debugging. Do not disable.
                            }
                            else
                            {
                                nmFamily = "";
                                qtUnit = ""; // may want function here
                                /// UPDATE[2018.05.30.08]
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
                                    // Debug.Print "Raw Stock Selection"
                                    // Debug.Print "  Current : " & prRawMatl.Value
                                    // Debug.Print "  Proposed: " & pnStock
                                    // Stop 'because we might not want to change existing stock setting
                                    // if
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
                                        System.Diagnostics.Debugger.Break(); // and check both so we DON'T
                                                                             // automatically "fix" the RMUNIT value

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
                else
                    System.Diagnostics.Debugger.Break();// because we might need

                // the design tracking property set,
                // and update the Cost Center Property
                if (invDoc.ComponentDefinition.IsContentMember)
                {
                }
                else if (Strings.Len(nmFamily) > 0)
                {
                    prFamily.Value = nmFamily;
                    rt = dcAddProp(prFamily, rt);
                }
            }

            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            dcGeniusPropsPartRev20200409 = rt;
        }
    }
    /// NOTE[2018.05.30.01]:

    /// Raw Material Quantity value

    /// SHOULD be set upon return

    /// We may need to review the process

    /// to find an appropriate place

    /// to set for NON sheet metal

    /// NOTE[2018.05.31.01]:

    /// At this point, we MAY wish

    /// to check for a valid flat pattern,

    /// and otherwise attempt to verify

    /// an actual sheet metal design.

    /// NOTE[2018.07.31.01][by AT]

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

    /// [2018.07.31.02][by AT]

    /// Added the following to try to

    /// preselect non-sheet metal stock

    /// [and then disabled the following]
    // .dbFamily.Value = "D-BAR"
    // .lbxFamily.Value = "D-BAR"
    /// Doesn't quite do it.
    // With New aiBoxData
    // bd = nuAiBoxData().UsingInches.UsingBox(invDoc.ComponentDefinition.RangeBox)
    // 
    // End With
    /// NOTE[2018.09.14.01]: ACTION ADVISED

    /// pnStock can probably be set to prRawMatl.Value

    /// and THEN checked for length to see if lookup needed.

    /// This might also allow us to check for machined

    /// or other non-sheet metal parts.

    /// NOTE[2018.09.14.02]: ACTION ADVISED

    /// Will need to address this situation

    /// in a more robust manner.

    /// A more thorough query above

    /// might also be called for.

    /// NOTE[2020.04.09.01]: This section should check

    /// for Purchased Part status in Genius, as well

    /// as the checks below. BOM Structure should also

    /// be checked, but SETTING it eventually needs

    /// to be shifted to a subsequent operation.

    /// NOTE[2020.04.09.02]:

    /// this is where Document's BOMStructure

    /// is set. should be moved to a later stage

    /// NOTE[2020.04.09.03]:

    /// [original date unknown]

    /// NON Content Center members

    /// might still be D-HDWR

    /// Additional checks might

    /// be recommended

    /// NOTE[2020.04.09.04]

    /// [original date unknown]

    /// We're going to need something here

    /// to make sure raw material gets added

    /// for non sheet metal parts, as well

    /// What we're going to need to do

    /// is refactor this whole bloody thing.

    /// NOTE[2020.04.09.05]

    /// [original date unknown]

    /// 

    /// The following If block is copied

    /// wholesale from sheet metal section above.

    /// Some changes (to be) made to accommodate

    /// machined or other non-sheet metal stock.

    /// 

    /// Ultimately, whole mess to require refactor.

    /// 

    /// NOTE[2020.04.09.06]

    /// [original date unknown]

    /// This enclosing With block should NOT be necessary

    /// since the newFmTest1 above takes care of collecting

    /// the Stock Family along with the Stock itself

    /// 

    /// NOTE[2020.04.09.07]

    /// [original date unknown]

    /// 

    /// Content formerly here moved BELOW and OUT of this section

    /// as it should only require results of newFmTest1 exchange above

    /// 
    // '''''
    // '''''
    // '''''
    /// UPDATE[2018.05.30.01]:

    /// Rename variable Family to nmFamily

    /// to minimize confusion between code

    /// and comment text in searches.

    /// Also add variable mtFamily

    /// for raw material Family name

    /// UPDATE[2018.05.30.02]:

    /// Value produced here

    /// will now be held for later processing,

    /// more toward the end of this function.

    /// UPDATE[2018.05.30.03]:

    /// Pulling ALL code/text from this section

    /// to get rid of excessive cruft.

    /// 

    /// In fact, reversing logic to go directly

    /// to User Prompt if no stock identified

    /// 

    /// IN DOUBLE FACT, hauling this WHOLE MESS

    /// RIGHT UP after initial pnStock assignment

    /// to prompt user IMMEDIATELY if no stock found

    /// UPDATE[2018.05.30.04]:

    /// Pulling some extraneous commented code

    /// from here and beginning of block

    /// UPDATE[2018.05.30.05]:

    /// Restoring original key check

    /// and adding code for debug

    /// Previously changed to "~OFFTHK"

    /// to avoid this block and its issues.

    /// (Might re-revert if not prepped to fix now)

    /// UPDATE[2018.05.30.06]: (two locations)

    /// Moving part family assignment

    /// to this section for better mapping

    /// and updating to new Family names

    /// as well as pulling up qtUnit assignment

    /// UPDATE[2018.05.30.07]: (two locations)

    /// As noted above

    /// Will keep Stop for now

    /// pending further review,

    /// hopefully soon

    /// UPDATE[2018.05.30.08]: As noted above

    /// However, might need more handling here.

    /// UPDATE[2018.05.30.09]:

    /// Pulling some extraneous commented code

    /// from here and beginning of block

    /// UPDATE[2018.05.31.01]:

    /// Combined both InStr checks

    /// by addition to generate a single test for > 0

    /// If EITHER string match succeeds, the total

    /// SHOULD exceed zero, so this SHOULD work.

    /// UPDATE[2018.02.06.01]:

    /// Using new UserForm

    /// UPDATE[2020.04.09.02]:

    /// Remove disabled/outdated code as follows

    /// UPDATE[2018.02.06]: Using same

    /// new UserForm as noted above.
    // ck = newFmTest2().AskAbout(invDoc, ,"Is this a Purchased Part?")
    // ElseIf InStr(1,"|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|","|" & prFamily.Value & "|") > 0 Then
    /// UPDATE[2018.02.06]: Using same
    /// new UserForm as noted above.
    /// UPDATE[2020.04.09.03]:
    /// Removed disabled/outdated code as follows
    // Request #4: Change Cost Center iProperty.
    // If BOMStructure = Normal, then Family = D-MTO,
    // else if BOMStructure = Purchased then Family = D-PTS.
    /// UPDATE[2020.04.09.04]:

    /// Adding Stop here to see if prRawMatl

    /// ever comes up missing inside a sheet metal part

    /// UPDATE[2020.04.09.05]: (multiple points)

    /// Removing disabled/obsolete code as follows
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
    /// 
    // Debug.Print "qtUnit = """; qtUnit; """"
    /// 
    // Debug.Print ""
    // Debug.Print ""
    // Debug.Print ""
    /// Actually, we might NOT need to stop here

    /// if bar stock is already selected,

    /// because quantities would presumably

    /// have been established already.

    /// Any D-BAR handler probably needs

    /// to be implemented in prior section(s)

    /// 

    /// UPDATE[2020.04.09.06]:

    /// Removing disabled/obsolete code as follows
    // Debug.Print "Raw Stock Selection"
    // Debug.Print "  Current : " & prRawMatl.Value
    // Debug.Print "  Proposed: " & pnStock
    // Stop 'because we might not want to change existing stock setting
    // if
    /// UPDATE[2020.04.09.07]:

    /// Removing disabled/obsolete code as follows
    // rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt) 'qtUnit WAS "FT2"
    /// Plan to remove commented line above,
    /// superceded by the one above that

    public Scripting.Dictionary dcGnsPartProps(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// NOTES[2021.03.12]
        /// Don't recall when this function block was created.
        /// Probably around 2020.04.09, with the generation
        /// of function dcGeniusPropsPartRev20200409, above.
        /// 
        /// As of this writing, no code present, so will
        /// use this to rebuild the Part Properties retrieval
        /// function more or less from the ground up.
        /// 
        /// Primary goal: reconstruct the basic process
        /// as faithfully as possible, but in a NONdestructive
        /// manner. That is, avoid changing the Part Document
        /// in any way, but simply collect as much information
        /// as is available, and generate whatever else is needed,
        /// and possible WITHOUT altering the Document.
        /// 
        /// NOTES[2021.03.22]
        /// Following review of process in functions
        /// dcGeniusPropsPartRev20180530 and dcFlatPatProps,
        /// added calls to aiGetProp to retrieve all Property
        /// items checked and/or set by those functions.
        /// 
        /// Again, this function should NOT attempt to create
        /// any missing/nonexistent Property items, in order
        /// to avoid altering the source Document at this stage.
        /// 

        /// 
        Scripting.Dictionary rt;
        // '
        // '  Property Sets
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;
        // '
        // '
        // '  Properties
        // Dim prPartNum   As Inventor.Property 'pnPartNum
        // Dim prFamily    As Inventor.Property 'pnFamily
        // Dim prRawMatl   As Inventor.Property 'pnRawMaterial
        // Dim prRmUnit    As Inventor.Property 'pnRmUnit
        // Dim prRmQty     As Inventor.Property 'pnRmQty
        // '
        // '
        // '  Property Values
        string pnModel;
        string nmFamily;
        string pnStock;
        string mtFamily;
        string qtUnit;
        // '
        // '
        // '
        Inventor.BOMStructureEnum bomStruct;
        VbMsgBoxResult ck;
        aiBoxData bd;

        rt = new Scripting.Dictionary();
        /// 

        {
            var withBlock = invDoc;
            // Property Sets
            {
                var withBlock1 = withBlock.PropertySets;
                aiPropsUser = withBlock1.Item(gnCustom);
                aiPropsDesign = withBlock1.Item(gnDesign);
            }
        }

        {
            var withBlock = rt;
            // Part Number and Family
            // Properties from Design set
            withBlock.Add(pnPartNum, aiGetProp(aiPropsDesign, pnPartNum)); withBlock.Add(pnFamily, aiGetProp(aiPropsDesign, pnFamily));
            // Custom Properties
            withBlock.Add(pnRawMaterial, aiGetProp(aiPropsUser, pnRawMaterial)); withBlock.Add(pnRmUnit, aiGetProp(aiPropsUser, pnRmUnit)); withBlock.Add(pnRmQty, aiGetProp(aiPropsUser, pnRmQty));         // NOTE[2021.03.12]: Removed 'create' flag
                                                                                                                                                                                                             // from these function calls to prevent
                                                                                                                                                                                                             // creation of nonexistent Properties,
                                                                                                                                                                                                             // which would alter the source Document.
                                                                                                                                                                                                             // NOTE ALSO: should try to obtain all other
                                                                                                                                                                                                             // custom Properties intended to generate,
                                                                                                                                                                                                             // in case they're already present.

            // Custom Mass/Dimensional Properties
            withBlock.Add(pnMass, aiGetProp(aiPropsUser, pnMass));         // .Add pnRmQty, aiGetProp(aiPropsUser, pnRmQty) 'prRmQty
                                                                           // this one already called above
            withBlock.Add(pnWidth, aiGetProp(aiPropsUser, pnWidth)); withBlock.Add(pnLength, aiGetProp(aiPropsUser, pnLength)); withBlock.Add(pnArea, aiGetProp(aiPropsUser, pnArea));         // .Add "OFFTHK", aiGetProp(aiPropsUser, "OFFTHK") '<prOffThk>
                                                                                                                                                                                               // disabled -- not sure if needed any longer
                                                                                                                                                                                               // and results in many fewer Prop Dicts
                                                                                                                                                                                               // with 'NoVal' Properties

            // prPartNum = .Item(pnPartNum)
            // pnModel = prPartNum.Value
            // prFamily = .Item(pnFamily)
            // prRawMatl = .Item(pnRawMaterial)
            // prRmUnit = .Item(pnRmUnit)
            // prRmQty = .Item(pnRmQty)

            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }

        /// 
        dcGnsPartProps = rt;
    }

    public Scripting.Dictionary dcGnsPartsWithProps(Inventor.Document invDoc)
    {
        /// function dcGnsPartsWithProps
        /// 
        /// returns Dictionary of Dictionaries
        /// containing Genius-related Properties
        /// for each Component of supplied
        /// Inventor Document, be it Part
        /// or Assembly.
        /// 
        /// NOTE: actual Dictionary processing
        /// removed to separate function
        /// dcGnsPartsWithPropsFromDc
        /// in order to support invocation
        /// from other functions w/o need
        /// for actual source Document
        /// 
        // Dim dc As Scripting.Dictionary
        Scripting.Dictionary rt;
        Variant ky;

        Inventor.PartDocument it;

        rt = dcGnsPartsWithPropsFromDc(dcAiDocComponents(invDoc, null/* Conversion error: Set to default value for this argument */, 0));

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        dcGnsPartsWithProps = rt;
    }

    public Scripting.Dictionary dcGnsPartsWithPropsFromDc(Scripting.Dictionary dc)
    {
        /// function dcGnsPartsWithPropsFromDc
        /// 
        /// returns Dictionary of Dictionaries
        /// containing Genius-related Properties
        /// for each Inventor Document in supplied
        /// Dictionary. Intended for invocation
        /// against a Dictionary of Inventor
        /// Documents generated by and/or within
        /// a separate function or procedure.
        /// 
        /// Initial creation intended to support
        /// companion function dcGnsPartsWithProps
        /// along with any others which might
        /// require it
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        Inventor.PartDocument it;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                it = aiDocPart(withBlock.Item(ky));
                if (it == null)
                {
                }
                else
                    rt.Add(ky, dcGnsPartProps(withBlock.Item(ky)));
            }
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        dcGnsPartsWithPropsFromDc = rt;
    }

    public Scripting.Dictionary dcOfDcAiPropVals(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        if (dc == null)
        {
        }
        else
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcAiPropValsFromDc(dcOb(withBlock.Item(ky))));
        }

        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        dcOfDcAiPropVals = rt;
    }

    public Scripting.Dictionary dcSansNoVals(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;
        Variant it;
        object ob;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                it = withBlock.Item(ky);
                {
                    var withBlock1 = rt;
                    if (IsObject(it))
                    {
                        // ob = obOf(it)
                        if (obOf(it) == null)
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        else
                        {
                            withBlock1.Add(ky, it);
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }
                    }
                    else if (IsNull(it))
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    else if (IsEmpty(it))
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    else
                    {
                        withBlock1.Add(ky, it);
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                }
            }
        }
        dcSansNoVals = rt;
    }

    public Scripting.Dictionary dcOfOnlyNoVals(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Variant ky;
        Variant it;
        object ob;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                it = withBlock.Item(ky);
                {
                    var withBlock1 = rt;
                    if (IsObject(it))
                    {
                        // ob = obOf(it)
                        if (obOf(it) == null)
                        {
                            withBlock1.Add(ky, it);
                            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                        }
                        else
                            Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
                    }
                    else if (IsNull(it))
                    {
                        withBlock1.Add(ky, it);
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                    else if (IsEmpty(it))
                    {
                        withBlock1.Add(ky, it);
                        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    }
                    else
                        Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
                }
            }
        }
        dcOfOnlyNoVals = rt;
    }

    public Scripting.Dictionary dc4noValStatus(Variant it, Scripting.Dictionary hasVal, Scripting.Dictionary noVal)
    {
        if (IsObject(it))
        {
            if (obOf(it) == null)
            {
                dc4noValStatus = noVal;
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }
            else
            {
                dc4noValStatus = hasVal;
                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }
        }
        else if (IsNull(it))
        {
            dc4noValStatus = noVal;
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
        else if (IsEmpty(it))
        {
            dc4noValStatus = noVal;
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
        else
        {
            dc4noValStatus = hasVal;
            Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
        }
    }

    public Scripting.Dictionary dcOfNoValStatus(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary nv;
        Scripting.Dictionary hv;
        Variant ky;
        Variant it;
        object ob;
        long ck;

        rt = new Scripting.Dictionary();
        hv = new Scripting.Dictionary();
        nv = new Scripting.Dictionary();
        rt.Add("HASVAL", hv);
        rt.Add("NOVAL", nv);
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                dc4noValStatus(withBlock.Item(ky), hv, nv).Add(ky, withBlock.Item(ky));
        }
        dcOfNoValStatus = rt;
    }

    public Scripting.Dictionary dcOfDcNoValStatus(Scripting.Dictionary dc)
    {
        /// dcOfDcNoValStatus
        /// 
        /// Given a Dictionary of Dictionaries,
        /// return a Dictionary of "No Value Status"
        /// Dictionaries for each Item in the
        /// source Dictionary
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dcOfNoValStatus(dcOb(withBlock.Item(ky))));
        }
        dcOfDcNoValStatus = rt;
    }

    public Scripting.Dictionary dcOfDcWithNoVals(Scripting.Dictionary dc)
    {
        /// dcOfDcWithNoVals
        /// 
        /// Given a Dictionary of Dictionaries,
        /// return a sub Dictionary of those
        /// with at least one "No Value" Item
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary wk;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcOfDcNoValStatus(dc);
            foreach (var ky in withBlock.Keys)
            {
                wk = dcOb(dcOb(withBlock.Item(ky)).Item("NOVAL"));
                if (wk.Count > 0)
                    rt.Add(ky, wk);
            }
        }
        dcOfDcWithNoVals = rt;
    }

    public Scripting.Dictionary mGr1g0f1(Inventor.PartDocument ob, Scripting.Dictionary dcIfIs, Scripting.Dictionary dcIfNot) // Object
    {
        /// 
        /// 
        if (ob == null)
            System.Diagnostics.Debugger.Break();
        else if (ob.ComponentDefinition.IsContentMember)
            mGr1g0f1 = dcIfIs;
        else
            mGr1g0f1 = dcIfNot;
    }

    public Scripting.Dictionary mGr1g0f2(Scripting.Dictionary dc)
    {
        /// 
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Scripting.Dictionary pr;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                pr = dcOb(withBlock.Item(ky));
                if (pr == null)
                    System.Diagnostics.Debugger.Break();
                else
                {
                }
            }
        }
        mGr1g0f2 = rt;
    }

    public Scripting.Dictionary dcGeniusPropsPartPre20180530(Inventor.PartDocument invDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        /// REV[2022.08.26.1204]
        /// moved to module modGPUpdateATrev
        /// from modGPUpdateAT to get it out
        /// of the way, while keeping it on
        /// hand for reference, just in case.
        /// 
        Scripting.Dictionary rt;
        // '
        Inventor.PropertySet aiPropsUser;
        Inventor.PropertySet aiPropsDesign;
        // '
        Inventor.Property prFamily;
        Inventor.Property prRawMatl; // pnRawMaterial
        Inventor.Property prRmUnit; // pnRmUnit
        Inventor.Property prRmQty; // pnRmQty
                                   // '
        string Family;
        string pnStock;
        string qtUnit;
        Inventor.BOMStructureEnum bomStruct;
        VbMsgBoxResult ck;

        if (dc == null)
            dcGeniusPropsPartPre20180530 = dcGeniusPropsPartPre20180530(invDoc, new Scripting.Dictionary());
        else
        {
            rt = dc;

            {
                var withBlock = invDoc;
                // the custom property set.
                aiPropsUser = withBlock.PropertySets.Item(gnCustom);
                aiPropsDesign = withBlock.PropertySets.Item(gnDesign);
                prRawMatl = aiGetProp(aiPropsUser, pnRawMaterial, 1);
                // '[2018-03-13:Add 1 to create RM property if not found]
                // '[2018-05-15:Add following to get props for RM Unit & Qty]
                prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1);
                prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1);
                // '
                prFamily = aiGetProp(aiPropsDesign, pnFamily);


                /// We should check HERE for possibly misidentified purchased parts
                /// UPDATE[2018.02.06]: Using new UserForm; see below
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    ck = Constants.vbNo;
                    if (InStr(1, invDoc.FullFileName, @"\Doyle_Vault\Designs\purchased\") > 0)
                        /// UPDATE[2018.02.06]: Using new UserForm
                        /// to show image and details
                        /// of part to be verified.
                        ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Is this a Purchased Part?");
                    else if (InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + prFamily.Value + "|") > 0)
                        /// This ElseIf condition should be combinable
                        /// with the initial If condition above
                        /// to simplify this check process.
                        /// All/most text in this clause should be
                        /// redundant and removable, once updated
                        /// check process has been validated.
                        /// UPDATE[2018.02.06]: Using same
                        /// new UserForm as noted above.
                        ck = newFmTest2().AskAbout(invDoc, null/* Conversion error: Set to default value for this argument */, "Is this a Purchased Part?");

                    /// Check process below replaces duplicate check/responses above.
                    /// Should be able to merge back into main branch
                    /// once code above is validated and refactored.
                    if (ck == Constants.vbYes)
                    {
                        if (withBlock1.BOMStructure != kPurchasedBOMStructure)
                            withBlock1.BOMStructure = kPurchasedBOMStructure;
                    }

                    bomStruct = withBlock1.BOMStructure;
                }

                // Request #1:  the Mass in Pounds and add to Custom Property GeniusMass
                {
                    var withBlock1 = withBlock.ComponentDefinition.MassProperties;
                    rt = dcWithProp(aiPropsUser, pnMass, Round(withBlock1.Mass * cvMassKg2LbM, 4), rt);
                }

                // ----------------------------------------------------'
                if (withBlock.SubType == guidSheetMetal)
                {
                    // ----------------------------------------------------'
                    // Request #4: Change Cost Center iProperty.
                    // If BOMStructure = Normal, then Family = D-MTO,
                    // else if BOMStructure = Purchased then Family = D-PTS.
                    // With .ComponentDefinition

                    if (bomStruct == kNormalBOMStructure)
                    {
                        // If prRawMatl.Value = "" Or cnGnsDoyle().Execute("select I.Family from vgMfiItems As I where I.Item='" & prRawMatl.Value & "';").Fields("Family").Value = "DSHEET" Then
                        // Request #3:
                        // sheet metal extent area
                        // and add to custom property "RMQTY"
                        rt = dcFlatPatProps(withBlock.ComponentDefinition, rt);

                        // Moved to start of block to check for NON sheet metal

                        // NOTE: THIS call might best be combined somehow
                        // with the flat pattern prop pickup above.
                        // Note especially that if dcFlatPatProps
                        // FINDS NO .FlatPattern, then there should
                        // BE NO sheet metal part number!
                        if (prRawMatl == null)
                        {
                            if (rt.Exists("~OFFTHK"))
                                pnStock = "";
                            else
                            {
                                System.Diagnostics.Debugger.Break(); // because we don't know IF this is sheet metal yet
                                pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                            }
                        }
                        else if (prRawMatl.Value == "")
                            // Stop 'because we're not sure what we have.
                            pnStock = ptNumShtMetal(withBlock.ComponentDefinition);
                        else
                        {
                            System.Diagnostics.Debugger.Break();
                            // With cnGnsDoyle().Execute(sqlOf_simpleSelWhere("vgMfiItems", "Family", "Item", prRawMatl.Value))
                            {
                                var withBlock1 = cnGnsDoyle().Execute("select i.Family from vgMfiItems i " + "where i.Item='" + prRawMatl.Value + "';");
                                {
                                    var withBlock2 = withBlock1.Fields("Family");
                                    if (withBlock2.Value == "DSHEET")
                                    {
                                    }
                                    else if (withBlock2.Value == "D-BAR")
                                        System.Diagnostics.Debugger.Break(); // because we might want a D-BAR handler
                                    else
                                        System.Diagnostics.Debugger.Break();// because we don't know WHAT do with it
                                }
                            }
                            pnStock = prRawMatl.Value;
                        }

                        if (Strings.Len(pnStock) > 0)
                            // Stop
                            Family = "D-MTO";
                        else
                        /// We MIGHT have an incorrectly marked PURCHASED part
                        // Stop
                        /// We'll want to see about fixing that here, maybe?
                        /// 

                        {
                            var withBlock1 = newFmTest1();
                            // aiSMdef.Document
                            if (!(invDoc.ComponentDefinition.Document == invDoc))
                                System.Diagnostics.Debugger.Break();
                            if (withBlock1.AskAbout(invDoc, "No Stock Found! Please Review") == Constants.vbYes)
                            {
                                // Join(Array(Join(Array("NO STOCK# for",Format(invSheetMetalThickness, "0.000") & "in",invSheetMetalMaterial), " "),"in " & docName, " ", "Stop/pause here?"), vbNewLine)
                                {
                                    var withBlock2 = withBlock1.ItemData // .Synch
   ;
                                    if (withBlock2.Exists(pnFamily))
                                    {
                                        Family = withBlock2.Item(pnFamily);
                                        Debug.Print(pnFamily + "=" + Family);
                                    }

                                    if (withBlock2.Exists(pnRawMaterial))
                                    {
                                        pnStock = withBlock2.Item(pnRawMaterial);
                                        Debug.Print(pnRawMaterial + "=" + pnStock);
                                    }
                                }
                            }
                        }

                        prRmQty = aiGetProp(aiPropsUser, pnRmQty, 1);
                        rt = dcAddProp(prRmUnit, rt);
                        qtUnit = "FT2";
                        prRawMatl.Value = pnStock;
                        rt = dcAddProp(prRawMatl, rt);
                        // rt = dcWithProp(aiPropsUser, pnRawMaterial, pnStock, rt)
                        /// 
                        // If aiGetProp(aiPropsUser, pnRmUnit) Is Nothing Then
                        // Stop
                        // Else
                        // If aiGetProp(aiPropsUser, pnRmUnit).Value <> "FT2" Then
                        /// prRmUnit = aiGetProp(aiPropsUser, pnRmUnit, 1)
                        if (Len(prRmUnit.Value) > 0)
                        {
                            if (prRmUnit.Value != qtUnit)
                                System.Diagnostics.Debugger.Break();// so we DON'T automatically "fix" the RMUNIT value
                        }
                        else
                        {
                        }
                        // End If
                        /// When this Stop activates, skip the next line
                        rt = dcWithProp(aiPropsUser, pnRmUnit, qtUnit, rt); // qtUnit WAS "FT2"
                        /// Want to change this part to allow for alternate RMUNIT values
                        /// When prior Stop is activated, use Ctrl-F9
                        /// to continue at the Stop line below.
                        if (0)
                            System.Diagnostics.Debugger.Break(); // to give us a skipover point
                    }
                    else if (bomStruct == kPurchasedBOMStructure)
                        Family = "D-PTS";
                    else
                        System.Diagnostics.Debugger.Break();// because we might need to do something else
                }
                else
                // --------------------------'
                // Request #2: Change Cost Center iProperty.
                // If BOMStructure = Purchased and not content center,
                // then Family = D-PTS, else Family = D-HDWR.
                {
                    var withBlock1 = withBlock.ComponentDefinition;
                    if (bomStruct == kPurchasedBOMStructureAnd.IsContentMember == false)
                        Family = "D-PTS";
                    else
                        // Family = "D-HDWR"
                        Family = "";
                } // Sheetmetal vs Part

                // the design tracking property set,
                // and update the Cost Center Property
                if (Strings.Len(Family) > 0)
                    rt = dcWithProp(aiPropsDesign, pnFamily, Family, rt);
            }

            iSyncPartFactory(invDoc); // Backport Properties to iPart Factory
            dcGeniusPropsPartPre20180530 = rt;
        }
    }
}