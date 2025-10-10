class app
{
    public void Update_Genius_Properties()
    {
        /// use Find (Ctrl-F or H) to jump to:
        /// WAYPOINT:UPDATE
        /// 
        /// NOTE[2023.06.20.1115]
        /// This Sub moved to module 'app' from 'modMacros'
        /// for better placement in Macro selection dialog.
        /// Also renamed to add underscores, thus spacing
        /// out the title for easier recognition.
        /// 
        // Dim invProgressBar As Inventor.ProgressBar

        gnsIfcAiDoc fc;
        Scripting.Dictionary dc;
        Scripting.Dictionary rt;
        VbMsgBoxResult goAhead;
        Document ActiveDoc;
        string txOut;
        Variant ky;
        Variant kyPt;
        long ct;

        // Dim dx As Long

        fmIfcTest05A fm;

        /// Enable the following Message Box
        /// when procedure in developmental transition.
        // goAhead = MsgBox(        Join(Array(            "Procedure Update_Genius_Properties is under MAJOR revision!",            "Significant changes are in effect which may",            "present issues in routine processing.","Watch for any problems, and be prepared","to respond appropriately."), " "),vbOKOnly + vbCritical,"!!!!! WARNING !!!!!")

        /// Confirm User Request
        /// to process active Document
        goAhead = Constants.vbYes; // '''MsgBox(''''        Join(Array(''''            "Are you sure you want to process this document?",''''            "The process may require a few minutes depending on assembly size.",''''            "Suppressed and excluded parts will not be processed."''''        ), " "),''''        vbYesNo + vbQuestion,''''        "Process Document Custom iProperties"''''    )
        if (goAhead == Constants.vbYes)
        {
            ActiveDoc = ThisApplication.ActiveDocument;
            if (ActiveDoc.DocumentType == kAssemblyDocumentObject)
            {
                /// Check whether User wants to process main document.
                /// Simple part/assembly collections generally
                /// should not be processed.
                goAhead = Constants.vbYes; // '''MsgBox(''''                Join(Array(''''                    "Do you want to process the primary assembly?",''''                    "If the main assembly document is just a collection",''''                    "of separate parts and assemblies to be processed,",''''                    "it's generally best not to include it in processing."''''                ), " "),''''                vbYesNo + vbQuestion,''''                "Include Main Assembly?"''''            )
                if (goAhead == Constants.vbYes)
                    ct = 1;
            }
            else if (ActiveDoc.DocumentType == kPartDocumentObject)
            {
                goAhead = Constants.vbYes;
                ct = 1;
            }

            /// Collect Components for Processing
            /// NOTE[2019.08.22]: Added call to dcRemapByPtNum
            /// against original result to remap Keys
            /// from file names to Part Numbers
            /// !!!WARNING!!! This presents a significant risk
            /// of Key collision, since different models
            /// MIGHT be assigned the same Part Number.
            /// This may be especially true if/when
            /// Bolted Connections become involved.
            /// 
            // dc = dcRemapByPtNum(dcAiDocComponents(ActiveDoc, , ct))
            /// REV[2022.05.24.0956]
            /// replacing the preceding call with the following
            /// section to more effectively manage the collection
            /// process, with potentially more flexibility
            {
                var withBlock = dcAiDocCompSetsByPtNum(ActiveDoc, ct) // AiDoc
       ;
                if (withBlock.Exists(""))
                    System.Diagnostics.Debugger.Break();// for now

                if (withBlock.Exists(2))
                {
                    // THIS situation IS known to occur,
                    // if not TERRIBLY frequently, so a
                    // handler here is a good idea.
                    // 
                    {
                        var withBlock1 = dcOb(withBlock.Item(2));
                        // fortunately, we have one ready made
                        // in the dcRemapByPtNum function this
                        // section is replacing (see above).
                        Debug.Print(MsgBox(Join(Array("The following Part Numbers are", "assigned to more than one Model:", "", Constants.vbTab + Join(withBlock1.Keys, Constants.vbNewLine + Constants.vbTab), ""), Constants.vbNewLine), Constants.vbOKOnly | Constants.vbInformation, "Duplicate Part Numbers!"));
                    }
                }

                // and HERE is the step which ACTUALLY
                // replaces the prior version above.
                // Key 1 is guaranteed to be present
                // in the Dictionary returned, so no
                // need to check for it here.
                dc = dcOb(withBlock.Item(1));
            }

            /// REV[2022.03.23.1344]
            /// disabling weedout filter of known
            /// purchased parts. now that a display
            /// and review form is used to present
            /// the set of processed parts, exclusion
            /// of parts known to be present may
            /// prove a source of confusion
            /// Filter out Purchased Parts Before Proceeding
            /// (implementation pending/in-progress)
            // If ActiveDoc.DocumentType = kAssemblyDocumentObject Then
            // goAhead = MsgBox(Join(Array('        "Some Items may already be recognized",'        "as Purchased Parts in Genius. Models",'        "for these components do not normally",'        "require further updates.",'        "",'        "Would you like to skip these parts?",'        ""'    ), vbNewLine), vbYesNo,'        "Skip Purchased Parts?"'    )
            // 
            // If goAhead = vbYes Then
            // MsgBox Join(Array('            "Purchased Parts in Genius",'            "will not be processed.",'            ""'        ), vbNewLine), vbOKOnly,'            "Skipping Known Purchased"
            // dc = dcKeysMissing(dc,'            dcAiPurch01fromDict('            dcRemapByPtNum(dc)'        ))
            // Else
            // MsgBox Join(Array('            "You will be prompted to verify",'            "all Purchased Parts, including",'            "any already in Genius.",'            ""'        ), vbNewLine), vbOKOnly,'        "Including ALL Purchased"
            // End If
            // Else 'nevermind
            // '   don't need to check single parts
            // '   for purchased components
            // End If

            /// REV[2022.03.14.1135]
            /// Adding subdivision of gathered
            /// Items into subgroups by form:
            /// MAYB - probable R-RTM #parts : D-BAR #rawStock
            /// #subtype #shtMetal indicates SHTM
            /// but #invalid #flatPattern
            /// suggests otherwise
            /// DBAR - definite R-RTM #parts : D-BAR #rawStock
            /// #subtype NOT #shtMetal
            /// SHTM - D-RTM #parts : DSHEET #rawStock
            /// #subtype #shtMetal
            /// with #valid #flatPattern
            /// ASSY - D/R-MTO #assemblies
            /// PRCH - D/R-PTS/O #purchased #items
            /// HDWR - D-HDWR #hardware #items
            // dc = dcAiDocGrpsByForm(dc)

            /// Create a new ProgressBar object.
            /// REV[2022.03.14.1137]
            /// Disabling Progress Bar to avoid
            /// complications arising from new
            /// subdivision. MIGHT restore later.
            // ct = .Count
            // dx = 0
            // invProgressBar = ThisApplication.CreateProgressBar(True, ct, "Progressing: ")


            fc = new gnsIfcAiDoc();
            /// REV[2022.03.16.1318]

            fm = nu_fmIfcTest05A(dc); // nu_fmTest05A
            /// REV[2022.03.22.1448]
            /// REV[2022.03.17.1324]
            /// REV[2022.02.09.0829]
            rt = new Scripting.Dictionary();

            fm.Show(vbModeless);
            /// REV[2022.03.17.1354]
            /// REV[2022.03.14.1526]
            {
                var withBlock = dcAiDocGrpsByForm(dc);
                /// Process the full Component Collection
                foreach (var ky in Array("ASSY", "SHTM", "MAYB", "DBAR", "PRCH"))
                {
                    /// note how we're also skipping "HDWR" entirely
                    /// also plan on handling "PRCH" items separately

                    /// REV[2022.02.09.1432]
                    if (withBlock.Exists(ky))
                    {
                        if (dcOb(withBlock.Item(ky)).Count > 0)
                        {
                            if (fm.InGroup(System.Convert.ToHexString(ky)).GroupNow == ky)
                            {
                                /// REV[2022.03.22.1225]
                                /// REV[2022.03.17.1339]
                                {
                                    var withBlock1 = dcOb(withBlock.Item(ky));
                                    foreach (var kyPt in withBlock1.Keys)
                                    {
                                        /// REV[2022.03.22.1246]

                                        /// Update message for the progress bar
                                        /// REV[2022.03.14.1140]
                                        /// Disabling Progress Bar updates
                                        /// per REV[2022.03.14.1137] (above)
                                        // dx = 1 + dx
                                        // With invProgressBar
                                        // .Message'        = "Processing - " & ky'        & " - " & dx'        & "/" & ct
                                        // .UpdateProgress
                                        // End With

                                        /// WAYPOINT:UPDATE
                                        /// Process Genius Properties for next Item
                                        /// THIS is where ALL the magic happens!

                                        /// REV[2022.03.22.1246]
                                        if (fm.OnItem(System.Convert.ToHexString(kyPt)).ItemNow == kyPt)
                                        {
                                            Information.Err.Clear();
                                            if (false)
                                                rt.Add(kyPt, dcGeniusProps(withBlock1.Item(kyPt)));
                                            else
                                                rt.Add(kyPt, fc.Props(aiDocument(withBlock1.Item(kyPt))));

                                            if (Information.Err.Number == 0)
                                            {
                                                {
                                                    var withBlock2 = dcOb(rt.Item(kyPt));
                                                    withBlock2.Add("FORM", ky);
                                                }
                                            }
                                        }
                                        else
                                            System.Diagnostics.Debugger.Break();
                                    }
                                }
                                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                            }
                            else
                                System.Diagnostics.Debugger.Break();
                        }
                    }
                }

                Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
            }

            /// REV[2022.02.09.0844]
            {
                var withBlock = rt;
                /// Dump all Processing Results
                /// REVISION[2021.08.09]
                foreach (var ky in withBlock.Keys)
                    withBlock.Item(ky) = dcPropVals(dcOb(withBlock.Item(ky)));
            }

            /// NOTE!![2022.03.14.1522]
            /// ''' following section to matching NOTE!! below
            /// ''' will require review to address changes
            /// ''' resulting from addition of FORM
            /// ''' subleveling
            /// REV[2022.02.09.0847]
            dc = dcKeysMissing(dc, rt);

            rt = dcRecSetDcDx4json(dcDxFromRecSetDc(rt));

            /// REV[2022.02.09.0847]
            if (dc.Count > 0)
            {
                System.Diagnostics.Debugger.Break(); // so we can check how this is going to work
                rt.Add("NOTPROCESSED", dc.Keys);
            }
            /// NOTE!![2022.03.14.1522]
            /// section above ENDS here

            txOut = ConvertToJson(Array("[[ DELETE THIS PLACEHOLDER (KEEP COMMA IF NEEDED) ]]", rt), Constants.vbTab); // dc
        }
    }

    public void Update_iPtAssy_Genius_Props()
    {
        Inventor.Document md;
        Scripting.Dictionary dc;
        VbMsgBoxResult ck;

        md = aiDocActive();
        dc = gnsUpdtAll_iFact(compDefOf(md));
        if (dc.Count > 0)
        {
        }
        else
            ck = MsgBox(Join(Array("", ""), Constants.vbNewLine), Constants.vbOKOnly, "");
    }

    /// 

    /// 
    private string app()
    {
        app = Array("module app version date 2023.06.20", "")(0);
    }
}