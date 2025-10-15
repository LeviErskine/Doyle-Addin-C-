using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class app
{
    public void Update_Genius_Properties()
    {
        // use Find (Ctrl-F or H) to jump to:
        // WAYPOINT:UPDATE
        // 
        // NOTE[2023.06.20.1115]
        // This Sub moved to module 'app' from 'modMacros'
        // for better placement in Macro selection dialog.
        // Also renamed to add underscores, thus spacing
        // out the title for easier recognition.
        // 
        // Dim invProgressBar As Inventor.ProgressBar

        // Dim dx As Long

        // Enable the following Message Box
        // when procedure in developmental transition.
        // goAhead = MessageBox.Show( Join(new string[] { "Procedure Update_Genius_Properties is under MAJOR revision!", "Significant changes are in effect which may", "present issues in routine processing.","Watch for any problems, and be prepared","to respond appropriately."), " "),vbOKOnly + vbCritical,"!!!!! WARNING !!!!!")
        // Confirm User Request
        // to process active Document
        var
            goAhead = Constants
                .vbYes; // '''MessageBox.Show('''' Join(new string[] {'''' "Are you sure you want to process this document?",'''' "The process may require a few minutes depending on assembly size.",'''' "Suppressed and excluded parts will not be processed."'''' ), " "),'''' vbYesNo + vbQuestion,'''' "Process Document Custom iProperties"'''' )
        if (false) return;
        {
            Document ActiveDoc = ThisApplication.ActiveDocument;
            long ct;
            switch (ActiveDoc.DocumentType)
            {
                case kAssemblyDocumentObject:
                {
                    // Check whether User wants to process main document.
                    // Simple part/assembly collections generally
                    // should not be processed.
                    goAhead = Constants
                        .vbYes; // '''MessageBox.Show('''' Join(new string[] {'''' "Do you want to process the primary assembly?",'''' "If the main assembly document is just a collection",'''' "of separate parts and assemblies to be processed,",'''' "it's generally best not to include it in processing."'''' ), " "),'''' vbYesNo + vbQuestion,'''' "Include Main Assembly?"'''' )
                    if (true)
                        ct = 1;
                    break;
                }
                case kPartDocumentObject:
                    goAhead = Constants.vbYes;
                    ct = 1;
                    break;
                case kUnknownDocumentObject:
                case kDrawingDocumentObject:
                case kPresentationDocumentObject:
                case kDesignElementDocumentObject:
                case kForeignModelDocumentObject:
                case kSATFileDocumentObject:
                case kNoDocument:
                case kNestingDocument:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            // Collect Components for Processing
            // NOTE[2019.08.22]: Added call to dcRemapByPtNum
            // against original result to remap Keys
            // from file names to Part Numbers
            // !!!WARNING!!! This presents a significant risk
            // of Key collision, since different models
            // MIGHT be assigned the same Part Number.
            // This may be especially true if/when
            // Bolted Connections become involved.
            // 
            // dc = dcRemapByPtNum(dcAiDocComponents(ActiveDoc, , ct))
            // REV[2022.05.24.0956]
            // replacing the preceding call with the following
            // section to more effectively manage the collection
            // process, with potentially more flexibility
            Dictionary dc;
            {
                var withBlock = dcAiDocCompSetsByPtNum(ActiveDoc, ct) // AiDoc
                    ;
                if (withBlock.Exists(""))
                    Debugger.Break(); // for now

                if (withBlock.Exists(2))
                {
                    // THIS situation IS known to occur,
                    // if not TERRIBLY frequently, so a
                    // handler here is a good idea.
                    // 
                    {
                        var withBlock1 = dcOb(withBlock.get_Item(2));
                        // fortunately, we have one ready made
                        // in the dcRemapByPtNum function this
                        // section is replacing (see above).
                        Debug.Print(MessageBox.Show(
                            Join(
                                new[]
                                {
                                    "The following Part Numbers are", "assigned to more than one Model:", "",
                                    Constants.vbTab + Join(withBlock1.Keys, Constants.vbCrLf + Constants.vbTab), ""
                                }, Constants.vbCrLf), Constants.vbOKOnly | Constants.vbInformation,
                            "Duplicate Part Numbers!"));
                    }
                }

                // and HERE is the step which ACTUALLY
                // replaces the prior version above.
                // Key 1 is guaranteed to be present
                // in the Dictionary returned, so no
                // need to check for it here.
                dc = dcOb(withBlock.get_Item(1));
            }

            // REV[2022.03.23.1344]
            // disabling weedout filter of known
            // purchased parts. now that a display
            // and review form is used to present
            // the set of processed parts, exclusion
            // of parts known to be present may
            // prove a source of confusion
            // Filter out Purchased Parts Before Proceeding
            // (implementation pending/in-progress)
            // If ActiveDoc.DocumentType = kAssemblyDocumentObject Then
            // goAhead = MessageBox.Show(Join(new string[] {' "Some Items may already be recognized",' "as Purchased Parts in Genius. Models",' "for these components do not normally",' "require further updates.",' "",' "Would you like to skip these parts?",' ""' ), vbCrLf), vbYesNo,' "Skip Purchased Parts?"' )
            // 
            // If goAhead = vbYes Then
            // MessageBox.Show Join(new string[] {' "Purchased Parts in Genius",' "will not be processed.",' ""' ), vbCrLf), vbOKOnly,' "Skipping Known Purchased"
            // dc = dcKeysMissing(dc,' dcAiPurch01fromDict(' dcRemapByPtNum(dc)' ))
            // Else
            // MessageBox.Show Join(new string[] {' "You will be prompted to verify",' "all Purchased Parts, including",' "any already in Genius.",' ""' ), vbCrLf), vbOKOnly,' "Including ALL Purchased"
            // End If
            // Else 'nevermind
            // ' don't need to check single parts
            // ' for purchased components
            // End If

            // REV[2022.03.14.1135]
            // Adding subdivision of gathered
            // Items into subgroups by form:
            // MAYB - probable R-RTM #parts : D-BAR #rawStock
            // #subtype #shtMetal indicates SHTM
            // but #invalid #flatPattern
            // suggests otherwise
            // DBAR - definite R-RTM #parts : D-BAR #rawStock
            // #subtype NOT #shtMetal
            // SHTM - D-RTM #parts : DSHEET #rawStock
            // #subtype #shtMetal
            // with #valid #flatPattern
            // ASSY - D/R-MTO #assemblies
            // PRCH - D/R-PTS/O #purchased #items
            // HDWR - D-HDWR #hardware #items
            // dc = dcAiDocGrpsByForm(dc)

            // Create a new ProgressBar dynamic.
            // REV[2022.03.14.1137]
            // Disabling Progress Bar to avoid
            // complications arising from new
            // subdivision. MIGHT restore later.
            // ct = .Count
            // dx = 0
            // invProgressBar = ThisApplication.CreateProgressBar(True, ct, "Progressing: ")

            var fc = new gnsIfcAiDoc();
            // REV[2022.03.16.1318]

            var fm = nu_fmIfcTest05A(dc);
            // REV[2022.03.22.1448]
            // REV[2022.03.17.1324]
            // REV[2022.02.09.0829]
            var rt = new Dictionary();

            fm.Show(vbModeless);
            // REV[2022.03.17.1354]
            // REV[2022.03.14.1526]
            {
                var withBlock = dcAiDocGrpsByForm(dc);
                // Process the full Component Collection
                foreach (var ky in new[] { "ASSY", "SHTM", "MAYB", "DBAR", "PRCH" })
                {
                    // note how we're also skipping "HDWR" entirely
                    // also plan on handling "PRCH" items separately

                    // REV[2022.02.09.1432]
                    if (!withBlock.Exists(ky)) continue;
                    if (dcOb(withBlock.get_Item(ky)).Count <= 0) continue;
                    if (fm.InGroup(Convert.ToHexString(ky)).GroupNow == ky)
                    {
                        // REV[2022.03.22.1225]
                        // REV[2022.03.17.1339]
                        {
                            var withBlock1 = dcOb(withBlock.get_Item(ky));
                            foreach (var kyPt in withBlock1.Keys)
                            {
                                // REV[2022.03.22.1246]

                                // Update message for the progress bar
                                // REV[2022.03.14.1140]
                                // Disabling Progress Bar updates
                                // per REV[2022.03.14.1137] (above)
                                // dx = 1 + dx
                                // With invProgressBar
                                // .Message' = "Processing - " & ky' & " - " & dx' & "/" & ct
                                // .UpdateProgress
                                // End With

                                // WAYPOINT:UPDATE
                                // Process Genius Properties for next Item
                                // THIS is where ALL the magic happens!

                                // REV[2022.03.22.1246]
                                if (fm.OnItem(Convert.ToString(kyPt as string)).ItemNow == (Func<string>)kyPt)
                                {
                                    Information.Err().Clear();
                                    if (false)
                                        rt.Add(kyPt, dcGeniusProps(withBlock1.get_Item(kyPt)));
                                    rt.Add(kyPt, fc.Props(aiDocument(withBlock1.get_Item(kyPt))));

                                    if (Information.Err().Number != 0) continue;
                                    {
                                        var withBlock2 = dcOb(rt.get_Item(kyPt));
                                        withBlock2.Add("FORM", ky);
                                    }
                                }
                                else
                                    Debugger.Break();
                            }
                        }
                        Debug.Print(""); // Breakpoint Landing
                    }
                    else
                        Debugger.Break();
                }

                Debug.Print(""); // Breakpoint Landing
            }

            // REV[2022.02.09.0844]
            {
                // Dump all Processing Results
                // REVISION[2021.08.09]
                foreach (var ky in rt.Keys)
                    rt.get_Item(ky) = dcPropVals(dcOb(rt.get_Item(ky)));
            }

            // NOTE!![2022.03.14.1522]
            // ''' following section to matching NOTE!! below
            // ''' will require review to address changes
            // ''' resulting from addition of FORM
            // ''' subleveling
            // REV[2022.02.09.0847]
            dc = dcKeysMissing(dc, rt);

            rt = dcRecSetDcDx4json(dcDxFromRecSetDc(rt));

            // REV[2022.02.09.0847]
            if (dc.Count > 0)
            {
                Debugger.Break(); // so we can check how this is going to work
                rt.Add("NOTPROCESSED", dc.Keys);
            }

            // NOTE!![2022.03.14.1522]
            // section above ENDS here
            var txOut = ConvertToJson(new[] { "[[ DELETE THIS PLACEHOLDER (KEEP COMMA IF NEEDED) ]]", rt },
                Constants.vbTab); // dc
        }
    }

    public void Update_iPtAssy_Genius_Props()
    {
        var md = aiDocActive();
        var dc = gnsUpdtAll_iFact(compDefOf(md));
        if (dc.Count > 0)
        {
        }
        else
            MsgBoxResult ck = MessageBox.Show(
                Join(Environment.NewLine, "", ""),
                @"",
                MessageBoxButtons.OK
            );
    }

    // 

    // 
    string app()
    {
        return new string[] { "module app version date 2023.06.20", "" }[0];
    }
}