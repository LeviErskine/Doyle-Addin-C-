class dvlGnsIfc201904
{
    public Scripting.Dictionary dgiG0t0()
    {
        Scripting.Dictionary dcTree;
        Scripting.Dictionary dcFlat;
        string nm;
        string dt;

        nm = nuSelAiDoc().GetReply();
        if (Strings.Len(Strings.Trim(nm)) > 0)
        {
            {
                var withBlock = ThisApplication.Documents;
                dcTree = dgiAiDocClassified(withBlock.ItemByName(nm));
                dt = dgiG2f2(dgiG2f1(dcTree));
                if (MsgBox("Send this text to the clipoard?" + Constants.vbNewLine + Constants.vbNewLine + dt, Constants.vbYesNo + Constants.vbQuestion, "Send to Clipboard?") == Constants.vbYes)
                {
                    Information.Err.Clear();
                    send2clipBdWin10(dt);
                    if (Information.Err.Number == 0)
                        // MsgBox "PROMPT", vbOKOnly, "TITLE"
                        MsgBox(System.Convert.ToHexString(Strings.Len(dt)) + " characters" + vbNewLine, Constants.vbOKOnly, "COPY SUCCESSFUL!");
                    else if (MsgBox("Error Code " + Hex(Information.Err.Number) + ":" + Constants.vbNewLine + Information.Err.Description + vbNewLine, Constants.vbYesNo, "COPY FAILED!") == Constants.vbYes)
                        System.Diagnostics.Debugger.Break();
                }
                else
                    MsgBox("No data sent to clipboard", Constants.vbOKOnly, "COPY CANCELED");
            }
        }
    }

    public Scripting.Dictionary dgiAiDocClassified(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        // '
        // '  Classify supplied Inventor Document
        // '  by basic Document Type. Retrieve or
        // '  generate sub Dictionary associated
        // '  with Document Type, and reference
        // '  Document there by its Full Name/Path
        // '
        Inventor.DocumentTypeEnum dt;
        string fp;
        // Dim st As String

        if (dc == null)
            dgiAiDocClassified = dgiAiDocClassified(AiDoc, new Scripting.Dictionary());
        else
        {
            {
                var withBlock = AiDoc;
                fp = withBlock.FullDocumentName;
                dt = withBlock.DocumentType;
            }

            if (Strings.Len(fp) > 0)
            {
                {
                    var withBlock = dc;
                    if (!withBlock.Exists(dt))
                        withBlock.Add(dt, new Scripting.Dictionary());
                    {
                        var withBlock1 = dcOb(withBlock.Item(dt));
                        if (!withBlock1.Exists(fp))
                            withBlock1.Add(fp, AiDoc);
                    }
                }
            }

            if (dt == kAssemblyDocumentObject)
                dgiAiDocClassified = dgiMembersClassified(AiDoc, dc);
            else
                dgiAiDocClassified = dc;
        }
    }

    public Scripting.Dictionary dgiMembersClassified(Inventor.AssemblyDocument AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        // '
        // '  Given an Assembly Document,
        // '  categorize its Components.
        // '
        Inventor.ComponentOccurrence oc;
        Scripting.Dictionary rt;

        rt = dc;
        {
            var withBlock = AiDoc.ComponentDefinition;
            foreach (var oc in withBlock.Occurrences)
            {
                {
                    var withBlock1 = oc.Definition;
                    rt = dgiAiDocClassified(withBlock1.Document, rt);
                }
            }
        }
        dgiMembersClassified = rt;
    }

    public Scripting.Dictionary dgiFlatListed(Scripting.Dictionary dc)
    {
        // '
        // '  Flatten Dictionary
        // '  of Dictionaries of
        // '  Inventor Documents
        // '  into one singular
        // '  Dictionary for rescan.
        // '
        Scripting.Dictionary rt;
        Variant ky;
        Variant fp;
        long ct;

        rt = new Scripting.Dictionary();

        if (dc == null)
        {
        }
        else
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcOb(withBlock.Item(ky));
                    foreach (var fp in withBlock1.Keys)
                        rt.Add(fp, withBlock1.Item(fp));
                }
            }
        }

        dgiFlatListed = rt;
    }

    public fmSelectorList nuSelAiDoc(object As = string == "%$#@*&!")
    {
        var nuSelAiDoc = new nuSelector()
            .SetHdrCancel("Cancel Operation?")
            .SetHdrNoSelection("No Document Selected!")
            .SetHdrOK("Proceed With Operation?")
            .SetMsgCancel("No changes will be applied to any open Document.")
            .SetMsgNoSelection(
                string.Join(
                    Environment.NewLine,
                    new[] { "Do you wish to cancel the Operation?", "(Click NO to return to list)" }
                )
            )
            .SetMsgOK(
                string.Join(
                    Environment.NewLine,
                    new[] { "The following Document(s) will be affected: ", "%%%", "(Click CANCEL to quit with no changes)" }
                )
            )
            .WithList(dcAiDocsVisible().Keys)
            .SelectIfIn(Default);
    }

    public Inventor.Document AskUser4aiDoc(object As = Inventor.Document == null/* TODO Change to default(_) if this is not a reference type */, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        string nm;

        if (dc == null)
            AskUser4aiDoc = AskUser4aiDoc(Default, dcAiDocsVisible());
        else
if (Default == null)
        {
        }
        result = AskUser4aiDoc(ThisApplication.ActiveDocument, dc);
    }
else
{
    string nm = d0g6f0(Default);
    nm = nuSelAiDoc()
            .WithList(dc.Keys)
            .SelectIfIn(nm)
            .GetReply();
        if (dc.Exists(nm)) 
    {
                result = dc[nm]; 
    }
    else
{
    result = null;
}
}
    /// 

    /// 
    public Variant dgiG0f0()
    {
        dgiG0f0 = Empty;
    }

    public Scripting.Dictionary dgiG0f1(Inventor.Document AiDoc, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        // '
        // '  "Junk" function originally intended to
        // '  collect and categorize Inventor Documents.
        // '  See following functions for preferred approach.
        // '
        if (dc == null)
            dgiG0f1 = dgiG0f1(AiDoc, new Scripting.Dictionary());
        else
        {
            {
                var withBlock = AiDoc;
                if (withBlock.DocumentType == kAssemblyDocumentObject)
                {
                    if (!dc.Exists(kAssemblyDocumentObject))
                        dc.Add(kAssemblyDocumentObject, new Scripting.Dictionary());
                    {
                        var withBlock1 = dcOb(dc.Item(kAssemblyDocumentObject));
                        if (!withBlock1.Exists(AiDoc.FullDocumentName))
                            withBlock1.Add(AiDoc.FullDocumentName, AiDoc);
                    }
                }
                else if (withBlock.DocumentType == kPartDocumentObject)
                {
                }
                else
                {
                    System.Diagnostics.Debugger.Break();
                    if (withBlock.DocumentType == kDesignElementDocumentObject)
                        System.Diagnostics.Debugger.Break();
                    else if (withBlock.DocumentType == kDrawingDocumentObject)
                        System.Diagnostics.Debugger.Break();
                    else if (withBlock.DocumentType == kForeignModelDocumentObject)
                        System.Diagnostics.Debugger.Break();
                    else if (withBlock.DocumentType == kNoDocument)
                        System.Diagnostics.Debugger.Break();
                    else if (withBlock.DocumentType == kPresentationDocumentObject)
                        System.Diagnostics.Debugger.Break();
                    else if (withBlock.DocumentType == kSATFileDocumentObject)
                        System.Diagnostics.Debugger.Break();
                    else if (withBlock.DocumentType == kUnknownDocumentObject)
                        System.Diagnostics.Debugger.Break();
                }
            }
            dgiG0f1 = dc;
        }
    }

    public long dgiG1f0(Scripting.Dictionary dc)
    {
        // '
        // '  Return the grand total count
        // '  of entries in all Dictionaries
        // '  within supplied Dictionary.
        // '
        // '  This is meant to check for
        // '  any additions to the collection
        // '  after each processing pass
        // '
        Variant ky;
        long ct;

        ct = 0;

        if (dc == null)
        {
        }
        else
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcOb(withBlock.Item(ky));
                    ct = ct + withBlock1.Count;
                }
            }
        }

        dgiG1f0 = ct;
    }

    public Scripting.Dictionary dgiG1f1(Scripting.Dictionary dc, long ck = -1)
    {
        // '
        // '  Build up Dictionary of Inventor
        // '  Part and Assembly Documents
        // '
        Inventor.Document AiDoc;
        Scripting.Dictionary rt;
        Variant ky;
        Variant fp;
        long ct;

        if (dc == null)
            dgiG1f1 = new Scripting.Dictionary();
        else if (ck < 0)
            dgiG1f1 = dgiG1f1(dc, dgiG1f0(dc));
        else
        {
            rt = dc;
            {
                var withBlock = dgiFlatListed(rt);
                foreach (var ky in withBlock.Keys)
                {
                    AiDoc = aiDocument(obOf(withBlock.Item(ky)));
                    rt = dgiAiDocClassified(AiDoc, rt);

                    if (AiDoc.DocumentType == kAssemblyDocumentObject)
                        rt = dgiMembersClassified(AiDoc, rt);
                }
            }
            ct = dgiG1f0(dc);

            if (ct > ck)
                dgiG1f1 = dgiG1f1(rt, ct);
            else if (ct == ck)
                dgiG1f1 = rt;
            else
                System.Diagnostics.Debugger.Break();// cuz something went wrong
        }
    }

    public Scripting.Dictionary dgiG2f0(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Inventor.Document id;
        Variant ky;
        string sb;
        string fp;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                id = aiDocument(withBlock.Item(ky));

                {
                    var withBlock1 = id;
                    fp = withBlock1.FullDocumentName;
                    sb = withBlock1.SubType;
                }

                {
                    var withBlock1 = rt;
                    if (!withBlock1.Exists(sb))
                        withBlock1.Add(sb, new Scripting.Dictionary());
                    {
                        var withBlock2 = dcOb(withBlock1.Item(sb));
                        if (!withBlock2.Exists(fp))
                            withBlock2.Add(fp, id);
                    }
                }
            }
        }
        dgiG2f0 = rt;
    }

    public Scripting.Dictionary dgiG2f1(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Inventor.Document id;
        Variant ky;
        string sb;
        string fp;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
                rt.Add(ky, dgiG2f0(dcOb(withBlock.Item(ky))));
        }
        dgiG2f1 = rt;
    }

    public string dgiG2f2(Scripting.Dictionary dc, string pfx = "", string dlm = "|", string brk = Constants.vbNewLine)
    {
        Scripting.Dictionary rt;
        Variant ky;
        Variant it;
        string tx;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                tx = System.Convert.ToHexString(ky);
                if (Strings.Len(pfx) > 0)
                    tx = pfx + dlm + tx;
                if (withBlock.Item(ky) is Scripting.Dictionary)
                    rt.Add(dgiG2f2(dcOb(obOf(withBlock.Item(ky))), tx, dlm, brk), 0);
                else
                    rt.Add(tx, 0);
            }
        }
        dgiG2f2 = Join(rt.Keys, brk);
    }
}