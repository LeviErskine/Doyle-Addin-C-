class lib1
{
    public string Repeat(long Count, string Text)
    {
        Repeat = Replace(Space(Count), " ", Text);
    }

    public string txBlk(long Lines, long Chars, string Use = "+")
    {
        txBlk = Mid(Repeat(Lines, Constants.vbNewLine + String(Chars, "+")), 1 + Strings.Len(Constants.vbNewLine));
    }

    public void MakeActivePurchased()
    {
        Inventor.Document md;
        VbMsgBoxResult ck;

        md = ThisApplication.ActiveDocument;
        if (md == ThisDocument)
            ck = Constants.vbNo;
        else
            ck = mkAiDocPurchased(md);

        if (ck == Constants.vbOK)
            ck = MsgBox(Join(Array("Model BOM Structure", "now Purchased."), Constants.vbNewLine), Constants.vbOKOnly + Constants.vbInformation, "Success!");
        else if (ck == Constants.vbNo)
            ck = MsgBox(Join(Array("Document is not", "a valid Model.", "", "Please select a", "Part or Assembly."), Constants.vbNewLine), Constants.vbOKOnly + Constants.vbExclamation, "No Model");
        else if (ck == Constants.vbAbort)
            ck = MsgBox(Join(Array("Failed to update", "model's BOM Structure!", "", "Check for locks", "or other issues."), Constants.vbNewLine), Constants.vbOKOnly + Constants.vbCritical, "Change Failed!");
        else
            ck = MsgBox(Join(Array("Change Operation returned", "unexpected result code.", "", "Please review model status."), Constants.vbNewLine), Constants.vbOKOnly + Constants.vbQuestion, "Result Unknown");
    }

    public VbMsgBoxResult mkAiDocPurchased(Inventor.Document AiDoc)
    {
        VbMsgBoxResult ck;

        if (AiDoc is Inventor.PartDocument)
            ck = mkAiPartPurchased(AiDoc);
        else if (AiDoc is Inventor.AssemblyDocument)
            ck = mkAiAssyPurchased(AiDoc);
        else
            ck = Constants.vbNo;

        mkAiDocPurchased = ck;
    }

    public VbMsgBoxResult mkAiPartPurchased(Inventor.PartDocument AiDoc)
    {
        if (AiDoc == null)
            mkAiPartPurchased = Constants.vbNo;
        else
        {
            var withBlock = AiDoc.ComponentDefinition;
            Information.Err.Clear();
            withBlock.BOMStructure = kPurchasedBOMStructure;
            if (Information.Err.Number == 0)
                mkAiPartPurchased = Constants.vbOK;
            else
                mkAiPartPurchased = Constants.vbAbort;
        }
    }

    public VbMsgBoxResult mkAiAssyPurchased(Inventor.AssemblyDocument AiDoc)
    {
        if (AiDoc == null)
            mkAiAssyPurchased = Constants.vbNo;
        else
        {
            var withBlock = AiDoc.ComponentDefinition;
            Information.Err.Clear();
            withBlock.BOMStructure = kPurchasedBOMStructure;
            if (Information.Err.Number == 0)
                mkAiAssyPurchased = Constants.vbOK;
            else
                mkAiAssyPurchased = Constants.vbAbort;
        }
    }

    public Scripting.Dictionary dcTemplate0A(Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;

        if (dc == null)
            rt = dcTemplate0A(new Scripting.Dictionary());
        else
            rt = dc;

        dcTemplate0A = rt;
    }

    public Variant send2clipBd_OBSOLETE(Variant src)
    {
        {
            var withBlock = new MSForms.DataObject();
            withBlock.SetText(src);
            withBlock.PutInClipboard();
        }
        send2clipBd_OBSOLETE = src;
    }
}