using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class dvlGnsIfc201904
{
    public static Dictionary dgiG0t0()
    {
        Dictionary dcFlat;

        var nm = nuSelAiDoc().GetReply();
        if (Strings.Len(Strings.Trim(nm)) > 0)
        {
            {
                var withBlock = ThisApplication.Documents;
                var dcTree = dgiAiDocClassified(withBlock.ItemByName(nm));
                var dt = dgiG2f2(dgiG2f1(dcTree));
                if (MessageBox.Show("Send this text to the clipoard?" + Constants.vbCrLf + Constants.vbCrLf + dt,
                        Constants.vbYesNo + Constants.vbQuestion, "Send to Clipboard?") == Constants.vbYes)
                {
                    Information.Err().Clear();
                    send2clipBdWin10(dt);
                    if (Information.Err().Number == 0)
                        // MessageBox.Show "PROMPT", vbOKOnly, "TITLE"
                        MessageBox.Show(Convert.ToHexString(Strings.Len(dt)) + " characters" + vbCrLf,
                            Constants.vbOKOnly, "COPY SUCCESSFUL!");
                    else if (MessageBox.Show(
                                 "Error Code " + Hex(Information.Err().Number) + ":" + Constants.vbCrLf +
                                 Information.Err().Description + vbCrLf, Constants.vbYesNo, "COPY FAILED!") ==
                             Constants.vbYes)
                        Debugger.Break();
                }
                else
                    MessageBox.Show(@"No data sent to clipboard", @"COPY CANCELED", MessageBoxButtons.OK);
            }
        }
    }

    public  static Dictionary dgiAiDocClassified(Document AiDoc, Dictionary dc = null)
    {
        while (true)
        {
            // '
            // ' Classify supplied Inventor Document
            // ' by basic Document Type. Retrieve or
            // ' generate sub Dictionary associated
            // ' with Document Type, and reference
            // ' Document there by its Full Name/Path
            // '
            DocumentTypeEnum dt;
            string fp;
            // Dim st As String

            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            {
                fp = AiDoc.FullDocumentName;
                dt = AiDoc.DocumentType;
            }

            if (Strings.Len(fp) <= 0) return dt == kAssemblyDocumentObject ? dgiMembersClassified(AiDoc, dc) : dc;
            {
                {
                    if (!dc.Exists(dt)) dc.Add(dt, new Dictionary());
                    {
                        var withBlock1 = dcOb(dc.get_Item(dt));
                        if (!withBlock1.Exists(fp)) withBlock1.Add(fp, AiDoc);
                    }
                }
            }

            return dt == kAssemblyDocumentObject ? dgiMembersClassified(AiDoc, dc) : dc;
        }
    }

    public  static Dictionary dgiMembersClassified(AssemblyDocument AiDoc, Dictionary dc = null)
    {
        // '
        // ' Given an Assembly Document,
        // ' categorize its Components.
        // '

        var rt = dc;
        {
            var withBlock = AiDoc.ComponentDefinition;
            foreach (ComponentOccurrence oc in withBlock.Occurrences)
            {
                {
                    var withBlock1 = oc.Definition;
                    rt = dgiAiDocClassified(withBlock1.Document, rt);
                }
            }
        }
        return rt;
    }

    public static  Dictionary dgiFlatListed(Dictionary dc)
    {
        // '
        // ' Flatten Dictionary
        // ' of Dictionaries of
        // ' Inventor Documents
        // ' into one singular
        // ' Dictionary for rescan.
        // '
        long ct;

        var rt = new Dictionary();

        if (dc == null)
        {
        }
        else
        {
            foreach (dynamic ky in dc.Keys)
            {
                {
                    var withBlock1 = dcOb(dc.get_Item(ky));
                    foreach (var fp in withBlock1.Keys)
                        rt.Add(fp, withBlock1.get_Item(fp));
                }
            }
        }

        return rt;
    }

    public static  fmSelectorList nuSelAiDoc(dynamic As = string == "%$#@*&!")
    {
        var nuSelAiDoc = new nuSelector()
            .SetHdrCancel("Cancel Operation?")
            .SetHdrNoSelection("No Document Selected!")
            .SetHdrOK("Proceed With Operation?")
            .SetMsgCancel("No changes will be applied to any open Document.")
            .SetMsgNoSelection(
                Join(
                    Environment.NewLine, "Do you wish to cancel the Operation?", "(Click NO to return to list)")
            )
            .SetMsgOK(
                Join(
                    Environment.NewLine, "The following Document(s) will be affected: ", "%%%",
                    "(Click CANCEL to quit with no changes)")
            )
            .WithList(dcAiDocsVisible().Keys)
            .SelectIfIn(Default);
    }

    public static  Document AskUser4aiDoc(dynamic As = Document == null, Dictionary dc = null)
    {
        while (true)
        {
            if (dc == null)
            {
                dc = dcAiDocsVisible();
                continue;
            }

            if (Default == null)
            {
            }

            continue;
            var nm = d0g6f0(Default);
            nm = nuSelAiDoc()
                .WithList(dc.Keys)
                .SelectIfIn(nm)
                .GetReply();
            return dc.Exists(nm) ? (Document)dc[nm] : null;
            break;
        }
    }
    // 

    // 
    public static  dynamic dgiG0f0()
    {
        return null;
    }

    public static  Dictionary dgiG0f1(Document AiDoc, Dictionary dc = null)
    {
        while (true)
        {
            // '
            // ' "Junk" function originally intended to
            // ' collect and categorize Inventor Documents.
            // ' See following functions for preferred approach.
            // '
            if (dc == null)
            {
                dc = new Dictionary();
                continue;
            }

            {
                switch (AiDoc.DocumentType)
                {
                    case kAssemblyDocumentObject:
                    {
                        if (!dc.Exists(kAssemblyDocumentObject)) dc.Add(kAssemblyDocumentObject, new Dictionary());
                        {
                            var withBlock1 = dcOb(dc.get_Item(kAssemblyDocumentObject));
                            if (!withBlock1.Exists(AiDoc.FullDocumentName))
                                withBlock1.Add(AiDoc.FullDocumentName, AiDoc);
                        }
                        break;
                    }
                    case kPartDocumentObject:
                        break;
                    case kUnknownDocumentObject:
                    case kDrawingDocumentObject:
                    case kPresentationDocumentObject:
                    case kDesignElementDocumentObject:
                    case kForeignModelDocumentObject:
                    case kSATFileDocumentObject:
                    case kNoDocument:
                    case kNestingDocument:
                    default:
                        Debugger.Break();
                        switch (AiDoc.DocumentType)
                        {
                            case kDesignElementDocumentObject:
                            case kDrawingDocumentObject:
                            case kForeignModelDocumentObject:
                            case kNoDocument:
                            case kPresentationDocumentObject:
                            case kSATFileDocumentObject:
                            case kUnknownDocumentObject:
                                Debugger.Break();
                                break;
                            case kPartDocumentObject:
                            case kAssemblyDocumentObject:
                            case kNestingDocument:
                                break;
                            default:
                                throw new ArgumentOutOfRangeException();
                        }

                        break;
                }
            }
            return dc;
        }
    }

    public static  long dgiG1f0(Dictionary dc)
    {
        // '
        // ' Return the grand total count
        // ' of entries in all Dictionaries
        // ' within supplied Dictionary.
        // '
        // ' This is meant to check for
        // ' any additions to the collection
        // ' after each processing pass
        // '

        long ct = 0;

        if (dc == null)
        {
        }
        else
        {
            foreach (dynamic ky in dc.Keys)
            {
                {
                    var withBlock1 = dcOb(dc.get_Item(ky));
                    ct = ct + withBlock1.Count;
                }
            }
        }

        return ct;
    }

    public static  Dictionary dgiG1f1(Dictionary dc, long ck = -1)
    {
        while (true)
        {
            // '
            // ' Build up Dictionary of Inventor
            // ' Part and Assembly Documents
            // '
            dynamic fp;

            if (dc == null) return new Dictionary();
            if (ck < 0)
            {
                ck = dgiG1f0(dc);
                continue;
            }

            var rt = dc;
            {
                var withBlock = dgiFlatListed(rt);
                foreach (dynamic ky in withBlock.Keys)
                {
                    Document AiDoc = aiDocument(obOf(withBlock.get_Item(ky)));
                    rt = dgiAiDocClassified(AiDoc, rt);

                    if (AiDoc.DocumentType == kAssemblyDocumentObject) rt = dgiMembersClassified(AiDoc, rt);
                }
            }
            var ct = dgiG1f0(dc);

            if (ct > ck)
            {
                dc = rt;
                ck = ct;
                continue;
            }

            if (ct == ck)
                Debugger.Break(); // cuz something went wrong

            return rt;
        }
    }

    public static  Dictionary dgiG2f0(Dictionary dc)
    {
        var rt = new Dictionary();
        {
            foreach (dynamic ky in dc.Keys)
            {
                Document id = aiDocument(dc.get_Item(ky));

                string sb;
                string fp;
                {
                    fp = id.FullDocumentName;
                    sb = id.SubType;
                }

                {
                    if (!rt.Exists(sb))
                        rt.Add(sb, new Dictionary());
                    {
                        var withBlock2 = dcOb(rt.get_Item(sb));
                        if (!withBlock2.Exists(fp))
                            withBlock2.Add(fp, id);
                    }
                }
            }
        }
        return rt;
    }

    public static  Dictionary dgiG2f1(Dictionary dc)
    {
        Document id;
        string sb;
        string fp;

        var rt = new Dictionary();
        {
            foreach (dynamic ky in dc.Keys)
                rt.Add(ky, dgiG2f0(dcOb(dc.get_Item(ky))));
        }
        return rt;
    }

    public static  string dgiG2f2(Dictionary dc, string pfx = "", string dlm = "|", string brk = Constants.vbCrLf)
    {
        dynamic it;
        var rt = new Dictionary();
        {
            foreach (dynamic ky in dc.Keys)
            {
                var tx = Convert.ToHexString(ky);
                if (Strings.Len(pfx) > 0)
                    tx = pfx + dlm + tx;
                rt.Add(dc.get_Item(ky) is Dictionary ? dgiG2f2(dcOb(obOf(dc.get_Item(ky))), tx, dlm, brk) : tx, 0);
            }
        }
        return Join(rt.Keys, brk);
    }
}