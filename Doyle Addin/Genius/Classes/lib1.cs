// For DataObject (Clipboard)

using Microsoft.VisualBasic;
using static Microsoft.VisualBasic.Strings; // VB helper functions: Replace, Space, Mid, Len, Join, StrDup
// Autodesk Inventor types
// For kPurchasedBOMStructure
using VbMsgBoxResult = Microsoft.VisualBasic.MsgBoxResult; // Align VB-style result type
using MsgBoxStyle = Microsoft.VisualBasic.MsgBoxStyle;

namespace Doyle_Addin.Genius.Classes;

public class lib1
{
    public static string Repeat(long Count, string Text)
    {
        return Replace(Space((int)Count), " ", Text);
    }

    public static string txBlk(long Lines, long Chars, string Use = "+")
    {
        // Build a block of Lines, each with Chars of the specified symbol, separated by CRLF
        return Mid(Repeat(Lines, Constants.vbCrLf + StrDup((int)Chars, Use)), 1 + Len(Constants.vbCrLf));
    }

    public static void MakeActivePurchased()
    {
        Document md = ThisApplication.ActiveDocument;
        var ck = md == ThisDocument ? Constants.vbNo : mkAiDocPurchased(md);

        switch (ck)
        {
            case Constants.vbOK:
                ck = Interaction.MsgBox(Join(new[]
                    {
                        "Model BOM Structure", "now Purchased."
                    }, Constants.vbCrLf),
                    Constants.vbOKOnly | Constants.vbInformation, "Success!");
                break;
            case Constants.vbNo:
                ck = Interaction.MsgBox(
                    Join(new[] { "Document is not", "a valid Model.", "", "Please select a", "Part or Assembly." },
                        Constants.vbCrLf), Constants.vbOKOnly | Constants.vbExclamation, "No Model");
                break;
            case Constants.vbAbort:
                ck = Interaction.MsgBox(
                    Join(new[] { "Failed to update", "model's BOM Structure!", "", "Check for locks", "or other issues." },
                        Constants.vbCrLf), Constants.vbOKOnly | Constants.vbCritical, "Change Failed!");
                break;
            case VbMsgBoxResult.Cancel:
            case VbMsgBoxResult.Retry:
            case VbMsgBoxResult.Ignore:
            case VbMsgBoxResult.Yes:
                break;
            default:
                ck = Interaction.MsgBox(
                    Join(
                        new[] { "Change Operation returned", "unexpected result code.", "", "Please review model status." },
                        Constants.vbCrLf), Constants.vbOKOnly | Constants.vbQuestion, "Result Unknown");
                break;
        }
    }

    public static VbMsgBoxResult mkAiDocPurchased(Document AiDoc)
    {
        var ck = AiDoc switch
        {
            PartDocument p => mkAiPartPurchased(p),
            AssemblyDocument a => mkAiAssyPurchased(a),
            _ => Constants.vbNo
        };

        return ck;
    }

    public static VbMsgBoxResult mkAiPartPurchased(PartDocument AiDoc)
    {
        if (AiDoc == null)
            return Constants.vbNo;
        var withBlock = AiDoc.ComponentDefinition;
        // Use VB Err object to detect COM errors
        Information.Err().Clear();
        withBlock.BOMStructure = kPurchasedBOMStructure;
        return Information.Err().Number == 0 ? Constants.vbOK : Constants.vbAbort;
    }

    public static VbMsgBoxResult mkAiAssyPurchased(AssemblyDocument AiDoc)
    {
        if (AiDoc == null)
            return Constants.vbNo;
        var withBlock = AiDoc.ComponentDefinition;
        Information.Err().Clear();
        withBlock.BOMStructure = kPurchasedBOMStructure;
        return Information.Err().Number == 0 ? Constants.vbOK : Constants.vbAbort;
    }

    public static Dictionary dcTemplate0A(Dictionary dc = null)
    {
        var rt = dc ?? dcTemplate0A(new Dictionary());

        return rt;
    }

    public static dynamic send2clipBd_OBSOLETE(dynamic src)
    {
        {
            var withBlock = new DataObject();
            withBlock.SetText(src);
            // VB-style DataObject.PutInClipboard is not available in WinForms; use Clipboard API instead
            Clipboard.SetDataObject(withBlock, true);
        }
        return src;
    }
}