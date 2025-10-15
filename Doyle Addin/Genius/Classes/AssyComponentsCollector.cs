using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

internal class AssyComponentsCollector
{
    // Purpose of this module is to provide an alternate
    // method of collecting assembly components
    // using the native VBA Collection instead of
    // the Scripting Runtime's Dictionary.
    // Though less powerful/convenient, it does avoid
    // the need for a reference to the Scripting Runtime.
    // 

    private static Collection CollectItem(dynamic Item, string Key, Collection coll = null)
    {
        var rt = coll ?? new Collection();

        {
            var withBlock = Information.Err;
            rt.Add(Item, Key);
            if (!withBlock.Number) return rt;
            if (withBlock.Number == 457)
            {
                if (Item != null)
                {
                    if (rt.get_Item(Key) is dynamic)
                    {
                        if (Item == rt.get_item(Key))
                        {
                        }
                        else
                            Debugger.Break(); // Different Objects!
                    }
                    else
                        Debugger.Break(); // dynamic vs non-dynamic
                }
                else if (rt.get_Item(Key) is dynamic)
                    Debugger.Break(); // dynamic vs non-dynamic
                else if ((dynamic)null == rt.get_Item(Key))
                {
                }
                else
                    Debugger.Break(); // Different Values!
            }
            else
                Debugger.Break();
        }

        return rt;
    }

    private static Collection CollectComponents(Document AiDoc, Collection coll = null)
    {
        Collection rt;

        if (coll == null)
            rt = CollectComponents(AiDoc, new Collection());
        else
        {
            rt = coll;
            var aiDType = AiDoc.DocumentType;
            switch (aiDType)
            {
                case kAssemblyDocumentObject:
                {
                    {
                        var withBlock = aiDocAssy(AiDoc).ComponentDefinition;
                        foreach (ComponentOccurrence aiOcc in withBlock.Occurrences)
                        {
                            if (aiOcc.Definition.Document == AiDoc)
                            {
                            }
                            else
                                rt = CollectComponents(aiOcc.Definition.Document, rt);
                        }
                    }
                    break;
                }
                case kPartDocumentObject:
                    rt = CollectItem(AiDoc, AiDoc.FullFileName, rt);
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
                    Debugger.Break(); // cuz we don't know what to do with this one.
                    break;
            }
        }

        return rt;
    }

    private static Collection ActiveDocsComponents(Application aiApp)
    {
        return CollectComponents(aiApp.ActiveDocument);
    }

    public string strActiveDocsComponents(Application aiApp)
    {
        Document AiDoc;

        return ActiveDocsComponents(aiApp).Cast<dynamic>()
            .Aggregate("", (current, AiDoc) => current + Constants.vbCrLf + AiDoc.FullFileName);
    }
}