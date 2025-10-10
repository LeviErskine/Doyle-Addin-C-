class AssyComponentsCollector
{
    /// Purpose of this module is to provide an alternate

    /// method of collecting assembly components

    /// using the native VBA Collection instead of

    /// the Scripting Runtime's Dictionary.

    /// Though less powerful/convenient, it does avoid

    /// the need for a reference to the Scripting Runtime.

    /// 

    public Collection CollectItem(Variant Item, Variant Key = , Collection coll = null)
    {
        Collection rt;

        if (coll == null)
            rt = new Collection();
        else
            rt = coll;


        {
            var withBlock = Information.Err;
            rt.Add(Item, Key);
            if (withBlock.Number)
            {
                if (withBlock.Number == 457)
                {
                    if (IsObject(Item))
                    {
                        if (IsObject(rt.Item(Key)))
                        {
                            if (Item == rt.Item(Key))
                            {
                            }
                            else
                                System.Diagnostics.Debugger.Break();// Different Objects!
                        }
                        else
                            System.Diagnostics.Debugger.Break();// Object vs non-Object
                    }
                    else if (IsObject(rt.Item(Key)))
                        System.Diagnostics.Debugger.Break(); // Object vs non-Object
                    else if (Item == rt.Item(Key))
                    {
                    }
                    else
                        System.Diagnostics.Debugger.Break();// Different Values!
                }
                else
                    System.Diagnostics.Debugger.Break();
            }
        }


        CollectItem = rt;
    }

    public Collection CollectComponents(Inventor.Document AiDoc, Collection coll = null)
    {
        Inventor.DocumentTypeEnum aiDType;
        Inventor.ComponentOccurrence aiOcc;
        Collection rt;

        if (coll == null)
            rt = CollectComponents(AiDoc, new Collection());
        else
        {
            rt = coll;
            aiDType = AiDoc.DocumentType;
            if (aiDType == kAssemblyDocumentObject)
            {
                {
                    var withBlock = aiDocAssy(AiDoc).ComponentDefinition;
                    foreach (var aiOcc in withBlock.Occurrences)
                    {
                        if (aiOcc.Definition.Document == AiDoc)
                        {
                        }
                        else
                            rt = CollectComponents(aiOcc.Definition.Document, rt);
                    }
                }
            }
            else if (aiDType == kPartDocumentObject)
                rt = CollectItem(AiDoc, AiDoc.FullFileName, rt);
            else
                System.Diagnostics.Debugger.Break();// cuz we dunno what to do with this one.
        }

        CollectComponents = rt;
    }

    public Collection ActiveDocsComponents(Inventor.Application aiApp)
    {
        ActiveDocsComponents = CollectComponents(aiApp.ActiveDocument);
    }

    public string strActiveDocsComponents(Inventor.Application aiApp)
    {
        Inventor.Document AiDoc;
        string rt;

        rt = "";
        foreach (var AiDoc in ActiveDocsComponents(aiApp))
            rt = rt + Constants.vbNewLine + AiDoc.FullFileName;
        strActiveDocsComponents = rt;
    }
}