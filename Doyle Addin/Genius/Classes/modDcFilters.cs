using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class modDcFilters
{
    public static Dictionary dcAiDocsVisible()
    {
        var rt = new Dictionary();
        foreach (Document AiDoc in ThisApplication.Documents.VisibleDocuments)
            // rt.Add aiDoc.FullDocumentName, aiDoc
            rt.Add(d0g6f0(AiDoc), AiDoc);
        return rt;
    }

    public static void lsAiDocsVisible()
    {
        Debug.Print(txDumpLs(dcAiDocsVisible().Keys));
    }

    public static Dictionary dcAiDocsByType(Dictionary dc)
    {
        // Split Dictionary of Inventor Documents
        // into separate "sub" Dictionaries,
        // keyed by Document Type
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                Document AiDoc = aiDocument(dc.get_Item(ky));
                DocumentTypeEnum tp;
                string fn;
                {
                    tp = AiDoc.DocumentType;
                    fn = AiDoc.FullFileName;
                }

                Dictionary gp;
                {
                    if (rt.Exists(tp))
                        gp = rt.get_Item(tp);
                    else
                    {
                        gp = new Dictionary();
                        rt.Add(tp, gp);
                    }
                }

                {
                    if (gp.Exists(fn))
                        Debugger.Break();
                    else
                        gp.Add(fn, AiDoc);
                }
            }
        }
        return rt;
    }
    // Debug.Print Join(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, ", ")

    public static Dictionary dcAiDocsOfType(DocumentTypeEnum tp, Dictionary dc = )
    {
        // Retrieve subDictionary for
        // given Inventor Document type
        // 
        {
            var withBlock = dcAiDocsByType(dc);
            return withBlock.Exists(tp) ? (Dictionary)withBlock.get_Item(tp) : new Dictionary();
        }
    }

    public static Dictionary dcAiPartDocs(Dictionary dc)
    {
        return dcAiDocsOfType(kPartDocumentObject, dc);
    }
    // Debug.Print Join(dcAiPartDocs(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbCrLf)

    public static Dictionary dcAiAssyDocs(Dictionary dc)
    {
        return dcAiDocsOfType(kAssemblyDocumentObject, dc);
    }

    public static Dictionary dcOf_iPartFactories(Dictionary dc = null)
    {
        Dictionary rt;

        if (dc == null)
            rt = dcOf_iPartFactories(dcAiDocsVisible());
        else
        {
            rt = new Dictionary();

            {
                var withBlock = dcAiPartDocs(dc);
                foreach (var ky in withBlock.Keys)
                {
                    PartDocument pt = aiDocPart(withBlock.get_Item(ky));
                    {
                        if (pt.ComponentDefinition.IsiPartFactory)
                            rt.Add.FullFileName(null /* Conversion error: Set to default value for this argument */,
                                pt);
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary dcOf_iAssyFactories(Dictionary dc = null)
    {
        Dictionary rt;

        if (dc == null)
            rt = dcOf_iAssyFactories(dcAiDocsVisible());
        else
        {
            rt = new Dictionary();

            {
                var withBlock = dcAiAssyDocs(dc);
                foreach (var ky in withBlock.Keys)
                {
                    AssemblyDocument sm = aiDocAssy(withBlock.get_Item(ky));
                    {
                        if (sm.ComponentDefinition.IsiAssemblyFactory)
                            rt.Add.FullFileName(null /* Conversion error: Set to default value for this argument */,
                                sm);
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary dcOf_iAll_Factories(Dictionary dc = null)
    {
        Dictionary rt;

        if (dc == null)
            rt = dcOf_iAll_Factories(dcAiDocsVisible());
        else
        {
            rt = dcOf_iPartFactories(dc);

            {
                var withBlock = dcOf_iAssyFactories(dc);
                foreach (var ky in withBlock.Keys)
                    rt.Add(ky, withBlock.get_Item(ky));
            }
        }

        return rt;
    }

    public static Dictionary dcAiSheetMetal(Dictionary dc)
    {
        PartDocument pt;

        var rt = new Dictionary();
        {
            var withBlock = dcAiPartDocs(dc);
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocPart(withBlock.get_Item(ky));
                    if (withBlock1.DocumentSubType.DocumentSubTypeID == guidSheetMetal)
                        rt.Add.FullFileName(null /* Conversion error: Set to default value for this argument */,
                            withBlock1.ComponentDefinition.Document);
                }
            }
        }
        return rt;
    }
    // Debug.Print Join(dcAiSheetMetal(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbCrLf)
    // 
    // Debug.Print 'dcAiPartDocs(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count
    // Debug.Print 'dcAiSheetMetal(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count

    public static Dictionary dcAssyPartsPrimary(AssemblyDocument aiAssy)
    {
        var rt = new Dictionary();
        {
            var withBlock = aiAssy.ComponentDefinition;
            foreach (ComponentOccurrence oc in withBlock.Occurrences)
            {
                {
                    var withBlock1 = aiDocument(oc.Definition.Document);
                    if (!rt.Exists(withBlock1.FullDocumentName))
                        rt.Add.FullDocumentName(null /* Conversion error: Set to default value for this argument */,
                            withBlock1.PropertySets.Parent);
                }
            }
        }
        return rt;
    }

    public static Dictionary dcAiDocsByPtNum(Dictionary dcIn)
    {
        // Dim oc As Inventor.ComponentOccurrence
        var rt = new Dictionary();
        {
            foreach (var ky in dcIn.Keys)
            {
                {
                    var withBlock1 = aiDocument(dcIn.get_Item(ky)).PropertySets.get_Item(gnDesign);
                    string pn = withBlock1.get_Item(pnPartNum).Value;
                    if (rt.Exists(pn))
                        Debugger.Break();
                    else
                        rt.Add(pn, withBlock1.Parent.Parent);
                }
            }
        }
        return rt;
    }
    // dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each ky In dc: Debug.Print txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.get_Item(ky)))).Keys, vbCrLf & vbTab): Next
    // tx = "": dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each 'ky In dc: tx = tx & vbCrLf & ky & vbCrLf & vbTab & 'txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.get_Item(ky)))).Keys, vbCrLf & 'vbTab): Next: send2clipBd tx: dc = Nothing

    public static Dictionary dcItemsNotInGenius(Dictionary dcPts)
    {
        // dcItemsNotInGenius --
        // takes a Dictionary of Items
        // (keyed by Item/Part Number)
        // and returns a Dictionary of
        // Items not yet found in Genius
        // 
        // NOTE: originally designed to take
        // a Dictionary of Inventor
        // Documents, it SHOULD be able
        // to process a Dictionary of
        // ANY sort of Items keyed
        // to Item/Part Number
        // 
        // dcPts = dcRemapByPtNum(dcAiDocComponents(aiDoc))

        return dcKeysMissing(dcPts,
            dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(q1g1x2v2(dcPts)))).get_Item("Item")));
    }

    public static Dictionary dcAiPartsNotInGenius(Document AiDoc)
    {
        // dcAiPartsNotInGenius --
        // calls dcItemsNotInGenius
        // against a Dictionary of Items
        // from supplied Inventor Document
        // to return a subset of Items
        // not yet added to Genius
        // 

        return dcItemsNotInGenius(dcRemapByPtNum(dcAiDocComponents(AiDoc)));
    }

    public static Dictionary mdf0g0f0(Document AiDoc)
    {
        const string bk = Constants.vbCrLf + Constants.vbTab;
        var rt = new Dictionary();
        {
            var withBlock = dcAiDocsByPtNum(dcAssyPartsPrimary(AiDoc));
            foreach (var ky in withBlock.Keys)
                rt.Add(
                    ky + bk + txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(withBlock.get_Item(ky)))).Keys,
                        bk), withBlock.get_Item(ky));
        }
        return rt;
    }
}