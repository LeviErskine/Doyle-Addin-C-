class SurroundingClass
{
    public Scripting.Dictionary dcAiDocsVisible()
    {
        Scripting.Dictionary rt;
        Inventor.Document AiDoc;

        rt = new Scripting.Dictionary();
        foreach (var AiDoc in ThisApplication.Documents.VisibleDocuments)
            // rt.Add aiDoc.FullDocumentName, aiDoc
            rt.Add(d0g6f0(AiDoc), AiDoc);
        dcAiDocsVisible = rt;
    }

    public void lsAiDocsVisible()
    {
        Debug.Print(txDumpLs(dcAiDocsVisible().Keys));
    }

    public Scripting.Dictionary dcAiDocsByType(Scripting.Dictionary dc)
    {
        /// Split Dictionary of Inventor Documents
        /// into separate "sub" Dictionaries,
        /// keyed by Document Type
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary gp;
        Inventor.Document AiDoc;
        Inventor.DocumentTypeEnum tp;
        string fn;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                AiDoc = aiDocument(withBlock.Item(ky));
                {
                    var withBlock1 = AiDoc;
                    tp = withBlock1.DocumentType;
                    fn = withBlock1.FullFileName;
                }

                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(tp))
                        gp = withBlock1.Item(tp);
                    else
                    {
                        gp = new Scripting.Dictionary();
                        withBlock1.Add(tp, gp);
                    }
                }

                {
                    var withBlock1 = gp;
                    if (withBlock1.Exists(fn))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(fn, AiDoc);
                }
            }
        }
        dcAiDocsByType = rt;
    }
    // Debug.Print Join(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences)).Keys, ", ")

    public Scripting.Dictionary dcAiDocsOfType(Inventor.DocumentTypeEnum tp, Scripting.Dictionary dc = )
    {
        /// Retrieve subDictionary for
        /// given Inventor Document type
        /// 
        {
            var withBlock = dcAiDocsByType(dc);
            if (withBlock.Exists(tp))
                dcAiDocsOfType = withBlock.Item(tp);
            else
                dcAiDocsOfType = new Scripting.Dictionary();
        }
    }

    public Scripting.Dictionary dcAiPartDocs(Scripting.Dictionary dc)
    {
        dcAiPartDocs = dcAiDocsOfType(kPartDocumentObject, dc);
    }
    // Debug.Print Join(dcAiPartDocs(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbNewLine)

    public Scripting.Dictionary dcAiAssyDocs(Scripting.Dictionary dc)
    {
        dcAiAssyDocs = dcAiDocsOfType(kAssemblyDocumentObject, dc);
    }

    public Scripting.Dictionary dcOf_iPartFactories(Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.PartDocument pt;
        Scripting.Dictionary rt;
        Variant ky;

        if (dc == null)
            rt = dcOf_iPartFactories(dcAiDocsVisible());
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = dcAiPartDocs(dc);
                foreach (var ky in withBlock.Keys)
                {
                    pt = aiDocPart(withBlock.Item(ky));
                    {
                        var withBlock1 = pt;
                        if (withBlock1.ComponentDefinition.IsiPartFactory)
                            rt.Add.FullFileName(null/* Conversion error: Set to default value for this argument */, pt);
                    }
                }
            }
        }

        dcOf_iPartFactories = rt;
    }

    public Scripting.Dictionary dcOf_iAssyFactories(Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.AssemblyDocument sm;
        Scripting.Dictionary rt;
        Variant ky;

        if (dc == null)
            rt = dcOf_iAssyFactories(dcAiDocsVisible());
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = dcAiAssyDocs(dc);
                foreach (var ky in withBlock.Keys)
                {
                    sm = aiDocAssy(withBlock.Item(ky));
                    {
                        var withBlock1 = sm;
                        if (withBlock1.ComponentDefinition.IsiAssemblyFactory)
                            rt.Add.FullFileName(null/* Conversion error: Set to default value for this argument */, sm);
                    }
                }
            }
        }

        dcOf_iAssyFactories = rt;
    }

    public Scripting.Dictionary dcOf_iAll_Factories(Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Variant ky;

        if (dc == null)
            rt = dcOf_iAll_Factories(dcAiDocsVisible());
        else
        {
            rt = dcOf_iPartFactories(dc);

            {
                var withBlock = dcOf_iAssyFactories(dc);
                foreach (var ky in withBlock.Keys)
                    rt.Add(ky, withBlock.Item(ky));
            }
        }

        dcOf_iAll_Factories = rt;
    }

    public Scripting.Dictionary dcAiSheetMetal(Scripting.Dictionary dc)
    {
        Variant ky;
        Inventor.PartDocument pt;
        Scripting.Dictionary rt;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAiPartDocs(dc);
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocPart(withBlock.Item(ky));
                    if (withBlock1.DocumentSubType.DocumentSubTypeID == guidSheetMetal)
                        rt.Add.FullFileName(null/* Conversion error: Set to default value for this argument */, withBlock1.ComponentDefinition.Document);
                }
            }
        }
        dcAiSheetMetal = rt;
    }
    // Debug.Print Join(dcAiSheetMetal(dcAiDocsByType(dcAssyCompAndSub(aiDocAssy(ThisApplication.ActiveDocument).ComponentDefinition.Occurrences))).Keys, vbNewLine)
    // 
    // Debug.Print 'dcAiPartDocs(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count
    // Debug.Print 'dcAiSheetMetal(dcAiDocsByType(dcAssyDocComponents(ThisApplication.ActiveDocument))).Count

    public Scripting.Dictionary dcAssyPartsPrimary(Inventor.AssemblyDocument aiAssy)
    {
        Scripting.Dictionary rt;
        Inventor.ComponentOccurrence oc;

        rt = new Scripting.Dictionary();
        {
            var withBlock = aiAssy.ComponentDefinition;
            foreach (var oc in withBlock.Occurrences)
            {
                {
                    var withBlock1 = aiDocument(oc.Definition.Document);
                    if (!rt.Exists(withBlock1.FullDocumentName))
                        rt.Add.FullDocumentName(null/* Conversion error: Set to default value for this argument */, withBlock1.PropertySets.Parent);
                }
            }
        }
        dcAssyPartsPrimary = rt;
    }

    public Scripting.Dictionary dcAiDocsByPtNum(Scripting.Dictionary dcIn)
    {
        Scripting.Dictionary rt;
        Variant ky;
        string pn;
        // Dim oc As Inventor.ComponentOccurrence

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcIn // aiAssy.ComponentDefinition
       ;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = aiDocument(withBlock.Item(ky)).PropertySets.Item(gnDesign);
                    pn = withBlock1.Item(pnPartNum).Value;
                    if (rt.Exists(pn))
                        System.Diagnostics.Debugger.Break();
                    else
                        rt.Add(pn, withBlock1.Parent.Parent);
                }
            }
        }
        dcAiDocsByPtNum = rt;
    }
    // dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each ky In dc: Debug.Print txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.Item(ky)))).Keys, vbNewLine & vbTab): Next
    // tx = "":  dc = dcAiDocsByPtNum(dcAssyPartsPrimary(ThisApplication.ActiveDocument)): For Each 'ky In dc: tx = tx & vbNewLine & ky & vbNewLine & vbTab & 'txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(dc.Item(ky)))).Keys, vbNewLine & 'vbTab): Next: send2clipBd tx:  dc = Nothing

    public Scripting.Dictionary dcItemsNotInGenius(Scripting.Dictionary dcPts)
    {
        /// dcItemsNotInGenius --
        /// takes a Dictionary of Items
        /// (keyed by Item/Part Number)
        /// and returns a Dictionary of
        /// Items not yet found in Genius
        /// 
        /// NOTE: originally designed to take
        /// a Dictionary of Inventor
        /// Documents, it SHOULD be able
        /// to process a Dictionary of
        /// ANY sort of Items keyed
        /// to Item/Part Number
        /// 
        // dcPts = dcRemapByPtNum(dcAiDocComponents(aiDoc))

        dcItemsNotInGenius = dcKeysMissing(dcPts, dcOb(dcDxFromRecSetDc(dcFromAdoRS(cnGnsDoyle().Execute(q1g1x2v2(dcPts)))).Item("Item")));
    }

    public Scripting.Dictionary dcAiPartsNotInGenius(Inventor.Document AiDoc)
    {
        /// dcAiPartsNotInGenius --
        /// calls dcItemsNotInGenius
        /// against a Dictionary of Items
        /// from supplied Inventor Document
        /// to return a subset of Items
        /// not yet added to Genius
        /// 

        dcAiPartsNotInGenius = dcItemsNotInGenius(dcRemapByPtNum(dcAiDocComponents(AiDoc)));
    }

    public Scripting.Dictionary mdf0g0f0(Inventor.Document AiDoc)
    {
        Scripting.Dictionary rt;
        Variant ky;
        string bk;

        bk = Constants.vbNewLine + Constants.vbTab;
        rt = new Scripting.Dictionary();
        {
            var withBlock = dcAiDocsByPtNum(dcAssyPartsPrimary(AiDoc));
            foreach (var ky in withBlock.Keys)
                rt.Add(ky + bk + txDumpLs(dcAiDocsByPtNum(dcAssyPartsPrimary(aiDocument(withBlock.Item(ky)))).Keys, bk), withBlock.Item(ky));
        }
        mdf0g0f0 = rt;
    }
}