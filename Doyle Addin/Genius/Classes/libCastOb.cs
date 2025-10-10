class libCastOb
{
    public object obOf(Variant vr)
    {
        if (IsObject(vr))
            obOf = vr;
        else
            obOf = null;
    }

    public Scripting.Dictionary dcOb(Variant vr)
    {
        if (IsObject(vr))
        {
            if (vr == null)
                dcOb = null/* TODO Change to default(_) if this is not a reference type */;
            else if (vr is Scripting.Dictionary)
                dcOb = vr;
            else
                dcOb = null/* TODO Change to default(_) if this is not a reference type */;
        }
        else
            dcOb = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public ADODB.Field fdOb(Variant vr)
    {
        if (IsObject(vr))
        {
            if (vr is ADODB.Field)
                fdOb = vr;
            else
                fdOb = null/* TODO Change to default(_) if this is not a reference type */;
        }
        else
            fdOb = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Document aiDocument(object doc)
    {
        if (doc == null)
            aiDocument = doc;
        else if (doc is Inventor.Document)
            aiDocument = doc;
        else
            aiDocument = null/* TODO Change to default(_) if this is not a reference type */;
    }
    // For Each itm In ActiveDocsComponents(ThisApplication): Debug.Print aiDocument(obOf(itm)).FullFileName: Next

    public Inventor.Document aiDocActive()
    {
        aiDocActive = ThisApplication.ActiveDocument;
    }

    public Inventor.PartDocument aiDocPart(Inventor.Document doc)
    {
        if (doc == null)
            aiDocPart = doc;
        else if (doc is Inventor.PartDocument)
            aiDocPart = doc;
        else
            aiDocPart = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.PartDocument aiDocPartFromCCtr(Inventor.Document doc)
    {
        Inventor.PartDocument rt;

        if (doc == null)
            aiDocPartFromCCtr = doc;
        else
        {
            rt = aiDocPart(doc);
            if (rt == null)
                aiDocPartFromCCtr = rt;
            else if (rt.ComponentDefinition.IsContentMember)
                aiDocPartFromCCtr = rt;
            else
                aiDocPartFromCCtr = null/* TODO Change to default(_) if this is not a reference type */;
        }
    }

    public Inventor.AssemblyDocument aiDocAssy(Inventor.Document doc)
    {
        if (doc == null)
            aiDocAssy = doc;
        else if (doc is Inventor.AssemblyDocument)
            aiDocAssy = doc;
        else
            aiDocAssy = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.DrawingDocument aiDocDwg(Inventor.Document doc)
    {
        if (doc == null)
            aiDocDwg = doc;
        else if (doc is Inventor.DrawingDocument)
            aiDocDwg = doc;
        else
            aiDocDwg = null/* TODO Change to default(_) if this is not a reference type */;
    }

    private Inventor.ComponentDefinition aiCompDefinition(object doc)
    {
        /// REV[2022.08.31.1313] OBSOLETED
        /// -   no calls found to this function
        /// -   aiCompDefOf serves same purpose
        /// in (slightly?) more robust manner
        /// -   changed scope to Private
        /// to prevent future usage
        /// outside local scope
        /// 
        if (doc is Inventor.ComponentDefinition)
            aiCompDefinition = doc;
        else
            aiCompDefinition = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.ComponentDefinition aiCompDefOf(object doc) // Inventor.Document
    {
        /// aiCompDefOf -- Return the ComponentDefinition
        /// of ANY Inventor Document which has one.
        /// NOTE: currently returns ComponentDefinition objects
        /// only from Part and Assembly Documents.
        /// NOTE[2022.08.31.1202]: copied comments from redundant
        /// function obAiCompDefAny prior to its deprecation
        /// 
        if (doc == null)
            aiCompDefOf = null/* TODO Change to default(_) if this is not a reference type */;
        else if (doc is Inventor.ComponentDefinition)
            aiCompDefOf = doc;
        else if (doc is Inventor.Document)
        {
            {
                var withBlock = aiDocument(doc);
                if (withBlock.DocumentType == kAssemblyDocumentObject)
                    aiCompDefOf = aiDocAssy(doc).ComponentDefinition;
                else if (withBlock.DocumentType == kPartDocumentObject)
                    aiCompDefOf = aiDocPart(doc).ComponentDefinition;
                else
                    aiCompDefOf = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }
        else
            aiCompDefOf = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.ComponentDefinition obAiCompDefAny(Inventor.Document AiDoc)
    {
        /// obAiCompDefAny -- Return the ComponentDefinition
        /// of ANY Inventor Document which has one.
        /// NOTE: currently returns ComponentDefinition objects
        /// only from Part and Assembly Documents.
        /// NOTE[2022.08.31.1203]: rediscovered original
        /// implementation aiCompDefOf; copied comments
        /// there prior to deprecation of this implementation
        /// 
        if (AiDoc == null)
            obAiCompDefAny = null/* TODO Change to default(_) if this is not a reference type */;
        else if (AiDoc.DocumentType == kAssemblyDocumentObject)
            obAiCompDefAny = aiDocAssy(AiDoc).ComponentDefinition;
        else if (AiDoc.DocumentType == kPartDocumentObject)
            obAiCompDefAny = aiDocPart(AiDoc).ComponentDefinition;
        else
            obAiCompDefAny = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.PartComponentDefinition aiCompDefPart(object doc)
    {
        /// REV[2022.08.31.1247]
        /// added ElseIf check for PartDocument
        /// to accept Inventor Document as well
        /// as ComponentDefinition
        /// applied same to functions {
        /// aiCompDefPart
        /// }
        /// 
        /// 
        if (doc == null)
            aiCompDefPart = null/* TODO Change to default(_) if this is not a reference type */;
        else if (doc is Inventor.PartComponentDefinition)
            aiCompDefPart = doc;
        else if (doc is Inventor.PartDocument)
            aiCompDefPart = aiDocPart(doc).ComponentDefinition;
        else
            aiCompDefPart = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.SheetMetalComponentDefinition aiCompDefShtMetal(object ob)
    {
        if (ob == null)
            aiCompDefShtMetal = null/* TODO Change to default(_) if this is not a reference type */;
        else if (ob is Inventor.SheetMetalComponentDefinition)
            aiCompDefShtMetal = ob;
        else if (ob is Inventor.PartDocument)
            aiCompDefShtMetal = aiCompDefShtMetal(aiDocPart(ob).ComponentDefinition);
        else
            aiCompDefShtMetal = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.AssemblyComponentDefinition aiCompDefAssy(object ob)
    {
        if (ob == null)
            aiCompDefAssy = null/* TODO Change to default(_) if this is not a reference type */;
        else if (ob is Inventor.AssemblyComponentDefinition)
            aiCompDefAssy = ob;
        else if (ob is Inventor.AssemblyDocument)
            aiCompDefAssy = aiDocAssy(ob).ComponentDefinition;
        else
            aiCompDefAssy = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Property aiProperty(object ob)
    {
        if (ob == null)
            // Stop
            aiProperty = null/* TODO Change to default(_) if this is not a reference type */;
        else if (ob is Inventor.Property)
            aiProperty = ob;
        else
            // Stop 'because this is NOT a Property!
            aiProperty = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Plane aiPlane(object ob)
    {
        if (ob is Inventor.Plane)
            aiPlane = ob;
        else
            aiPlane = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.ComponentOccurrence aiCompOcc(object ob)
    {
        if (ob is Inventor.ComponentOccurrence)
            aiCompOcc = ob;
        else
            aiCompOcc = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Property obAiProp(object ob)
    {
        if (ob is Inventor.Property)
            obAiProp = ob;
        else
            obAiProp = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public Inventor.Parameter obAiParam(object ob)
    {
        if (ob is Inventor.Parameter)
            obAiParam = ob;
        else
            obAiParam = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public VBIDE.VBProject obVbProject(object ob)
    {
        if (ob is VBIDE.VBProject)
            obVbProject = ob;
        else
            obVbProject = null/* TODO Change to default(_) if this is not a reference type */;
    }

    public VBIDE.CodeModule obVbCodeMod(object ob)
    {
        if (ob is VBIDE.CodeModule)
            obVbCodeMod = ob;
        else
            obVbCodeMod = null/* TODO Change to default(_) if this is not a reference type */;
    }
}