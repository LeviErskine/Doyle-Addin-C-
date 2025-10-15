using ADODB;
using VBIDE;
using Scripting;
using Parameter = Inventor.Parameter;
using Property = Inventor.Property;

namespace Doyle_Addin.Genius.Classes;

/// <summary>
/// 
/// </summary>
public class libCastOb
{
    public static dynamic obOf(dynamic vr)
    {
        return vr;
    }

    public static Dictionary dcOb(dynamic vr)
    {
        if (vr != null)
        {
            return vr switch
            {
                Dictionary => vr,
                _ => null
            };
        }

        return null;
    }

    public static Field fdOb(dynamic vr)
    {
        var field = vr as Field;
        return field;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static Document aiDocument(dynamic doc)
    {
        return doc is null or Document ? (Document)doc : null;
    }
    // For Each itm In ActiveDocsComponents(ThisApplication): Debug.Print aiDocument(obOf(itm)).FullFileName: Next

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    public static Document aiDocActive()
    {
        return ThisApplication.ActiveDocument;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static PartDocument aiDocPart(Document doc)
    {
        return doc is null or PartDocument ? (PartDocument)doc : null;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static PartDocument aiDocPartFromCCtr(Document doc)
    {
        if (doc == null)
            return null;
        var rt = aiDocPart(doc);
        if (rt == null || rt.ComponentDefinition.IsContentMember)
            return rt;
        return null;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static AssemblyDocument aiDocAssy(Document doc)
    {
        return doc is null or AssemblyDocument ? (AssemblyDocument)doc : null;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static DrawingDocument aiDocDwg(Document doc)
    {
        return doc is null or DrawingDocument ? (DrawingDocument)doc : null;
    }

    private static ComponentDefinition aiCompDefinition(dynamic doc)
    {
        // REV[2022.08.31.1313] OBSOLETED
        // - no calls found to this function
        // - aiCompDefOf serves same purpose
        // in (slightly?) more robust manner
        // - changed scope to Private
        // to prevent future usage
        // outside local scope
        // 
        return doc as ComponentDefinition;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static ComponentDefinition aiCompDefOf(dynamic doc) // Inventor.Document
    {
        switch (doc)
        {
            // aiCompDefOf -- Return the ComponentDefinition
            // of ANY Inventor Document which has one.
            // NOTE: currently returns ComponentDefinition objects
            // only from Part and Assembly Documents.
            // NOTE[2022.08.31.1202]: copied comments from redundant
            // function obAiCompDefAny prior to its deprecation
            // 
            case null:
                return null;
            case ComponentDefinition:
                return doc;
            case Document:
            {
                {
                    var withBlock = aiDocument(doc);
                    return withBlock.DocumentType switch
                    {
                        kAssemblyDocumentObject => (ComponentDefinition)aiDocAssy(doc).ComponentDefinition,
                        kPartDocumentObject => (ComponentDefinition)aiDocPart(doc).ComponentDefinition,
                        _ => null
                    };
                }
            }
            default:
                return null;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="AiDoc"></param>
    /// <returns></returns>
    public static ComponentDefinition obAiCompDefAny(Document AiDoc)
    {
        // obAiCompDefAny -- Return the ComponentDefinition
        // of ANY Inventor Document which has one.
        // NOTE: currently returns ComponentDefinition objects
        // only from Part and Assembly Documents.
        // NOTE[2022.08.31.1203]: rediscovered original
        // implementation aiCompDefOf; copied comments
        // there prior to deprecation of this implementation
        // 
        if (AiDoc == null)
            return null;
        return AiDoc.DocumentType switch
        {
            kAssemblyDocumentObject => (ComponentDefinition)aiDocAssy(AiDoc).ComponentDefinition,
            kPartDocumentObject => (ComponentDefinition)aiDocPart(AiDoc).ComponentDefinition,
            _ => null
        };
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="doc"></param>
    /// <returns></returns>
    public static PartComponentDefinition aiCompDefPart(dynamic doc)
    {
        // REV[2022.08.31.1247]
        // added ElseIf check for PartDocument
        // to accept Inventor Document as well
        // as ComponentDefinition
        // applied same to functions {
        // aiCompDefPart
        // }
        // 
        // 
        if (doc == null)
            return null;
        if (doc is PartComponentDefinition)
            return doc;
        return doc is PartDocument ? (PartComponentDefinition)aiDocPart(doc).ComponentDefinition : null;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="ob"></param>
    /// <returns></returns>
    public static SheetMetalComponentDefinition aiCompDefShtMetal(dynamic ob)
    {
        if (ob == null)
            return null;
        return ob switch
        {
            SheetMetalComponentDefinition => ob,
            PartDocument => aiCompDefShtMetal(aiDocPart(ob).ComponentDefinition),
            _ => null
        };
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="ob"></param>
    /// <returns></returns>
    public static AssemblyComponentDefinition aiCompDefAssy(dynamic ob)
    {
        if (ob == null)
            return null;
        return ob switch
        {
            AssemblyComponentDefinition => ob,
            AssemblyDocument => aiDocAssy(ob).ComponentDefinition,
            _ => null
        };
    }

    public static Property aiProperty(dynamic ob)
    {
        if (ob == null)
            // Stop
            return null;
        return ob is Property property
            ? property
            :
            // Stop 'because this is NOT a Property!
            null;
    }

    public static Plane aiPlane(dynamic ob)
    {
        return ob as Plane;
    }

    public static ComponentOccurrence aiCompOcc(dynamic ob)
    {
        return ob as ComponentOccurrence;
    }

    public Property obAiProp(dynamic ob)
    {
        return ob as Property;
    }

    public static Parameter obAiParam(dynamic ob)
    {
        return ob as Parameter;
    }

    public static VBProject obVbProject(dynamic ob)
    {
        return ob as VBProject;
    }

    public static CodeModule obVbCodeMod(dynamic ob)
    {
        return ob as CodeModule;
    }
}