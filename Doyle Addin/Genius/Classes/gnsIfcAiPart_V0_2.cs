namespace Doyle_Addin.Genius.Classes;

class gnsIfcAiPart_V0_2 : gnsIfcAiDoc
{
    private IDictionary gnsIfcAiDoc_Props(Document AiDoc, IDictionary dc = null)
    {
        return dcGeniusPropsPartRev20180530(aiDocPart(AiDoc), dc);
    }
}