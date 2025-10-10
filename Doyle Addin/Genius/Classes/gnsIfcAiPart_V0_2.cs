class gnsIfcAiPart_V0_2 : gnsIfcAiDoc
{
    private Scripting.IDictionary gnsIfcAiDoc_Props(Inventor.Document AiDoc, Scripting.IDictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        gnsIfcAiDoc_Props = dcGeniusPropsPartRev20180530(aiDocPart(AiDoc), dc);
    }
}