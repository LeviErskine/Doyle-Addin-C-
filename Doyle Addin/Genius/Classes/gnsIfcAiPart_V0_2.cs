

Implements gnsIfcAiDoc

Private Function gnsIfcAiDoc_Props( _
    AiDoc As Inventor.Document, _
    Optional dc As Scripting.IDictionary = Nothing _
) As Scripting.IDictionary
    Set gnsIfcAiDoc_Props = _
    dcGeniusPropsPartRev20180530( _
    aiDocPart(AiDoc), dc)
End Function

