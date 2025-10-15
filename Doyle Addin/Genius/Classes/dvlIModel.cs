namespace Doyle_Addin.Genius.Classes;

class dvlIModel
{
    public PartDocument dImG1f1iPart(PartDocument md)
    {
        if (md == null)
            return null;
        var withBlock = md.ComponentDefinition;
        if (withBlock.IsiPartFactory)
        {
            if (withBlock.iPartFactory != null) return aiDocPart(withBlock.iPartFactory.Parent);
            Debugger.Break();
            return null;
        }

        if (!withBlock.IsiPartMember) return null;
        if (withBlock.iPartMember != null) return dImG1f1iPart(withBlock.iPartMember.ParentFactory.Parent);
        Debugger.Break();
        return null;
    }

    public Dictionary dImG1f2iPart(Document md)
    {
        var rt = new Dictionary();

        {
            var withBlock = dcAiPartDocs(dcAiDocComponents(md));
            foreach (var ky in withBlock.Keys)
            {
                PartDocument pt = aiDocPart(withBlock.get_Item(ky));
                if (pt == null)
                {
                }
                else
                {
                    var ck = dImG1f1iPart(pt);
                    if (ck == null)
                    {
                    }
                    else
                        // ck.File.FullFileName
                        // Stop
                        Debug.Print(""); // Breakpoint Landing
                }
            }
        }

        return rt;
    }

    public string dImG0f0()
    {
        return "REV[2023.01.19.1046]";
    }
}