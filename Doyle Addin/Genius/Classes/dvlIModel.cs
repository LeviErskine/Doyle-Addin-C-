class dvlIModel
{
    public Inventor.PartDocument dImG1f1iPart(Inventor.PartDocument md)
    {
        if (md == null)
            dImG1f1iPart = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            var withBlock = md.ComponentDefinition;
            if (withBlock.IsiPartFactory)
            {
                if (withBlock.iPartFactory == null)
                {
                    System.Diagnostics.Debugger.Break();
                    dImG1f1iPart = null/* TODO Change to default(_) if this is not a reference type */;
                }
                else
                    dImG1f1iPart = aiDocPart(withBlock.iPartFactory.Parent);
            }
            else if (withBlock.IsiPartMember)
            {
                if (withBlock.iPartMember == null)
                {
                    System.Diagnostics.Debugger.Break();
                    dImG1f1iPart = null/* TODO Change to default(_) if this is not a reference type */;
                }
                else
                    dImG1f1iPart = dImG1f1iPart(withBlock.iPartMember.ParentFactory.Parent);
            }
            else
                dImG1f1iPart = null/* TODO Change to default(_) if this is not a reference type */;
        }
    }

    public Scripting.Dictionary dImG1f2iPart(Inventor.Document md)
    {
        Scripting.Dictionary rt;
        Variant ky;
        Inventor.PartDocument pt;
        Inventor.PartDocument ck;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dcAiPartDocs(dcAiDocComponents(md));
            foreach (var ky in withBlock.Keys)
            {
                pt = aiDocPart(withBlock.Item(ky));
                if (pt == null)
                {
                }
                else
                {
                    ck = dImG1f1iPart(pt);
                    if (ck == null)
                    {
                    }
                    else
                        // ck.File.FullFileName
                        // Stop
                        Debug.Print();/* TODO ERROR: Skipped SkippedTokensTrivia */// Breakpoint Landing
                }
            }
        }

        dImG1f2iPart = rt;
    }

    public string dImG0f0()
    {
        dImG0f0 = "REV[2023.01.19.1046]";
    }
}