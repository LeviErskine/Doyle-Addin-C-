class dvlContentCenterCheck
{
    public long ckCtCtr()
    {
    }

    public Inventor.Document ccc0g0f0(Inventor.Document ck)
    {
        Inventor.PartDocument pt;

        pt = aiDocPart(ck);
        if (pt == null)
            ccc0g0f0 = pt;
        else if (pt.ComponentDefinition.IsContentMember)
            ccc0g0f0 = pt;
        else if (pt.PropertySets.Count > 4)
            ccc0g0f0 = pt;
        else if (0)
        {
        }
        else
            ccc0g0f0 = null/* TODO Change to default(_) if this is not a reference type */;
    }
}