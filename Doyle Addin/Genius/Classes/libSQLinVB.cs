class SurroundingClass
{
    public string sqlTextInVBcode(string nm)
    {
        Variant ar;

        ar = Split(nm, "'''SQL'''");
        if (UBound(ar) < 1)
            sqlTextInVBcode = "";
        else
            // sqlTextInVBcode = ar(1)
            sqlTextInVBcode = Join(Split(ar(1), Constants.vbNewLine + "'"), Constants.vbNewLine);
    }

    public string sqlTextInDict(string nm, Scripting.Dictionary dc)
    {
        sqlTextInDict = sqlTextInVBcode(vbTextOfProcInDict(nm, dc));
    }

    public string sqlTextInProject(string nm, VBIDE.VBProject pj)
    {
        sqlTextInProject = sqlTextInDict(nm, dcOfVbProcsFlat(pj));
    }
}