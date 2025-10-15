using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

internal class libSQLinVB
{
    private string sqlTextInVBcode(string nm)
    {
        dynamic ar = Split(nm, "'''SQL'''");
        return UBound(ar) < 1
            ? ""
            :
            // sqlTextInVBcode = ar(1)
            Join(Split(ar(1), Constants.vbCrLf + "'"), Constants.vbCrLf);
    }

    private string sqlTextInDict(string nm, Dictionary dc)
    {
        return sqlTextInVBcode(vbTextOfProcInDict(nm, dc));
    }

    public string sqlTextInProject(string nm, VBIDE.VBProject pj)
    {
        return sqlTextInDict(nm, dcOfVbProcsFlat(pj));
    }
}