namespace Doyle_Addin.Genius.Classes;

public class kyPick
{
    private Dictionary dcGrpIn;
    private Dictionary dcGrpOut;

    private void Class_Initialize()
    {
        dcGrpIn = new Dictionary();
        dcGrpOut = new Dictionary();
    }

    public kyPick Itself()
    {
        return this;
    }

    public kyPick WithInDc(Dictionary Dict)
    {
        dcGrpIn = dcNewIfNone(Dict);
        return this;
    }

    public kyPick WithOutDc(Dictionary Dict)
    {
        dcGrpOut = dcNewIfNone(Dict);
        return this;
    }

    public kyPick AfterScanning(Dictionary dSrc)
    {
        {
            foreach (var ky in dSrc.Keys)
            {
                {
                    var withBlock1 = dcFor(dSrc.get_Item(ky));
                    if (withBlock1.Exists(ky))
                        Debugger.Break();
                    else
                        withBlock1.Add(ky, dSrc.get_Item(ky));
                }
            }
        }
        return this;
    }

    public Dictionary dcIn()
    {
        return dcGrpIn;
    }

    public Dictionary dcOut()
    {
        return dcGrpOut;
    }

    public Dictionary dcFor(dynamic Item)
    {
        return Item is not null ? dcGrpIn : dcGrpOut;
    }
}