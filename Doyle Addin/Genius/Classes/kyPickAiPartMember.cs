namespace Doyle_Addin.Genius.Classes;

public class kyPickAiPartMember 
{
    private kyPick pk;

    private void Class_Initialize()
    {
        pk = new kyPick();
    }

    public static IDictionary dcFor(dynamic Item)
    {
        PartDocument ob = aiDocPart(obOf(Item));
        if (ob == null)
            return pk.dcFor(0);
        return ob.ComponentDefinition.IsiPartMember ? pk.dcFor(ob) : pk.dcFor(0);
    }

    public static kyPick WithInDc(Dictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        return this;
    }

    public static kyPick WithOutDc(Dictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        return this;
    }

    public static Dictionary dcIn()
    {
        return pk.dcIn;
    }

    public static Dictionary dcOut()
    {
        return pk.dcOut;
    }

    public static kyPick AfterScanning(Dictionary dSrc)
    {
        return kyPick_AfterScanning(dSrc);
    }

    private kyPick kyPick_AfterScanning(IDictionary dSrc)
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

    // kyPick Implementation code follows

    // 
    private IDictionary kyPick_DcFor(dynamic Item)
    {
        return dcFor(Item);
    }

    private IDictionary kyPick_DcIn()
    {
        return dcIn();
    }

    private IDictionary kyPick_DcOut()
    {
        return dcOut();
    }

    public static kyPick Itself()
    {
        return this;
    }

    private kyPick kyPick_Itself()
    {
        return Itself();
    }

    private kyPick kyPick_WithInDc(IDictionary Dict)
    {
        return WithInDc(Dict);
    }

    private kyPick kyPick_WithOutDc(IDictionary Dict)
    {
        return WithOutDc(Dict);
    }
}