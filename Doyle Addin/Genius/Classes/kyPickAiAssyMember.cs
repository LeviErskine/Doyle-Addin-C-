namespace Doyle_Addin.Genius.Classes;

class kyPickAiAssyMember : kyPick
{
    private kyPick pk;

    private void Class_Initialize()
    {
        pk = new kyPick();
    }

    private new IDictionary dcFor(dynamic Item)
    {
        AssemblyDocument ob = aiDocAssy(obOf(Item));
        if (ob == null)
            return pk.dcFor(0);
        return ob.ComponentDefinition.IsiAssemblyMember ? pk.dcFor(ob) : pk.dcFor(0);
    }

    public new kyPick WithInDc(Dictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        return this;
    }

    private new kyPick WithOutDc(Dictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        return this;
    }

    private new IDictionary dcIn()
    {
        return pk.dcIn;
    }

    private new IDictionary dcOut()
    {
        return pk.dcOut;
    }

    public new kyPick AfterScanning(Dictionary dSrc)
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

    private new kyPick Itself()
    {
        return this;
    }

    private kyPick kyPick_Itself()
    {
        return Itself();
    }

    private kyPick kyPick_WithInDc(Dictionary Dict)
    {
        return WithInDc(Dict);
    }

    private kyPick kyPick_WithOutDc(Dictionary Dict)
    {
        return WithOutDc(Dict);
    }
}