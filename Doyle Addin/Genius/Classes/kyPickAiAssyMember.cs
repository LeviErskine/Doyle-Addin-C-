class kyPickAiAssyMember : kyPick
{
    private kyPick pk;

    private void Class_Initialize()
    {
        pk = new kyPick();
    }

    public Scripting.IDictionary dcFor(Variant Item)
    {
        Inventor.AssemblyDocument ob;

        ob = aiDocAssy(obOf(Item));
        if (ob == null)
            dcFor = pk.dcFor(0);
        else if (ob.ComponentDefinition.IsiAssemblyMember)
            dcFor = pk.dcFor(ob);
        else
            dcFor = pk.dcFor(0);
    }

    public kyPick WithInDc(Scripting.Dictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        WithInDc = this;
    }

    public kyPick WithOutDc(Scripting.Dictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        WithOutDc = this;
    }

    public Scripting.Dictionary dcIn()
    {
        dcIn = pk.dcIn;
    }

    public Scripting.Dictionary dcOut()
    {
        dcOut = pk.dcOut;
    }

    public kyPick AfterScanning(Scripting.Dictionary dSrc)
    {
        AfterScanning = kyPick_AfterScanning(dSrc);
    }

    private kyPick kyPick_AfterScanning(Scripting.IDictionary dSrc)
    {
        Variant ky;

        {
            var withBlock = dSrc;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcFor(withBlock.Item(ky));
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(ky, dSrc.Item(ky));
                }
            }
        }
        kyPick_AfterScanning = this;
    }

    /// kyPick Implementation code follows

    /// 
    private Scripting.IDictionary kyPick_DcFor(Variant Item)
    {
        kyPick_DcFor = dcFor(Item);
    }

    private Scripting.IDictionary kyPick_DcIn()
    {
        kyPick_DcIn = dcIn();
    }

    private Scripting.IDictionary kyPick_DcOut()
    {
        kyPick_DcOut = dcOut();
    }

    public kyPick Itself()
    {
        Itself = this;
    }

    private kyPick kyPick_Itself()
    {
        kyPick_Itself = this.Itself();
    }

    private kyPick kyPick_WithInDc(Scripting.IDictionary Dict)
    {
        kyPick_WithInDc = WithInDc(Dict);
    }

    private kyPick kyPick_WithOutDc(Scripting.IDictionary Dict)
    {
        kyPick_WithOutDc = WithOutDc(Dict);
    }
}