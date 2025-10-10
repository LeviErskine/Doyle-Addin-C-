class kyPick
{
    private Scripting.Dictionary dcGrpIn;
    private Scripting.Dictionary dcGrpOut;

    private void Class_Initialize()
    {
        dcGrpIn = new Scripting.Dictionary();
        dcGrpOut = new Scripting.Dictionary();
    }

    public kyPick Itself()
    {
        Itself = this;
    }

    public kyPick WithInDc(Scripting.Dictionary Dict)
    {
        dcGrpIn = dcNewIfNone(Dict);
        WithInDc = this;
    }

    public kyPick WithOutDc(Scripting.Dictionary Dict)
    {
        dcGrpOut = dcNewIfNone(Dict);
        WithOutDc = this;
    }

    public kyPick AfterScanning(Scripting.Dictionary dSrc)
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
        AfterScanning = this;
    }

    public Scripting.Dictionary dcIn()
    {
        dcIn = dcGrpIn;
    }

    public Scripting.Dictionary dcOut()
    {
        dcOut = dcGrpOut;
    }

    public Scripting.Dictionary dcFor(Variant Item)
    {
        if (IsObject(Item))
            dcFor = dcGrpIn;
        else
            dcFor = dcGrpOut;
    }
}