class kyPickAiSheetMetal : kyPick
{
    private kyPick pk;
    private const string txVersion = "kyPickAiSheetMetal v0.0.0.1 [2022.03.08.1332]";
    /// prior Versions

    /// ""

    /// 

    /// kyPick Implementation code follows

    /// 

    private kyPick kyPick_Itself()
    {
        kyPick_Itself = this;
    }


    private kyPick kyPick_WithInDc(Scripting.IDictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        kyPick_WithInDc = this;
    }

    private kyPick kyPick_WithOutDc(Scripting.IDictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        kyPick_WithOutDc = this;
    }


    private kyPick kyPick_AfterScanning(Scripting.IDictionary dSrc)
    {
        Variant ky;

        {
            var withBlock = dSrc;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = kyPick_DcFor(withBlock.Item(ky));
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(ky, dSrc.Item(ky));
                }
            }
        }
        kyPick_AfterScanning = this;
    }


    private Scripting.IDictionary kyPick_DcIn()
    {
        kyPick_DcIn = dcIn();
    }

    private Scripting.IDictionary kyPick_DcOut()
    {
        kyPick_DcOut = dcOut();
    }


    private Scripting.IDictionary kyPick_DcFor(Variant Item)
    {
        Inventor.PartDocument ob; // .Document

        ob = aiDocPart(aiDocument(obOf(Item)));
        if (ob == null)
            kyPick_DcFor = pk.dcFor(0);
        else
            kyPick_DcFor = g0f1(ob.ComponentDefinition);
    }
    /// 

    /// General Class handling code follows

    /// 

    private void Class_Initialize()
    {
        pk = new kyPick();
    }
    /// 

    /// kyPickAiSheetMetal Class

    /// implementation code follows

    /// 

    public kyPick Itself()
    {
        Itself = this;
    }


    public kyPick WithInDc(Scripting.Dictionary Dict)
    {
        WithInDc = kyPick_WithInDc(Dict);
    }

    public kyPick WithOutDc(Scripting.Dictionary Dict)
    {
        WithOutDc = kyPick_WithOutDc(Dict);
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


    public Scripting.IDictionary dcFor(Variant Item)
    {
        dcFor = kyPick_DcFor(Item);
    }
    /// 

    /// Internal support code follows

    /// 

    private Scripting.Dictionary g0f0(Inventor.PartDocument ob)
    {
        if (ob == null)
            g0f0 = pk.dcFor(0);
        else
            g0f0 = g0f1(ob.ComponentDefinition);
    }

    private Scripting.Dictionary g0f1(Inventor.PartComponentDefinition ob)
    {
        if (ob is Inventor.SheetMetalComponentDefinition)
            g0f1 = pk.dcFor(ob.Document);
        else
            g0f1 = pk.dcFor(0);
    }
    /// 

    /// Version code follows

    /// 

    public string Version()
    {
        Version = txVersion;
    }
}