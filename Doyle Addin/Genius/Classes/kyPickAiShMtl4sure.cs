class kyPickAiShMtl4sure : kyPick
{
    private kyPick pk;
    private const string txVersion = "kyPickAiShMtl4sure v0.0.0.1 [2022.03.08.1336]";
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
        Scripting.Dictionary dcI;
        Scripting.Dictionary dcO;
        Scripting.Dictionary dCk;
        Variant ky;

        {
            var withBlock = pk.AfterScanning(dSrc);
            dcI = withBlock.dcIn;
            dcO = withBlock.dcOut;
        }

        {
            var withBlock = dcI;
            foreach (var ky in withBlock.Keys)
            {
                dCk = kyPick_DcFor(withBlock.Item(ky));
                if (dCk == dcI)
                    // don't need to do anything here
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                else
                {
                    var withBlock1 = dCk;
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(ky, dcI.Item(ky));
                }
            }
        }
        pk = pk.WithInDc(dcKeysMissing(dcI, dcO));

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
        pk = new kyPickAiSheetMetal(); // kyPick
    }
    /// 

    /// kyPickAiShMtl4sure Class

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
            g0f1 = g0f2(ob);
        else
            g0f1 = pk.dcFor(0);
    }

    private Scripting.Dictionary g0f2(Inventor.SheetMetalComponentDefinition ob)
    {
        double smThk;
        double fpHgt;
        double dfRns;

        if (ob == null)
            g0f2 = pk.dcFor(0);
        else
        {
            var withBlock = ob;
            if (withBlock.HasFlatPattern)
            {
                // '  check stated thickness...
                smThk = withBlock.Thickness.Value;
                // Debug.Print "Thickness: " & CStr(smThk)

                // '  against flat pattern height
                {
                    var withBlock1 = nuAiBoxData().UsingBox(withBlock.FlatPattern.RangeBox);
                    // Debug.Print .Dump()
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                    fpHgt = withBlock1.SpanZ;
                }

                dfRns = fpHgt - smThk;
                if (Abs(dfRns) < 0.001)
                    // '  assume it's valid
                    g0f2 = pk.dcFor(ob.Document);
                else
                    // '  assume likely not
                    // Stop
                    g0f2 = pk.dcFor(0);
            }
            else
                g0f2 = pk.dcFor(0);
        }
    }
    /// 

    /// Version code follows

    /// 

    public string Version()
    {
        Version = txVersion;
    }
}