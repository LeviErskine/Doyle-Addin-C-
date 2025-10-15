namespace Doyle_Addin.Genius.Classes;

class kyPickAiShMtl4sure : kyPick
{
    private kyPick pk;
    private const string txVersion = "kyPickAiShMtl4sure v0.0.0.1 [2022.03.08.1336]";
    // prior Versions

    // ""

    // 

    // kyPick Implementation code follows

    // 

    private kyPick kyPick_Itself()
    {
        return this;
    }

    private kyPick kyPick_WithInDc(Dictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        return this;
    }

    private kyPick kyPick_WithOutDc(Dictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        return this;
    }

    private kyPick kyPick_AfterScanning(Dictionary dSrc)
    {
        Dictionary dcI;
        Dictionary dcO;

        {
            var withBlock = pk.AfterScanning(dSrc);
            dcI = withBlock.dcIn;
            dcO = withBlock.dcOut;
        }

        {
            foreach (var ky in dcI.Keys)
            {
                Dictionary dCk = kyPick_DcFor(dcI.get_Item(ky));
                if (dCk == dcI)
                    // don't need to do anything here
                    Debug.Print(""); // Breakpoint Landing
                else
                {
                    if (dCk.Exists(ky))
                        Debugger.Break();
                    else
                        dCk.Add(ky, dcI.get_Item(ky));
                }
            }
        }
        pk = pk.WithInDc(dcKeysMissing(dcI, dcO));

        return this;
    }

    private IDictionary kyPick_DcIn()
    {
        return dcIn();
    }

    private IDictionary kyPick_DcOut()
    {
        return dcOut();
    }

    private IDictionary kyPick_DcFor(dynamic Item)
    {
        PartDocument ob = // .Document
            aiDocPart(aiDocument(obOf(Item)));
        return ob == null ? pk.dcFor(0) : g0f1(ob.ComponentDefinition);
    }
    // 

    // General Class handling code follows

    // 

    private void Class_Initialize()
    {
        pk = new kyPickAiSheetMetal(); // kyPick
    }
    // 

    // kyPickAiShMtl4sure Class

    // implementation code follows

    // 

    public kyPick Itself()
    {
        return this;
    }

    public kyPick WithInDc(Dictionary Dict)
    {
        return kyPick_WithInDc(Dict);
    }

    public kyPick WithOutDc(Dictionary Dict)
    {
        return kyPick_WithOutDc(Dict);
    }

    public Dictionary dcIn()
    {
        return pk.dcIn;
    }

    public Dictionary dcOut()
    {
        return pk.dcOut;
    }

    public kyPick AfterScanning(Dictionary dSrc)
    {
        return kyPick_AfterScanning(dSrc);
    }

    public IDictionary dcFor(dynamic Item)
    {
        return kyPick_DcFor(Item);
    }
    // 

    // Internal support code follows

    // 

    private Dictionary g0f0(PartDocument ob)
    {
        return ob == null ? pk.dcFor(0) : g0f1(ob.ComponentDefinition);
    }

    private Dictionary g0f1(PartComponentDefinition ob)
    {
        return ob is SheetMetalComponentDefinition ? g0f2(ob) : pk.dcFor(0);
    }

    private Dictionary g0f2(SheetMetalComponentDefinition ob)
    {
        if (ob is not { HasFlatPattern: true })
            return pk.dcFor(0);
        // '  check stated thickness...
        double smThk = ob.Thickness.Value;
        // Debug.Print "Thickness: " & CStr(smThk)

        // '  against flat pattern height
        double fpHgt;
        {
            var withBlock1 = nuAiBoxData().UsingBox(ob.FlatPattern.RangeBox);
            // Debug.Print .Dump()
            Debug.Print(""); // Breakpoint Landing
            fpHgt = withBlock1.SpanZ;
        }

        var dfRns = fpHgt - smThk;
        return pk.dcFor(double.Abs(dfRns) < 0.001
            ?
            // '  assume it's valid
            ob.Document
            :
            // '  assume likely not
            // Stop
            0);
    }
    // 

    // Version code follows

    // 

    public string Version()
    {
        return txVersion;
    }
}