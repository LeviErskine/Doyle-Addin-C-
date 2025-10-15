namespace Doyle_Addin.Genius.Classes;

class kyPickAiSheetMetal : kyPick
{
    private kyPick pk;
    private const string txVersion = "kyPickAiSheetMetal v0.0.0.1 [2022.03.08.1332]";
    // prior Versions

    // ""

    // 

    // kyPick Implementation code follows

    // 

    private kyPick kyPick_Itself()
    {
        return this;
    }

    private kyPick kyPick_WithInDc(IDictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        return this;
    }

    private kyPick kyPick_WithOutDc(IDictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        return this;
    }

    private kyPick kyPick_AfterScanning(IDictionary dSrc)
    {
        {
            foreach (var ky in dSrc.Keys)
            {
                {
                    var withBlock1 = kyPick_DcFor(dSrc.get_Item(ky));
                    if (withBlock1.Exists(ky))
                        Debugger.Break();
                    else
                        withBlock1.Add(ky, dSrc.get_Item(ky));
                }
            }
        }
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
        pk = new kyPick();
    }
    // 

    // kyPickAiSheetMetal Class

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
        return pk.dcFor(ob is SheetMetalComponentDefinition ? ob.Document : 0);
    }
    // 

    // Version code follows

    // 

    public string Version()
    {
        return txVersion;
    }
}