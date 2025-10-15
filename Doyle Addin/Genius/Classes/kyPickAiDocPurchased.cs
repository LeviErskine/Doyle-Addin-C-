namespace Doyle_Addin.Genius.Classes;

class kyPickAiDocPurchased : kyPick
{
    private kyPick pk;
    // 

    // kyPick Implementation code follows

    // 

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
        return dcFor(Item);
    }
    // 

    // General Class handling code follows

    // 

    private void Class_Initialize()
    {
        pk = new kyPick();
    }
    // 

    // kyPickAiDocPurchased Class

    // implementation code follows

    // 

    public kyPick Itself()
    {
        return this;
    }

    public kyPick WithInDc(Dictionary Dict)
    {
        pk = pk.WithInDc(Dict);
        return this;
    }

    public kyPick WithOutDc(Dictionary Dict)
    {
        pk = pk.WithOutDc(Dict);
        return this;
    }

    public kyPick AfterScanning(Dictionary dSrc)
    {
        return kyPick_AfterScanning(dSrc);
    }

    public Dictionary dcIn()
    {
        return pk.dcIn;
    }

    public Dictionary dcOut()
    {
        return pk.dcOut;
    }

    public IDictionary dcFor(dynamic Item)
    {
        BOMStructureEnum ck;

        // REV[2022.03.08.1021]
        // Added BOMStructureEnum variable ck
        // to collect BOMStructureEnum for each
        // relevant Document type, and consolidate
        // BOMStructureEnum check to one block
        // following Doc type accommodation.
        Document ob = aiDocument(obOf(Item));

        if (ob == null)
            ck = kDefaultBOMStructure;
        else
            ck = ob.DocumentType switch
            {
                kPartDocumentObject => aiDocPart(ob).ComponentDefinition.BOMStructure,
                kAssemblyDocumentObject => aiDocAssy(ob).ComponentDefinition.BOMStructure,
                _ => kDefaultBOMStructure
            };

        if (ck == kPurchasedBOMStructure)
            return pk.dcFor(ob);
        // REV[2022.03.08.1038]
        // Additional checks on Item
        // Family and File Location
        // NOTE that this is more of
        // a "soft" identification
        // of likely purchased parts,
        // and might or might not be
        // appropriate to apply.
        var pr = ob.PropertySets.get_Item(gnDesign).get_Item(pnFamily);
        return InStr(1, ob.FullFileName, @"\Doyle_Vault\Designs\purchased\") +
            InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + pr.Value + "|") > 0
                ? pk.dcFor(ob)
                : pk.dcFor(0);
    }
}