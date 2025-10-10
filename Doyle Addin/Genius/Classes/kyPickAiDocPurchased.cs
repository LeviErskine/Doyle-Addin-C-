class kyPickAiDocPurchased : kyPick
{
    private kyPick pk;
    /// 

    /// kyPick Implementation code follows

    /// 

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
        kyPick_DcFor = dcFor(Item);
    }
    /// 

    /// General Class handling code follows

    /// 

    private void Class_Initialize()
    {
        pk = new kyPick();
    }
    /// 

    /// kyPickAiDocPurchased Class

    /// implementation code follows

    /// 

    public kyPick Itself()
    {
        Itself = this;
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


    public kyPick AfterScanning(Scripting.Dictionary dSrc)
    {
        AfterScanning = kyPick_AfterScanning(dSrc);
    }


    public Scripting.Dictionary dcIn()
    {
        dcIn = pk.dcIn;
    }

    public Scripting.Dictionary dcOut()
    {
        dcOut = pk.dcOut;
    }


    public Scripting.IDictionary dcFor(Variant Item)
    {
        Inventor.BOMStructureEnum ck;
        Inventor.Document ob;
        Inventor.Property pr;
        /// REV[2022.03.08.1021]
        /// Added BOMStructureEnum variable ck
        /// to collect BOMStructureEnum for each
        /// relevant Document type, and consolidate
        /// BOMStructureEnum check to one block
        /// following Doc type accommodation.

        ob = aiDocument(obOf(Item));

        if (ob == null)
            ck = kDefaultBOMStructure;
        else if (ob.DocumentType == kPartDocumentObject)
            ck = aiDocPart(ob).ComponentDefinition.BOMStructure;
        else if (ob.DocumentType == kAssemblyDocumentObject)
            ck = aiDocAssy(ob).ComponentDefinition.BOMStructure;
        else
            ck = kDefaultBOMStructure;

        if (ck == kPurchasedBOMStructure)
            dcFor = pk.dcFor(ob);
        else
        /// REV[2022.03.08.1038]
        /// Additional checks on Item
        /// Family and File Location
        /// NOTE that this is more of
        /// a "soft" identification
        /// of likely purchased parts,
        /// and might or might not be
        /// appropriate to apply.
        {
            var withBlock = ob;
            pr = withBlock.PropertySets.Item(gnDesign).Item(pnFamily);
            if (InStr(1, ob.FullFileName, @"\Doyle_Vault\Designs\purchased\") + InStr(1, "|D-HDWR|D-PTO|D-PTS|R-PTO|R-PTS|", "|" + pr.Value + "|") > 0)
                dcFor = pk.dcFor(ob);
            else
                dcFor = pk.dcFor(0);
        }
    }
}