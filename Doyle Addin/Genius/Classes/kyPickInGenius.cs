class kyPickInGenius : kyPick
{

    // Private Const sqlA As String = "select ItemId from vgMfiItems where Item='"
    private const string sqlA = "select count(ItemId) as ct from vgMfiItems where Item='";

    private kyPick pk;
    private ADODB.Connection cn;
    // Private cm As ADODB.Command

    private void Class_Initialize()
    {
        pk = new kyPick();
        cn = cnGnsDoyle;
    }

    public Scripting.IDictionary dcFor(Variant Item)
    {
        Inventor.Document ob;
        string pn;

        ob = aiDocument(obOf(Item));
        if (ob == null)
            dcFor = pk.dcFor(0);
        else if (ob.DocumentType == kPartDocumentObjectOr) /* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
        {
            pn = ob.PropertySets.Item(gnDesign).Item(pnPartNum).Value;
            if (Strings.Len(pn) > 0)
            {
                {
                    var withBlock = cn.Execute(sqlA + pn + "'");
                    // If .BOF Or .EOF Then
                    if (withBlock.Fields("ct").Value > 0)
                        dcFor = pk.dcFor(ob);
                    else
                        dcFor = pk.dcFor(0);
                }
            }
            else
                dcFor = pk.dcFor(0);
        }
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