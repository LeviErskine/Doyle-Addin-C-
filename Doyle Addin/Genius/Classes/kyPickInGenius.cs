using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

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

    public IDictionary dcFor(dynamic Item)
    {
        Document ob = aiDocument(obOf(Item));
        if (ob == null || ob.DocumentType != kPartDocumentObjectOr)
            return pk.dcFor(0);
        string pn = ob.PropertySets.get_Item(gnDesign).get_Item(pnPartNum).Value;
        if (Strings.Len(pn) <= 0) return pk.dcFor(0);
        {
            var withBlock = cn.Execute(sqlA + pn + "'");
            // If .BOF Or .EOF Then
            return withBlock.Fields("ct").Value > 0 ? pk.dcFor(ob) : pk.dcFor(0);
        }
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

    public IDictionary dcIn()
    {
        return pk.dcIn;
    }

    public IDictionary dcOut()
    {
        return pk.dcOut;
    }

    public kyPick AfterScanning(Dictionary dSrc)
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

    public kyPick Itself()
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