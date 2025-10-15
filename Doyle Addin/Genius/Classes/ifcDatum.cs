namespace Doyle_Addin.Genius.Classes;

public class ifcDatum
{
    private dynamic obNow;
    private dynamic valNow;
    private dynamic valWas;
    // 

    // 

    public ifcDatum Connect(dynamic This)
    {
        switch (This)
        {
            case null:
                return this;
            case ifcDatum:
                return This;
            default:
            {
                if (false)
                {
                }

                return obIfcDatum(This);
            }
        }
    }

    public ifcDatum MakeValue(dynamic This)
    {
        return this;
    }

    public ifcDatum Commit()
    {
        return this;
    }

    public ifcDatum Itself()
    {
        return this;
    }

    public bool Connected(dynamic ToThis = null)
    {
        if (ToThis == null)
            return false;
        return obNow == ToThis;
    }

    public dynamic Value()
    {
        return valNow;
    }

    private void Class_Initialize()
    {
        valNow = "";
    }
}