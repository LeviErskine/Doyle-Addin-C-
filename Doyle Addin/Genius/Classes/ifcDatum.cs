class ifcDatum
{
    private object obNow;
    private Variant valNow;
    private Variant valWas;
    /// 

    /// 

    public ifcDatum Connect(object This)
    {
        if (This == null)
            Connect = this;
        else if (This is ifcDatum)
            Connect = This;
        else if (false)
        {
        }
        else
            Connect = obIfcDatum(This);
    }

    public ifcDatum MakeValue(Variant This)
    {
        MakeValue = this;
    }

    public ifcDatum Commit()
    {
        Commit = this;
    }

    public ifcDatum Itself()
    {
        Itself = this;
    }

    public bool Connected(object ToThis = null)
    {
        if (ToThis == null)
            Connected = !obNow == null;
        else
            Connected = obNow == ToThis;
    }

    public Variant Value()
    {
        Value = valNow;
    }

    private void Class_Initialize()
    {
        valNow = "";
    }
}