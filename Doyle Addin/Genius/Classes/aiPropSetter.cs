class aiPropSetter
{
    private string[] ls;

    public string[] PropList()
    {
        PropList = ls;
    }

    public Scripting.Dictionary dcPropsIn(Inventor.Document ad, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.PropertySet ps;
        Inventor.Property pr;
        Variant ky;

        ps = ad.PropertySets.Item(gnCustom);

        if (dc == null)
            dcPropsIn = dcPropsIn(ad, new Scripting.Dictionary());
        else if (Information.IsArray(ls))
        {
            foreach (var ky in ls)
            {
                pr = aiGetProp(ps, System.Convert.ToHexString(ky), 1);
                if (pr == null)
                {
                }
                else
                    dc.Add(ky, pr);
            }
            dcPropsIn = dc;
        }
        else if (VarType(ls) == Constants.vbString)
        {
            System.Diagnostics.Debugger.Break(); // shouldn't wind up here
            dcPropsIn = dcPropsIn(ad, dc);
        }
        else
            System.Diagnostics.Debugger.Break();// or here, either
    }

    private void Class_Initialize()
    {
        ls = Split("andrew patrick thompson", " ");
    }
}