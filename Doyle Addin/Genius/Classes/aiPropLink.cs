class aiPropLink
{
    private Scripting.Dictionary dc;

    private void Class_Initialize()
    {
        dc = new Scripting.Dictionary();
    }

    private void Class_Terminate()
    {
        dc.RemoveAll();
        dc = null;
    }

    public aiPropLink PrepFrom(Inventor.Document AiDoc)
    {
        Inventor.PropertySet ps;
        Inventor.Property pr;
        string psName;
        string prName;

        {
            var withBlock = AiDoc;
            foreach (var ps in withBlock.PropertySets)
            {
                psName = ps.Name;
                foreach (var pr in ps)
                {
                    prName = pr.Name;
                    if (dc.Exists(prName))
                        System.Diagnostics.Debugger.Break();
                    else
                        dc.Add(prName, psName);
                }
            }
        }

        PrepFrom = this;
    }
}