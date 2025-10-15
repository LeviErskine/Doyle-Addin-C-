namespace Doyle_Addin.Genius.Classes;

class aiPropLink
{
    private Dictionary dc;

    private void Class_Initialize()
    {
        dc = new Dictionary();
    }

    private void Class_Terminate()
    {
        dc.RemoveAll();
        dc = null;
    }

    public aiPropLink PrepFrom(Document AiDoc)
    {
        {
            foreach (PropertySet ps in AiDoc.PropertySets)
            {
                var psName = ps.Name;
                foreach (var prName in from Property pr in ps select pr.Name)
                {
                    if (dc.Exists(prName))
                        Debugger.Break();
                    else
                        dc.Add(prName, psName);
                }
            }
        }

        return this;
    }
}