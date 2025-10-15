using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

class aiPropSetter
{
    private string[] ls;

    public string[] PropList()
    {
        return ls;
    }

    public Dictionary dcPropsIn(Document ad, Dictionary dc = null)
    {
        while (true)
        {
            var ps = ad.PropertySets.GetType(gnCustom);

            if (dc == null)
                return dcPropsIn(ad, new Dictionary());
            if (Information.IsArray(ls))
            {
                foreach (var ky in ls)
                {
                    var pr = aiGetProp(ps, Convert.ToString(ky), 1);
                    if (pr == null)
                    {
                    }
                    else
                        dc.Add(ky, pr);
                }

                return dc;
            }

            if (VarType(ls) == Constants.vbString)
            {
                Debugger.Break(); // shouldn't wind up here
                continue;
            }

            Debugger.Break(); // or here, either

            break;
        }
    }

    private void Class_Initialize()
    {
        ls = string.Split("andrew patrick thompson", " ");
    }
}