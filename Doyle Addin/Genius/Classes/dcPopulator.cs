using System.Collections;

namespace Doyle_Addin.Genius.Classes;

class dcPopulator
{
    private Dictionary dc;

    private Dictionary checkDC(Dictionary dcIn = null, long Opts = 0)
    {
        if (dcIn == null)
        {
            dc ??= new Dictionary();
        }
        else if (dc == null)
            dc = dcIn;
        else
        {
            Debugger.Break();
            if ((Opts & 1) != 0)
            {
            }
        }

        return dc;
    }

    public dcPopulator Using(Dictionary Dict, long Opts = 0)
    {
        // '
        // '
        {
            checkDC(Dict, Opts);
        }
        return this;
    }

    public dcPopulator Setting(dynamic Key, dynamic Item)
    {
        {
            var withBlock = checkDC();
            if (withBlock.Exists(Key))
                withBlock.Remove(Key);
            withBlock.Add(Key, Item);
        }

        return this;
    }

    public long Count()
    {
        return checkDC().Count;
    }

    public Dictionary Dictionary(Dictionary Dict = null)
    {
        return checkDC(Dict);
    }

    public Dictionary Matching(dynamic KeySet)
    {
        Dictionary rt;

        if (KeySet is Array)
        {
            rt = new Dictionary();

            {
                var withBlock = checkDC();
                foreach (var ky in (IEnumerable)KeySet)
                {
                    if (!withBlock.Exists(ky)) continue;
                    if (rt.Exists(ky))
                    {
                    }
                    else
                        rt.Add(ky, withBlock.get_Item(ky));
                }
            }
        }
        else
            rt = Matching(new[] { KeySet });

        return rt;
    }

    public bool Exists(dynamic Key)
    {
        {
            var withBlock = Dictionary();
            return withBlock.Exists(Key);
        }
    }

    public dynamic Item(dynamic Key)
    {
        {
            var withBlock = Dictionary();
            if (!withBlock.Exists(Key)) return null;
            dynamic rt = new[] { withBlock.get_Item(Key) };
            if (IsObject(rt(0)))
                return rt(0);
            return rt(0);
        }
    }

    // OPTIONAL section for Inventor.NameValueMap
    // the following functions will ONLY compile
    // within the Autodesk Inventor VBA environment,
    // or other environment which supports the same
    // NameValueMap classes and structures.
    // 
    // It should be disabled or deleted for use
    // outside any such environment.
    // 
    public dcPopulator UsingNameValMap(NameValueMap NVMap, long Opts = 0)
    {
        return Using(dcFromAiNameValMap(NVMap), Opts);
    }

    public NameValueMap NameValMap(NameValueMap NVMap = null)
    {
        return dc2aiNameValMap(Dictionary(), NVMap);
    }
}