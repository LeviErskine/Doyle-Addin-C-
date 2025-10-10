class dcPopulator
{
    private Scripting.Dictionary dc;

    private Scripting.Dictionary checkDC(Scripting.Dictionary dcIn = null/* TODO Change to default(_) if this is not a reference type */, long Opts = 0)
    {
        // '
        // '
        if (dcIn == null)
        {
            if (dc == null)
                dc = new Scripting.Dictionary();
        }
        else if (dc == null)
            dc = dcIn;
        else
        {
            System.Diagnostics.Debugger.Break();
            if (Opts & 1)
            {
            }
        }

        checkDC = dc;
    }

    public dcPopulator Using(Scripting.Dictionary Dict, long Opts = 0)
    {
        // '
        // '
        {
            var withBlock = checkDC(Dict, Opts);
        }

        using ( == this)
        {
        }
    }

    public dcPopulator Setting(Variant Key, Variant Item)
    {
        {
            var withBlock = checkDC();
            if (withBlock.Exists(Key))
                withBlock.Remove(Key);
            withBlock.Add(Key, Item);
        }

        ting = this;
    }

    public long Count()
    {
        Count = checkDC().Count;
    }

    public Scripting.Dictionary Dictionary(Scripting.Dictionary Dict = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Dictionary = checkDC(Dict);
    }

    public Scripting.Dictionary Matching(Variant KeySet)
    {
        Scripting.Dictionary rt;
        Variant ky;

        if (IsArray(KeySet))
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = checkDC();
                foreach (var ky in KeySet)
                {
                    if (withBlock.Exists(ky))
                    {
                        if (rt.Exists(ky))
                        {
                        }
                        else
                            rt.Add(ky, withBlock.Item(ky));
                    }
                }
            }
        }
        else
            rt = Matching(Array(KeySet));

        Matching = rt;
    }

    public bool Exists(Variant Key)
    {
        {
            var withBlock = Dictionary();
            Exists = withBlock.Exists(Key);
        }
    }

    public Variant Item(Variant Key)
    {
        Variant rt;

        {
            var withBlock = Dictionary();
            if (withBlock.Exists(Key))
            {
                rt = Array(withBlock.Item(Key));
                if (IsObject(rt(0)))
                    Item = rt(0);
                else
                    Item = rt(0);
            }
            else
                Item = Empty;
        }
    }

    /// OPTIONAL section for Inventor.NameValueMap

    /// the following functions will ONLY compile

    /// within the Autodesk Inventor VBA environment,

    /// or other environment which supports the same

    /// NameValueMap classes and structures.

    /// 

    /// It should be disabled or deleted for use

    /// outside of any such environment.

    /// 
    public dcPopulator UsingNameValMap(Inventor.NameValueMap NVMap, long Opts = 0)
    {
        // '
        // '
        UsingNameValMap = ;
    }

    public Inventor.NameValueMap NameValMap(Inventor.NameValueMap NVMap = null/* TODO Change to default(_) if this is not a reference type */)
    {
        NameValMap = dc2aiNameValMap(Dictionary(), NVMap);
    }
}