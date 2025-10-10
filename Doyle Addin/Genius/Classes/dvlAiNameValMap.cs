class dvlAiNameValMap
{
    /// dvlAiNameValMap -- functions to streamline

    /// translation of data from Dictionary

    /// Objects to Inventor NameValueMap

    /// Objects, and vice-versa.

    /// 

    /// NOTE: these functions MIGHT be supplanted by

    /// their addition to or implementation in

    /// Class Module dcPopulator

    /// 

    public Inventor.NameValueMap dc2aiNameValMap(Scripting.Dictionary dc, Inventor.NameValueMap mp = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Inventor.NameValueMap rt;
        Variant ky;
        Variant it;
        string nm;
        VbMsgBoxResult ck;

        if (mp == null)
        {
            {
                var withBlock = ThisApplication.TransientObjects;
                rt = dc2aiNameValMap(dc, withBlock.CreateNameValueMap);   // NameValueMap cannot
            }
        }
        else
        {
            rt = mp;
            {
                var withBlock = dc;
                foreach (var ky in withBlock.Keys)
                {
                    nm = System.Convert.ToHexString(ky);
                    it = Array(withBlock.Item(ky));

                    if (IsObject(it(0)))
                    {
                        /// Object handling not
                        /// implemented as yet.
                        /// A general solution is
                        /// not likely possible.

                        /// UPDATE[2022.07.05.1319]
                        /// it appears that a NameValueMap
                        /// CAN include another NameValueMap
                        /// as a Value, thus enabling multi-
                        /// level NameValueMaps. Whether
                        /// other Objects can be so contained
                        /// is likely a question best left
                        /// unexplored for now, but it seems
                        /// at least sub-Dictionaries and
                        /// NameValueMaps can be processed.
                        if (it(0) is Scripting.Dictionary)
                            rt.Add(nm, dc2aiNameValMap(obOf(it(0))));
                        else if (it(0) is Inventor.NameValueMap)
                            rt.Add(nm, it(0));
                        else
                        {
                        }
                    }
                    else
                    {
                        Information.Err.Clear();
                        rt.Add(nm, it(0));
                        if (Information.Err.Number)
                        {
                            ck = MsgBox(Join(Array("Key \"" + nm, "\" Value (" + System.Convert.ToHexString(it(0)) + ")", "could not be set.", "The Key will not", "be assigned.", "", "Click OK to continue", "(Cancel to debug)"), Constants.vbNewLine), Constants.vbOKCancel + Constants.vbExclamation, "Assignment Error!");
                            if (ck == Constants.vbCancel)
                                System.Diagnostics.Debugger.Break();// to Debug
                        }
                    }
                }
            }
        }

        dc2aiNameValMap = rt;
    }

    public Scripting.Dictionary dcFromAiNameValMap(Inventor.NameValueMap mp, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        string nm;
        long mx;
        long dx;

        if (dc == null)
            rt = dcFromAiNameValMap(mp, new Scripting.Dictionary());
        else
        {
            rt = dc; // mp
            {
                var withBlock = mp // dc
       ;
                mx = withBlock.Count;
                for (dx = 1; dx <= mx; dx++)
                {
                    nm = withBlock.Name(dx);

                    rt.Add(nm, itemForDcOr(withBlock.Item(nm)));
                }
            }
        }

        dcFromAiNameValMap = rt;
    }

    public Variant itemForDcOr(Variant it, long tp = 0)
    {
        /// itemForDcOr -- given item it, return
        /// transformation according to type,
        /// and type of result desired for
        /// Dictionary and NameValueMap Objects,
        /// according to value of tp:
        /// NameValueMap for tp = 1
        /// Dictionary for any other
        /// value, including default 0
        /// all other types of item are returned
        /// as is, including Objects other than
        /// Dictionary and NameValueMap
        /// 
        Variant rt;
        Variant ck;
        long mx;
        long dx;
        Scripting.Dictionary dc;

        if (IsArray(it))
        {
            mx = UBound(it);
            if (mx < LBound(it))
                rt = Array();
            else
            {
                rt(mx);
                for (dx = 0; dx <= mx; dx++)
                {
                    ck = Array(itemForDcOr(it(dx), tp));
                    if (IsObject(ck(0)))
                        rt(dx) = ck(0);
                    else
                        rt(dx) = ck(0);
                }
            }
        }
        else if (IsObject(it))
        {
            if (it is Inventor.NameValueMap)
                rt = dcFromAiNameValMap(obOf(it));
            else if (it is Scripting.Dictionary)
            {
                dc = it;
                {
                    var withBlock = new dcPopulator();
                    foreach (var ck in dc.Keys)
                        withBlock.Setting(ck, itemForDcOr(dc.Item(ck), tp));

                    if (tp == 1)
                        rt = withBlock.NameValMap();
                    else
                        rt = withBlock.Dictionary();
                }
            }
            else
                rt = it;
        }
        else
            rt = it;

        if (IsObject(rt))
            itemForDcOr = rt;
        else
            itemForDcOr = rt;
    }

    public Inventor.NameValueMap nuAiNameValMap(Inventor.Document Using = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if ()
        {
            {
                var withBlock = ThisApplication.TransientObjects;
                nuAiNameValMap = withBlock.CreateNameValueMap;
            }
        }
        else
            nuAiNameValMap = dc2aiNameValMap;
    }
}