using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class dvlAiNameValMap
{
    // dvlAiNameValMap -- functions to streamline

    // translation of data from Dictionary

    // Objects to Inventor NameValueMap

    // Objects, and vice-versa.

    // 

    // NOTE: these functions MIGHT be supplanted by

    // their addition to or implementation in

    // Class Module dcPopulator

    // 

    public static Func<NameValueMap> dc2aiNameValMap(Dictionary dc, Func<NameValueMap> mp = null)
    {
        Func<NameValueMap> rt;

        if (mp == null)
        {
            {
                var withBlock = ThisApplication.TransientObjects;
                rt = dc2aiNameValMap(dc, withBlock.CreateNameValueMap); // NameValueMap cannot
            }
        }
        else
        {
            rt = mp;
            {
                foreach (var ky in dc.Keys)
                {
                    var nm = Convert.ToString(ky as string);
                    var it = new[] { dc.get_Item(ky) };

                    if (IsObject(it(0)))
                    {
                        // dynamic handling not
                        // implemented as yet.
                        // A general solution is
                        // not likely possible.

                        // UPDATE[2022.07.05.1319]
                        // it appears that a NameValueMap
                        // CAN include another NameValueMap
                        // as a Value, thus enabling multi-
                        // level NameValueMaps. Whether
                        // other Objects can be so contained
                        // is likely a question best left
                        // unexplored for now, but it seems
                        // at least sub-Dictionaries and
                        // NameValueMaps can be processed.
                        if (it(0) is Dictionary)
                            rt.Add(nm, dc2aiNameValMap(obOf(it(0))));
                        else if (it(0) is NameValueMap)
                            rt.Add(nm, it(0));
                    }
                    else
                    {
                        Information.Err.Clear();
                        rt.Add(nm, it(0));
                        if (!Information.Err.Number) continue;
                        VbMsgBoxResult ck = MessageBox.Show(
                            Join(
                                new[]
                                {
                                    "Key \"" + nm, "\" Value (" + Convert.ToHexString(it(0)) + ")",
                                    "could not be set.", "The Key will not", "be assigned.", "",
                                    "Click OK to continue", "(Cancel to debug)"
                                }, Constants.vbCrLf),
                            Constants.vbOKCancel & Constants.vbExclamation, "Assignment Error!");
                        if (ck == Constants.vbCancel)
                            Debugger.Break(); // to Debug
                    }
                }
            }
        }

        return rt;
    }

    public static Dictionary dcFromAiNameValMap(NameValueMap mp, Dictionary dc = null)
    {
        Dictionary rt;

        if (dc == null)
            rt = dcFromAiNameValMap(mp, new Dictionary());
        else
        {
            rt = dc; // mp
            {
                long mx = mp.Count;
                for (long dx = 1; dx <= mx; dx++)
                {
                    string nm = mp.Name(dx);

                    rt.Add(nm, itemForDcOr(mp.get_Item(nm)));
                }
            }
        }

        return rt;
    }

    public static dynamic itemForDcOr(dynamic it, bool tp = false)
    {
        // itemForDcOr -- given item it, return
        // transformation according to type,
        // and type of result desired for
        // Dictionary and NameValueMap Objects,
        // according to value of tp:
        // NameValueMap for tp = 1
        // Dictionary for any other
        // value, including default 0
        // all other types of item are returned
        // as is, including Objects other than
        // Dictionary and NameValueMap
        // 

        switch (it)
        {
            case Array:
            {
                long mx = UBound(it);
                dynamic rt;
                if (mx < LBound(it))
                    rt = Array.Empty<dynamic>();
                else
                {
                    rt(mx);
                    for (long dx = 0; dx <= mx; dx++)
                    {
                        dynamic ck = new[] { itemForDcOr(it(dx), tp) };
                        ;
                        if (IsObject(ck(0)))
                            rt(dx) = ck(0);
                        else
                            rt(dx) = ck(0);
                    }
                }

                break;
            }
            case not null:
                switch (it)
                {
                    case NameValueMap:
                        rt = dcFromAiNameValMap(obOf(it));
                        break;
                    case Dictionary:
                    {
                        Dictionary dc = it;
                        {
                            var withBlock = new dcPopulator();
                            foreach (var ck in dc.Keys)
                                withBlock.Setting(ck, itemForDcOr(dc.get_Item(ck), tp));

                            if (tp == 1)
                                rt = withBlock.NameValMap();
                            else
                                rt = withBlock.Dictionary();
                        }
                        break;
                    }
                    default:
                        rt = it;
                        break;
                }

                break;
            default:
                rt = it;
                break;
        }

        return rt;
    }

    public static NameValueMap nuAiNameValMap(Document Using = null)
    {
        if ()
        {
            {
                var withBlock = ThisApplication.TransientObjects;
                return withBlock.CreateNameValueMap;
            }
        }

        return dc2aiNameValMap;
    }
}