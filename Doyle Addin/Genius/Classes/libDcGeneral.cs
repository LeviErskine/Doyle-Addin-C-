using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class libDcGeneral
{
    public  static  Dictionary dcCopy(Dictionary dc)
    {
        // dcCopy -- return a new Dictionary
        // copying the contents of the
        // one supplied, including COPIES
        // of any Dictionary Objects within.
        // 

        var rt = new Dictionary();

        if (dc == null)
        {
        }
        else
        {
            foreach (var ky in dc.Keys)
            {
                dynamic bx = new[] { dc.get_Item(ky) };
                Dictionary ck = dcOb(obOf(bx(0)));

                rt.Add(ky, ck == null ? bx(0) : dcCopy(ck));
            }

            ;
        }

        return rt;
    }

    public static  Dictionary dcWith(dynamic ky, dynamic it, Dictionary dc = null)
    {
        if (dc == null)
            return dcWith(ky, it, new Dictionary());
        {
            if (dc.Exists(ky))
                dc.Remove(ky);
            dc.Add(ky, it);
        }
        return dc;
    }

    public  static dcPopulator nuDcPopulator(Dictionary Dict = null, long Opts = 0)
    {
        {
            var withBlock = new dcPopulator();
            return withBlock.Using(Dict, Opts);
        }
    }

    public static  dynamic dcItemIfPresent(Dictionary dc, dynamic ky, VbVarType vtMissing = )
    {
        if (dc == null)
            return noVal(vtMissing);
        if (dc.Exists(ky))
        {
            return IsObject(dc.get_Item(ky)) ? dc.get_Item(ky) : dc.get_Item(ky);
        }

        if (vtMissing == Constants.vbObject)
            return noVal(vtMissing);
        return noVal(vtMissing);
    }

    public static  Dictionary dcInDc(string dcKey, Dictionary inDc)
    {
        return dcOb(obOf(dcItemIfPresent(inDc, dcKey, Constants.vbObject)));
    }

    public static  Dictionary dcInDcMk(dynamic ky, Dictionary dc)
    {
        // dcInDcMk --
        // 
        Dictionary rt;

        {
            if (dc.Exists(ky))
                rt = dcOb(dc.get_Item(ky));
            else
            {
                rt = new Dictionary();
                dc.Add(ky, rt);
            }
        }

        return rt;
    }

    public static  Dictionary dcOfKeys2match(dynamic ls)
    {
        // dcOfKeys2match -- generate a Dictionary
        // mapping a supplied Key, or Array
        // of Keys to itself or themselves
        // '
        // primary purpose is to provide
        // a 'reference' Dictionary of Keys
        // to be sought in other Dictionaries
        // '
        // (formerly d4g4f2)
        // 

        var rt = new Dictionary();

        if (IsArray(ls))
        {
            foreach (var ky in ls)
                rt.Add(ky, ky);
        }
        else
            rt.Add(ls, ls);

        return rt;
    }

    public static  Dictionary dcKeysInCommon(Dictionary d0, Dictionary d1, long pk = 0)
    {
        // dcKeysInCommon -- return intersection
        // of two Dictionary Objects based on
        // matching keys. Use optional pk value
        // to select which Dictionary's Items
        // to return in result:
        // 
        // 0 returns an array of Items from both
        // this is the default
        // 1 returns only Items from the first
        // 2 returns only Items from the second
        // 
        // NOTE that if either Dictionary dynamic
        // is not supplied (is Nothing), then
        // an null Dictionary is returned,
        // just as if an null Dictionary
        // had been supplied.
        // 
        Dictionary rt;

        if (d0 == null)
            rt = dcKeysInCommon(new Dictionary(), d1);
        else if (d1 == null)
            rt = dcKeysInCommon(d0, new Dictionary());
        else
        {
            rt = new Dictionary();
            {
                foreach (var ky in d0.Keys)
                {
                    if (!d1.Exists(ky)) continue;
                    dynamic ls = new[] { d0.get_Item(ky), d1.get_Item(ky) };
                    rt.Add(ky, new[] { ls, ls(0), ls(1)}(pk) ;
                }
            }
        }

        return rt;
    }

    public  static Dictionary dcKeysMissing(Dictionary dcWith, Dictionary dcWout)
    {
        // dcKeysMissing -- return difference
        // of first Dictionary dynamic minus
        // those keys found in the second.
        // 

        var rt = new Dictionary();
        {
            foreach (var ky in dcWith.Keys)
            {
                if (dcWout.Exists(ky))
                {
                }
                else
                    rt.Add(ky, dcWith.get_Item(ky));
            }
        }
        return rt;
    }

    public static Dictionary dcKeysCombined(Dictionary d0, Dictionary d1, long pk = 0)
    {
        // dcKeysCombined -- return union
        // of two Dictionary Objects. For
        // keys in both, use optional pk
        // value to select which Dictionary's
        // Items to return in result:
        // 
        // 0 returns an array of Items from both
        // this is the default
        // 1 returns only Items from the first
        // 2 returns only Items from the second
        // 
        Dictionary rt;

        if (pk > 1)
            rt = dcKeysCombined(d1, d0, 1);
        else
        {
            rt = dcKeysInCommon(d0, d1, pk);

            foreach (var ls in new[] { d0, d1 })
            {
                {
                    var withBlock = dcKeysMissing(ls, rt) // d0
                        ;
                    foreach (var ky in withBlock.Keys)
                        rt.Add(ky, withBlock.get_Item(ky));
                }
            }
        }

        // rt = dcKeysInCommon()

        if (d0 == null)
            Debugger.Break();
        else if (d1 == null)
            Debugger.Break();
        else
            // NOTE[2023.04.10.1449]
            // need to review what's going on here
            // this LOOKS like an effor to include items
            // left out of earlier operation, however,
            // it's not clear this stage isn't redundant.
        {
            foreach (var ky in d0.Keys)
            {
                if (!d1.Exists(ky)) continue;
                dynamic ls = new[] { d0.get_Item(ky), d1.get_Item(ky) };
                if (rt.Exists(ky))
                {
                    if (ConvertToJson(new[] { ls, ls(0), ls(1))(pk)) != ConvertToJson(rt.get_Item(ky) });
                    Debugger.Break();
                }
                else
                    rt.Add(ky, new[] { ls, ls(0), ls(1))(pk) };
            }
        }

        return rt;
    }

    public static  Dictionary dcOfIdent(dynamic src)
    {
        var rt = new Dictionary();

        if (src is Array)
        {
            foreach (var ky in src)
                rt.Add(ky, ky);
        }
        else
            rt.Add(src, src);

        return rt;
    }

    public static  Dictionary dcTransposed(Dictionary dc)
    {
        // Transpose Key values of supplied
        // Dictionary with matching Item values.
        // 
        // As written, will ONLY work against
        // a Dictionary whose Item values,
        // like its Keys, are unique.
        // 

        var rt = new Dictionary();
        {
            foreach (var fm in dc.Keys)
            {
                if (rt.Exists(dc.get_Item(fm)))
                    Debugger.Break();
                else
                    rt.Add.get_Item(fm);
            }
        }
        return rt;
    }

    public static  Dictionary dcTransGrouped(Dictionary dc)
    {
        // dcTransGrouped
        // (derived from dcTransposed)
        // 
        // generate new Dictionary "tramsposing"
        // Key values with matching Item values
        // in supplied Dictionary.
        // 
        // because more than one Key might map to
        // the same Item, the returned Dictionary
        // maps each (Item) Key to a Dictionary of
        // the Keys which originally mapped to it.
        // 
        // each of these Dictionaries in turn
        // maps the original Key back to its
        // corresponding Item once more,
        // since, HEY, it might as well!
        // 
        // obviously, an effor to work around
        // the main limitation of the original
        // dcTransposed, which can only work
        // against a Dictionary whose Items,
        // like its Keys, are unique.
        // 
        Dictionary it;

        var rt = new Dictionary();
        {
            foreach (var ky in dc.Keys)
            {
                dynamic ar = new[] { dc.get_Item(ky) };
                if (!rt.Exists(ar(0)))
                    rt.Add(ar(0), new Dictionary());
                dcOb(rt.get_Item(ar(0))).Add(ky, ar(0));
            }

            ;
        }
        return rt;
    }

    public static  long dcDepth(Dictionary dc)
    {
        // this function extracts the "depth"
        // of the supplied Dictionary dynamic,
        // that is, how many "levels" of
        // Dictionary objects it contains,
        // counting the supplied Dictionary
        // itself. When an actual Dictionary
        // is supplied, the value returned
        // will be at least 1. It will only
        // be zero when Nothing is supplied.
        // 

        if (dc == null)
            return 0;
        long rt = 0;

        {
            foreach (var ky in dc.Keys)
            {
                var ck = dcDepth(dcOb(obOf(dc.get_Item(ky))));
                if (ck > rt)
                    rt = ck;
            }
        }

        return 1 + rt;
    }

    public static  Dictionary dcFlattenDown(Dictionary Dict, long DownTo = 1)
    {
        // this function partially "flattens"
        // a hierarchy of Dictionary objects
        // (a Dictionary of Dictionaries,
        // potentially of more Dictionaries)
        // starting from the top, working
        // down to 'DownTo' levels
        // 

        if (Dict == null)
            return null;
        var rt = new Dictionary();

        {
            foreach (var ky in Dict.Keys)
            {
                dynamic it = new[] { Dict.get_Item(ky) };

                var sd = DownTo > 0 ? dcFlattenDown(dcOb(obOf(it(0))), DownTo - 1) : null;

                if (sd == null)
                    rt.Add(ky, it(0));
                else
                {
                    foreach (var sk in sd.Keys)
                        rt.Add(ky + "." + sk, sd.get_Item(sk));
                }
            }
        }

        return rt;
    }

    public static  Dictionary dcFlattenUp(Dictionary Dict, long DownFrom = 0)
    {
        // this function partially "flattens"
        // a hierarchy of Dictionary objects
        // (a Dictionary of Dictionaries,
        // potentially of more Dictionaries)
        // starting BELOW the top, skipping
        // DownFrom levels before "flattening"
        // Dictionaries below and at that level
        // 

        if (Dict == null)
            return null;
        var rt = new Dictionary();

        {
            foreach (var ky in Dict.Keys)
            {
                dynamic it = new[] { Dict.get_Item(ky) };
                Dictionary sd = dcOb(obOf(it(0)));

                if (sd == null)
                    rt.Add(ky, it(0));
                else if (DownFrom > 0)
                    rt.Add(ky, dcFlattenUp(sd, DownFrom - 1));
                else
                {
                    var withBlock1 = dcFlattenUp(sd, 0);
                    foreach (var sk in withBlock1.Keys)
                        rt.Add(ky + "." + sk, withBlock1.get_Item(sk));
                }
            }
        }

        return rt;
    }

    public static  Dictionary dcOfDcRekeyedSecToPri(Dictionary dc)
    {
        // dcOfDcRekeyedSecToPri - take Dictionary of Dictionaries
        // and return Dictionary of Dictionaries
        // with Secondary Keys promoted to Primary,
        // and Primary demoted to Secondary
        // 

        var rt = new Dictionary();

        {
            foreach (var kp in dc.Keys)
            {
                Dictionary sb = dcOb(dc.get_Item(kp));
                if (sb == null)
                    // Stop
                    Debug.Print(""); // Breakpoint Landing
                else
                {
                    foreach (var ks in sb.Keys)
                    {
                        dynamic ar = new[] { sb.get_Item(ks) };
                        {
                            if (!rt.Exists(ks))
                                rt.Add(ks, new Dictionary());

                            {
                                var withBlock3 = dcOb(rt.get_Item(ks));
                                if (withBlock3.Exists(kp))
                                    Debugger.Break();
                                else
                                    withBlock3.Add(kp, ar(0));
                            }
                        }
                    }
                }
            }
        }

        return rt;
    }

    public static  Dictionary dcCmpTextOf2items(string id0, string id1, dynamic it0, dynamic it1)
    {
        Dictionary rt;
        Dictionary dc0;
        Dictionary dc1;
        dynamic ob0;
        dynamic ob1;

        long ck = IIf(IsObject(it0), 1, 0) + IIf(IsObject(it1), 2, 0) + IIf(IsEmpty(it0), 4, 0) +
                  IIf(IsEmpty(it1), 8, 0);
        if (ck & 1)
            ck = ck + Interaction.IIf(it0 == null, 4, 0);
        if (ck & 2)
            ck = ck + Interaction.IIf(it1 == null, 8, 0);

        // UPDATE[2021.12.09]:
        // Added trap for Nothing Objects
        // to treat them as null as well
        var c2 = ck & 12; // $11XX
        if (c2)
        {
            // is null/Nothing
            {
                var withBlock = nuDcPopulator();
                rt = c2 switch
                {
                    // are equally null
                    12 => withBlock.Setting("==", "").Dictionary(),
                    // and id0 needs processed
                    // If ck And 1 Then
                    // rt = dcCmpTextOf2dc(' dcOb(it0), .Dictionary()' )
                    // Else
                    8 => withBlock.Setting(id0, it0).Dictionary(),
                    _ => withBlock.Setting(id1, it1).Dictionary()
                };
            }
        }
        else
            switch (ck)
            {
                // of comparable Dictionaries.
                // compare these recursively
                case 3:
                    rt = dcCmpTextOf2dc(dcOb(it0), dcOb(it1));
                    break;
                case 0:
                {
                    // just a couple of String
                    // or (hopefully) String-
                    // compatible values.
                    // compare them directly.
                    var tx0 = Convert.ToHexString(it0);
                    var tx1 = Convert.ToHexString(it1);

                    // rt = New Scripting.Dictionary
                    // With rt
                    {
                        var withBlock = nuDcPopulator();
                        rt = tx0 == tx1
                            ? (Dictionary)withBlock.Setting("==", tx0).Dictionary()
                            : (Dictionary)withBlock.Setting(id0, tx0).Setting(id1, tx1).Dictionary();
                    }
                    break;
                }
                // can't compare them in any way
                // just add each on its own side
                default:
                {
                    var withBlock = nuDcPopulator();
                    rt = withBlock.Setting(id0, it0).Setting(id1, it1).Dictionary();
                    break;
                }
            }

        return rt;
    }

    public static  Dictionary dcCmpTextOf2dc(Dictionary dc0, Dictionary dc1)
    {
        Dictionary rt;
        Dictionary qi;
        string tx0;
        string tx1;
        string nm;

        const string nm0 = "src0";
        const string nm1 = "src1";

        if (dc0 == null)
            rt = dcCmpTextOf2dc(new Dictionary(), dc1);
        else if (dc1 == null)
            rt = dcCmpTextOf2dc(dc0, new Dictionary());
        else
        {
            rt = new Dictionary();

            {
                // and matching from wb1
                foreach (var ky in dc0.Keys)
                {
                    // tx0 = CStr(.get_Item(ky))

                    // qi = New Scripting.Dictionary
                    // rt.Add ky, qi

                    rt.Add(ky,
                        dc1.Exists(ky)
                            ? dcCmpTextOf2items(nm0, nm1, dc0.get_Item(ky), dc1.get_Item(ky))
                            // qi.Add nm0, tx0
                            : dcCmpTextOf2items(nm0, nm1, dc0.get_Item(ky), null));
                }
            }

            {
                foreach (var ky in dc1.Keys)
                {
                    if (rt.Exists(ky))
                    {
                    }
                    else
                        // so add it now

                        // tx1 = CStr(.get_Item(ky))
                        // 
                        // qi = New Scripting.Dictionary
                        // qi.Add nm1, tx1
                        // rt.Add ky, qi
                        rt.Add(ky, dcCmpTextOf2items(nm0, nm1, null, dc1.get_Item(ky)));
                }
            }
        }

        return rt;
    }

    public static  Dictionary dcCmpTextOf2subDc(Dictionary dc, string k0, string k1)
    {
        var rt = dcCmpTextOf2dc(dcInDc(k0, dc), dcInDc(k1, dc));
        Debug.Print("");
        return rt;
    }

    public static  Dictionary dcWBQbyCmpResult(Dictionary dc)
    {
        var rt = new Dictionary();

        {
            foreach (var ky in dc.Keys)
            {
                Dictionary ck = dc.get_Item(ky);

                dynamic gk;
                switch (ck.Count)
                {
                    case > 1:
                        gk = "!=";
                        break;
                    case < 1:
                        Debugger.Break(); // because SOMETHING went wrong
                        gk = "**";
                        break;
                    default:
                        gk = ck.Keys(0);
                        break;
                }

                {
                    Dictionary gp;
                    if (rt.Exists(gk))
                        gp = rt.get_Item(gk);
                    else
                    {
                        gp = new Dictionary();
                        rt.Add(gk, gp);
                    }

                    gp.Add(ky, ck);
                }
            }
        }

        return rt;
    }
}