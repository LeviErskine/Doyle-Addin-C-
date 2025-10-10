class libDcGeneral
{
    public Scripting.Dictionary dcCopy(Scripting.Dictionary dc)
    {
        /// dcCopy -- return a new Dictionary
        /// copying the contents of the
        /// one supplied, including COPIES
        /// of any Dictionary Objects within.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary ck;
        Variant bx;
        Variant ky;

        rt = new Scripting.Dictionary();

        if (dc == null)
        {
        }
        else
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                bx = Array(withBlock.Item(ky));
                ck = dcOb(obOf(bx(0)));

                if (ck == null)
                    rt.Add(ky, bx(0));
                else
                    rt.Add(ky, dcCopy(ck));
            }
        }

        dcCopy = rt;
    }

    public Scripting.Dictionary dcWith(Variant ky, Variant it, Scripting.Dictionary dc = null/* TODO Change to default(_) if this is not a reference type */)
    {
        if (dc == null)
            dcWith = dcWith(ky, it, new Scripting.Dictionary());
        else
        {
            {
                var withBlock = dc;
                if (withBlock.Exists(ky))
                    withBlock.Remove(ky);
                withBlock.Add(ky, it);
            }
            dcWith = dc;
        }
    }

    public dcPopulator nuDcPopulator(Scripting.Dictionary Dict = null/* TODO Change to default(_) if this is not a reference type */, long Opts = 0)
    {
        {
            var withBlock = new dcPopulator();
            nuDcPopulator = withBlock.Using(Dict, Opts);
        }
    }

    public Variant dcItemIfPresent(Scripting.Dictionary dc, Variant ky, VbVarType vtMissing = )
    {
        if (dc == null)
            dcItemIfPresent = noVal(vtMissing);
        else
        {
            var withBlock = dc;
            if (withBlock.Exists(ky))
            {
                if (IsObject(withBlock.Item(ky)))
                    dcItemIfPresent = withBlock.Item(ky);
                else
                    dcItemIfPresent = withBlock.Item(ky);
            }
            else if (vtMissing == Constants.vbObject)
                dcItemIfPresent = noVal(vtMissing);
            else
                dcItemIfPresent = noVal(vtMissing);
        }
    }

    public Scripting.Dictionary dcInDc(string dcKey, Scripting.Dictionary inDc)
    {
        dcInDc = dcOb(obOf(dcItemIfPresent(inDc, dcKey, Constants.vbObject)));
    }

    public Scripting.Dictionary dcInDcMk(Variant ky, Scripting.Dictionary dc)
    {
        /// dcInDcMk --
        /// 
        Scripting.Dictionary rt;

        {
            var withBlock = dc;
            if (withBlock.Exists(ky))
                rt = dcOb(withBlock.Item(ky));
            else
            {
                rt = new Scripting.Dictionary();
                withBlock.Add(ky, rt);
            }
        }

        dcInDcMk = rt;
    }

    public Scripting.Dictionary dcOfKeys2match(Variant ls)
    {
        /// dcOfKeys2match -- generate a Dictionary
        /// mapping a supplied Key, or Array
        /// of Keys to itself or themselves
        /// '
        /// primary purpose is to provide
        /// a 'reference' Dictionary of Keys
        /// to be sought in other Dictionaries
        /// '
        /// (formerly d4g4f2)
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        if (IsArray(ls))
        {
            foreach (var ky in ls)
                rt.Add(ky, ky);
        }
        else
            rt.Add(ls, ls);

        dcOfKeys2match = rt;
    }

    public Scripting.Dictionary dcKeysInCommon(Scripting.Dictionary d0, Scripting.Dictionary d1, long pk = 0)
    {
        /// dcKeysInCommon -- return intersection
        /// of two Dictionary Objects based on
        /// matching keys. Use optional pk value
        /// to select which Dictionary's Items
        /// to return in result:
        /// 
        /// 0 returns an array of Items from both
        /// this is the default
        /// 1 returns only Items from the first
        /// 2 returns only Items from the second
        /// 
        /// NOTE that if either Dictionary Object
        /// is not supplied (is Nothing), then
        /// an empty Dictionary is returned,
        /// just as if an empty Dictionary
        /// had been supplied.
        /// 
        Scripting.Dictionary rt;
        Variant ky;
        Variant ls;

        if (d0 == null)
            rt = dcKeysInCommon(new Scripting.Dictionary(), d1);
        else if (d1 == null)
            rt = dcKeysInCommon(d0, new Scripting.Dictionary());
        else
        {
            rt = new Scripting.Dictionary();
            {
                var withBlock = d0;
                foreach (var ky in withBlock.Keys)
                {
                    if (d1.Exists(ky))
                    {
                        ls = Array(withBlock.Item(ky), d1.Item(ky));
                        rt.Add(ky, Array(ls, ls(0), ls(1))(pk));
                    }
                }
            }
        }
        dcKeysInCommon = rt;
    }

    public Scripting.Dictionary dcKeysMissing(Scripting.Dictionary dcWith, Scripting.Dictionary dcWout)
    {
        /// dcKeysMissing -- return difference
        /// of first Dictionary Object minus
        /// those keys found in the second.
        /// 
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dcWith;
            foreach (var ky in withBlock.Keys)
            {
                if (dcWout.Exists(ky))
                {
                }
                else
                    rt.Add(ky, withBlock.Item(ky));
            }
        }
        dcKeysMissing = rt;
    }

    public Scripting.Dictionary dcKeysCombined(Scripting.Dictionary d0, Scripting.Dictionary d1, long pk = 0)
    {
        /// dcKeysCombined -- return union
        /// of two Dictionary Objects. For
        /// keys in both, use optional pk
        /// value to select which Dictionary's
        /// Items to return in result:
        /// 
        /// 0 returns an array of Items from both
        /// this is the default
        /// 1 returns only Items from the first
        /// 2 returns only Items from the second
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary ob;
        Variant ky;
        Variant ls;

        if (pk > 1)
            rt = dcKeysCombined(d1, d0, 1);
        else
        {
            rt = dcKeysInCommon(d0, d1, pk);

            foreach (var ls in Array(d0, d1))
            {
                ob = ls;
                {
                    var withBlock = dcKeysMissing(ob, rt) // d0
           ;
                    foreach (var ky in withBlock.Keys)
                        rt.Add(ky, withBlock.Item(ky));
                }
            }
        }

        // rt = dcKeysInCommon()

        if (d0 == null)
            System.Diagnostics.Debugger.Break();
        else if (d1 == null)
            System.Diagnostics.Debugger.Break();
        else
        /// NOTE[2023.04.10.1449]
        /// need to review what's going on here
        /// this LOOKS like an effor to include items
        /// left out of earlier operation, however,
        /// it's not clear this stage isn't redundant.
        {
            var withBlock = d0;
            foreach (var ky in withBlock.Keys)
            {
                if (d1.Exists(ky))
                {
                    ls = Array(withBlock.Item(ky), d1.Item(ky));
                    if (rt.Exists(ky))
                    {
                        if (ConvertToJson(Array(ls, ls(0), ls(1))(pk)) != ConvertToJson(rt.Item(ky)))
                            System.Diagnostics.Debugger.Break();
                    }
                    else
                        rt.Add(ky, Array(ls, ls(0), ls(1))(pk));
                }
            }
        }
        dcKeysCombined = rt;
    }

    public Scripting.Dictionary dcOfIdent(Variant src)
    {
        Scripting.Dictionary rt;
        Variant ky;

        rt = new Scripting.Dictionary();

        if (IsArray(src))
        {
            foreach (var ky in src)
                rt.Add(ky, ky);
        }
        else
            rt.Add(src, src);

        dcOfIdent = rt;
    }

    public Scripting.Dictionary dcTransposed(Scripting.Dictionary dc)
    {
        /// Transpose Key values of supplied
        /// Dictionary with matching Item values.
        /// 
        /// As written, will ONLY work against
        /// a Dictionary whose Item values,
        /// like its Keys, are unique.
        /// 
        Scripting.Dictionary rt;
        Variant fm;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var fm in withBlock.Keys)
            {
                if (rt.Exists(withBlock.Item(fm)))
                    System.Diagnostics.Debugger.Break();
                else
                    rt.Add.Item(fm);/* TODO ERROR: Skipped SkippedTokensTrivia *//* TODO ERROR: Skipped SkippedTokensTrivia */
            }
        }
        dcTransposed = rt;
    }

    public Scripting.Dictionary dcTransGrouped(Scripting.Dictionary dc)
    {
        /// dcTransGrouped
        /// (derived from dcTransposed)
        /// 
        /// generate new Dictionary "tramsposing"
        /// Key values with matching Item values
        /// in supplied Dictionary.
        /// 
        /// because more than one Key might map to
        /// the same Item, the returned Dictionary
        /// maps each (Item) Key to a Dictionary of
        /// the Keys which originally mapped to it.
        /// 
        /// each of these Dictionaries in turn
        /// maps the original Key back to its
        /// corresponding Item once more,
        /// since, HEY, it might as well!
        /// 
        /// obviously, an effor to work around
        /// the main limitation of the original
        /// dcTransposed, which can only work
        /// against a Dictionary whose Items,
        /// like its Keys, are unique.
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary it;
        Variant ky;
        Variant ar;

        rt = new Scripting.Dictionary();
        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ar = Array(withBlock.Item(ky));
                if (!rt.Exists(ar(0)))
                    rt.Add(ar(0), new Scripting.Dictionary());
                dcOb(rt.Item(ar(0))).Add(ky, ar(0));
            }
        }
        dcTransGrouped = rt;
    }

    public long dcDepth(Scripting.Dictionary dc)
    {
        /// this function extracts the "depth"
        /// of the supplied Dictionary object,
        /// that is, how many "levels" of
        /// Dictionary objects it contains,
        /// counting the supplied Dictionary
        /// itself. When an actual Dictionary
        /// is supplied, the value returned
        /// will be at least 1. It will only
        /// be zero when Nothing is supplied.
        /// 
        long rt;
        long ck;
        Variant ky;

        if (dc == null)
            dcDepth = 0;
        else
        {
            rt = 0;

            {
                var withBlock = dc;
                foreach (var ky in withBlock.Keys)
                {
                    ck = dcDepth(dcOb(obOf(withBlock.Item(ky))));
                    if (ck > rt)
                        rt = ck;
                }
            }

            dcDepth = 1 + rt;
        }
    }

    public Scripting.Dictionary dcFlattenDown(Scripting.Dictionary Dict, long DownTo = 1)
    {
        /// this function partially "flattens"
        /// a hierarchy of Dictionary objects
        /// (a Dictionary of Dictionaries,
        /// potentially of more Dictionaries)
        /// starting from the top, working
        /// down to 'DownTo' levels
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary sd;
        Variant ky;
        Variant it;
        Variant sk;

        if (Dict == null)
            dcFlattenDown = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = Dict;
                foreach (var ky in withBlock.Keys)
                {
                    it = Array(withBlock.Item(ky));

                    if (DownTo > 0)
                        sd = dcFlattenDown(dcOb(obOf(it(0))), DownTo - 1);
                    else
                        sd = null/* TODO Change to default(_) if this is not a reference type */;

                    if (sd == null)
                        rt.Add(ky, it(0));
                    else
                    {
                        var withBlock1 = sd;
                        foreach (var sk in withBlock1.Keys)
                            rt.Add(ky + "." + sk, withBlock1.Item(sk));
                    }
                }
            }

            dcFlattenDown = rt;
        }
    }

    public Scripting.Dictionary dcFlattenUp(Scripting.Dictionary Dict, long DownFrom = 0)
    {
        /// this function partially "flattens"
        /// a hierarchy of Dictionary objects
        /// (a Dictionary of Dictionaries,
        /// potentially of more Dictionaries)
        /// starting BELOW the top, skipping
        /// DownFrom levels before "flattening"
        /// Dictionaries below and at that level
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary sd;
        Variant ky;
        Variant it;
        Variant sk;

        if (Dict == null)
            dcFlattenUp = null/* TODO Change to default(_) if this is not a reference type */;
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = Dict;
                foreach (var ky in withBlock.Keys)
                {
                    it = Array(withBlock.Item(ky));
                    sd = dcOb(obOf(it(0))); // dcFlattenUp(, DownFrom - 1)

                    if (sd == null)
                        rt.Add(ky, it(0));
                    else if (DownFrom > 0)
                        rt.Add(ky, dcFlattenUp(sd, DownFrom - 1));
                    else
                    {
                        var withBlock1 = dcFlattenUp(sd, 0);
                        foreach (var sk in withBlock1.Keys)
                            rt.Add(ky + "." + sk, withBlock1.Item(sk));
                    }
                }
            }

            dcFlattenUp = rt;
        }
    }

    public Scripting.Dictionary dcOfDcRekeyedSecToPri(Scripting.Dictionary dc)
    {
        /// dcOfDcRekeyedSecToPri - take Dictionary of Dictionaries
        /// and return Dictionary of Dictionaries
        /// with Secondary Keys promoted to Primary,
        /// and Primary demoted to Secondary
        /// 
        Scripting.Dictionary rt;
        Scripting.Dictionary sb;
        Variant kp;
        Variant ks;
        Variant ar;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var kp in withBlock.Keys)
            {
                sb = dcOb(withBlock.Item(kp));
                if (sb == null)
                    // Stop
                    Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */ // Breakpoint Landing
                else
                {
                    var withBlock1 = sb;
                    foreach (var ks in withBlock1.Keys)
                    {
                        ar = Array(withBlock1.Item(ks));
                        {
                            var withBlock2 = rt;
                            if (!withBlock2.Exists(ks))
                                withBlock2.Add(ks, new Scripting.Dictionary());

                            {
                                var withBlock3 = dcOb(withBlock2.Item(ks));
                                if (withBlock3.Exists(kp))
                                    System.Diagnostics.Debugger.Break();
                                else
                                    withBlock3.Add(kp, ar(0));
                            }
                        }
                    }
                }
            }
        }

        dcOfDcRekeyedSecToPri = rt;
    }

    public Scripting.Dictionary dcCmpTextOf2items(string id0, string id1, Variant it0, Variant it1)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary dc0;
        Scripting.Dictionary dc1;
        object ob0;
        object ob1;
        string tx0;
        string tx1;
        long ck;
        long c2;

        ck = IIf(IsObject(it0), 1, 0) + IIf(IsObject(it1), 2, 0) + IIf(IsEmpty(it0), 4, 0) + IIf(IsEmpty(it1), 8, 0);
        if (ck & 1)
            ck = ck + Interaction.IIf(it0 == null, 4, 0);
        if (ck & 2)
            ck = ck + Interaction.IIf(it1 == null, 8, 0);
        /// UPDATE[2021.12.09]:
        /// Added trap for Nothing Objects
        /// to treat them as Empty as well

        c2 = ck & 12; // $11XX
        if (c2)
        {
            // is Empty/Nothing
            {
                var withBlock = nuDcPopulator();
                if (c2 == 12)
                    // are equally Empty
                    rt = withBlock.Setting("==", "").Dictionary();
                else if (c2 == 8)
                    // and id0 needs processed
                    // If ck And 1 Then
                    // rt = dcCmpTextOf2dc('        dcOb(it0), .Dictionary()'    )
                    // Else
                    rt = withBlock.Setting(id0, it0).Dictionary();
                else
                    // and id1 needs processed
                    // If ck And 2 Then
                    // rt = dcCmpTextOf2dc('        .Dictionary(), dcOb(it1)'    )
                    // Else
                    rt = withBlock.Setting(id1, it1).Dictionary();
            }
        }
        else if (ck == 3)
            // of comparable Dictionaries.
            // compare these recursively
            rt = dcCmpTextOf2dc(dcOb(it0), dcOb(it1));
        else if (ck == 0)
        {
            // just a couple of String
            // or (hopefully) String-
            // compatible values.
            // compare them directly.
            tx0 = System.Convert.ToHexString(it0);
            tx1 = System.Convert.ToHexString(it1);

            // rt = New Scripting.Dictionary
            // With rt
            {
                var withBlock = nuDcPopulator();
                if (tx0 == tx1)
                    rt = withBlock.Setting("==", tx0).Dictionary();
                else
                    rt = withBlock.Setting(id0, tx0).Setting(id1, tx1).Dictionary();
            }
        }
        else
        // can't compare them in any way
        // just add each on its own side
        {
            var withBlock = nuDcPopulator();
            rt = withBlock.Setting(id0, it0).Setting(id1, it1).Dictionary();
        }

        dcCmpTextOf2items = rt;
    }

    public Scripting.Dictionary dcCmpTextOf2dc(Scripting.Dictionary dc0, Scripting.Dictionary dc1)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary qi;
        string nm0;
        string nm1;
        string tx0;
        string tx1;
        Variant ky;
        string nm;

        nm0 = "src0";
        nm1 = "src1";

        if (dc0 == null)
            rt = dcCmpTextOf2dc(new Scripting.Dictionary(), dc1);
        else if (dc1 == null)
            rt = dcCmpTextOf2dc(dc0, new Scripting.Dictionary());
        else
        {
            rt = new Scripting.Dictionary();

            {
                var withBlock = dc0 // add all from wb0
       ;
                // and matching from wb1
                foreach (var ky in withBlock.Keys)
                {
                    // tx0 = CStr(.Item(ky))

                    // qi = New Scripting.Dictionary
                    // rt.Add ky, qi

                    if (dc1.Exists(ky))
                        rt.Add(ky, dcCmpTextOf2items(nm0, nm1, withBlock.Item(ky), dc1.Item(ky)));
                    else
                        // qi.Add nm0, tx0
                        rt.Add(ky, dcCmpTextOf2items(nm0, nm1, withBlock.Item(ky), Empty));
                }
            }

            {
                var withBlock = dc1 // add any missed from wb1
       ;
                foreach (var ky in withBlock.Keys)
                {
                    if (rt.Exists(ky))
                    {
                    }
                    else
                        // so add it now

                        // tx1 = CStr(.Item(ky))
                        // 
                        // qi = New Scripting.Dictionary
                        // qi.Add nm1, tx1
                        // rt.Add ky, qi
                        rt.Add(ky, dcCmpTextOf2items(nm0, nm1, Empty, withBlock.Item(ky)));
                }
            }
        }

        dcCmpTextOf2dc = rt;
    }

    public Scripting.Dictionary dcCmpTextOf2subDc(Scripting.Dictionary dc, string k0, string k1)
    {
        Scripting.Dictionary rt;

        rt = dcCmpTextOf2dc(dcInDc(k0, dc), dcInDc(k1, dc));
        Debug.Print(); /* TODO ERROR: Skipped SkippedTokensTrivia */
        dcCmpTextOf2subDc = rt;
    }

    public Scripting.Dictionary dcWBQbyCmpResult(Scripting.Dictionary dc)
    {
        Scripting.Dictionary rt;
        Scripting.Dictionary ck;
        Scripting.Dictionary gp;
        Variant ky;
        Variant gk;

        rt = new Scripting.Dictionary();

        {
            var withBlock = dc;
            foreach (var ky in withBlock.Keys)
            {
                ck = withBlock.Item(ky);

                if (ck.Count > 1)
                    gk = "!=";
                else if (ck.Count < 1)
                {
                    System.Diagnostics.Debugger.Break(); // because SOMETHING went wrong
                    gk = "**";
                }
                else
                    gk = ck.Keys(0);

                {
                    var withBlock1 = rt;
                    if (withBlock1.Exists(gk))
                        gp = withBlock1.Item(gk);
                    else
                    {
                        gp = new Scripting.Dictionary();
                        rt.Add(gk, gp);
                    }

                    gp.Add(ky, ck);
                }
            }
        }

        dcWBQbyCmpResult = rt;
    }
}