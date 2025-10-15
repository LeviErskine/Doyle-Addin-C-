namespace Doyle_Addin.Genius.Classes;

public class libDict
{
    /// <summary>
    /// 
    /// </summary>
    /// <param name="Dict"></param>
    /// <returns></returns>
    public static Dictionary dcNewIfNone(Dictionary Dict)
    {
        return Dict ?? new Dictionary();
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="rs"></param>
    /// <returns></returns>
    public static Dictionary dcOfRsFields(ADODB.Recordset rs)
    {
        var rt = new Dictionary();
        {
            if (rs.State != adStateOpen) return rt;
            foreach (ADODB.Field fd in rs.Fields)
                rt.Add(fd.Name, fd);
        }
        return rt;
    }

    /// <summary>
    ///dcDotted -- return Dictionary with links to itself, under key ".", and under "..", either itself, or, if supplied, an optional "parent" Dictionary this mimics the traditional linkage within POSIX-compliant and other file systems, where the "." and ".." names in each directory are assigned to itself and its parent, respecrively !!WARNING!! this self- and back-linkage WILL cause endless loops in Dictionary traversal routines not prepared to deal with them! Be sure to review any procedure BEFORE calling against a Dictionary using this linkage!
    /// </summary>
    /// <param name="Under"></param>
    /// <param name="Using"></param>
    /// <returns></returns>
    public static Dictionary dcDotted(Dictionary Under = null, Dictionary Using = null)
    {
        // 
        Dictionary rt;

        return dcNewIfNone(Using);

        if (rt.Exists("."))
        {
            if (rt.get_Item(".") == rt)
            {
            }
            else
                Debugger.Break();
        }
        else
            rt.Add(".", rt);

        if (rt.Exists(".."))
        {
            if (rt.get_Item("..") is Dictionary)

                if (Using is null) ;
                else
                    stop
                        .get_Item("..") = Using;

            else
                Debugger.Break();
        }
        else
            rt.Add("..", IIf(Under == null, rt, Under));

        return rt;
    }

    /// <summary>
    /// dcUnDotted -- remove Keys "." and ".." from supplied Dictionary dc no checks are made of the Items under these Keys. the Dictionary is assumed to have originated from or passed through a prior call to dcDotted, and thus include self- and back-linkage thereunder. (a check system was considered and attempted, but deemed too unweildy, and so abandoned) 
    /// </summary>
    /// <param name="dc"></param>
    /// <returns></returns>
    public static Dictionary dcUnDotted(Dictionary dc)
    {
        {
            if (dc.Exists("."))
                dc.Remove(".");
            if (dc.Exists(".."))
                dc.Remove("..");
        }
        return dc;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="rs"></param>
    /// <param name="fnKey"></param>
    /// <param name="fnVal"></param>
    /// <param name="flt"></param>
    /// <returns></returns>
    public static Dictionary dcFrom2Fields(ADODB.Recordset rs, string fnKey, string fnVal, string flt = "")
    {
        var rt = new Dictionary();
        {
            ADODB.Field fdKey;
            ADODB.Field fdVal;
            {
                var withBlock1 = rs.Fields;
                fdKey = withBlock1.get_Item(fnKey);
                fdVal = withBlock1.get_Item(fnVal);
            }

            rs.Filter = flt;
            while (!rs.BOF | rs.EOF)
            {
                {
                    if (rt.Exists(fdKey.Value))
                        Debugger.Break();
                    else
                        rt.Add(fdKey.Value, fdVal.Value);
                }
                rs.MoveNext();
            }
        }
        return rt;
    }

    /// <summary>
    /// , fnKey As String, fnVal As String
    /// , Optional ovr As Long = -1
    /// dcFromAdoRS - return a Dictionary
    /// of tuples (rows) from an ADODB
    /// Recordset, keyed on order of
    /// encounter and processing.
    /// NOTE that this Dictionary is NOT
    /// keyed on any particular Field.
    /// The wide range of situations which
    /// might be encountered suggests that
    /// indexing and keying on field values
    /// is best addressed in a separate,
    /// dedicated process.
    /// Dim fdKey As ADODB.Field 
    /// </summary>
    /// <param name="rs"></param>
    /// <param name="flt"></param>
    /// <returns></returns>
    public static Dictionary dcFromAdoRS(ADODB.Recordset rs, string flt = "")
    {
        var rt = new Dictionary();
        {
            // With .Fields
            // fdKey = .get_Item(fnKey)
            // End With

            rs.Filter = flt;
            while (!rs.BOF | rs.EOF)
            {
                dynamic ky = rt.Count; // fdKey.Value
                Dictionary tp;
                {
                    // If .Exists(ky) Then 'we have a collision!
                    // Stop 'and figure out what to do!
                    // Else
                    // .Add ky, dcFromAdoRSrow(rs)
                    rt.Add(ky, new Dictionary());
                    // End If
                    tp = dcOb(rt.get_Item(ky));
                }

                foreach (ADODB.Field fdVal in rs.Fields)
                {
                    dynamic vl;
                    string nm;
                    {
                        nm = fdVal.Name;
                        vl = fdVal.Value;
                    }

                    {
                        // If .Exists(nm) Then
                        // If ovr Then 'change if needed
                        // If .get_Item(nm) <> vl Then
                        // .get_Item(nm) = vl
                        // End If
                        // Else 'fuhgeddaboudit!
                        // End If
                        // Else
                        tp.Add(nm, vl);
                    }
                }

                rs.MoveNext();
            }
        }
        return rt;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="rs"></param>
    /// <param name="nullVal"></param>
    /// <returns></returns>
    public static Dictionary dcFromAdoRSrow(ADODB.Recordset rs, dynamic nullVal = null)
    {
        var rt = new Dictionary();
        {
            var ck = rs.BOF | rs.EOF;
            foreach (ADODB.Field fd in rs.Fields)
            {
                {
                    var nm = fd.Name;
                    rt.Add(nm, ck ? nullVal : fd.Value);
                }
            }
        }
        return rt;
    }

    /// <summary>
    /// dcDxFromRecSetDc -- Generate Dictionary of Indices from "RecordSet" Dictionary as returned by dcFromAdoRS
    /// </summary>
    /// <param name="Dict"></param>
    /// <returns></returns>
    public static Dictionary dcDxFromRecSetDc(Dictionary Dict)
    {
        dynamic k2;
        // '

        // rt = New Scripting.Dictionary
        var dcDx = new Dictionary();

        // ' the Dictionary of Indices
        {
            // ' Start scanning primary Keys
            // ' to begin overall process
            foreach (var k0 in Dict.Keys)
            {
                // ' Retrieve "record" Dictionary
                // ' for next/current Key
                Dictionary tp = dcOb(Dict.get_Item(k0));
                if (tp == null)
                    // Stop
                    Debug.Print(""); // Breakpoint Landing
                else
                {
                    // ' Scan "field" Keys of current "record"
                    // ' to identify and populate Index Dictionaries
                    foreach (var k1 in tp.Keys)
                    {
                        // ' Retrieve "index" Dictionary for current
                        // ' "field". Generate new one, if not present.
                        // '
                        // ' (might want to support Key filtering
                        // ' to either exclude some "fields",
                        // ' or limit indexing to a list)
                        Dictionary dcVl;
                        {
                            if (dcDx.Exists(k1))
                            {
                            }
                            else
                                dcDx.Add(k1, new Dictionary());

                            dcVl = dcOb(dcDx.get_Item(k1));
                        }

                        // ' Retrieve current "field" value, and return its
                        // ' Dictionary from the "field index" Dictionary.
                        // '
                        // ' Again, generate a new one, if needed.
                        var vl = tp.get_Item(k1);
                        Dictionary dcTp;
                        {
                            if (dcVl.Exists(vl))
                            {
                            }
                            else
                                dcVl.Add(vl, new Dictionary());

                            dcTp = dcOb(dcVl.get_Item(vl));
                        }

                        // ' Add the current "record" to the recovered
                        // ' "field value" Dictionary. This SHOULD only
                        // ' add a link to the same "record" Dictionary,
                        // ' rather than duplicate the whole thing.
                        // '
                        // ' However, converting to JSON generates
                        // ' a new dump of the Dictionary structure
                        // ' wherever it appears, thus replicating it
                        // ' multiple times in the output.
                        // '
                        {
                            if (dcTp.Exists(k0))
                                Debugger.Break(); // for now. might still be okay
                            else
                                dcTp.Add(k0, tp);
                        }
                    }
                }
            }
        }

        {
            if (dcDx.Exists(""))
                Debugger.Break(); // because we have
            else
                dcDx.Add("", Dict);
        }

        return dcDx;
    }

    /// <summary>
    /// dcRecSetDcDx4json -- Prep RecordSet Index Dictionary for JSON export. Replaces each field/value index Dictionary with its Keys for export to JSON, to avoid replicating each original "record" Item in its entirety wherever it's referenced in the indices.
    /// </summary>
    /// <param name="Dict"></param>
    /// <returns></returns>
    public static Dictionary dcRecSetDcDx4json(Dictionary Dict)
    {
        // 
        // '

        var rt = new Dictionary();

        // ' the Dictionary of Indices
        {
            // ' Start scanning field
            // ' names (top level Keys)
            // ' to begin transformation
            foreach (var k0 in Dict.Keys)
            {
                // ' Retrieve next "field index" Dictionary
                Dictionary dcFdIn = dcOb(Dict.get_Item(k0));

                // ' Check for original RecordSet Dictionary
                if (k0 == "")
                    rt.Add("", dcFdIn);
                else
                {
                    // ' Generate corresponding "field index"
                    // ' output Dictionary
                    Dictionary dcFdOut;
                    {
                        if (rt.Exists(k0))
                            Debugger.Break(); // because it should NOT
                        else
                            rt.Add(k0, new Dictionary());

                        dcFdOut = dcOb(rt.get_Item(k0));
                    }

                    // ' Scan value Keys of current "field"
                    // ' to retrieve index Dictionaries
                    {
                        foreach (var vl in dcFdIn.Keys)
                        {
                            // ' Retrieve Dictionary for current value
                            Dictionary dcVl = dcOb(dcFdIn.get_Item(vl));

                            {
                                if (dcFdOut.Exists(vl))
                                    Debugger.Break(); // because it should
                                else
                                    // ' Dump record Keys to output
                                    // ' field value index Dictionary
                                    dcFdOut.Add(vl, dcVl.Keys);
                            }
                        }
                    }
                }
            }
        }

        return rt;
    }

    /// <summary>
    /// dcOfSubDict -- intended to return a "flat" Dictionary containing the supplied Dictionary and all Dictionary objects within it. DO NOT ATTEMPT TO USE AT THIS TIME!!! Need to work out a way to tell if the supplied Dictionary is already in the returned
    /// </summary>
    /// <param name="dc"></param>
    /// <param name="rt"></param>
    /// <returns></returns>
    public static Dictionary dcOfSubDict(Dictionary dc, Dictionary rt = null)
    {
        while (true)
        {
            // 
            dynamic ky;

            if (rt == null)
            {
                rt = new Dictionary();
                continue;
            }

            if (dc == null) return rt;
            break;
        }
    }
}