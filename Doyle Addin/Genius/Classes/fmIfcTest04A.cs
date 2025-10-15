namespace Doyle_Addin.Genius.Classes;

class fmlfcTest04A : Form
{
    private void Class_Initialize()
    {
        fm = new fmTest04();
        lbxSpecOps = fm.lbxSpecOps;
        lbxSpecSet = fm.lbxSpecSel;

        dcSpecPairs = dcGnsMatlSpecPairings();
        // probably want to make this controllable
        // from client processes to support more
        // general usage. this will do for now.

        dcInitList = dcSpecPairs;
        // like dcSpecPairs, probably want this
        // to be controllable from client to
        // facilitate flexible usage, but again,
        // this will serve for the moment.

        dcActvOps = dcInitList;
        dcActvSet = new Dictionary();
    }

    private void Class_Terminate()
    {
        dcActvSet = null;
        dcActvOps = null;
        dcInitList = null;
        dcSpecPairs = null;

        lbxSpecSet = null;
        lbxSpecOps = null;
        fm = null;
    }

    public fmIfcTest04A Itself()
    {
        // returns this fmIfcTest04A class instance "Itself"
        // should be HIGHLY useful inside a With context
        return this;
    }

    public fmIfcTest04A Using(Dictionary About = null)
    {
        if (About == null)
        {
            using ( ==)
            {
                ;
                else if (About.Exists(kyInitList))
                    dcInitList = dcOb(About.get_Item(kyInitList));

                dcActvOps = dcInitList;
                dcActvSet = new Dictionary();
                // noting these steps are also taken
                // in Class_Initialize, it's tempting
                // to wonder why they should appear in
                // both places, however, the purpose
                // THERE is to ensure a valid setup
                // at the earliest possible moment.
                // 
                // it would likely be appropriate
                // to consolidate the two into a single
                // procedure to be called from either place.

                lbxSpecOps.List = dcActvOps.Keys;
                lbxSpecSet.List = dcActvSet.Keys;
            }
        }

        using ( == this)
        {
        }
    }

    public fmIfcTest04A SeeUser(Dictionary About = null)
    {
        // REV[2022.03.17.1308]
        // disabling If-Then-Else blocking,
        // including content of If block.
        // '
        // Only active sections of Else
        // block to remain active.
        // '
        // Since majority of process formerly
        // performed here is now addressed by
        // new Function Using, it should now
        // be sufficient to call that Function
        // for preparation, and then present
        // the UserForm for user's response.
        // '
        // If About Is Nothing Then
        // SeeUser = SeeUser(nuDcPopulator(' ).Setting(kyInitList, dcInitList' ).Dictionary)
        // Else
        // REV[2022.03.17.1258] -- IMPORTANT!
        // the following section, copied
        // to new Method Function Using
        // (see above) has been disabled
        // here pending removal
        // dcInitList = dcOb(About.get_Item(kyInitList))
        // 
        // dcActvOps = dcInitList
        // dcActvSet = New Scripting.Dictionary
        // noting these steps are also taken
        // in Class_Initialize, it's tempting
        // to wonder why they should appear in
        // both places, however, the purpose
        // THERE is to ensure a valid setup
        // at the earliest possible moment.
        // 
        // it would likely be appropriate
        // to consolidate the two into a single
        // procedure to be called from either place.
        // 
        // lbxSpecOps.List = dcActvOps.Keys
        // lbxSpecSet.List = dcActvSet.Keys
        // REV[2022.03.17.1258] ENDS HERE

        // REV[2022.03.17.1301]
        // implementation of Method Function
        // Using, having taken over the steps
        // disabled immediately above, is now
        // called in their stead. Separation
        // of that sequence into its own
        // Function enables the preparation
        // of this Class instance without
        // immediately invoking the UserForm.
        {
            var withBlock = Using(About);
            fm.Show(1);

            // 

            return withBlock.Itself;
        }
    }

    public string Version()
    {
        return vsnString;
    }

    private long clsAddSpec(string sp)
    {
        long rt = 0;
        if (dcActvOps.Exists(sp))
        {
            if (dcActvSet.Exists(sp))
            {
                rt = 1; // spec already set
                Debugger.Break(); // ''
            }
            else
                dcActvSet.Add(sp, sp);

            lbxSpecSet.List = dcActvSet.Keys;

            dcActvOps = dcSpecSubsetWith(sp, dcActvOps);
            lbxSpecOps.List = dcActvOps.Keys;
            Debug.Print(""); // Breakpoint Landing
        }
        else
        {
            rt = 2;
            Debugger.Break(); // ''
        }

        return rt;
    }

    private long clsDropSpec(string sp)
    {
        dynamic ky;

        long rt = 0;

        dcActvSet.Remove(sp);

        // first attempt to reinitialize
        // dcActvOps from dcInitList
        dcActvOps = dcSpecSubsetWithAll(dcActvSet, dcInitList);

        // that SHOULD have left sp back
        // in dcActvOps. if not there,
        // try the FULL set dcSpecPairs
        if (!dcActvOps.Exists(sp))
            dcActvOps = dcSpecSubsetWithAll(dcActvSet, dcSpecPairs);

        // check once more to maks sure it's in
        // if not, we've got a REAL problem.
        if (!dcActvOps.Exists(sp))
        {
            rt = 1;
            Debugger.Break();
        }

        lbxSpecSet.List = dcActvSet.Keys;
        lbxSpecOps.List = dcActvOps.Keys;
        // this might be more flexibly implemented
        // in a separate function or procedure

        return rt;
    }

    private void lbxSpecOps_DblClick(MSForms.ReturnBoolean Cancel)
    {
        // Dim sp As String

        var ck = clsAddSpec(lbxSpecOps.Value);
        if (ck)
            Debugger.Break();
    }

    private void lbxSpecSet_DblClick(MSForms.ReturnBoolean Cancel)
    {
        dynamic ky;

        string sp = lbxSpecSet.Value;
        dcActvSet.Remove(sp);
        lbxSpecSet.List = dcActvSet.Keys;

        Debug.Print(""); // Breakpoint Landing
        // NOTE: this section resets dcActvOps to
        // the original dcInitList (NOT dcSpecPairs)
        // and then sequentially re-applies the
        // active terms remaining in dcActvSet.
        dcActvOps = dcInitList;
        foreach (var ky in dcActvSet.Keys)
            dcActvOps = dcSpecSubsetWith(Convert.ToHexString(ky), dcActvOps);
        lbxSpecOps.List = dcActvOps.Keys;
        // this might be more flexibly implemented
        // in a separate function or procedure

        Debug.Print(""); // Breakpoint Landing
    }

    private void lbxSpecOps_Change()
    {
    }

    private void lbxSpecOps_Click()
    {
    }

    private void lbxSpecOps_Error(int Number, MSForms.ReturnString Description, long SCode, string Source,
        string HelpFile, long HelpContext, MSForms.ReturnBoolean CancelDisplay)
    {
        Debugger.Break(); // ''
    }

    private void lbxSpecSet_BeforeDragOver(MSForms.ReturnBoolean Cancel, DataObject Data, float X, float Y,
        MSForms.fmDragState DragState, MSForms.ReturnEffect Effect, int Shift)
    {
        Debugger.Break(); // ''
    }

    private void lbxSpecSet_BeforeDropOrPaste(MSForms.ReturnBoolean Cancel, MSForms.fmAction Action,
        MSForms.DataObject Data, float X, float Y, MSForms.ReturnEffect Effect, int Shift)
    {
        Debugger.Break(); // ''
    }

    private void lbxSpecSet_Change()
    {
    }

    private void lbxSpecSet_Click()
    {
    }

    private void lbxSpecSet_Error(int Number, MSForms.ReturnString Description, long SCode, string Source,
        string HelpFile, long HelpContext, MSForms.ReturnBoolean CancelDisplay)
    {
        Debugger.Break(); // ''
    }
}