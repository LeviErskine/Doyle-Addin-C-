class dcSplitter
{
    // Private dcGrpIn As Scripting.Dictionary
    // Private dcGrpOut As Scripting.Dictionary
    private kyPick dcPicker;

    /// Sample Usage:
    // Debug.Print txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).OutGroup().Keys)
    // Debug.Print txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).WithSel(New kyPickAiDocWithRM, 1).SubScanning().OutGroup().Keys)

    private void Class_Initialize()
    {
        // dcGrpIn = New Scripting.Dictionary
        // dcGrpOut = New Scripting.Dictionary
        dcPicker = new kyPick();
    }

    public dcSplitter WithInDc(Scripting.Dictionary Dict)
    {
        dcPicker = dcPicker.WithInDc(Dict);
        WithInDc = this;
    }

    public dcSplitter WithOutDc(Scripting.Dictionary Dict)
    {
        dcPicker = dcPicker.WithOutDc(Dict);
        WithOutDc = this;
    }

    public dcSplitter WithSel(kyPick Selector, long KeepData = 0)
    {
        if (KeepData == 0)
            dcPicker = Selector;
        else
            dcPicker = Selector.WithInDc(dcPicker.dcIn).WithOutDc(dcPicker.dcOut);
        WithSel = this;
    }

    public Scripting.Dictionary InGroup()
    {
        InGroup = dcPicker.dcIn;
    }

    public Scripting.Dictionary OutGroup()
    {
        OutGroup = dcPicker.dcOut;
    }

    public dcSplitter Scanning(Scripting.Dictionary SrcDict)
    {
        Variant ky;

        {
            var withBlock = SrcDict;
            foreach (var ky in withBlock.Keys)
            {
                {
                    var withBlock1 = dcPicker.dcFor(withBlock.Item(ky));
                    if (withBlock1.Exists(ky))
                        System.Diagnostics.Debugger.Break();
                    else
                        withBlock1.Add(ky, SrcDict.Item(ky));
                }
            }
        }
        Scanning = this;
    }

    public dcSplitter SubScanning(long WantOut = 0)
    {
        Scripting.Dictionary dcSub;

        if (WantOut == 0)
            dcSub = dcPicker.dcIn;
        else
            dcSub = dcPicker.dcOut;
        dcPicker = dcPicker.WithInDc(new Scripting.Dictionary()).WithOutDc(new Scripting.Dictionary());

        SubScanning = Scanning(dcSub);
    }
}