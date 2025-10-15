namespace Doyle_Addin.Genius.Classes;

public class dcSplitter
{
    // Private dcGrpIn As Scripting.Dictionary
    // Private dcGrpOut As Scripting.Dictionary
    private kyPick dcPicker;

    // Sample Usage:
    // Debug.Print txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).OutGroup().Keys)
    // Debug.Print txDumpLs(nuSplitter().WithSel(New kyPickAiPartVsAssy).Scanning(dcAiDocComponents(aiDocActive())).WithSel(New kyPickAiDocWithRM, 1).SubScanning().OutGroup().Keys)

    private void Class_Initialize()
    {
        // dcGrpIn = New Scripting.Dictionary
        // dcGrpOut = New Scripting.Dictionary
        dcPicker = new kyPick();
    }

    public dcSplitter WithInDc(Dictionary dict)
    {
        dcPicker = dcPicker.WithInDc(dict);
        return this;
    }

    public dcSplitter WithOutDc(Dictionary dict)
    {
        dcPicker = dcPicker.WithOutDc(dict);
        return this;
    }

    public dcSplitter WithSel(kyPick selector, bool keepData = false)
    {
        dcPicker = !keepData ? selector : selector.WithInDc(dcPicker.dcIn).WithOutDc(dcPicker.dcOut);
        return this;
    }

    public Dictionary InGroup()
    {
        return dcPicker.dcIn;
    }

    public Dictionary OutGroup()
    {
        return dcPicker.dcOut;
    }

    public dcSplitter Scanning(Dictionary srcDict)
    {
        {
            foreach (var ky in srcDict.Keys)
            {
                {
                    var withBlock1 = dcPicker.dcFor(srcDict.get_Item(ky));
                    if (withBlock1.Exists(ky))
                        Debugger.Break();
                    else
                        withBlock1.Add(ky, srcDict.get_Item(ky));
                }
            }
        }
        return this;
    }

    public dcSplitter SubScanning(bool wantOut = false)
    {
        Dictionary dcSub;

        if (!wantOut)
            dcSub = dcPicker.dcIn;
        else
            dcSub = dcPicker.dcOut;
        dcPicker = dcPicker.WithInDc(new Dictionary()).WithOutDc(new Dictionary());

        return Scanning(dcSub);
    }
}