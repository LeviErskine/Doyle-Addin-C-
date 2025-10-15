namespace Doyle_Addin.Genius.Classes;

class wkgCls0
{
    private Dictionary dcWkg;
    private Dictionary dcFiled;
    private Dictionary dcIndex;

    private fmIfcTest05A fm;

    private void Class_Initialize()
    {
        // 
        dcWkg = new Dictionary();
        dcFiled = new Dictionary();
        dcIndex = new Dictionary();

        fm = new fmIfcTest05A();
    }

    private void Class_Terminate()
    {
        // 
        fm = null;

        dcWkg.RemoveAll();
        dcFiled.RemoveAll();
        dcIndex.RemoveAll();

        dcWkg = null;
        dcFiled = null;
        dcIndex = null;
    }

    public wkgCls0 Itself()
    {
        return this;
    }

    public wkgCls0 Using(Document AiDoc = null)
    {
        if
            (AiDoc == null)
            using ( ==
        this)
        else
        using ( == Collect(AiDoc));
    }

    public Dictionary Process(Document AiDoc = null)
    {
        Dictionary rt;

        if (AiDoc == null)
        {
            // THIS is where we start processing
            // Inventor Documents collected
            // in the internal Dictionary

            // at the moment, we simply pull up
            // the standard form and present it
            // with the current collection
            // of Inventor Documents
            {
                var withBlock = fm.Using(dcWkg);
                withBlock.Show(1);
                Debugger.Break();
            }

            rt = dcCopy(dcWkg);
        }
        else
            rt = Collect(AiDoc).Process();

        return rt;
    }

    public wkgCls0 Collect(Document AiDoc = null) // Scripting.Dictionary
    {
        // ,optional dcWkg as Scripting.Dictionary=nothing
        // Method Function Collect
        // 
        // given a valid Inventor Document
        // (usually an assembly), gather
        // any and all Parts in it into
        // the internal Dictionary dcWkg,
        // and return a copy.
        // 
        Dictionary rt;

        if (AiDoc == null)
            Collect(ThisApplication.ActiveDocument);
        else if (AiDoc == ThisDocument)
        {
        }
        else
            dcWkg = dcAiDocGrpsByForm(dcRemapByPtNum(dcAiDocComponents(AiDoc, null, 0)));

        return this; // dcCopy(dcWkg)
    }
}