namespace Doyle_Addin.Genius.Classes;

class modAiVba
{
    // ' Will be using this one to work on VB Extensibility code

    public long m2g0f0()
    {
        {
            var withBlock = ThisApplication;
            foreach (InventorVBAProject pj in withBlock.VBAProjects)
                Debug.Print(pj.VBProject);
            return withBlock.VBAProjects.Count;
        }
    }

    public InventorVBAProject m2g1f0()
    {
        return ThisApplication.VBAProjects.get_item(1);
    }

    public string fnOfDefaultVBAproject()
    {
        return ThisApplication.FileOptions.DefaultVBAProjectFileFullFilename;
    }

    public VBIDE.VBProject m2g1f2(InventorVBAProject ob)
    {
        return ob.VBProject;
    }
    // Debug.Print m2g1f2(dcInVBAprojects(ThisApplication).get_item(fnOfDefaultVBAproject)).BuildFileName

    public Dictionary dcInVBAprojects(Application ap)
    {
        var rt = new Dictionary();
        {
            var withBlock = ap.VBAProjects;
            long mx = withBlock.Count;
            for (long dx = 1; dx <= mx; dx++)
            {
                InventorVBAProject pj = withBlock.get_item(dx);
                rt.Add(m2g1f2(pj).FileName, pj);
            }
        }
        return rt;
    }
}