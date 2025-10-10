class SurroundingClass
{
    // ' Will be using this one to work on VB Extensibility code

    public long m2g0f0()
    {
        Inventor.InventorVBAProject pj;

        {
            var withBlock = ThisApplication;
            foreach (var pj in withBlock.VBAProjects)
                Debug.Print(pj.VBProject);
            m2g0f0 = withBlock.VBAProjects.Count;
        }
    }

    public Inventor.InventorVBAProject m2g1f0()
    {
        m2g1f0 = ThisApplication.VBAProjects.Item(1);
    }

    public string fnOfDefaultVBAproject()
    {
        fnOfDefaultVBAproject = ThisApplication.FileOptions.DefaultVBAProjectFileFullFilename;
    }

    public VBIDE.VBProject m2g1f2(Inventor.InventorVBAProject ob)
    {
        m2g1f2 = ob.VBProject;
    }
    // Debug.Print m2g1f2(dcInVBAprojects(ThisApplication).Item(fnOfDefaultVBAproject)).BuildFileName


    public Scripting.Dictionary dcInVBAprojects(Inventor.Application ap)
    {
        Scripting.Dictionary rt;
        Inventor.InventorVBAProject pj;
        long mx;
        long dx;

        rt = new Scripting.Dictionary();
        {
            var withBlock = ap.VBAProjects;
            mx = withBlock.Count;
            for (dx = 1; dx <= mx; dx++)
            {
                pj = withBlock.Item(dx);
                rt.Add(m2g1f2(pj).Filename, pj);
            }
        }
        dcInVBAprojects = rt;
    }
}