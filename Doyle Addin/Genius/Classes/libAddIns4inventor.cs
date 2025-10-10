class libAddIns4inventor
{
    public Scripting.Dictionary dcAddIns4inventor(Inventor.Application app = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Dictionary rt;
        Inventor.ApplicationAddIn adIn;

        if (app == null)
            dcAddIns4inventor = dcAddIns4inventor(ThisApplication);
        else
        {
            rt = new Scripting.Dictionary();
            foreach (var adIn in app.ApplicationAddIns)
            {
                rt.Add(adIn.ClassIdString, adIn);
                DoEvents();
            }
            dcAddIns4inventor = rt;
        }
    }

    public Inventor.ApplicationAddIn addIn4inventor(string clsId)
    {
        {
            var withBlock = dcAddIns4inventor();
            if (withBlock.Exists(clsId))
                addIn4inventor = withBlock.Item(clsId);
            else
                addIn4inventor = null/* TODO Change to default(_) if this is not a reference type */;
        }
    }

    public Inventor.ApplicationAddIn addInILogic()
    {
        addInILogic = addIn4inventor("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}");
    }
}