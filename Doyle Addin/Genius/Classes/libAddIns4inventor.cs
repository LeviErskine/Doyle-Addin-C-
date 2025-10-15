namespace Doyle_Addin.Genius.Classes;

class libAddIns4inventor
{
    public Dictionary dcAddIns4inventor(Application app = null)
    {
        while (true)
        {
            if (app == null)
            {
                app = ThisApplication;
            }
            else
            {
                var rt = new Dictionary();
                foreach (ApplicationAddIn adIn in app.ApplicationAddIns)
                {
                    rt.Add(adIn.ClassIdString, adIn);
                }

                return rt;
            }
        }
    }

    public ApplicationAddIn addIn4inventor(string clsId)
    {
        {
            var withBlock = dcAddIns4inventor();
            return withBlock.Exists(clsId) ? (ApplicationAddIn)withBlock.get_Item(clsId) : null;
        }
    }

    public ApplicationAddIn addInILogic()
    {
        return addIn4inventor("{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}");
    }
}