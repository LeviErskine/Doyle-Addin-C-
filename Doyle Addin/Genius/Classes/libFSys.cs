class SurroundingClass
{
    public Scripting.FileSystemObject nuFso()
    {
        nuFso = new Scripting.FileSystemObject();
    }

    public Scripting.Folder fdUserHome()
    {
        fdUserHome = nuFso.GetFolder(Interaction.Environ("AppData")).ParentFolder;
    }

    public string fnExt(string fn)
    {
        Variant ar;
        long mx;

        ar = Split(fn, ".");
        mx = UBound(ar);
        if (mx < 0)
            fnExt = "";
        else
            fnExt = ar(mx);
    }

    public Scripting.Folder folderIfPresent(string Path, Scripting.Folder Base = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.Folder rt;

        if (Base == null)
        {
            {
                var withBlock = nuFso();
                if (withBlock.FolderExists(Path))
                    rt = withBlock.GetFolder(Path);
                else
                    rt = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }
        else
            rt = folderIfPresent(Base.Path + @"\" + Path);

        folderIfPresent = rt;
    }

    public Scripting.File fileIfPresent(string Path, Scripting.Folder Base = null/* TODO Change to default(_) if this is not a reference type */)
    {
        Scripting.File rt;

        if (Base == null)
        {
            {
                var withBlock = nuFso();
                if (withBlock.FileExists(Path))
                    rt = withBlock.GetFile(Path);
                else
                    rt = null/* TODO Change to default(_) if this is not a reference type */;
            }
        }
        else
            rt = fileIfPresent(Base.Path + @"\" + Path);

        fileIfPresent = rt;
    }

    public Scripting.Dictionary dcFilesIn(Scripting.Folder fd)
    {
        Scripting.Dictionary rt;
        Scripting.File fl;

        rt = new Scripting.Dictionary();
        if (!fd == null)
        {
            foreach (var fl in fd.Files)
                rt.Add(fl.Name, fl);
        }
        dcFilesIn = rt;
    }
    // send2clipBd dcFilesIn(nuFso.GetFolder(""))
    // send2clipBd Join(dcFoldersIn(nuFso.GetFolder("C:\Doyle_Vault\Designs\doyle")).Keys, vbNewLine)
    // send2clipBd Join(dcFilesIn(nuFso.GetFolder("W:\Parts Lists")).Keys, vbNewLine)

    public Scripting.Dictionary dcFoldersIn(Scripting.Folder fd)
    {
        Scripting.Dictionary rt;
        Scripting.Folder fl;

        rt = new Scripting.Dictionary();
        if (!fd == null)
        {
            foreach (var fl in fd.SubFolders)
                rt.Add(fl.Name, fl);
        }
        dcFoldersIn = rt;
    }
}