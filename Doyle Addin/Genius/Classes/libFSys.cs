using Microsoft.VisualBasic;

namespace Doyle_Addin.Genius.Classes;

public class libFSys
{
    public static FileSystemObject nuFso()
    {
        return new FileSystemObject();
    }

    public static Folder fdUserHome()
    {
        return nuFso().GetFolder(Interaction.Environ("AppData")).ParentFolder;
    }

    public static string fnExt(string fn)
    {
        var ar = Strings.Split(fn, ".");
        var mx = ar.Length - 1;
        return mx < 0 ? "" : ar[mx];
    }

    public static Folder folderIfPresent(string Path, Folder Base = null)
    {
        Folder rt;

        if (Base == null)
        {
            var withBlock = nuFso();
            rt = withBlock.FolderExists(Path) ? withBlock.GetFolder(Path) : null;
        }
        else
            rt = folderIfPresent(Base.Path + @"\" + Path);

        return rt;
    }

    public static Scripting.File fileIfPresent(string Path, Folder Base = null)
    {
        Scripting.File rt;

        if (Base == null)
        {
            var withBlock = nuFso();
            rt = withBlock.FileExists(Path) ? withBlock.GetFile(Path) : null;
        }
        else
            rt = fileIfPresent(Base.Path + @"\" + Path);

        return rt;
    }

    public static Dictionary dcFilesIn(Folder fd)
    {
        var rt = new Dictionary();
        if (fd == null) return rt;
        foreach (Scripting.File fl in fd.Files)
            rt.Add(fl.Name, fl);
        return rt;
    }
    // send2clipBd dcFilesIn(nuFso.GetFolder(""))
    // send2clipBd Join(dcFoldersIn(nuFso.GetFolder("C:\Doyle_Vault\Designs\doyle")).Keys, vbCrLf)
    // send2clipBd Join(dcFilesIn(nuFso.GetFolder("W:\Parts Lists")).Keys, vbCrLf)

    public static Dictionary dcFoldersIn(Folder fd)
    {
        var rt = new Dictionary();
        if (fd == null) return rt;
        foreach (Folder fl in fd.SubFolders)
            rt.Add(fl.Name, fl);
        return rt;
    }
}