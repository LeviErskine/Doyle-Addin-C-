// Clipboard APIs

// For Constants if needed

namespace Doyle_Addin.Genius.Classes;

public class libClipboardWin10 // : libClipboardAPI (base class is sealed; using standalone implementation)
{
    // Re-implemented using .NET Clipboard APIs. Original Win32/VB code kept below (commented)
    public static void SetClipboard(string sUniText)
    {
        // .NET Unicode clipboard
        try
        {
            Clipboard.SetText(sUniText ?? Empty, TextDataFormat.UnicodeText);
        }
        catch
        {
            // If not in STA or clipboard is locked, fail silently to preserve behavior
            // Consider retry logic or STA marshaling in future revision
        }

        /*
        // Original (VB/Win32) implementation - does not compile in C#; preserved for reference
        long i;
        OpenClipboard(0 &);
        EmptyClipboard();
        long iLen = LenB(sUniText) + 2 &;
        var iStrPtr = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, iLen);
        var iLock = GlobalLock(iStrPtr);
        lstrcpy(iLock, StrPtr(sUniText));
        GlobalUnlock(iStrPtr);
        ClipboardData(CF_UNICODETEXT, iStrPtr);
        CloseClipboard();
        */
    }

    public static string GetClipboard()
    {
        // Return Unicode text if present; else empty string
        try
        {
            if (Clipboard.ContainsText(TextDataFormat.UnicodeText))
                return Clipboard.GetText(TextDataFormat.UnicodeText);
            if (Clipboard.ContainsText())
                return Clipboard.GetText();
        }
        catch
        {
            // Clipboard unavailable or thread not STA
        }
        return Empty;

        /*
        // Original (VB/Win32) implementation - does not compile in C#; preserved for reference
        long iLen;
        string sUniText;

        OpenClipboard(0 &);
        if (IsClipboardFormatAvailable(CF_UNICODETEXT))
        {
            var iStrPtr = GetClipboardData(CF_UNICODETEXT);
            if (!iStrPtr) return Left(sUniText, Convert.ToInt32(iLen));
            var iLock = GlobalLock(iStrPtr);
            iLen = GlobalSize(iStrPtr);
            sUniText = String(iLen / 2 & -1 &, Constants.vbNullChar);
            if (Strings.Len(sUniText) > 0)
            {
                lstrcpy(StrPtr(sUniText), iLock);
                iLen = InStr(1, sUniText, Constants.vbNullChar);
                if (iLen == 0)
                    iLen = Strings.Len(sUniText);
            }
            else
            {
                Debugger.Break();
                iLen = 0;
            }

            GlobalUnlock(iStrPtr);
            return Left(sUniText, Convert.ToInt32(iLen));
        }
        CloseClipboard();
        */
    }

    public static dynamic send2clipBdWin10(dynamic src)
    {
        // Convert bytes to hex; otherwise ToString
        string tx;
        try
        {
            if (src is byte[] bytes)
            {
                tx = Convert.ToHexString(bytes);
            }
            else
            {
                tx = Convert.ToString(src) ?? Empty;
            }

            try { Clipboard.SetText(tx, TextDataFormat.UnicodeText); } catch { }
        }
        catch
        {
            // Swallow conversion errors to match legacy behavior
        }
        return src;

        /*
        var tx = Convert.ToHexString(src);
        Clipboard(tx);
        return src;
        */
    }

    public static dynamic getFromClipBdWin10(dynamic fmt = null)
    {
        // NOTE: fmt retained for compatibility (VB used CF_TEXT=1). Here we return text only.
        var rt = GetClipboard();
        return rt;

        /*
        // '  1 is the value of CF_TEXT, one of the clipboard format
        // '  enums which SHOULD be defined, but apparently aren't.
        // '  That is the effective default format used by GetText,
        // '  if none is given
        dynamic rt = GetClipboard(); // fmt)
        return rt;
        */
    }
}