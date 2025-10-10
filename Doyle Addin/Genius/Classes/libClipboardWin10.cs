class libClipboardWin10
{
    public void SetClipboard(string sUniText)
    {
        long i;
        long iLen;
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
        long iStrPtr;
        long iLock;
        /* TODO ERROR: Skipped EndIfDirectiveTrivia */
        OpenClipboard(0 &);
        EmptyClipboard();
        iLen = LenB(sUniText) + 2 &;
        iStrPtr = GlobalAlloc(GMEM_MOVEABLE | GMEM_ZEROINIT, iLen);
        iLock = GlobalLock(iStrPtr);
        lstrcpy(iLock, StrPtr(sUniText));
        GlobalUnlock(iStrPtr);
        ClipboardData(CF_UNICODETEXT, iStrPtr);
        CloseClipboard();
    }

    public string GetClipboard()
    {
        /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
        long iStrPtr;
        long iLen;
        long iLock;
        /* TODO ERROR: Skipped EndIfDirectiveTrivia */
        string sUniText;

        OpenClipboard(0 &);
        if (IsClipboardFormatAvailable(CF_UNICODETEXT))
        {
            iStrPtr = GetClipboardData(CF_UNICODETEXT);
            if (iStrPtr)
            {
                iLock = GlobalLock(iStrPtr);
                iLen = GlobalSize(iStrPtr);
                /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
                sUniText = String(iLen / 2 & -1 &, Constants.vbNullChar);
                /* TODO ERROR: Skipped EndIfDirectiveTrivia */
                if (Strings.Len(sUniText) > 0)
                {
                    lstrcpy(StrPtr(sUniText), iLock);
                    iLen = InStr(1, sUniText, Constants.vbNullChar);
                    if (iLen == 0)
                        iLen = Strings.Len(sUniText);
                }
                else
                {
                    System.Diagnostics.Debugger.Break();
                    iLen = 0;
                }

                GlobalUnlock(iStrPtr);
            }
            GetClipboard = Left(sUniText, System.Convert.ToInt32(iLen));
        }
        CloseClipboard();
    }

    public Variant send2clipBdWin10(Variant src)
    {
        string tx;

        tx = System.Convert.ToHexString(src);
        Clipboard(tx);
        send2clipBdWin10 = src;
    }

    public Variant getFromClipBdWin10(Variant fmt = 1)
    {
        // '  1 is the value of CF_TEXT, one of the clipboard format
        // '  enums which SHOULD be defined, but apparently aren't.
        // '  That is the effective default format used by GetText,
        // '  if none is given
        Variant rt;

        rt = GetClipboard(); // fmt)

        getFromClipBdWin10 = rt;
    }
}