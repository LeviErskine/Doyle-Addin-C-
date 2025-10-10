class libClipboardAPI
{
    public const var GHND = 0x42;
    public const var MAXSIZE = 4096;
    public const long GMEM_MOVEABLE = 0x2;
    public const long GMEM_ZEROINIT = 0x40;

    public const var CF_TEXT = 1;
    public const long CF_UNICODETEXT = 13;

    /* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia *//* TODO ERROR: Skipped IfDirectiveTrivia *//* TODO ERROR: Skipped DisabledTextTrivia *//* TODO ERROR: Skipped ElseDirectiveTrivia */
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern long OpenClipboard(long hwnd);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern long CloseClipboard();
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern long IsClipboardFormatAvailable(long wFormat);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern long GetClipboardData(long wFormat);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern long SetClipboardData(long wFormat, long hMem);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern long EmptyClipboard();

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    public static extern long GlobalAlloc(long wFlags, long dwBytes);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    public static extern long GlobalLock(long hMem);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    public static extern long GlobalUnlock(long hMem);
    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    public static extern long GlobalSize(long hMem);

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    public static extern long lstrcpy(Any lpString1, Any lpString2);
}