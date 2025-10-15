using System;
using System.Runtime.InteropServices;

namespace Doyle_Addin.Genius.Classes;

public static class libClipboardAPI
{
    // Global memory flags
    /// <summary>
    /// 
    /// </summary>
    public const uint GMEM_MOVEABLE = 0x0002;
    /// <summary>
    /// 
    /// </summary>
    public const uint GMEM_ZEROINIT = 0x0040;
    public const uint GHND = GMEM_MOVEABLE | GMEM_ZEROINIT;

    /// <summary>
    /// 
    /// </summary>
    public const int MAXSIZE = 4096;

    
    /// <summary>
    /// Clipboard formats
    /// </summary>
    public const uint CF_TEXT = 1;
    /// <summary>
    /// 
    /// </summary>
    public const uint CF_UNICODETEXT = 13;

    // user32.dll
    /// <summary>
    /// 
    /// </summary>
    /// <param name="hWndNewOwner"></param>
    /// <returns></returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool OpenClipboard(IntPtr hWndNewOwner);

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool CloseClipboard();

    /// <summary>
    /// 
    /// </summary>
    /// <param name="format"></param>
    /// <returns></returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool IsClipboardFormatAvailable(uint format);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="uFormat"></param>
    /// <returns></returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetClipboardData(uint uFormat);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="uFormat"></param>
    /// <param name="hMem"></param>
    /// <returns></returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

    /// <summary>
    /// 
    /// </summary>
    /// <returns></returns>
    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool EmptyClipboard();

    // kernel32.dll
    /// <summary>
    /// 
    /// </summary>
    /// <param name="uFlags"></param>
    /// <param name="dwBytes"></param>
    /// <returns></returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMem"></param>
    /// <returns></returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern IntPtr GlobalLock(IntPtr hMem);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMem"></param>
    /// <returns></returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern bool GlobalUnlock(IntPtr hMem);

    /// <summary>
    /// 
    /// </summary>
    /// <param name="hMem"></param>
    /// <returns></returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    public static extern UIntPtr GlobalSize(IntPtr hMem);

    // String copy (Unicode)
    /// <summary>
    /// 
    /// </summary>
    /// <param name="lpString1"></param>
    /// <param name="lpString2"></param>
    /// <returns></returns>
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "lstrcpyW")]
    public static extern IntPtr lstrcpy(IntPtr lpString1, string lpString2);
}