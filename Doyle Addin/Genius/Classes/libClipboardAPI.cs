

Public Const GHND = &H42
Public Const MAXSIZE = 4096
Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40

Public Const CF_TEXT = 1
Public Const CF_UNICODETEXT As Long = 13

#If Mac Then
    ' ignore
#Else
#If VBA7 Then

Public Declare PtrSafe Function OpenClipboard Lib "user32.dll" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function CloseClipboard Lib "user32.dll" () As Long
Public Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As LongPtr
Public Declare PtrSafe Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function EmptyClipboard Lib "user32.dll" () As Long

Public Declare PtrSafe Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalLock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalSize Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As LongPtr) As LongPtr

Public Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr

#Else

Public Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32.dll" () As Long
Public Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
Public Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function EmptyClipboard Lib "user32.dll" () As Long

Public Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Public Declare Function GlobalSize Lib "kernel32.dll" (ByVal hMem As Long) As Long

Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long

#End If
#End If

