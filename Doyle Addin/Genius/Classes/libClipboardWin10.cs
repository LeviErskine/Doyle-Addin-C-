

Public Sub SetClipboard(ByVal sUniText As String)
    Dim i As Long
    Dim iLen As Long
#If VBA7 Then
    Dim iStrPtr As LongPtr
    Dim iLock As LongPtr
#Else
    Dim iStrPtr As Long
    Dim iLock As Long
#End If

    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(sUniText) + 2&
    iStrPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(iStrPtr)
    lstrcpy iLock, StrPtr(sUniText)
    GlobalUnlock iStrPtr
    SetClipboardData CF_UNICODETEXT, iStrPtr
    CloseClipboard
End Sub

Public Function GetClipboard() As String
#If VBA7 Then
    Dim iStrPtr As LongPtr
    Dim iLen As LongPtr
    Dim iLock As LongPtr
#Else
    Dim iStrPtr As Long
    Dim iLen As Long
    Dim iLock As Long
#End If
    Dim sUniText As String

    OpenClipboard 0&
    If IsClipboardFormatAvailable(CF_UNICODETEXT) Then
        iStrPtr = GetClipboardData(CF_UNICODETEXT)
        If iStrPtr Then
            iLock = GlobalLock(iStrPtr)
            iLen = GlobalSize(iStrPtr)
#If VBA7 Then
            If iLen > 4294967272^ Then
                sUniText = ""
            Else
                sUniText = String$(CLng(iLen \ 2^ - 1^), vbNullChar)
            End If
#Else
            sUniText = String$(iLen \ 2& - 1&, vbNullChar)
#End If
            If Len(sUniText) > 0 Then
                lstrcpy StrPtr(sUniText), iLock
                iLen = InStr(1, sUniText, vbNullChar)
                If iLen = 0 Then iLen = Len(sUniText)
            Else
                Stop
                iLen = 0
            End If
            
            GlobalUnlock iStrPtr
        End If
        GetClipboard = Left$(sUniText, CLng(iLen))
    End If
    CloseClipboard
End Function

Public Function send2clipBdWin10(src As Variant) As Variant
    Dim tx As String
    
    tx = CStr(src)
    SetClipboard tx
    send2clipBdWin10 = src
End Function

Public Function getFromClipBdWin10(Optional fmt As Variant = 1) As Variant
    ''  1 is the value of CF_TEXT, one of the clipboard format
    ''  enums which SHOULD be defined, but apparently aren't.
    ''  That is the effective default format used by GetText,
    ''  if none is given
    Dim rt As Variant
    
    rt = GetClipboard() 'fmt)
    
    getFromClipBdWin10 = rt
End Function

