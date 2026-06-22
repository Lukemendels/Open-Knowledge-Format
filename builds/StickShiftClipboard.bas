Attribute VB_Name = "StickShiftClipboard"
' =====================================================================
'  StickShift Clipboard -- OKF-compliant
'  Shared Win-API clipboard reader.
'
'  Provides GetClipboardText(), a reliable Unicode-aware clipboard
'  reader used by both StickShiftWriteApply and StickShiftContextBundle.
'
'  All Win32 Declare statements and format constants are Private to
'  this module; only GetClipboardText is Public.
' =====================================================================

Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As LongPtr
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
#End If

Private Const CF_UNICODETEXT As Long = 13&
Private Const CF_TEXT        As Long = 1&

Public Function GetClipboardText() As String
    Dim hData As LongPtr
    Dim pData As LongPtr
    Dim sizeBytes As Long
    Dim tmp As String
    Dim charsCount As Long
    Dim nulPos As Long

    ' Try Unicode text first.
    If OpenClipboard(0) = 0 Then Exit Function

    On Error GoTo CleanExit

    If IsClipboardFormatAvailable(CF_UNICODETEXT) <> 0 Then
        hData = GetClipboardData(CF_UNICODETEXT)
        If hData <> 0 Then
            pData = GlobalLock(hData)
            If pData <> 0 Then
                sizeBytes = CLng(GlobalSize(hData))
                If sizeBytes > 0 Then
                    ' Each Unicode character is 2 bytes.
                    charsCount = sizeBytes \ 2
                    tmp = String$(charsCount, vbNullChar)
                    lstrcpyW StrPtr(tmp), pData

                    ' Trim at first null terminator.
                    nulPos = InStr(1, tmp, vbNullChar)
                    If nulPos > 0 Then
                        GetClipboardText = Left$(tmp, nulPos - 1)
                    Else
                        GetClipboardText = tmp
                    End If
                End If
                GlobalUnlock hData
            End If
        End If

    ElseIf IsClipboardFormatAvailable(CF_TEXT) <> 0 Then
        ' Fallback to ANSI text and convert to Unicode.
        hData = GetClipboardData(CF_TEXT)
        If hData <> 0 Then
            pData = GlobalLock(hData)
            If pData <> 0 Then
                sizeBytes = CLng(GlobalSize(hData))
                If sizeBytes > 0 Then
                    tmp = String$(sizeBytes, vbNullChar)
                    lstrcpyW StrPtr(tmp), pData
                    nulPos = InStr(1, tmp, vbNullChar)
                    If nulPos > 0 Then
                        tmp = Left$(tmp, nulPos - 1)
                    End If
                    GetClipboardText = StrConv(tmp, vbUnicode)
                End If
                GlobalUnlock hData
            End If
        End If
    End If

CleanExit:
    CloseClipboard
End Function
