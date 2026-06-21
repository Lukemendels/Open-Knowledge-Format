Attribute VB_Name = "OKFWriteApply"
' =====================================================================
'  OKF Write Apply  (conformant OKF v0.1 writer)
'
'  Reads a <VBA_WRITE> envelope from the clipboard (output of the
'  Portfolio Writer DHSChat Assistant), parses each ### FILE: block,
'  and writes every file directly (creating parent dirs as needed).
'
'  After applying, calls GenerateOKFIndexes so new builds appear in
'  the index immediately.
'
'  Requires: OKFClipboard module (GetClipboardText), OKFConfig module,
'  OKFIndexGenerator module.
'  Microsoft ActiveX Data Objects 2.x (ADODB.Stream for UTF-8 I/O).
'
'  *** BUNDLE_ROOT must match the value in OKFIndexGenerator.bas ***
' =====================================================================

Option Explicit

Private m_BundleRoot As String

Private fso As Object


Sub ApplyOKFWrite()
    m_BundleRoot = OKFConfig.BundleRoot()
    If m_BundleRoot = "" Then
        MsgBox "Bundle root not set - click Set Bundle Root.", vbExclamation, "OKF Write Apply"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(m_BundleRoot) Then
        MsgBox "Bundle root not found: " & m_BundleRoot, vbCritical, "OKF Write Apply"
        Exit Sub
    End If

    ' --- 1. Read clipboard ---
    Dim clip As String
    On Error Resume Next
    clip = ReadClipboard()
    On Error GoTo 0
    If clip = "" Then
        MsgBox "Clipboard is empty. Copy the Portfolio Writer's output and try again.", vbExclamation
        Exit Sub
    End If

    ' --- 2. Require <VBA_WRITE> block ---
    ' Normalise line endings before searching.
    clip = Replace(Replace(clip, vbCrLf, vbLf), vbCr, vbLf)

    Dim startPos As Long, endPos As Long
    startPos = InStr(clip, "<VBA_WRITE>")
    endPos = InStr(clip, "</VBA_WRITE>")
    If startPos = 0 Or endPos = 0 Or endPos <= startPos Then
        MsgBox "No <VBA_WRITE> block found in clipboard." & vbLf & _
               "Copy the Portfolio Writer's full output and try again.", vbExclamation
        Exit Sub
    End If

    Dim body As String
    Dim tagLen As Long: tagLen = Len("<VBA_WRITE>")
    body = Mid(clip, startPos + tagLen, endPos - (startPos + tagLen))

    ' --- 3. Parse ### FILE: ... ### END FILE pairs ---
    Dim filePaths(0 To 99) As String
    Dim fileContents(0 To 99) As String
    Dim fileCount As Long: fileCount = 0

    Dim searchFrom As Long: searchFrom = 1
    Do
        Dim fileTagPos As Long
        fileTagPos = InStr(searchFrom, body, "### FILE:")
        If fileTagPos = 0 Then Exit Do

        Dim eolPos As Long
        eolPos = InStr(fileTagPos, body, vbLf)
        If eolPos = 0 Then Exit Do

        Dim relPath As String
        relPath = Trim(Mid(body, fileTagPos + Len("### FILE:"), eolPos - (fileTagPos + Len("### FILE:"))))

        Dim endFilePos As Long
        endFilePos = InStr(eolPos, body, "### END FILE")
        If endFilePos = 0 Then Exit Do

        Dim contents As String
        contents = Mid(body, eolPos + 1, endFilePos - eolPos - 1)
        If Right(contents, 1) = vbLf Then contents = Left(contents, Len(contents) - 1)

        If relPath <> "" And fileCount <= 99 Then
            filePaths(fileCount) = relPath
            fileContents(fileCount) = contents
            fileCount = fileCount + 1
        End If

        searchFrom = endFilePos + Len("### END FILE")
    Loop

    If fileCount = 0 Then
        MsgBox "No ### FILE: blocks found inside <VBA_WRITE>. Nothing to apply.", vbExclamation
        Exit Sub
    End If

    ' --- 4. Write all files (create parent dir if needed) ---
    Dim writeCount As Long: writeCount = 0

    Dim i As Long
    For i = 0 To fileCount - 1
        Dim absPath As String
        absPath = ResolvePath(filePaths(i))

        Dim parentDir As String
        parentDir = fso.GetParentFolderName(absPath)
        If Not fso.FolderExists(parentDir) Then
            fso.CreateFolder parentDir
        End If

        WriteUtf8 absPath, fileContents(i)
        writeCount = writeCount + 1
    Next i

    ' --- 5. Summary ---
    MsgBox writeCount & " file(s) written.", vbInformation, "OKF Write Apply"

    ' --- 6. Regenerate index so new builds appear immediately ---
    GenerateOKFIndexes

End Sub


Private Function ResolvePath(ByVal relPath As String) As String
    Dim p As String
    p = Trim(relPath)
    p = Replace(p, "/", "\")
    If Left(p, 1) = "\" Then p = Mid(p, 2)
    ResolvePath = m_BundleRoot & p
End Function

Private Function ReadClipboard() As String
    On Error GoTo FailSafe

    ReadClipboard = OKFClipboard.GetClipboardText()
    Exit Function

FailSafe:
    ' If anything goes wrong, return empty string so caller can handle.
    ReadClipboard = ""
End Function

Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal content As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText content
    st.SaveToFile path, 2
    st.Close
End Sub
