Attribute VB_Name = "StickShiftWriteApply"
' =====================================================================
'  StickShift Write Apply  (conformant OKF v0.1 writer)
'
'  Reads a <VBA_WRITE> envelope from the clipboard (output of the
'  Portfolio Writer DHSChat Assistant), parses each ### FILE: block,
'  and writes every file directly (creating parent dirs as needed).
'
'  Machine-owned files (log.md, index.md) are never overwritten by a
'  write block. Every write is appended to the bundle-root log.md as
'  an append-only audit trail: action = new | edit.
'
'  After applying, calls GenerateStickShiftIndexes so new builds appear
'  in the index immediately.
'
'  Requires: StickShiftClipboard module (GetClipboardText),
'            StickShiftConfig module, StickShiftIndexGenerator module.
'  Microsoft ActiveX Data Objects 2.x (ADODB.Stream for UTF-8 I/O).
'
'  StickShiftConfig         -> BundleRoot, BundleRootRaw, SetBundleRoot
'  StickShiftWriteApply     -> ApplyStickShiftWrite, ApplyWriteEnvelopeText
'  StickShiftIndexGenerator -> GenerateStickShiftIndexes
'  StickShiftBootstrap      -> BootstrapBundle
' =====================================================================

Option Explicit

Private m_BundleRoot As String

Private fso As Object


' Parses a <VBA_WRITE> envelope string, writes its ### FILE blocks (honoring the
' index.md/log.md reserved guard), logs each write to log.md, and returns a
' summary string "(N written, M skipped)". Self-contained: sets m_BundleRoot
' from StickShiftConfig.BundleRoot() so the private helpers (ResolvePath /
' WriteUtf8 / AppendEditLog) work regardless of caller.
Public Function ApplyWriteEnvelopeText(ByVal envelope As String, _
                                        ByRef writeCount As Long, _
                                        ByRef skipCount As Long) As Boolean
    m_BundleRoot = StickShiftConfig.BundleRoot()
    If m_BundleRoot = "" Then
        MsgBox "Bundle root not set - click Switch Context.", vbExclamation, "StickShift"
        ApplyWriteEnvelopeText = False
        Exit Function
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(m_BundleRoot) Then
        MsgBox "Bundle root not found: " & m_BundleRoot, vbCritical, "StickShift"
        ApplyWriteEnvelopeText = False
        Exit Function
    End If

    ' --- Normalise line endings before searching ---
    Dim body As String
    body = Replace(Replace(envelope, vbCrLf, vbLf), vbCr, vbLf)

    ' --- Require <VBA_WRITE> block ---
    Dim startPos As Long, endPos As Long
    startPos = InStr(body, "<VBA_WRITE>")
    endPos = InStr(body, "</VBA_WRITE>")
    If startPos = 0 Or endPos = 0 Or endPos <= startPos Then
        MsgBox "No <VBA_WRITE> block found in clipboard." & vbLf & _
               "Copy the Portfolio Writer's full output and try again.", vbExclamation
        ApplyWriteEnvelopeText = False
        Exit Function
    End If

    Dim tagLen As Long: tagLen = Len("<VBA_WRITE>")
    body = Mid(body, startPos + tagLen, endPos - (startPos + tagLen))

    ' --- Parse ### FILE: ... ### END FILE pairs ---
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
        ApplyWriteEnvelopeText = False
        Exit Function
    End If

    ' --- Write all files (create parent dir if needed) ---
    writeCount = 0
    skipCount = 0
    Dim logLines As String: logLines = ""

    Dim i As Long
    Dim leaf As String
    Dim absPath As String
    Dim existed As Boolean
    Dim parentDir As String
    Dim action As String
    Dim relFwd As String

    For i = 0 To fileCount - 1
        ' Guard: never overwrite machine-owned files.
        leaf = LCase(fso.GetFileName(filePaths(i)))
        If leaf = "log.md" Or leaf = "index.md" Then
            skipCount = skipCount + 1
        Else
            absPath = ResolvePath(filePaths(i))
            existed = fso.FileExists(absPath)

            parentDir = fso.GetParentFolderName(absPath)
            If Not fso.FolderExists(parentDir) Then EnsureFolderTree parentDir

            WriteUtf8 absPath, fileContents(i)
            writeCount = writeCount + 1

            action = IIf(existed, "edit", "new")
            relFwd = Replace(filePaths(i), "\", "/")
            If Left(relFwd, 1) = "/" Then relFwd = Mid(relFwd, 2)
            logLines = logLines & "- " & Format(Now, "yyyy-mm-dd hh:nn:ss") & _
                       "  " & action & "  " & relFwd & vbLf
        End If
    Next i

    ' --- Append one batch of log entries ---
    If logLines <> "" Then AppendEditLog logLines

    ApplyWriteEnvelopeText = True
End Function


Sub ApplyStickShiftWrite()
    ' Early root + folder check so the clipboard is never read unnecessarily.
    Dim rootCheck As String
    rootCheck = StickShiftConfig.BundleRoot()
    If rootCheck = "" Then
        MsgBox "Bundle root not set - click Switch Context.", vbExclamation, "StickShift"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(rootCheck) Then
        MsgBox "Bundle root not found: " & rootCheck, vbCritical, "StickShift"
        Exit Sub
    End If

    ' --- Read clipboard ---
    Dim clip As String
    On Error Resume Next
    clip = ReadClipboard()
    On Error GoTo 0
    If clip = "" Then
        MsgBox "Clipboard is empty. Copy the Portfolio Writer's output and try again.", vbExclamation
        Exit Sub
    End If

    Dim w As Long, s As Long
    If ApplyWriteEnvelopeText(clip, w, s) Then
        Dim summary As String
        Dim ws As Object
        Dim ts As String

        summary = w & " file(s) written. Logged to log.md."
        If s > 0 Then
            summary = summary & vbLf & s & " reserved file(s) skipped (log.md / index.md)."
        End If

        ts = Format(Now(), "yyyy-mm-dd hh:mm:ss")

        ' Write success summary (with timestamp) into StickShift!D9 instead of showing a MsgBox.
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets("StickShift")
        If Not ws Is Nothing Then
            ws.Range("D9").Value = ts & "  |  " & Replace(summary, vbLf, "  ")
        End If
        On Error GoTo 0

        ' Regenerate index so new builds appear immediately.
        StickShiftIndexGenerator.GenerateStickShiftIndexes
    End If
End Sub



Private Sub AppendEditLog(ByVal entries As String)
    Dim logPath As String
    logPath = m_BundleRoot & "log.md"

    Dim existing As String
    If fso.FileExists(logPath) Then
        existing = ReadUtf8(logPath)
    Else
        existing = "# Log" & vbLf & vbLf
    End If

    WriteUtf8 logPath, existing & entries
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

    ReadClipboard = StickShiftClipboard.GetClipboardText()
    Exit Function

FailSafe:
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

' Create every missing folder in the chain down to dirPath (absolute, under m_BundleRoot).
Private Sub EnsureFolderTree(ByVal dirPath As String)
    If dirPath = "" Then Exit Sub
    If fso.FolderExists(dirPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(dirPath)
    If parent <> "" And Not fso.FolderExists(parent) Then EnsureFolderTree parent
    On Error Resume Next
    fso.CreateFolder dirPath
    On Error GoTo 0
End Sub
