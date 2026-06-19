Attribute VB_Name = "OKFLint"
' =====================================================================
'  OKF Lint  (deterministic integrity checker, no model, no mutation)
'
'  Scans every concept .md file under BUNDLE_ROOT and writes a
'  colour-coded findings table to an "OKF Lint Report" worksheet.
'  The report file is the worksheet — it lives inside the workbook,
'  never as a .md file, so the generator never indexes it.
'
'  Checks performed:
'    1. Missing / empty `type` field          (OKF conformance)
'    2. Missing `status`, `effort`, `impact`  (schema required)
'    3. Broken cross-links                    (.md links whose target is absent)
'    4. WIP violation                         (> 1 build at status:working)
'    5. Stalls                                (working + missing/old last_touched,
'                                              oldest first in report)
'    6. Pending .proposed files               (staged edits awaiting review)
'    7. Active-to-archived links              (non-archived build → archived build)
'
'  Each finding: severity (error / warning), file, one-line description.
'
'  *** BUNDLE_ROOT must match the value in OKFIndexGenerator.bas ***
'
'  Requires (Tools -> References): Microsoft ActiveX Data Objects 2.x
' =====================================================================

Option Explicit

Private Const BUNDLE_ROOT As String = "C:\Users\YourName\OneDrive\build-portfolio\"
Private Const STALE_DAYS  As Long   = 14

Private fso As Object

Sub RunOKFLint()
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(BUNDLE_ROOT) Then
        MsgBox "Bundle root not found: " & BUNDLE_ROOT, vbCritical
        Exit Sub
    End If

    Dim findings As Collection
    Set findings = New Collection

    ScanBundle findings

    WriteReport findings

    Dim summary As String
    If findings.count = 0 Then
        summary = "No findings — bundle is clean."
    Else
        summary = findings.count & " finding(s). See 'OKF Lint Report' sheet."
    End If
    MsgBox summary, vbInformation, "OKF Lint"
End Sub


Private Sub ScanBundle(ByVal findings As Collection)
    Dim buildsPath As String
    buildsPath = BUNDLE_ROOT & "builds\"

    If Not fso.FolderExists(buildsPath) Then
        AddFinding findings, "error", "(bundle)", "builds\ directory not found at " & buildsPath
        Exit Sub
    End If

    Dim allFm(0 To 999) As Variant
    Dim allFmCount As Long: allFmCount = 0

    Dim allLinks(0 To 999) As Variant
    Dim allLinksCount As Long: allLinksCount = 0

    Dim workingBuilds As Collection
    Set workingBuilds = New Collection

    Dim f As Object
    Dim buildsFolder As Object
    Set buildsFolder = fso.GetFolder(buildsPath)

    For Each f In buildsFolder.Files
        If IsConceptFile(f.Name) Then
            Dim content As String
            content = ReadUtf8(f.path)

            Dim cType As String, cTitle As String, cDesc As String
            Dim cStatus As String, cLastTouched As String
            Dim cEffort As String, cImpact As String
            ParseFrontmatterFull content, cType, cTitle, cDesc, _
                                  cStatus, cLastTouched, cEffort, cImpact

            If cType = "" Then
                AddFinding findings, "error", f.Name, "missing or empty `type` field"
            End If

            If cStatus = "" Then AddFinding findings, "error", f.Name, "missing required field `status`"
            If cEffort = "" Then AddFinding findings, "error", f.Name, "missing required field `effort`"
            If cImpact = "" Then AddFinding findings, "error", f.Name, "missing required field `impact`"

            allFm(allFmCount) = Array(f.path, LCase(cStatus), f.Name)
            allFmCount = allFmCount + 1

            Dim links As Collection
            Set links = ExtractLinks(content)
            allLinks(allLinksCount) = Array(f.path, links, f.Name)
            allLinksCount = allLinksCount + 1

            If LCase(cStatus) = "working" Then
                workingBuilds.Add Array(cLastTouched, f.path, f.Name)
            End If
        End If
    Next f

    Dim i As Long
    For i = 0 To allLinksCount - 1
        Dim filePath As String: filePath = allLinks(i)(0)
        Dim fileName As String: fileName = allLinks(i)(2)
        Dim fileLinks As Collection: Set fileLinks = allLinks(i)(1)

        Dim lnk As Variant
        For Each lnk In fileLinks
            Dim resolved As String
            resolved = ResolveLink(CStr(lnk), filePath)
            If resolved <> "" Then
                If Not fso.FileExists(resolved) Then
                    AddFinding findings, "error", fileName, "broken link: " & CStr(lnk)
                End If
            End If
        Next lnk
    Next i

    If workingBuilds.count > 1 Then
        Dim wipNames As String: wipNames = ""
        Dim wb As Variant
        For Each wb In workingBuilds
            If wipNames <> "" Then wipNames = wipNames & ", "
            wipNames = wipNames & CStr(wb(2))
        Next wb
        AddFinding findings, "error", "(portfolio)", _
            "WIP violation: " & workingBuilds.count & _
            " builds at status:working (" & wipNames & ")"
    End If

    Dim staleArr(0 To 999) As Variant
    Dim staleCount As Long: staleCount = 0

    For Each wb In workingBuilds
        Dim lt As String: lt = CStr(wb(0))
        Dim staleMsg As String: staleMsg = ""

        If lt = "" Then
            staleMsg = "working build has missing last_touched"
        Else
            Dim ltDate As Date
            ltDate = DateFromISO(lt)
            If ltDate = 0 Then
                staleMsg = "working build has invalid last_touched: " & lt
            ElseIf DateDiff("d", ltDate, Date) >= STALE_DAYS Then
                staleMsg = "stale working build: last_touched " & lt & _
                            " (" & DateDiff("d", ltDate, Date) & " days ago)"
            End If
        End If

        If staleMsg <> "" Then
            Dim sortKey As String
            sortKey = IIf(lt = "", "0000-00-00", lt)
            staleArr(staleCount) = Array(sortKey, CStr(wb(2)), staleMsg)
            staleCount = staleCount + 1
        End If
    Next wb

    If staleCount > 1 Then SortByFirstKey staleArr, staleCount

    For i = 0 To staleCount - 1
        AddFinding findings, "warning", CStr(staleArr(i)(1)), CStr(staleArr(i)(2))
    Next i

    ScanProposed BUNDLE_ROOT, findings

    Dim archivedSet As Object
    Set archivedSet = CreateObject("Scripting.Dictionary")
    archivedSet.CompareMode = vbTextCompare

    Dim statusMap As Object
    Set statusMap = CreateObject("Scripting.Dictionary")
    statusMap.CompareMode = vbTextCompare

    For i = 0 To allFmCount - 1
        Dim fmPath As String: fmPath = CStr(allFm(i)(0))
        Dim fmStatus As String: fmStatus = CStr(allFm(i)(1))
        If fmStatus = "archived" Then archivedSet(fmPath) = True
        statusMap(fmPath) = fmStatus
    Next i

    For i = 0 To allLinksCount - 1
        filePath = CStr(allLinks(i)(0))
        fileName = CStr(allLinks(i)(2))
        Set fileLinks = allLinks(i)(1)

        Dim srcStatus As String: srcStatus = ""
        If statusMap.Exists(filePath) Then srcStatus = CStr(statusMap(filePath))
        If srcStatus = "archived" Then GoTo NextLinkFile

        For Each lnk In fileLinks
            Dim resolvedAA As String
            resolvedAA = ResolveLink(CStr(lnk), filePath)
            If resolvedAA <> "" Then
                If archivedSet.Exists(resolvedAA) Then
                    AddFinding findings, "warning", fileName, _
                        "active build links to archived build: " & CStr(lnk)
                End If
            End If
        Next lnk

NextLinkFile:
    Next i
End Sub


Private Sub AddFinding(ByVal findings As Collection, ByVal severity As String, _
                       ByVal fileName As String, ByVal message As String)
    findings.Add Array(severity, fileName, message)
End Sub


Private Sub ScanProposed(ByVal folderPath As String, ByVal findings As Collection)
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)

    Dim f As Object
    For Each f In folder.Files
        If LCase(Right(f.Name, 9)) = ".proposed" Then
            Dim relPath As String
            relPath = MakeRelative(f.path)
            AddFinding findings, "warning", f.Name, _
                "pending .proposed awaiting review: " & relPath
        End If
    Next f

    Dim d As Object
    For Each d In folder.SubFolders
        ScanProposed d.path, findings
    Next d
End Sub


Private Function MakeRelative(ByVal absPath As String) As String
    If Left(absPath, Len(BUNDLE_ROOT)) = BUNDLE_ROOT Then
        MakeRelative = Mid(absPath, Len(BUNDLE_ROOT) + 1)
    Else
        MakeRelative = absPath
    End If
End Function


Private Function ResolveLink(ByVal link As String, ByVal fromFile As String) As String
    Dim hashPos As Long: hashPos = InStr(link, "#")
    If hashPos > 0 Then link = Left(link, hashPos - 1)
    link = Trim(link)

    If link = "" Then ResolveLink = "": Exit Function

    Dim ll As String: ll = LCase(link)
    If Left(ll, 7) = "http://" Or Left(ll, 8) = "https://" Then
        ResolveLink = "": Exit Function
    End If

    If LCase(Right(link, 3)) <> ".md" Then ResolveLink = "": Exit Function

    Dim normalized As String
    If Left(link, 1) = "/" Then
        normalized = BUNDLE_ROOT & Replace(Mid(link, 2), "/", "\")
    Else
        normalized = fso.GetParentFolderName(fromFile) & "\" & Replace(link, "/", "\")
    End If

    Do While InStr(normalized, "\\") > 0
        normalized = Replace(normalized, "\\", "\")
    Loop

    ResolveLink = normalized
End Function


Private Function ExtractLinks(ByVal content As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\]\(([^)]+)\)"
    re.Global = True

    On Error Resume Next
    Dim matches As Object
    Set matches = re.Execute(content)
    On Error GoTo 0
    If matches Is Nothing Then Set ExtractLinks = result: Exit Function

    Dim m As Object
    For Each m In matches
        Dim url As String
        url = Trim(m.SubMatches(0))
        Dim ul As String: ul = LCase(url)
        If Left(ul, 7) <> "http://" And Left(ul, 8) <> "https://" Then
            Dim baseUrl As String
            Dim h As Long: h = InStr(url, "#")
            If h > 0 Then baseUrl = Left(url, h - 1) Else baseUrl = url
            If LCase(Right(Trim(baseUrl), 3)) = ".md" Then
                result.Add url
            End If
        End If
    Next m

    Set ExtractLinks = result
End Function


Private Sub ParseFrontmatterFull(ByVal content As String, _
                                  ByRef cType As String, ByRef cTitle As String, _
                                  ByRef cDesc As String, ByRef cStatus As String, _
                                  ByRef cLastTouched As String, _
                                  ByRef cEffort As String, ByRef cImpact As String)
    cType = "": cTitle = "": cDesc = "": cStatus = ""
    cLastTouched = "": cEffort = "": cImpact = ""

    content = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String
    lines = Split(content, vbLf)
    If UBound(lines) < 1 Then Exit Sub
    If Trim(lines(0)) <> "---" Then Exit Sub

    Dim i As Long
    For i = 1 To UBound(lines)
        If Trim(lines(i)) = "---" Then Exit Sub
        Dim colon As Long: colon = InStr(lines(i), ":")
        If colon > 0 Then
            Dim key As String: key = LCase(Trim(Left(lines(i), colon - 1)))
            Dim val As String: val = Unquote(Trim(Mid(lines(i), colon + 1)))
            Select Case key
                Case "type":         cType = val
                Case "title":        cTitle = val
                Case "description":  cDesc = val
                Case "status":       cStatus = val
                Case "last_touched": cLastTouched = val
                Case "effort":       cEffort = val
                Case "impact":       cImpact = val
            End Select
        End If
    Next i
End Sub


Private Function IsConceptFile(ByVal name As String) As Boolean
    Dim ln As String: ln = LCase(name)
    IsConceptFile = (Right(ln, 3) = ".md") And (ln <> "index.md") And (ln <> "log.md")
End Function


Private Function DateFromISO(ByVal s As String) As Date
    On Error GoTo fail
    Dim parts() As String
    parts = Split(Trim(s), "-")
    If UBound(parts) = 2 Then
        Dim y As Long: y = CLng(parts(0))
        Dim mo As Long: mo = CLng(parts(1))
        Dim d As Long: d = CLng(parts(2))
        If y >= 1900 And y <= 2100 And mo >= 1 And mo <= 12 And d >= 1 And d <= 31 Then
            DateFromISO = DateSerial(y, mo, d)
            Exit Function
        End If
    End If
fail:
    DateFromISO = 0
End Function


Private Function Unquote(ByVal s As String) As String
    If Len(s) >= 2 Then
        If (Left(s, 1) = Chr(34) And Right(s, 1) = Chr(34)) _
           Or (Left(s, 1) = "'" And Right(s, 1) = "'") Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If
    Unquote = s
End Function


Private Sub SortByFirstKey(ByRef arr() As Variant, ByVal n As Long)
    Dim i As Long, j As Long
    Dim tmp As Variant
    For i = 1 To n - 1
        tmp = arr(i)
        j = i - 1
        Do While j >= 0
            If LCase(CStr(arr(j)(0))) <= LCase(CStr(tmp(0))) Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub


Private Sub WriteReport(ByVal findings As Collection)
    Dim sheetName As String: sheetName = "OKF Lint Report"

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Severity"
    ws.Cells(1, 2).Value = "File"
    ws.Cells(1, 3).Value = "Finding"
    ws.Rows(1).Font.Bold = True

    Dim row As Long: row = 2
    Dim item As Variant
    For Each item In findings
        ws.Cells(row, 1).Value = item(0)
        ws.Cells(row, 2).Value = item(1)
        ws.Cells(row, 3).Value = item(2)
        row = row + 1
    Next item

    Dim r As Long
    For r = 2 To row - 1
        Select Case LCase(ws.Cells(r, 1).Value)
            Case "error":   ws.Rows(r).Interior.Color = RGB(255, 200, 200)
            Case "warning": ws.Rows(r).Interior.Color = RGB(255, 235, 156)
        End Select
    Next r

    ws.Columns("A:C").AutoFit
    ws.Activate
End Sub


Private Function ReadUtf8(ByVal path As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText
    st.Close
End Function
