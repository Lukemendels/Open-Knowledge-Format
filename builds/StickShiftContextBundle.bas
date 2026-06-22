Attribute VB_Name = "StickShiftContextBundle"
' =====================================================================
'  StickShift Context Bundle -- OKF-compliant
'  (conformant OKF v0.1 bundle assembler)
'
'  Entry point: Sub BuildContextBundle()
'
'  Reads a <CONTEXT_REQUEST> from the clipboard, resolves it against
'  the bundle (graph traversal included), assembles an anchored bundle,
'  and writes it to a stable StickShift-context.md file in the -dist
'  sibling folder of the bundle root.
'
'  Modes:
'    index  - Hop-1 opener: foundation + /index.md + /skills/index.md.
'    bundle - BFS expansion from seed paths at configurable depth and
'             direction (outbound links or inbound backlinks), with
'             optional via: heading scoping for outbound.
'
'  Default (no CONTEXT_REQUEST in clipboard): mode index.
'
'  Requires (Tools -> References): Microsoft ActiveX Data Objects 2.x
'  (Scripting.FileSystemObject / Dictionary / VBScript.RegExp are
'  late-bound; no other references needed.)
'
'  Module-level dependencies: StickShiftConfig (BundleRoot, DistDir),
'  StickShiftClipboard (GetClipboardText).  All leaf helpers (UTF-8 I/O,
'  link parsing, etc.) are private copies in this module.
' =====================================================================

Option Explicit

Private Const OKF_VERSION    As String = "0.1"
Private Const FOUNDATION_DIR As String = "_foundation"
Private Const SKILLS_INDEX   As String = "skills/index.md"
Private Const OUT_FILENAME   As String = "StickShift-context.md"

Private fso As Object


' -- Entry point ------------------------------------------------------------------

Sub BuildContextBundle()

    ' 1. Bundle root via StickShiftConfig (prompts picker if not yet set).
    Dim root As String
    root = StickShiftConfig.BundleRoot()
    If root = "" Then
        MsgBox "Bundle root not set - click Switch Context.", vbExclamation, "StickShift"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(root) Then
        MsgBox "Bundle root not found: " & root, vbCritical, "StickShift"
        Exit Sub
    End If

    ' 2. Read clipboard; normalise line endings.
    Dim clip As String
    On Error Resume Next
    clip = ReadClipboard()
    On Error GoTo 0
    clip = Replace(Replace(clip, vbCrLf, vbLf), vbCr, vbLf)

    ' 3. Detect CONTEXT_REQUEST block; default to index mode if absent.
    Dim startPos As Long: startPos = InStr(clip, "<CONTEXT_REQUEST>")
    Dim endPos   As Long: endPos = InStr(clip, "</CONTEXT_REQUEST>")

    Dim mode      As String: mode = "index"
    Dim depth     As Long:   depth = 1
    Dim direction As String: direction = "outbound"
    Dim viaH      As String: viaH = ""
    Dim seeds(0 To 199) As String
    Dim seedCount As Long: seedCount = 0

    If startPos > 0 And endPos > startPos Then
        Dim blockBody As String
        blockBody = Mid(clip, startPos + Len("<CONTEXT_REQUEST>"), _
                        endPos - (startPos + Len("<CONTEXT_REQUEST>")))
        Dim blockLines() As String: blockLines = Split(blockBody, vbLf)

        Dim inInclude As Boolean: inInclude = False
        Dim bl As Long
        For bl = 0 To UBound(blockLines)
            Dim bline As String: bline = Trim(blockLines(bl))
            If Left(bline, 6) = "mode: " Then
                mode = Trim(Mid(bline, 7)): inInclude = False
            ElseIf Left(bline, 7) = "depth: " Then
                On Error Resume Next
                depth = CLng(Trim(Mid(bline, 8)))
                On Error GoTo 0
                inInclude = False
            ElseIf Left(bline, 11) = "direction: " Then
                direction = Trim(Mid(bline, 12)): inInclude = False
            ElseIf Left(bline, 5) = "via: " Then
                viaH = Trim(Mid(bline, 6)): inInclude = False
            ElseIf bline = "include:" Then
                inInclude = True
            ElseIf inInclude And Left(bline, 2) = "- " Then
                If seedCount <= UBound(seeds) Then
                    seeds(seedCount) = Trim(Mid(bline, 3))
                    seedCount = seedCount + 1
                End If
            ElseIf bline <> "" And Left(bline, 1) <> "-" Then
                inInclude = False
            End If
        Next bl
    End If

    ' 4. Dispatch to index or bundle assembly.
    Dim foundCount As Long: foundCount = 0
    Dim mapCount   As Long: mapCount = 0
    Dim selCount   As Long: selCount = 0
    Dim assembled  As String

    If mode = "index" Then
        assembled = AssembleIndex(root, foundCount, mapCount, selCount)
    Else
        If seedCount = 0 Then
            MsgBox "No include: paths found in request.", vbExclamation, "StickShift"
            Exit Sub
        End If
        Dim seedSlice() As String
        ReDim seedSlice(0 To seedCount - 1)
        Dim si As Long
        For si = 0 To seedCount - 1: seedSlice(si) = seeds(si): Next si
        assembled = AssembleBundle(root, seedSlice, depth, direction, viaH, _
                                   foundCount, mapCount, selCount)
    End If

    ' 5. Build the OKF-CONTEXT-BUNDLE header.
    Dim totalConcepts As Long: totalConcepts = foundCount + mapCount + selCount
    Dim approxTokens  As Long: approxTokens = CLng(Len(assembled)) \ 4
    Dim ts As String
    ts = Format(Now(), "yyyy-mm-dd") & "T" & Format(Now(), "hh:mm:ss") & "Z"

    Dim header As String
    header = "<!-- OKF-CONTEXT-BUNDLE" & vbLf & _
             "mode: " & mode & vbLf & _
             "okf_version: " & OKF_VERSION & vbLf & _
             "assembled: " & ts & vbLf & _
             "concepts: " & totalConcepts & " (" & foundCount & " foundation, " & _
                 mapCount & " map, " & selCount & " selected)" & vbLf & _
             "approx_tokens: " & approxTokens & vbLf & _
             "-->" & vbLf & vbLf

    Dim fullOutput As String: fullOutput = header & assembled

    ' 6. Write to -dist sibling folder (never inside the bundle root).
    Dim DistDir As String: DistDir = StickShiftConfig.DistDir()
    If DistDir = "" Then
        MsgBox "Bundle root not set - cannot determine output directory.", _
               vbCritical, "StickShift"
        Exit Sub
    End If

    Dim outPath As String: outPath = DistDir & OUT_FILENAME
    WriteUtf8 outPath, fullOutput

    ' 7. Open/focus the dist folder in Explorer (reuse existing window).
    Dim charCount  As Long: charCount = Len(fullOutput)
    Dim tokenCount As Long: tokenCount = charCount \ 4

    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    Dim distNorm As String
    distNorm = LCase(TrimTrailSlash(DistDir))
    Dim openWin As Boolean: openWin = True
    Dim w As Object
    For Each w In shellApp.Windows
        Dim wFolderPath As String: wFolderPath = ""
        On Error Resume Next
        wFolderPath = LCase(TrimTrailSlash(w.Document.folder.Self.path))
        On Error GoTo 0
        If wFolderPath = distNorm Then
            w.Visible = True
            On Error Resume Next
            AppActivate w.LocationName
            On Error GoTo 0
            openWin = False
            Exit For
        End If
    Next w
    If openWin Then Shell "explorer.exe """ & DistDir & """", vbNormalFocus

    MsgBox "Context bundle written:" & vbLf & outPath & vbLf & vbLf & _
           charCount & " chars   ~" & tokenCount & " tokens   " & _
           totalConcepts & " concepts", _
           vbInformation, "StickShift"
End Sub


' -- Mode: index ------------------------------------------------------------------

Private Function AssembleIndex(ByVal root As String, _
                                ByRef foundCount As Long, _
                                ByRef mapCount As Long, _
                                ByRef selCount As Long) As String
    Dim sb As String: sb = ""

    ' 1. All concept files under _foundation/, sorted by filename ascending.
    Dim foundDir As String: foundDir = root & FOUNDATION_DIR
    If fso.FolderExists(foundDir) Then
        Dim foundFiles() As String
        Dim ffc As Long: ffc = 0
        ReDim foundFiles(0 To 99)
        CollectConceptsRecursive fso.GetFolder(foundDir), root, foundFiles, ffc
        If ffc > 0 Then
            ReDim Preserve foundFiles(0 To ffc - 1)
            InsertionSortByFilename foundFiles
            Dim i As Long
            For i = 0 To ffc - 1
                Dim fContent As String: fContent = ReadUtf8(root & Replace(foundFiles(i), "/", "\"))
                sb = sb & MakeAnchor(foundFiles(i), fContent, "foundation")
                foundCount = foundCount + 1
            Next i
        End If
    End If

    ' 2. Root index.md.
    Dim rootIdx As String: rootIdx = root & "index.md"
    If fso.FileExists(rootIdx) Then
        sb = sb & MakeAnchor("index.md", ReadUtf8(rootIdx), "map")
        mapCount = mapCount + 1
    End If

    ' 3. skills/index.md.
    Dim skillsIdx As String: skillsIdx = root & Replace(SKILLS_INDEX, "/", "\")
    If fso.FileExists(skillsIdx) Then
        sb = sb & MakeAnchor(SKILLS_INDEX, ReadUtf8(skillsIdx), "map")
        mapCount = mapCount + 1
    End If

    AssembleIndex = sb
End Function


' -- Mode: bundle (BFS graph expansion) ------------------------------------------

Private Function AssembleBundle(ByVal root As String, _
                                 ByRef seeds() As String, _
                                 ByVal depth As Long, _
                                 ByVal direction As String, _
                                 ByVal viaHeading As String, _
                                 ByRef foundCount As Long, _
                                 ByRef mapCount As Long, _
                                 ByRef selCount As Long) As String

    ' visited preserves insertion order (keys iterated in Add order).
    Dim visited As Object
    Set visited = CreateObject("Scripting.Dictionary")
    visited.CompareMode = vbTextCompare

    ' Initialise frontier with seeds that actually exist in the bundle.
    Dim frontier() As String
    ReDim frontier(0 To UBound(seeds) + 1)
    Dim fi As Long: fi = 0

    Dim s As Long
    For s = 0 To UBound(seeds)
        Dim seedRel As String: seedRel = NormalizeRelPath(seeds(s))
        If Not visited.Exists(seedRel) Then
            If fso.FileExists(root & Replace(seedRel, "/", "\")) Then
                visited.Add seedRel, True
                frontier(fi) = seedRel: fi = fi + 1
            End If
            ' Absent seeds skipped silently per OKF Sec. 9.
        End If
    Next s

    If fi = 0 Then AssembleBundle = "": Exit Function
    ReDim Preserve frontier(0 To fi - 1)

    ' BFS expansion up to 'depth' layers.
    Dim depthIdx As Long
    For depthIdx = 0 To depth - 1
        If fi = 0 Then Exit For

        Dim nextFrontier() As String
        ReDim nextFrontier(0 To 999)
        Dim nfi As Long: nfi = 0

        If LCase(direction) = "outbound" Then
            Dim fw As Long
            For fw = 0 To fi - 1
                Dim fromRel As String: fromRel = frontier(fw)
                Dim fromAbs As String: fromAbs = root & Replace(fromRel, "/", "\")
                If Not fso.FileExists(fromAbs) Then GoTo NextFrontierItem

                Dim fLinks As Object
                Set fLinks = ExtractLinksScoped(ReadUtf8(fromAbs), viaHeading)

                Dim lnk As Variant
                For Each lnk In fLinks.keys
                    Dim resolved As String
                    resolved = ResolveLinkToRel(CStr(lnk), fromRel, root)
                    If resolved <> "" Then
                        If Not visited.Exists(resolved) Then
                            visited.Add resolved, True
                            If nfi > UBound(nextFrontier) Then _
                                ReDim Preserve nextFrontier(0 To nfi + 999)
                            nextFrontier(nfi) = resolved: nfi = nfi + 1
                        End If
                    End If
                Next lnk
NextFrontierItem:
            Next fw

        ElseIf LCase(direction) = "inbound" Then
            ' Build a lookup set from the current frontier.
            Dim frontSet As Object
            Set frontSet = CreateObject("Scripting.Dictionary")
            frontSet.CompareMode = vbTextCompare
            Dim fsx As Long
            For fsx = 0 To fi - 1
                If Not frontSet.Exists(frontier(fsx)) Then _
                    frontSet.Add frontier(fsx), True
            Next fsx

            ' Scan every link source in the bundle for links to frontier members.
            ' Uses CollectLinkSourcesRecursive (includes index.md, excludes log.md).
            Dim allConcepts() As String
            Dim allCount As Long: allCount = 0
            ReDim allConcepts(0 To 999)
            CollectLinkSourcesRecursive fso.GetFolder(root), root, allConcepts, allCount

            Dim ac As Long
            For ac = 0 To allCount - 1
                Dim acRel As String: acRel = allConcepts(ac)
                If Not visited.Exists(acRel) Then
                    Dim acAbs As String: acAbs = root & Replace(acRel, "/", "\")
                    If fso.FileExists(acAbs) Then
                        Dim acLinks As Object
                        Set acLinks = ExtractLinksScoped(ReadUtf8(acAbs), "")

                        ' Copy keys to a local array to avoid error 10 on For Each.
                        Dim acKeys As Variant
                        Dim acLnk As Variant
                        Dim kIdx As Long

                        If Not acLinks Is Nothing Then
                            acKeys = acLinks.keys

                            If IsArray(acKeys) Then
                                For kIdx = LBound(acKeys) To UBound(acKeys)
                                    acLnk = acKeys(kIdx)

                                    Dim acResolved As String
                                    acResolved = ResolveLinkToRel(CStr(acLnk), acRel, root)
                                    If acResolved <> "" Then
                                        If frontSet.Exists(acResolved) Then
                                            visited.Add acRel, True
                                            If nfi > UBound(nextFrontier) Then _
                                                ReDim Preserve nextFrontier(0 To nfi + 999)
                                            nextFrontier(nfi) = acRel: nfi = nfi + 1
                                            GoTo NextConcept
                                        End If
                                    End If
                                Next kIdx
                            End If
                        End If
                    End If
                End If
NextConcept:
            Next ac
        End If


        If nfi > 0 Then
            ReDim Preserve nextFrontier(0 To nfi - 1)
            frontier = nextFrontier: fi = nfi
        Else
            fi = 0
        End If
    Next depthIdx

    ' Assemble all visited concepts in BFS insertion order.
    Dim sb As String: sb = ""
    Dim k As Variant
    For Each k In visited.keys
        Dim kRel As String: kRel = CStr(k)
        Dim kAbs As String: kAbs = root & Replace(kRel, "/", "\")
        If fso.FileExists(kAbs) Then
            Dim kLayer As String: kLayer = GetLayer(kRel)
            sb = sb & MakeAnchor(kRel, ReadUtf8(kAbs), kLayer)
            Select Case kLayer
                Case "foundation": foundCount = foundCount + 1
                Case "map":        mapCount = mapCount + 1
                Case Else:         selCount = selCount + 1
            End Select
        End If
    Next k

    AssembleBundle = sb
End Function


' -- Helpers - file collection ---------------------------------------------------

' Collect all concept files under folder, root-relative with forward slashes.
Private Sub CollectConceptsRecursive(ByVal folder As Object, ByVal root As String, _
                                      ByRef files() As String, ByRef count As Long)
    Dim f As Object
    For Each f In folder.files
        If IsConceptFile(f.name) Then
            Dim rel As String: rel = Mid(f.path, Len(root) + 1)
            rel = Replace(rel, "\", "/")
            If count > UBound(files) Then ReDim Preserve files(0 To count + 999)
            files(count) = rel: count = count + 1
        End If
    Next f
    Dim d As Object
    For Each d In folder.SubFolders
        CollectConceptsRecursive d, root, files, count
    Next d
End Sub

' Collect all .md files except log.md (includes index.md) for inbound link scanning.
Private Sub CollectLinkSourcesRecursive(ByVal folder As Object, ByVal root As String, _
                                         ByRef files() As String, ByRef count As Long)
    Dim f As Object
    For Each f In folder.files
        If IsLinkSourceFile(f.name) Then
            Dim rel As String: rel = Mid(f.path, Len(root) + 1)
            rel = Replace(rel, "\", "/")
            If count > UBound(files) Then ReDim Preserve files(0 To count + 999)
            files(count) = rel: count = count + 1
        End If
    Next f
    Dim d As Object
    For Each d In folder.SubFolders
        CollectLinkSourcesRecursive d, root, files, count
    Next d
End Sub


' Insertion sort: primary key = filename (leaf), secondary = full path.
Private Sub InsertionSortByFilename(ByRef arr() As String)
    Dim i As Long, j As Long, tmp As String, tmpKey As String
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i): tmpKey = LCase(LeafName(tmp))
        j = i - 1
        Do While j >= LBound(arr)
            If LCase(LeafName(arr(j))) <= tmpKey Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub

Private Function LeafName(ByVal p As String) As String
    Dim i As Long, last As Long: last = 0
    For i = Len(p) To 1 Step -1
        If Mid(p, i, 1) = "\" Or Mid(p, i, 1) = "/" Then last = i: Exit For
    Next i
    If last > 0 Then LeafName = Mid(p, last + 1) Else LeafName = p
End Function


' -- Helpers - link extraction and resolution ------------------------------------

' Extract .md links, optionally scoped to a named heading's section.
Private Function ExtractLinksScoped(ByVal content As String, _
                                     ByVal viaHeading As String) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    content = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)

    Dim scanContent As String
    If viaHeading = "" Then
        scanContent = content
    Else
        scanContent = ExtractHeadingSection(content, viaHeading)
        If scanContent = "" Then Set ExtractLinksScoped = result: Exit Function
    End If

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\]\(([^)]+)\)"
    re.Global = True

    On Error Resume Next
    Dim matches As Object: Set matches = re.Execute(scanContent)
    On Error GoTo 0
    If matches Is Nothing Then Set ExtractLinksScoped = result: Exit Function

    Dim m As Object
    For Each m In matches
        Dim url As String: url = Trim(m.SubMatches(0))
        If Not (Left(LCase(url), 7) = "http://" Or Left(LCase(url), 8) = "https://") Then
            Dim base As String
            Dim h As Long: h = InStr(url, "#")
            If h > 0 Then base = Left(url, h - 1) Else base = url
            base = Trim(base)
            If LCase(Right(base, 3)) = ".md" And base <> "" Then
                If Not result.Exists(url) Then result.Add url, True
            End If
        End If
    Next m

    Set ExtractLinksScoped = result
End Function


' Return the text content under a named heading until the next same-or-higher heading.
Private Function ExtractHeadingSection(ByVal content As String, _
                                        ByVal headingName As String) As String
    Dim lines() As String: lines = Split(content, vbLf)
    Dim hStart As Long: hStart = -1
    Dim hLevel As Long: hLevel = 0

    Dim i As Long
    For i = 0 To UBound(lines)
        Dim lvl As Long: lvl = HeadingLevel(lines(i))
        If lvl > 0 Then
            If LCase(Trim(HeadingText(lines(i)))) = LCase(Trim(headingName)) Then
                hStart = i: hLevel = lvl: Exit For
            End If
        End If
    Next i

    If hStart = -1 Then ExtractHeadingSection = "": Exit Function

    Dim sb As String: sb = ""
    For i = hStart + 1 To UBound(lines)
        Dim cl As Long: cl = HeadingLevel(lines(i))
        If cl > 0 And cl <= hLevel Then Exit For
        sb = sb & lines(i) & vbLf
    Next i
    ExtractHeadingSection = sb
End Function

Private Function HeadingLevel(ByVal line As String) As Long
    Dim s As String: s = Trim(line)
    If Left(s, 1) <> "#" Then HeadingLevel = 0: Exit Function
    Dim i As Long: i = 1
    Do While i <= Len(s) And Mid(s, i, 1) = "#": i = i + 1: Loop
    HeadingLevel = i - 1
End Function

Private Function HeadingText(ByVal line As String) As String
    Dim s As String: s = Trim(line)
    Dim i As Long: i = 1
    Do While i <= Len(s) And Mid(s, i, 1) = "#": i = i + 1: Loop
    HeadingText = Trim(Mid(s, i))
End Function


' Resolve a markdown link (relative or /-rooted) to a root-relative forward-slash path.
' Returns "" if the link is external, non-.md, or the target file doesn't exist.
Private Function ResolveLinkToRel(ByVal link As String, _
                                   ByVal fromRel As String, _
                                   ByVal root As String) As String
    Dim h As Long: h = InStr(link, "#")
    If h > 0 Then link = Left(link, h - 1)
    link = Trim(link)

    If link = "" Then ResolveLinkToRel = "": Exit Function
    If Left(LCase(link), 7) = "http://" Or Left(LCase(link), 8) = "https://" Then _
        ResolveLinkToRel = "": Exit Function
    If LCase(Right(link, 3)) <> ".md" Then ResolveLinkToRel = "": Exit Function

    Dim absPath As String
    If Left(link, 1) = "/" Then
        absPath = root & Replace(Mid(link, 2), "/", "\")
    Else
        Dim fromDir As String: fromDir = ""
        Dim p As Long
        For p = Len(fromRel) To 1 Step -1
            If Mid(fromRel, p, 1) = "\" Or Mid(fromRel, p, 1) = "/" Then
                fromDir = Replace(Left(fromRel, p), "/", "\"): Exit For
            End If
        Next p
        absPath = root & fromDir & Replace(link, "/", "\")
    End If

    Do While InStr(absPath, "\\") > 0: absPath = Replace(absPath, "\\", "\"): Loop

    If Not fso.FileExists(absPath) Then ResolveLinkToRel = "": Exit Function

    If Left(LCase(absPath), Len(LCase(root))) = LCase(root) Then
        ResolveLinkToRel = Replace(Mid(absPath, Len(root) + 1), "\", "/")
    Else
        ResolveLinkToRel = ""
    End If
End Function


' -- Helpers - anchor and path ---------------------------------------------------

Private Function MakeAnchor(ByVal relPath As String, ByVal content As String, _
                              ByVal layer As String) As String
    Dim p As String: p = Replace(relPath, "\", "/")
    Dim a As String
    a = "<!-- OKF:BEGIN concept=" & p & " layer=" & layer & " -->" & vbLf
    a = a & content
    If Len(content) = 0 Or Right(content, 1) <> vbLf Then a = a & vbLf
    a = a & "<!-- OKF:END concept=" & p & " -->" & vbLf & vbLf
    MakeAnchor = a
End Function

Private Function GetLayer(ByVal relPath As String) As String
    Dim p As String: p = Replace(relPath, "\", "/")
    If Left(LCase(p), Len(FOUNDATION_DIR) + 1) = LCase(FOUNDATION_DIR) & "/" Then
        GetLayer = "foundation"
    ElseIf LCase(LeafName(p)) = "index.md" Then
        GetLayer = "map"
    Else
        GetLayer = "selected"
    End If
End Function

Private Function NormalizeRelPath(ByVal p As String) As String
    p = Replace(p, "\", "/")
    If Left(p, 1) = "/" Then p = Mid(p, 2)
    NormalizeRelPath = p
End Function

Private Function IsConceptFile(ByVal name As String) As Boolean
    Dim ln As String: ln = LCase(name)
    IsConceptFile = (Right(ln, 3) = ".md") And (ln <> "index.md") And (ln <> "log.md")
End Function

Private Function IsLinkSourceFile(ByVal name As String) As Boolean
    Dim ln As String: ln = LCase(name)
    IsLinkSourceFile = (Right(ln, 3) = ".md") And (ln <> "log.md")
End Function

Private Function TrimTrailSlash(ByVal s As String) As String
    If Len(s) > 0 And Right(s, 1) = "\" Then
        TrimTrailSlash = Left(s, Len(s) - 1)
    Else
        TrimTrailSlash = s
    End If
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


' -- Helpers - I/O ---------------------------------------------------------------

Private Function ReadClipboard() As String
    On Error GoTo FailSafe

    ReadClipboard = StickShiftClipboard.GetClipboardText()
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
