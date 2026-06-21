Attribute VB_Name = "OKFIndexGenerator"
' =====================================================================
'  OKF Index Generator  (conformant OKF v0.1 producer)
'
'  Walks a bundle tree and (re)writes an index.md in every directory,
'  per SPEC.md Sec.6:
'    - index.md files carry NO frontmatter ...
'    - ... EXCEPT the bundle-root index.md, which MAY carry exactly one
'      frontmatter key: okf_version (SPEC Sec.11).
'    - Body = sections grouping concepts under headings, each entry
'      "* [Title](relative-url) - description".
'    - Entries include the description from the linked concept's frontmatter.
'    - Reserved filenames (index.md, log.md) are never listed as concepts.
'
'  Grouping is inferred from the concepts' own frontmatter (not declared):
'    - The first field in GROUP_BY_CANDIDATES that any concept in the folder
'      carries wins as the grouping axis -> lifecycle board with headings.
'    - If no candidate field is present in any concept -> one flat alphabetical
'      list, no group headings (e.g. skills/ folder which carries no status).
'    - Group headings follow GROUP_ORDER, then any extras alpha.
'    - The STALL_GROUP ("working") sorts oldest-last_touched first and shows
'      the date inline, so stalled builds float to the top of the active section.
'    - Subdirectories listed under a "# Subdirectories" heading per Sec.6.
'    - UTF-8 in and out (SPEC Sec.4: concepts are UTF-8 markdown).
'
'  Requires (Tools -> References): Microsoft ActiveX Data Objects 2.x
'  (Scripting.FileSystemObject / Dictionary are late-bound below.)
' =====================================================================

Option Explicit

Private m_BundleRoot As String
Private Const OKF_VERSION As String = "0.1"

' --- Grouping candidates (EXTENSION SEAM) ---
' Ordered list of fields the generator can group by; first one PRESENT in a
' folder's concepts wins. None present -> flat alphabetical list, no headings.
' EXTENSION SEAM: append a field name (e.g. "domain") to support a new axis
' without per-folder config. See _meta/okf-roadmap.md before adding one.
Private Const GROUP_BY_CANDIDATES As String = "status"   ' extend: "status,domain"
Private Const STALL_GROUP As String = "working"  ' this group sorts oldest-last_touched first
Private Const GROUP_ORDER As String = "working,boilerplate,spec,idea,parked,production,archived"

Private fso As Object

Sub GenerateOKFIndexes()
    m_BundleRoot = OKFConfig.BundleRoot()
    If m_BundleRoot = "" Then
        MsgBox "Bundle root not set - click Set Bundle Root.", vbExclamation, "OKF Index Generator"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(m_BundleRoot) Then
        MsgBox "Bundle root not found: " & m_BundleRoot, vbCritical, "OKF Index Generator"
        Exit Sub
    End If

    Dim count As Long
    count = ProcessDir(fso.GetFolder(m_BundleRoot), True)

    MsgBox "Regenerated " & count & " index.md file(s).", vbInformation
End Sub


' Returns count of index files written (this dir + all descendants).
Private Function ProcessDir(ByVal folder As Object, ByVal isRoot As Boolean) As Long
    Dim written As Long: written = 0

    ' --- Phase 1: collect concept data into parallel arrays ---
    Const MAX_C As Long = 499
    Dim cNm(0 To MAX_C) As String
    Dim cTi(0 To MAX_C) As String
    Dim cDe(0 To MAX_C) As String
    Dim cSt(0 To MAX_C) As String
    Dim cLT(0 To MAX_C) As String
    Dim cc As Long: cc = 0

    Dim f As Object
    For Each f In folder.Files
        If IsConceptFile(f.Name) And cc <= MAX_C Then
            Dim cType As String, ti As String, de As String, st As String, lt As String
            ParseFrontmatter ReadUtf8(f.path), cType, ti, de, st, lt
            If ti = "" Then ti = BaseName(f.Name)   ' SPEC Sec.4.1: derive title from filename
            cNm(cc) = f.Name: cTi(cc) = ti: cDe(cc) = de
            cSt(cc) = st: cLT(cc) = lt
            cc = cc + 1
        End If
    Next f

    ' --- Phase 2: detect grouping field ---
    ' First candidate in GROUP_BY_CANDIDATES present in ANY concept wins.
    ' None present -> groupField = "" -> flat alphabetical list.
    Dim groupField As String: groupField = ""
    Dim candidates() As String: candidates = Split(GROUP_BY_CANDIDATES, ",")
    Dim ci As Long, k As Long
    Dim found As Boolean: found = False

    For ci = 0 To UBound(candidates)
        If found Then Exit For
        Dim cand As String: cand = Trim(candidates(ci))
        If cand <> "" Then
            For k = 0 To cc - 1
                Dim fv As String
                Select Case LCase(cand)
                    Case "status": fv = cSt(k)
                    Case Else: fv = ""   ' future: map additional fields here
                End Select
                If fv <> "" Then
                    groupField = cand
                    found = True
                    Exit For
                End If
            Next k
        End If
    Next ci

    ' --- subdirectories ---
    Dim subdirs As Collection
    Set subdirs = New Collection
    Dim d As Object
    For Each d In folder.SubFolders
        subdirs.Add d
    Next d

    ' --- Phase 3: assemble index.md body ---
    Dim sb As String: sb = ""

    If isRoot Then
        ' Only permitted frontmatter in any index.md (SPEC Sec.11).
        sb = "---" & vbLf & "okf_version: """ & OKF_VERSION & """" & vbLf & "---" & vbLf & vbLf
    End If

    If cc > 0 Then
        If groupField = "" Then
            ' Flat alphabetical list -- no group headings.
            Dim flatColl As Collection
            Set flatColl = New Collection
            For k = 0 To cc - 1
                Dim fLine As String
                fLine = "* [" & cTi(k) & "](" & cNm(k) & ")"
                If cDe(k) <> "" Then fLine = fLine & " - " & cDe(k)
                flatColl.Add Array(LCase(cTi(k)), fLine)
            Next k
            Dim flatLines() As String
            flatLines = SortedEntryLines(flatColl)
            Dim fl As Long
            For fl = 0 To UBound(flatLines)
                sb = sb & flatLines(fl) & vbLf
            Next fl
            sb = sb & vbLf

        Else
            ' Grouped under headings (lifecycle-board behavior).
            Dim byGroup As Object
            Set byGroup = CreateObject("Scripting.Dictionary")
            byGroup.CompareMode = vbTextCompare

            For k = 0 To cc - 1
                Dim grpVal As String
                Select Case LCase(groupField)
                    Case "status": grpVal = cSt(k)
                    Case Else: grpVal = ""
                End Select
                If grpVal = "" Then grpVal = "(unset)"

                Dim gLine As String
                gLine = "* [" & cTi(k) & "](" & cNm(k) & ")"
                If cDe(k) <> "" Then gLine = gLine & " - " & cDe(k)

                Dim sortKey As String
                If LCase(grpVal) = LCase(STALL_GROUP) Then
                    ' show staleness inline; oldest (or never-touched) floats to top
                    gLine = gLine & "  _(last touched " & IIf(cLT(k) = "", "never", cLT(k)) & ")_"
                    sortKey = IIf(cLT(k) = "", "0000-00-00", cLT(k))
                Else
                    sortKey = LCase(cTi(k))
                End If

                If Not byGroup.Exists(grpVal) Then byGroup.Add grpVal, New Collection
                byGroup(grpVal).Add Array(sortKey, gLine)
            Next k

            Dim groupKeys() As String
            groupKeys = OrderedGroupKeys(byGroup)
            Dim gi As Long
            For gi = 0 To UBound(groupKeys)
                If groupKeys(gi) <> "" Then
                    sb = sb & "# " & groupKeys(gi) & vbLf
                    Dim gEntries() As String
                    gEntries = SortedEntryLines(byGroup(groupKeys(gi)))
                    Dim ge As Long
                    For ge = 0 To UBound(gEntries)
                        sb = sb & gEntries(ge) & vbLf
                    Next ge
                    sb = sb & vbLf
                End If
            Next gi
        End If
    End If

    ' Subdirectories section.
    If subdirs.Count > 0 Then
        sb = sb & "# Subdirectories" & vbLf
        Dim sdNames() As String
        sdNames = SubfolderNamesSorted(subdirs)
        Dim s As Long
        For s = 0 To UBound(sdNames)
            sb = sb & "* [" & sdNames(s) & "](" & sdNames(s) & "/)" & vbLf
        Next s
        sb = sb & vbLf
    End If

    ' Write index.md (overwrite). Skip writing an empty index in an empty dir.
    If Len(Trim(sb)) > 0 Then
        WriteUtf8 fso.BuildPath(folder.path, "index.md"), sb
        written = written + 1
    End If

    ' Recurse.
    For Each d In folder.SubFolders
        written = written + ProcessDir(d, False)
    Next d

    ProcessDir = written
End Function


' --- A non-reserved markdown file is a concept (SPEC Sec.3.1). ---
Private Function IsConceptFile(ByVal name As String) As Boolean
    Dim ln As String
    ln = LCase(name)
    IsConceptFile = (Right(ln, 3) = ".md") And (ln <> "index.md") And (ln <> "log.md")
End Function


' --- Lightweight frontmatter parse: top-level scalars. ---
Private Sub ParseFrontmatter(ByVal content As String, ByRef cType As String, _
                             ByRef cTitle As String, ByRef cDesc As String, _
                             ByRef cStatus As String, ByRef cLastTouched As String)
    cType = "": cTitle = "": cDesc = "": cStatus = "": cLastTouched = ""

    content = Replace(Replace(content, vbCrLf, vbLf), vbCr, vbLf)
    Dim lines() As String
    lines = Split(content, vbLf)
    If UBound(lines) < 1 Then Exit Sub
    If Trim(lines(0)) <> "---" Then Exit Sub          ' no frontmatter block

    Dim i As Long
    For i = 1 To UBound(lines)
        If Trim(lines(i)) = "---" Then Exit Sub        ' end of frontmatter

        Dim colon As Long
        colon = InStr(lines(i), ":")                   ' split on FIRST colon only
        If colon > 0 Then
            Dim key As String, val As String
            key = LCase(Trim(Left(lines(i), colon - 1)))
            val = Unquote(Trim(Mid(lines(i), colon + 1)))
            Select Case key
                Case "type":         cType = val
                Case "title":        cTitle = val
                Case "description":  cDesc = val
                Case "status":       cStatus = val
                Case "last_touched": cLastTouched = val
            End Select
        End If
    Next i
End Sub


Private Function Unquote(ByVal s As String) As String
    If Len(s) >= 2 Then
        If (Left(s, 1) = """" And Right(s, 1) = """") _
           Or (Left(s, 1) = "'" And Right(s, 1) = "'") Then
            s = Mid(s, 2, Len(s) - 2)
        End If
    End If
    Unquote = s
End Function


Private Function BaseName(ByVal fileName As String) As String
    Dim s As String
    s = fileName
    If LCase(Right(s, 3)) = ".md" Then s = Left(s, Len(s) - 3)
    BaseName = s
End Function


' --- Sorting helpers (insertion sort; bundles are small, clarity > speed). ---

' Group headings: GROUP_ORDER tokens that exist (in that order), then any
' remaining groups alphabetically.
Private Function OrderedGroupKeys(ByVal dict As Object) As String()
    Dim result() As String
    ReDim result(0 To dict.count - 1)
    Dim cnt As Long: cnt = 0

    Dim used As Object: Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare

    Dim toks() As String: toks = Split(GROUP_ORDER, ",")
    Dim t As Long, tok As String
    For t = 0 To UBound(toks)
        tok = Trim(toks(t))
        If tok <> "" Then
            If dict.Exists(tok) And Not used.Exists(tok) Then
                result(cnt) = tok: cnt = cnt + 1
                used.Add tok, True
            End If
        End If
    Next t

    Dim leftover() As String, rc As Long: rc = 0
    ReDim leftover(0 To dict.count)
    Dim kk As Variant
    For Each kk In dict.Keys
        If Not used.Exists(CStr(kk)) Then leftover(rc) = CStr(kk): rc = rc + 1
    Next kk
    If rc > 0 Then
        ReDim Preserve leftover(0 To rc - 1)
        SortStringArray leftover
        Dim m As Long
        For m = 0 To rc - 1: result(cnt) = leftover(m): cnt = cnt + 1: Next m
    End If

    OrderedGroupKeys = result
End Function

' Entries arrive as Array(sortKey, line); sort by sortKey ascending, return lines.
Private Function SortedEntryLines(ByVal col As Collection) As String()
    Dim n As Long: n = col.count
    Dim keys() As String, lines() As String
    ReDim keys(0 To n - 1): ReDim lines(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        keys(i - 1) = col(i)(0)
        lines(i - 1) = col(i)(1)
    Next i

    Dim j As Long, tk As String, tl As String
    For i = 1 To n - 1
        tk = keys(i): tl = lines(i): j = i - 1
        Do While j >= 0
            If LCase(keys(j)) <= LCase(tk) Then Exit Do
            keys(j + 1) = keys(j): lines(j + 1) = lines(j): j = j - 1
        Loop
        keys(j + 1) = tk: lines(j + 1) = tl
    Next i

    SortedEntryLines = lines
End Function

Private Function SubfolderNamesSorted(ByVal col As Collection) As String()
    Dim arr() As String
    ReDim arr(0 To col.count - 1)
    Dim i As Long
    For i = 1 To col.count: arr(i - 1) = col(i).name: Next i
    SortStringArray arr
    SubfolderNamesSorted = arr
End Function

Private Sub SortStringArray(ByRef arr() As String)
    Dim i As Long, j As Long, tmp As String
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= LBound(arr)
            If LCase(arr(j)) <= LCase(tmp) Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
End Sub


' --- UTF-8 I/O (Private: no clash if imported alongside OKFContextBundle). ---
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
    st.SaveToFile path, 2          ' overwrite
    st.Close
End Sub
