Attribute VB_Name = "OKFBootstrap"
' =====================================================================
'  OKF Bootstrap  - one-click bundle initialiser
'
'  BootstrapBundle creates the standard seed concepts in an empty (or
'  partial) bundle root. It is idempotent: files that already exist
'  are never overwritten.
'
'  Flow:
'    1. Require bundle root (same guard as ApplyOKFWrite).
'    2. Build seed list: (_foundation/00-operating-profile.md,
'       builds/example-build.md, skills/skill-md-authoring.md).
'    3. Filter to files that do NOT yet exist (create-if-absent).
'    4. Build a <VBA_WRITE> envelope from survivors.
'    5. Apply via OKFWriteApply.ApplyWriteEnvelopeText (shared path).
'    6. Regenerate indexes via OKFIndexGenerator.GenerateOKFIndexes.
'    7. Summary MsgBox.
'
'  Seed content lives in the private Seed* functions below so the
'  bulky strings stay out of the other modules.
'
'  Requires: OKFConfig, OKFWriteApply, OKFIndexGenerator.
' =====================================================================

Option Explicit

Public Sub BootstrapBundle()
    Dim root As String
    root = OKFConfig.BundleRoot()
    If root = "" Then
        MsgBox "Set Bundle Root first (button 5).", vbExclamation, "OKF Bootstrap"
        Exit Sub
    End If

    ' --- 1. Build the full seed list ---
    Dim seedPaths(0 To 2) As String
    Dim seedContents(0 To 2) As String

    seedPaths(0) = "_foundation/00-operating-profile.md"
    seedContents(0) = SeedOperatingProfile()

    seedPaths(1) = "builds/example-build.md"
    seedContents(1) = SeedExampleBuild()

    seedPaths(2) = "skills/skill-md-authoring.md"
    seedContents(2) = SeedSkillMdAuthoring()

    ' --- 2. Create-if-absent filter ---
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Build a root path with trailing separator for ResolvePath parity.
    Dim rootFwd As String
    rootFwd = root
    If Right(rootFwd, 1) <> "\" And Right(rootFwd, 1) <> "/" Then
        rootFwd = rootFwd & "\"
    End If

    Dim keepPaths(0 To 2) As String
    Dim keepContents(0 To 2) As String
    Dim keepCount As Long: keepCount = 0

    Dim i As Long
    Dim absPath As String
    For i = 0 To 2
        Dim rel As String
        rel = Replace(seedPaths(i), "/", "\")
        absPath = rootFwd & rel
        If Not fso.FileExists(absPath) Then
            keepPaths(keepCount) = seedPaths(i)
            keepContents(keepCount) = seedContents(i)
            keepCount = keepCount + 1
        End If
    Next i

    If keepCount = 0 Then
        MsgBox "Bundle already initialized - nothing to add.", vbInformation, "OKF Bootstrap"
        Exit Sub
    End If

    ' --- 3. Build <VBA_WRITE> envelope from survivors ---
    Dim envelope As String
    envelope = "<VBA_WRITE>" & vbLf

    For i = 0 To keepCount - 1
        envelope = envelope & "### FILE: " & keepPaths(i) & vbLf
        envelope = envelope & keepContents(i) & vbLf
        envelope = envelope & "### END FILE" & vbLf
    Next i

    envelope = envelope & "</VBA_WRITE>"

    ' --- 4. Apply via the shared write path ---
    Dim w As Long, s As Long
    If Not OKFWriteApply.ApplyWriteEnvelopeText(envelope, w, s) Then
        Exit Sub
    End If

    ' --- 5. Generate indexes (includes skills manifest) ---
    OKFIndexGenerator.GenerateOKFIndexes

    ' --- 6. Summary ---
    MsgBox "Bundle initialized: " & w & " file(s) created. Indexes generated.", _
           vbInformation, "OKF Bootstrap"
End Sub


' ---------------------------------------------------------------------------
'  Seed content
' ---------------------------------------------------------------------------

Private Function SeedOperatingProfile() As String
    Dim s As String
    s = "---" & vbLf
    s = s & "type: Foundation" & vbLf
    s = s & "title: Operating Profile" & vbLf
    s = s & "description: Who I am and what this bundle is for." & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "<!-- TODO: Describe who you are, your role, and what this bundle is for." & vbLf
    s = s & "     The LLM reads this file first, every session, to orient itself to" & vbLf
    s = s & "     your context. Include: your name/role, the domain this portfolio" & vbLf
    s = s & "     covers, key constraints or priorities the assistant should know, and" & vbLf
    s = s & "     any standing preferences (tone, output format, decision style). -->"
    SeedOperatingProfile = s
End Function


Private Function SeedExampleBuild() As String
    Dim s As String
    s = "---" & vbLf
    s = s & "type: Build" & vbLf
    s = s & "title: Example Build" & vbLf
    s = s & "description: Illustrates the full field set for a Build concept." & vbLf
    s = s & "status: parked" & vbLf
    s = s & "effort: M" & vbLf
    s = s & "impact: medium" & vbLf
    s = s & "last_touched: 2026-01-01" & vbLf
    s = s & "Dependencies:" & vbLf
    s = s & "  - skills/skill-md-authoring.md" & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "<!-- TODO: Replace this example with a real build idea." & vbLf
    s = s & "     Field notes:" & vbLf
    s = s & "     - status drives the builds/ index grouping:" & vbLf
    s = s & "         working | idea | parked | production | archived" & vbLf
    s = s & "     - status: parked keeps this out of the single 'working' WIP slot" & vbLf
    s = s & "       and avoids tripping the stall detector." & vbLf
    s = s & "     - Dependencies lists paths to other concepts this build relies on." & vbLf
    s = s & "     - last_touched (YYYY-MM-DD) is used by the stall detector." & vbLf
    s = s & "     - effort: XS | S | M | L | XL" & vbLf
    s = s & "     - impact: low | medium | high -->"
    SeedExampleBuild = s
End Function


Private Function SeedSkillMdAuthoring() As String
    Dim s As String
    s = "---" & vbLf
    s = s & "type: Skill" & vbLf
    s = s & "title: Skill - Markdown Authoring" & vbLf
    s = s & "description: Teaches the LLM how to author OKF-conformant" & vbLf
    s = s & "  concept files (Builds, Skills, Foundation) in this bundle." & vbLf
    s = s & "trigger: Use whenever the assistant needs to create or edit a" & vbLf
    s = s & "  concept .md file." & vbLf
    s = s & "---" & vbLf
    s = s & "" & vbLf
    s = s & "<!-- TODO: Paste the full skill-md-authoring.md body here." & vbLf
    s = s & "     This is the skill that lets the LLM author all other skills," & vbLf
    s = s & "     so it should be substantive, not left as a stub." & vbLf
    s = s & "     The canonical text lives in your existing bundle." & vbLf
    s = s & "     Steps:" & vbLf
    s = s & "       1. Open your bundle's skills/skill-md-authoring.md." & vbLf
    s = s & "       2. Copy its full body (below the closing --- of the frontmatter)." & vbLf
    s = s & "       3. Replace this comment block with that text." & vbLf
    s = s & "       4. Re-import OKFBootstrap.bas into the workbook. -->"
    SeedSkillMdAuthoring = s
End Function
