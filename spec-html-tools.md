---
type: Spec
title: HTML Tools — local browser tools that slot into StickShift
description: Lets a self-contained HTML tool (e.g. a JS port of AutoReviewer) ship with StickShift, install via a cockpit macro, and open via a cockpit button that reads an <HTML_OPEN> block from the clipboard, launches the local tool, and loads the companion skill into -dist. One block, one button, two effects.
status: spec
okf_version: "0.1"
---

# HTML Tools

## 0. Model + the resolved constraint

DHSChat is **GPT-5.1**; the skill and prompt below use explicit imperative steps and a literal
output contract.

The open question from the earlier draft is now answered: **a local HTML file on the C: drive
cannot be opened from a DHSChat hyperlink or a browser favorite** (the chat UI will not navigate
a `file://` link). **A VBA macro can open it.** So delivery is a cockpit button that launches the
tool in the operator's default browser (via the Win32 `ShellExecute` "open" verb, which routes
through the OS file association - not `explorer.exe`, which can open a file window instead), not a
clickable link. That collapses the earlier four-branch link
decision to a single path and, as a bonus, removes the portability problem: the skill only needs
the tool's **filename**, and the button resolves it against the local `-html` folder, so the same
skill works on any machine (Paul's button resolves against Paul's folder).

## 1. Goal

Make a self-contained browser tool a first-class StickShift citizen:

- It **ships** alongside StickShift (a single `.html` file, no server, no build step).
- It **installs** into the active context with one cockpit click.
- It **opens** with one cockpit click that also loads its companion skill into `-dist`.

First instance: the **AutoReviewer JS port** - tracked changes are OOXML markup inside the
`.docx` zip, so JavaScript can manufacture them client-side with no Word/COM. A standalone
`file://` page takes a dropped `.docx`, rewrites the zip, and triggers a download, entirely in
the browser.

## 2. Two firm constraints

1. **HTML never flows through the token stream.** Install is a local file-copy via the cockpit,
   not a chat round-trip; pasting a multi-KB tool through `<VBA_WRITE>` would re-create the exact
   friction we are removing.
2. **Open is a VBA launch, not a link.** Section 0.

## 3. HTML home: an `-html` sibling

Store installed HTML in a **sibling** of the context root, parallel to the existing `-dist`
sibling:

```
C:\StickShift\          <- context root (BundleRoot)
C:\StickShift-dist\     <- assembled StickShift-context.md (existing)
C:\StickShift-html\     <- installed HTML tools (new)
```

A sibling has **zero indexer/bundler footprint** (both only walk `BundleRoot`). Add `HtmlDir()`
to `StickShiftConfig`, a near-copy of `DistDir()`:

```vba
Public Function HtmlDir() As String
    Dim root As String
    root = BundleRoot()
    If root = "" Then HtmlDir = "": Exit Function

    Dim stripped As String: stripped = Left(root, Len(root) - 1)   ' drop trailing "\"
    Dim h As String:        h = stripped & "-html\"

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(h) Then
        On Error Resume Next
        fso.CreateFolder h
        On Error GoTo 0
    End If
    HtmlDir = h
End Function
```

## 4. Tool packaging: one self-contained `.html` with an embedded skill

A compliant tool is a **single self-contained `.html`** that embeds its own companion skill (and
a StickShift onboarding panel). The full authoring standard is `_meta/spec-html-tool-compliance.md`;
the host side only needs to know two things about the file:

```
builds/html-tools/autoreviewer/
  autoreviewer.html          <- the self-contained tool; embeds its skill + onboarding panel
```

1. **The skill is embedded**, verbatim, in a non-rendering block the install macro can read:
   `<script type="text/markdown" id="stickshift-skill" data-skill-slug="autoreviewer-tool">...</script>`.
   It is a real StickShift skill (internal format: `type: Skill`, `title`, `description`, `tags`;
   no `status`, so it files flat-alphabetical in `skills/index.md`). Its body instructs GPT-5.1 to
   emit the tool's `<HTML_OPEN>` block and tells the operator to click Open HTML Tool.
2. **The skill names the tool by filename** (in its `<HTML_OPEN> tool:` line), never a path, so
   it stays portable; `OpenHtmlTool` resolves the filename against the local `-html` folder.

Because the skill is embedded, no separate `.skill.md` file is distributed - handing someone the
one `.html` is enough, and `Install HTML Tool` extracts the skill from it (section 7). The
AutoReviewer skill is also provided as a standalone `autoreviewer.skill.md` for reference and for
pasting into the tool's `#stickshift-skill` block.

## 5. The `<HTML_OPEN>` block (the unified trigger)

GPT-5.1 emits this when a tool-skill applies; the operator copies it and clicks Open HTML Tool.

```
<HTML_OPEN>
tool: autoreviewer.html
include:
- skills/autoreviewer-tool.md
</HTML_OPEN>
```

Grammar (line-oriented, same style as `<CONTEXT_REQUEST>`):

- `tool:` **required** - the html filename, resolved against `HtmlDir()`.
- `include:` *optional* - seed paths to assemble into `-dist` (typically the companion skill so
  a cold/fresh chat can be primed). Omit to open the tool only.
- `depth:` *optional* - bundle depth for the seeds; default `0` (a tool skill is self-contained).

## 6. Runtime change A: `OpenHtmlTool` (co-located in `StickShiftContextBundle`)

This is the button's handler. It must live **in `StickShiftContextBundle`** because it reuses
that module's private assembly path (`ReadClipboard`, `AssembleBundle`, `WriteUtf8`, the header
build, `OUT_FILENAME`, `fso`). It opens the tool and, if seeds are present, writes their bundle
to `-dist`.

First, a tiny refactor so the header is shared: extract the existing header construction (the
block that builds `header` inside `BuildContextBundle`) into a private function and call it from
both places.

```vba
Private Function BundleHeader(ByVal mode As String, ByVal foundCount As Long, _
                              ByVal mapCount As Long, ByVal selCount As Long, _
                              ByVal assembled As String) As String
    Dim totalConcepts As Long: totalConcepts = foundCount + mapCount + selCount
    Dim approxTokens  As Long: approxTokens = CLng(Len(assembled)) \ 4
    Dim ts As String
    ts = Format(Now(), "yyyy-mm-dd") & "T" & Format(Now(), "hh:mm:ss") & "Z"
    BundleHeader = "<!-- OKF-CONTEXT-BUNDLE" & vbLf & _
        "mode: " & mode & vbLf & _
        "okf_version: " & OKF_VERSION & vbLf & _
        "assembled: " & ts & vbLf & _
        "concepts: " & totalConcepts & " (" & foundCount & " foundation, " & _
            mapCount & " map, " & selCount & " selected)" & vbLf & _
        "approx_tokens: " & approxTokens & vbLf & _
        "-->" & vbLf & vbLf
End Function
```

In `BuildContextBundle`, replace the inline header build with
`header = BundleHeader(mode, foundCount, mapCount, selCount, assembled)`. No behavior change.

Then add the public Sub:

```vba
' Reads an <HTML_OPEN> block from the clipboard, launches the named local tool from the
' -html sibling, and (if include: seeds are given) assembles their bundle into -dist.
Public Sub OpenHtmlTool()
    Dim root As String
    root = StickShiftConfig.BundleRoot()
    If root = "" Then
        MsgBox "Bundle root not set - click Switch Context.", vbExclamation, "StickShift"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim clip As String
    On Error Resume Next
    clip = ReadClipboard()
    On Error GoTo 0
    clip = Replace(Replace(clip, vbCrLf, vbLf), vbCr, vbLf)

    Dim startPos As Long: startPos = InStr(clip, "<HTML_OPEN>")
    Dim endPos   As Long: endPos = InStr(clip, "</HTML_OPEN>")
    If startPos = 0 Or endPos <= startPos Then
        MsgBox "No <HTML_OPEN> block on the clipboard." & vbLf & _
               "Copy the block the assistant gave you, then try again.", _
               vbExclamation, "StickShift"
        Exit Sub
    End If

    Dim body As String
    body = Mid(clip, startPos + Len("<HTML_OPEN>"), endPos - (startPos + Len("<HTML_OPEN>")))

    Dim toolFile As String: toolFile = ""
    Dim depth As Long: depth = 0
    Dim seeds(0 To 199) As String
    Dim seedCount As Long: seedCount = 0

    Dim lines() As String: lines = Split(body, vbLf)
    Dim inInclude As Boolean: inInclude = False
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String: ln = Trim(lines(i))
        If Left(ln, 6) = "tool: " Then
            toolFile = Trim(Mid(ln, 7)): inInclude = False
        ElseIf Left(ln, 7) = "depth: " Then
            On Error Resume Next
            depth = CLng(Trim(Mid(ln, 8)))
            On Error GoTo 0
            inInclude = False
        ElseIf ln = "include:" Then
            inInclude = True
        ElseIf inInclude And Left(ln, 2) = "- " Then
            If seedCount <= UBound(seeds) Then
                seeds(seedCount) = Trim(Mid(ln, 3)): seedCount = seedCount + 1
            End If
        ElseIf ln <> "" And Left(ln, 1) <> "-" Then
            inInclude = False
        End If
    Next i

    If toolFile = "" Then
        MsgBox "<HTML_OPEN> block has no `tool:` line.", vbExclamation, "StickShift"
        Exit Sub
    End If

    ' Resolve + launch the tool from the -html sibling.
    Dim htmlPath As String
    htmlPath = StickShiftConfig.HtmlDir() & toolFile
    If Not fso.FileExists(htmlPath) Then
        MsgBox "Tool not found:" & vbLf & htmlPath & vbLf & vbLf & _
               "Install it first with Install HTML Tool.", vbExclamation, "StickShift"
        Exit Sub
    End If
    OpenInDefaultBrowser htmlPath   ' default browser via ShellExecute "open" (see helper below)

    ' If seeds were given, assemble their bundle into -dist.
    Dim msg As String: msg = "Opened: " & toolFile
    If seedCount > 0 Then
        Dim seedSlice() As String
        ReDim seedSlice(0 To seedCount - 1)
        For i = 0 To seedCount - 1: seedSlice(i) = seeds(i): Next i

        Dim fC As Long, mC As Long, sC As Long
        Dim assembled As String
        assembled = AssembleBundle(root, seedSlice, depth, "outbound", "", fC, mC, sC)

        Dim fullOutput As String
        fullOutput = BundleHeader("bundle", fC, mC, sC, assembled) & assembled

        Dim distDir As String: distDir = StickShiftConfig.DistDir()
        If distDir <> "" Then
            WriteUtf8 distDir & OUT_FILENAME, fullOutput
            msg = msg & vbLf & "Skill context written to -dist\" & OUT_FILENAME & _
                  " (paste it only if you reopen the tool from a fresh chat)."
        End If
    End If

    MsgBox msg, vbInformation, "StickShift"
End Sub
```

Default-browser launch (replaces `explorer.exe`). Add the Win32 declare at the top of the module
(VBA7-guarded, like the declares in `StickShiftClipboard`) and a small helper. The `"open"` verb
routes through the OS file association, i.e. the operator's default browser:

```vba
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecuteA Lib "shell32" ( _
        ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecuteA Lib "shell32" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

Private Sub OpenInDefaultBrowser(ByVal filePath As String)
    ' SW_SHOWNORMAL = 1. ANSI ShellExecute is fine for ASCII install paths.
    ShellExecuteA 0, "open", filePath, vbNullString, vbNullString, 1
End Sub
```

Notes:
- It deliberately does **not** pop the `-dist` Explorer window (the browser is the focus; popping
  Explorer too is clutter). The MsgBox states the path.
- ASCII only, no `ChrW`.
- Optional UX upgrade (separate, not required): add a `SetClipboardText` Win32 helper to
  `StickShiftClipboard` (currently read-only) and copy `fullOutput` to the clipboard instead of
  only writing `-dist`, so the operator's next paste into a fresh chat just works. Defer unless
  the cold-start paste proves annoying.

## 7. Runtime change B: `InstallHtmlTool` (`StickShiftHtmlTools.bas`)

A new module + public Sub, modeled on `BootstrapBundle` (same root guard, same FSO, finish with
reindex). Flow:

1. Root guard (copy from `BootstrapBundle`).
2. File picker (`Application.FileDialog`, file-picker) filtered to `*.html`; operator points at
   `<tool>.html`. (No paired file is needed - the skill is embedded in the HTML.)
3. Copy the HTML into `HtmlDir()` (`fso.CopyFile srcHtml, HtmlDir() & leafName, True`).
4. Read the copied HTML and extract the embedded skill: find the
   `<script type="text/markdown" id="stickshift-skill" ...>` block, read its `data-skill-slug`
   attribute as `skillSlug`, and take the inner text as the skill markdown. If no such block is
   found, message and abort ("not a StickShift-compliant tool"). For rename-safety, rewrite the
   skill's `tool:` line to the actual copied filename (`leafName`) before writing.
5. Write the extracted skill to `skills/<skillSlug>.md` via the shared writer
   (`StickShiftWriteApply.ApplyWriteEnvelopeText` with a one-file envelope, exactly as
   `BootstrapBundle` does - this also logs the write and honors the reserved-file guard).
6. `StickShiftIndexGenerator.GenerateStickShiftIndexes`.
7. Summary MsgBox: tool name, where the HTML landed, the skill path.

ASCII only, no `ChrW`. The HTML is a file copy, never embedded in a `.bas` (the guards do not
touch the `.html`, and embedding would also blow the ASCII guard). Re-install is idempotent
(`CopyFile ..., True` overwrites; the skill rewrite goes through the normal path).

## 8. Cockpit buttons

`CreateStickShiftDashboard` rebuilds the sheet each run, so adding a button = adding rows + one
`Make*Button` call + shifting later row indices. Add two **demoted** buttons (setup-class) near
`Initialize Context` / `Restart Setup Wizard`, using `MakeDemotedButton` and the existing
palette:

- **Install HTML Tool** -> `StickShiftHtmlTools.InstallHtmlTool`
- **Open HTML Tool** -> `StickShiftContextBundle.OpenHtmlTool`

Leave the exact row arithmetic to the implementer against the current layout.

## 9. System-prompt addition (`okf-context-assistant.md`)

Additive block for your current (edited, sidecar-free) instructions - merge by meaning. The
three existing modes stay RETRIEVE / WRITE / ANSWER; this adds one output the assistant emits
when a tool-skill applies.

> **Opening a local tool.** Some skills describe a local HTML tool that the operator launches
> from StickShift (a chat hyperlink cannot open a local file). When such a skill applies, emit
> its `<HTML_OPEN>` block exactly as the skill specifies - `tool:` names the tool file, optional
> `include:` names skill paths to load into `-dist` - followed by a one-line instruction to click
> Open HTML Tool, and nothing else. Emit at most one envelope per turn (an `<HTML_OPEN>` counts
> as that turn's envelope).

## 10. Portability (now solved)

The skill stores only the tool **filename**; `OpenHtmlTool` resolves it against the local
`HtmlDir()`. So a context shared to another machine works as long as that machine has installed
the tool (its `-html` folder has the file). No absolute paths, no per-machine link baking, no
bundle-time rewrite.

## 11. Non-goals

- **No HTML through `<VBA_WRITE>`** (section 2).
- **No clickable links anywhere** - launch is via the button (section 0).
- **No bundled tool catalog / auto-update** in v1 (the file picker is enough; a curated
  `builds/html-tools/` chooser can come later).
- **No clipboard auto-copy** in v1 unless the cold-start paste proves annoying (section 6 note).

## 12. Tests / CI

- **The `<HTML_OPEN>` parser is twin-testable** (pure logic, like `parse_write_envelope`). Add a
  small Python twin (e.g. `tests/test_html_open.py`): input an `<HTML_OPEN>` block, expect
  `(tool, depth, seeds[])`. Golden vectors: tool only; tool + include list; tool + depth +
  include; ignores lines after a non-list line; no `tool:` -> tool is empty (caller errors).
- **`OpenHtmlTool`, `InstallHtmlTool`, `HtmlDir`, the cockpit, and bundle assembly are
  filesystem/Excel-bound** - verify by **manual test**: install a tool (confirm the `.html` lands
  in `<root>-html\`, the skill appears in `skills/` and `skills/index.md`), then copy an
  `<HTML_OPEN>` block and click Open HTML Tool (confirm the tool opens in the browser and, with
  an `include:`, `StickShift-context.md` is written to `-dist`).
- `pytest tests/ -q` must still pass; the `BundleHeader` extraction is behavior-preserving so the
  existing context-bundle twin is unaffected. ASCII + ChrW guards pass on
  `StickShiftConfig.bas`, `StickShiftContextBundle.bas`, and the new `StickShiftHtmlTools.bas`.

## 13. Build prompt for Claude Code

> Read `_meta/spec-html-tools.md` and implement it in full. Scope:
>
> 1. `builds/StickShiftConfig.bas`: add `HtmlDir()` exactly as in section 3 (the `-html` sibling,
>    parallel to `DistDir()`). Pure ASCII, no `ChrW`.
> 2. `builds/StickShiftContextBundle.bas`: extract the inline header construction in
>    `BuildContextBundle` into the private `BundleHeader` function from section 6 and call it from
>    `BuildContextBundle` (behavior-preserving). Add the public `OpenHtmlTool` Sub exactly as in
>    section 6, plus the `ShellExecuteA` declare and `OpenInDefaultBrowser` helper (open the tool
>    in the default browser via the `"open"` verb - NOT `explorer.exe`). It reuses `ReadClipboard`,
>    `AssembleBundle`, `WriteUtf8`, `OUT_FILENAME`, `fso`, `DistDir`, and the new `HtmlDir`. Do not
>    change `AssembleBundle`, `AssembleIndex`, or the `BuildContextBundle` public contract. Pure
>    ASCII, no `ChrW`.
> 3. Create `builds/StickShiftHtmlTools.bas` with a public `InstallHtmlTool` following section 7
>    (root guard from `BootstrapBundle`, `*.html` picker, copy the HTML to `HtmlDir()`, **extract
>    the embedded skill** from the copied HTML's `<script id="stickshift-skill" data-skill-slug>`
>    block, rewrite its `tool:` line to the actual filename, write it to `skills/<slug>.md` via
>    `StickShiftWriteApply.ApplyWriteEnvelopeText`, then `GenerateStickShiftIndexes`, then a
>    summary MsgBox). Pure ASCII, no `ChrW`.
> 4. `builds/StickShiftDashboard.bas`: add demoted buttons "Install HTML Tool" ->
>    `InstallHtmlTool` and "Open HTML Tool" -> `OpenHtmlTool` via `MakeDemotedButton`; adjust the
>    later row indices so the layout still renders.
> 5. The AutoReviewer tool `.html` itself is built separately per
>    `_meta/spec-html-tool-compliance.md` (it embeds its skill). For reference, the standalone
>    `autoreviewer.skill.md` is the skill that belongs in its `#stickshift-skill` block. Do NOT
>    author the `.html` in this task.
> 6. Add `tests/test_html_open.py` - the `<HTML_OPEN>` parser twin + golden vectors from
>    section 12.
> 7. Apply the section 9 system-prompt block to `okf-context-assistant.md` by meaning (the file
>    is edited locally; merge additively).
>
> Constraints: `pytest tests/ -q` passes (the `BundleHeader` extraction is behavior-preserving;
> only the new `test_html_open.py` is added). ASCII + ChrW guards pass on all edited/new `.bas`.
> Do not route HTML content through `<VBA_WRITE>`, do not add clickable links, do not add a
> clipboard writer (section 11). Summarize the diff and list the end-to-end manual test (install
> a tool, then open it via an `<HTML_OPEN>` block).
