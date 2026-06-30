---
type: Spec
title: Edit-Diff Capture — human edits as a skill-tuning signal
description: Gates the VBA write so a tracked write (one carrying a reason) snapshots before/after the file and stores the pair in an -edits sibling store, keyed on a stable concept id. The accumulated diffs are labeled data ("the bundle said X, I changed it to Y") for tuning skills and prompts. Capture is global; per-skill vs. taste attribution is a downstream analysis pass, not a field at the gate.
tags:
  - markdown-editor
  - html-tool
  - concept-viewer
status: spec
okf_version: "0.1"
---

# Edit-Diff Capture

## 0. Model note

This is a deterministic capture spec — it is entirely VBA + a small viewer change, no LLM in the loop at capture time. The judgment work (which skill does a diff implicate, what kind of edit is it) is deliberately pushed to a later analysis pass over the store. Keep that split: the gate captures, a model attributes.

## 1. The problem

When you hand-edit a context-bundle file in the viewer and commit it, the only trace today is one line in `log.md`: `edit  builds/precheck-ria.md`. That records *that* a write happened, not *what changed* or *why*. So the single richest signal the system produces — your deliberate correction of what DHSChat drafted — evaporates at the moment it is most informative.

Those corrections are labeled data. "The bundle said X, I changed it to Y" is exactly the supervision a skill-tuning loop needs: capture enough of them with a reason attached and you can see which parts of a bundle you keep fixing by hand, which tells you what to bake into the skill (or the system prompt) so the fix stops being manual. The capture tool is the front end of a skill-improvement flywheel.

## 2. What already exists (so we do not rebuild it)

1. **A single commit point.** `ApplyWriteEnvelopeText` in `StickShiftWriteApply.bas` is the one place every write — viewer or DHSChat — crosses from packet to disk. Its per-file loop (currently ~line 128) already computes `absPath`, checks `existed = fso.FileExists(absPath)`, and then calls `WriteUtf8 absPath, fileContents(i)`. **The old file is still on disk in the window between that check and that write.** That window is the entire capture opportunity — no new commit point is needed.

2. **The sibling-store convention.** `StickShiftConfig` already exposes `DistDir()` (`-dist`) and `HtmlDir()` (`-html`), each stripping the root's trailing slash and appending a suffix. Siblings of the bundle root are never indexed or retrieved as concepts. The diff store is a third sibling, `-edits`, built the same way.

3. **The reserved-file guard.** The loop already skips `log.md` / `index.md`. Capture inherits that skip for free — machine-owned files are never diffed.

## 3. Design decisions (carried from the 2026-06-27 session, plus the attribution decision)

1. **Gate at the VBA write button** — the commit point, where "what it was" (disk) and "what it's becoming" (packet) coexist in one hand. Instrument the write, not the writer.
2. **Capture is global, keyed on a stable concept `id`.** No `skill_id` field at the gate. Per-skill drift and taste-level patterns are both *queries over the store*, not fields committed at write time — because at human-edit time there is no skill in the loop, so per-skill attribution is often unavailable and would mean storing a guessed or empty field. (Resolved 2026-06-30: capture global, attribute downstream.)
3. **Stamp the coarse bucket for free.** The edited file's own frontmatter `type` (Skill / Build / Foundation / Note) is on hand during the same parse that finds the id. Stamp it into the diff record as `concept_type`. That buys the single most useful cut — skill vs. content — with zero attribution work.
4. **Store fat now, compute lean later.** Full before/after snapshots per write, not deltas. You can derive deltas from snapshots; you cannot recover snapshots from deltas. Disk is free at this scale.
5. **Capture the reason at write time.** A half-sentence (git-commit style), prompted once per save, stored in the diff record's front matter. The diff records *that* an edit happened; the reason records *why*, while it is still in your head.
6. **Crash-safe order:** write the diff record (AFTER then BEFORE) *first*, then overwrite the live file. If the process dies mid-write, the live file is still the BEFORE and the diff holds both versions — nothing is lost. Overwriting first would risk losing BEFORE.

## 4. Identity: a lazily-assigned `id` in frontmatter

OKF frontmatter has no stable id today. Rather than retrofit ids across the whole bundle (a migration with no present user), assign one **lazily, on first tracked edit**, and persist it in the file's own frontmatter so it travels with the content through rename or move.

`EnsureConceptId(ByRef content As String) As String`:

- Parse the leading `---` … `---` frontmatter block.
- If an `id:` line exists, return its value, leave `content` unchanged.
- Else generate `id = "c-" & Format(Now, "yyyymmddhhnnss") & "-" & seq` (a within-batch sequence counter guarantees uniqueness when several new files are written in the same second), insert `id: <new>` as a new line immediately after the opening `---`, mutate `content` byref, and return the new id.
- If the file has **no** leading frontmatter at all (e.g. a sharding stub): do not invent a frontmatter block. Fall back to `id = "path:" & relFwd` and set `id_source = "path"` in the diff record (note: a path-keyed id does not survive rename — acceptable for the rare frontmatter-less file).

Because the id is injected into `content` *before* both the live write and the diff's AFTER section, the live file and the snapshot stay byte-consistent.

## 5. The trigger: a `reason:` line, not a new envelope verb

A tracked write is an ordinary `<VBA_WRITE>` envelope with **one optional first line**, `reason: <text>`, immediately after the opening tag:

```
<VBA_WRITE>
reason: tightened wage-source citation to BLS OEWS
### FILE: builds/precheck-ria.md
...full file...
### END FILE
</VBA_WRITE>
```

**Presence of `reason:` is the capture signal.** DHSChat's routine writes omit it and stay on the existing fast path untouched; the viewer's human-edit writes include it and trigger capture. This keeps one envelope grammar, one parser, one write loop, and the load-bearing invariant ("every write restates a whole, schema-valid file") — capture is a side effect gated by a boolean, not a second code path.

> **Alternative considered:** a distinct `<VBA_WRITE_TRACKED>` verb. Rejected for now — it doubles the envelope surface to express what one optional line expresses, and the review-surface concept (§8) can introduce a real verb later if it earns one.

The reason is **per-batch** (one prompt per save, applied to every file in the packet), matching the git-commit mental model. Per-file reasons are a deferred refinement (§8).

## 6. VBA changes (`StickShiftWriteApply.bas`)

All additions pure ASCII, no `ChrW`, to pass CI guards.

**6.1 Parse the reason.** In `ApplyWriteEnvelopeText`, after extracting the envelope body and before the `### FILE:` parse loop, detect a leading `reason:` line:

```vba
Dim captureOn As Boolean: captureOn = False
Dim reasonText As String:  reasonText = ""
' (body is the text between <VBA_WRITE> and </VBA_WRITE>)
Dim firstLine As String
firstLine = Trim(Left(body, InStr(body & vbLf, vbLf) - 1))
If Left(LCase(firstLine), 7) = "reason:" Then
    captureOn = True
    reasonText = Trim(Mid(firstLine, 8))
    ' strip the reason line from body so the FILE parser is unaffected
    body = Mid(body, InStr(body & vbLf, vbLf) + 1)
End If
```

**6.2 Capture in the loop.** Inside the per-file loop, in the non-reserved branch, replace the current `existed`/`WriteUtf8` sequence with capture-then-write:

```vba
absPath = ResolvePath(filePaths(i))
existed = fso.FileExists(absPath)

parentDir = fso.GetParentFolderName(absPath)
If Not fso.FolderExists(parentDir) Then EnsureFolderTree parentDir

' --- tracked-write capture (gated) ---
Dim afterContent As String: afterContent = fileContents(i)
If captureOn Then
    Dim cid As String: cid = EnsureConceptId(afterContent)   ' may inject id into afterContent
    Dim beforeContent As String
    beforeContent = ""
    If existed Then beforeContent = ReadUtf8(absPath)
    CaptureEditDiff cid, filePaths(i), afterContent, beforeContent, _
                    IIf(existed, "edit", "new"), reasonText, i
End If

WriteUtf8 absPath, afterContent     ' afterContent, not fileContents(i): carries any injected id
writeCount = writeCount + 1
```

(`i` is passed only as the within-batch sequence seed for `EnsureConceptId` / filename uniqueness.)

**6.3 `CaptureEditDiff`.** New private sub. Writes one diff file to the `-edits` store, AFTER then BEFORE, crash-safe order already satisfied because this runs *before* the live `WriteUtf8`:

```vba
Private Sub CaptureEditDiff(ByVal cid As String, ByVal relPath As String, _
                            ByVal afterContent As String, ByVal beforeContent As String, _
                            ByVal action As String, ByVal reasonText As String, _
                            ByVal seq As Long)
    Dim dir As String: dir = EditStoreDir()       ' the -edits sibling
    If dir = "" Then Exit Sub
    Dim relFwd As String: relFwd = Replace(relPath, "\", "/")
    If Left(relFwd, 1) = "/" Then relFwd = Mid(relFwd, 2)

    Dim ctype As String: ctype = FrontmatterValue(afterContent, "type")  ' "" if none
    Dim stamp As String: stamp = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim fname As String
    fname = SanitizeId(cid) & "__" & Format(Now, "yyyymmddhhnnss") & "-" & seq & ".md"

    Dim out As String
    out = "---" & vbLf
    out = out & "okf_version: ""0.1""" & vbLf
    out = out & "type: EditDiff" & vbLf
    out = out & "id: " & cid & vbLf
    out = out & "path_at_write: " & relFwd & vbLf
    out = out & "concept_type: " & ctype & vbLf
    out = out & "action: " & action & vbLf
    out = out & "reason: " & reasonText & vbLf
    out = out & "captured: " & stamp & vbLf
    out = out & "---" & vbLf & vbLf
    out = out & "## AFTER" & vbLf & vbLf & afterContent & vbLf & vbLf
    If action = "edit" Then
        out = out & "## BEFORE" & vbLf & vbLf & beforeContent & vbLf
    End If

    WriteUtf8 dir & fname, out
End Sub
```

**6.4 Helpers.** `EditStoreDir()` in `StickShiftConfig` (copy `DistDir`, swap `-dist` for `-edits`). `FrontmatterValue(content, key)` reads one frontmatter scalar (returns "" if absent). `SanitizeId(s)` strips characters illegal in a filename. `EnsureConceptId` per §4.

## 7. Viewer change (`stickshift-viewer.html`)

The Save-Changes export must (a) prompt once for a reason and (b) prepend the `reason:` line.

- On "Save Changes (VBA Envelope)": open a small modal with a single text field ("Why this change? (one line)") and a confirm button. Do **not** use blocking `prompt()` — it is ugly under `file://` and easy to dismiss without a value.
- On confirm, build the envelope with the reason line first:
  ```js
  var env = "<VBA_WRITE>\nreason: " + reason.replace(/\r?\n/g, " ").trim() + "\n";
  dirty.forEach(function(f){ env += "### FILE: " + f.path + "\n" + normalize(f.text) + "\n### END FILE\n"; });
  env += "</VBA_WRITE>";
  ```
- Require a non-empty reason (capture is the point); if blank, keep the modal open with a hint.

This is the only viewer change; the existing dirty-tracking and clipboard path are unchanged.

## 8. Non-goals (per roadmap discipline)

- **No review surface / PENDING-APPROVED UI.** The crystallized "agent proposes -> rendered for review -> approved -> committed" surface rides *on top of* this capture once it exists. Capture is the foundation; build the review UI only when routine vs. consequential writes actually need different handling. (Roadmap candidate.)
- **No automatic skill tuning.** This spec produces the store. Mining it — clustering diffs, naming recurring edits, proposing skill/prompt changes — is a separate analysis tool, and a separate decision about per-skill vs. global *reporting* (the capture is already global; only the report view is open).
- **No per-file reasons.** Per-batch only. Add per-file if a single save routinely mixes unrelated edits.
- **No delta storage.** Full snapshots (decision 4).
- **No bundle-wide id migration.** Ids are assigned lazily on first tracked edit (§4).

## 9. Tests / CI

Twin-testable (pure logic — add to `tests/`):
- **Reason detection:** body with leading `reason: foo` -> `captureOn = True`, `reasonText = "foo"`, and the FILE parser sees the body with that line removed; body without it -> `captureOn = False`, parse unchanged.
- **`EnsureConceptId`:** frontmatter with `id:` present -> returned unchanged, content unchanged; absent -> id inserted exactly once, immediately after the opening `---`, and returned; no frontmatter -> `path:`-keyed fallback, content untouched.
- **Diff body format:** `action = "edit"` -> record has both `## AFTER` and `## BEFORE`; `action = "new"` -> `## AFTER` only, no `## BEFORE`; frontmatter carries `id`, `concept_type`, `reason`, `action`.

Manual test (filesystem-bound, not twin-testable):
- Tracked write of an existing file: confirm the diff file lands in `-edits` with AFTER+BEFORE **before** the live file is overwritten (set a breakpoint after `CaptureEditDiff`, confirm live file still holds BEFORE), and that the live file ends as AFTER with the injected id.
- Untracked write (no `reason:`): confirm no diff file is produced and the live write + log are byte-identical to today.

Guards: the existing twins and golden vectors must pass unchanged (the `### FILE:` parse/compute contract does not change once the reason line is stripped). ASCII guard + ChrW guard must pass on the edited `.bas`.

## 10. Build prompt for Claude Code

> Read `_meta/spec-edit-capture.md` and implement it. Scope:
>
> 1. `StickShiftConfig.bas`: add `EditStoreDir()` by copying `DistDir()` and substituting the `-edits` suffix for `-dist`. Same create-if-absent behavior.
> 2. `StickShiftWriteApply.bas`: add the reason-line detection (§6.1), the gated capture in the per-file loop (§6.2), and the private subs/helpers `CaptureEditDiff` (§6.3), `EnsureConceptId` (§4), `FrontmatterValue`, and `SanitizeId` (§6.4). The capture must run *before* the live `WriteUtf8`, and `WriteUtf8` must write the id-injected `afterContent`, not the raw `fileContents(i)`. Do not change the reserved-file guard, the `index.md`/`log.md` skip, the log format, or the `ApplyWriteEnvelopeText` / `ApplyStickShiftWrite` public contract. Untracked writes (no `reason:`) must behave exactly as today. Pure ASCII, no `ChrW`.
> 3. Add the twin tests in §9 to `tests/` (reason detection, `EnsureConceptId`, diff-body format). Existing twins and golden vectors must pass unchanged.
>
> Do NOT build the viewer change, the review surface, or any analysis/tuning tool — those are out of scope (§7, §8). The viewer change in §7 is specced for a separate pass.
