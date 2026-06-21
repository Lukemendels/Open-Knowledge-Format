# OKF — Manual Smoke Test

Runs the VBA-runtime behavior the 76 twin tests can't reach: registry read/write, the folder
picker, the clipboard, Explorer open, and real file I/O. Pair this with the **okf-smoke-bundle**
(unzip it somewhere on OneDrive). Each step has an **Expect**; tick them off in order.

---

## Setup

**A. Import the modules.** New macro-enabled workbook (`.xlsm`) → Alt+F11 → File → Import File →
import all six: `OKFConfig`, `OKFIndexGenerator`, `OKFLint`, `OKFWriteApply`, `OKFContextBundle`,
`OKFDashboard`. Save as `.xlsm`.
- If macros are blocked: File → Options → Trust Center → enable macros / trust the file's location.
- **Clipboard gotcha (note now, fix only if step 7 fails):** if Build Context Bundle later errors
  reading the clipboard, add Tools → References → **Microsoft Forms 2.0 Object Library** (or insert
  one blank UserForm), then retry. `MSForms.DataObject` late-binding can fail without it.

**B. Unzip the smoke bundle** to a folder on OneDrive (e.g. `…\OneDrive\okf-smoke-bundle`). Note
the full path. Put nothing else in it.

---

## Dashboard + config (registry + picker)

**1. Build the dashboard.** Alt+F8 → `CreateOKFDashboard` → Run.
**Expect:** a formatted "OKF Dashboard" sheet, **5 buttons**, and a current-root cell at **row 21**
showing `(not set)`.

**2. Set the root.** Click **Set Bundle Root** → folder picker → select the `okf-smoke-bundle` folder.
**Expect:** a confirmation, and row 21 updates to the path (ending in `\`).

**3. Persistence.** Save + close the workbook, reopen it, re-run `CreateOKFDashboard` (or just read
row 21).
**Expect:** the path is still there — proving it persisted to the registry, not the workbook.

---

## Existing pipeline (confirms `BundleRoot()` resolves everywhere)

**4. Regenerate Indexes.** Click it.
**Expect:** `index.md` written in the root and in `_foundation/`, `skills/`, `builds/` (4 files).
Open `builds/index.md` → a bullet per build (`title — description`); open the root `index.md` →
links to the subdirectories.

**5. Run Linter.** Click it.
**Expect:** a report with **zero errors**. (The bundle is valid: exactly one `working` build with a
fresh `last_touched`, and every dependency link resolves.) Any findings still confirm the linter runs.

---

## Context bundle — INDEX mode (Hop 1)

**6. Empty-request run.** Make sure the clipboard has **no** `<CONTEXT_REQUEST>` (copy a random word,
or nothing). Click **Build Context Bundle**.
**Expect:** Explorer opens to `…\okf-smoke-bundle-dist\OKF-context.md`, and a MsgBox reports char /
token counts. Open `OKF-context.md`:
- a header block (mode, concept counts, `approx_tokens`),
- the foundation file → then root `index.md` → then `skills/index.md`,
- each wrapped in `<!-- OKF:BEGIN concept=… layer=… -->` / `<!-- OKF:END -->`, with `layer=foundation`
  and `layer=map`.

Note the output lives in the **`-dist` sibling**, never inside the bundle.

---

## Context bundle — BUNDLE mode (Hop 2+)

For each, copy the block to the clipboard, click **Build Context Bundle**, then open `OKF-context.md`.

**7. Depth-1, all links.**
```
<CONTEXT_REQUEST>
mode: bundle
depth: 1
include:
- builds/alpha.md
</CONTEXT_REQUEST>
```
**Expect:** `alpha` + `beta` + `delta` (depth-1 outbound — beta from Dependencies, delta from Notes),
each `layer=selected`.

**8. via-scoping (the key one).**
```
<CONTEXT_REQUEST>
mode: bundle
depth: 1
direction: outbound
via: Dependencies
include:
- builds/alpha.md
</CONTEXT_REQUEST>
```
**Expect:** `alpha` + `beta` **only** — `delta` is excluded because its link is under `# Notes`, not
`# Dependencies`. This proves `via:` works.

**9. Depth-2 chain.**
```
<CONTEXT_REQUEST>
mode: bundle
depth: 2
direction: outbound
via: Dependencies
include:
- builds/alpha.md
</CONTEXT_REQUEST>
```
**Expect:** `alpha` + `beta` + `gamma` — the full chain, two layers out.

**10. Inbound (backlinks).**
```
<CONTEXT_REQUEST>
mode: bundle
direction: inbound
include:
- builds/gamma.md
</CONTEXT_REQUEST>
```
**Expect:** `gamma` + `beta` (beta is the one that links to gamma). Swap in `builds/beta.md` →
expect `beta` + `alpha`.

---

## Write path (recommended)

**11. Stage an edit.** Copy this whole block (it bumps Delta `idea → spec`):
```
<VBA_WRITE>
### FILE: builds/delta.md
---
type: Build
title: Delta Experiment
description: Smoke-test idea referenced from Alpha's notes (for inbound + via tests).
status: spec
effort: M
impact: med
domain: Personal
timestamp: 2026-06-20T00:00:00Z
last_touched: 2026-06-20
tags: [smoke]
---

# What it is
An idea referenced from Alpha's Notes — used to verify via-scoping and inbound backlinks.

# Next action
Decide whether to pursue.

# Dependencies
(none)

# Notes
(none)
### END FILE
</VBA_WRITE>
```
Click **Apply Write Envelope**.
**Expect:** a `builds/delta.md.proposed` sidecar is created — the original is **not** overwritten.
Review it, rename it over `delta.md` to merge, then **Run Linter** again → still clean.

---

## Pass criteria

- [ ] All 5 buttons run with no VBA error.
- [ ] Root persists across a close/reopen (registry).
- [ ] Indexes generate; linter reports zero errors.
- [ ] Index mode and bundle mode each write a correct `OKF-context.md` to the `-dist` sibling.
- [ ] via / inbound / depth-2 behave exactly as the Expects above.
- [ ] The write envelope stages a `.proposed` — no silent overwrite.

## Troubleshooting quickref

- **Clipboard read fails** → add the Microsoft Forms 2.0 reference (or insert a UserForm).
- **"Bundle root not set"** → click Set Bundle Root.
- **Explorer doesn't open but the file exists** → harmless; open the `-dist` folder manually.
- **Linter flags a WIP violation** → you have more than one `working` build; the smoke bundle ships
  with exactly one, so check nothing got edited.
