# Build Portfolio — Usage Loop

Six steps. Do them in order. Stop after step 4 until the portfolio is large
enough that a step stops feeling instant.

---

## Step 1 — Populate `builds/`

Brain-dump every project to DHSChat. Ask it to return one schema-conformant
`.md` file per project. Copy each file into `builds/` with a short slug as the
filename (e.g. `briefing-builder.md`).

One file = one build. That's the whole data model.

## Step 2 — Run the Index Generator

Open `StickShiftIndexGenerator.bas` in Excel's VBA editor (Alt+F11 → Import File).

Before first run:
1. Set `BUNDLE_ROOT` in the constants block to your actual OneDrive path
   (e.g. `C:\Users\YourName\OneDrive\build-portfolio\`).
2. Verify Tools → References includes **Microsoft ActiveX Data Objects 2.x**.
3. Test against a 3-file subfolder before running on the full bundle.
4. Run `GenerateStickShiftIndexes`. Confirm `index.md` appears at the bundle root
   and reads as a status board with `working` items at the top, sorted oldest
   `last_touched` first.

## Step 3 — Paste the index into the Portfolio Strategist

Open the Portfolio Strategist DHSChat Assistant (bookmark the URL after first
setup). Paste the contents of `index.md` as your first message. Then ask:
"What should I work on?"

The Assistant holds the reasoning frame. The index is the data. Fresh paste
each session = always current, no retrieval needed.

## Step 4 — Act

Do the one thing the Strategist recommends. When it's done, use the write path
below to update the build file, re-run the generator, and repeat from Step 3.

---

## Step 5 — Write path: propose a change behind a human gate

The Portfolio Writer DHSChat Assistant (`portfolio-writer-assistant.md`) authors
edits. It takes the pasted `index.md` plus a request and returns a write
envelope. The macro applies it — but only after you review it.

### Write a change

1. Open the Portfolio Writer Assistant (bookmark the URL after first setup).
2. Paste `index.md` as your first message, then state your request:
   - "Mark briefing-builder done and set last_touched to today"
   - "Add a new build for the OKF write side"
   - "Archive the abandoned X"
3. The Assistant returns **only** a `<VBA_WRITE>` envelope. Copy the entire
   output to the clipboard.

### Apply the change via macro

4. Import `builds/StickShiftWriteApply.bas` into the same Excel workbook as
   `StickShiftIndexGenerator.bas`. Before first run, set `BUNDLE_ROOT` in the
   constants block — it **must** match `StickShiftIndexGenerator.bas`.
5. Run `ApplyStickShiftWrite`. The macro reads the clipboard and applies the gate:
   - **New file** (path does not exist) → written directly.
   - **Edit to existing file** → written as `<path>.md.proposed` sidecar.
     The original is untouched.
6. A summary message reports how many files were written and how many were
   staged as `.proposed`.
7. The macro calls `GenerateStickShiftIndexes` automatically. New builds appear in
   the index immediately. `.proposed` sidecars are ignored by the generator
   (they do not end in `.md`).

### Review and merge a staged edit

8. Open the `.proposed` file alongside the original in any text editor.
9. Verify the change is exactly what was intended.
10. Rename `<slug>.md.proposed` → `<slug>.md` to apply the edit, or delete
    the `.proposed` file to discard it.
11. Re-run `GenerateStickShiftIndexes` after renaming.

The rename is the merge. No edit ever lands without a human approving it.

---

## Step 6 — Run the linter to check integrity

Import `builds/StickShiftLint.bas` into the same workbook. Set `BUNDLE_ROOT` to match
the other macros. Run `RunStickShiftLint`.

The linter checks:
- **Missing type / required fields** — every build needs `type`, `status`,
  `effort`, and `impact`.
- **Broken cross-links** — `.md` links whose target file does not exist.
- **WIP violation** — more than one build at `status: working`.
- **Stalls** — `working` builds with missing or old `last_touched` (oldest
  first, so the most-stalled build is easiest to act on).
- **Pending `.proposed` files** — staged edits that have not been reviewed yet.
- **Active-to-archived links** — a non-archived build linking to an archived
  one (retarget the dependency before depending on it).

Findings are written to a **StickShift Lint Report** worksheet (colour-coded errors
and warnings). Run the linter after any batch of edits, and before committing
the portfolio to version control.

---

That's the shipped tool. Do not add to it until you've outgrown it.
