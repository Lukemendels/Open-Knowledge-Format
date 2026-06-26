---
artifact: ModelBuilder
status: draft-v1
layer: StickShift HTML tooling
authors_supported: human (manual)  # no LLM in the loop for v1
persistence: download (terminal); write-gate route deferred
---

# ModelBuilder — Functional Spec (v1)

## One-line

Drop N source `.xlsx` files into an inert HTML page; it consolidates every
source sheet into a single self-contained workbook and generates an **input
sheet** that wires those sources in via live formulas, then emits one
downloadable `.xlsx`. No network, no API, no LLM. Pure client-side
manufacturing.

---

## 1. Why this exists (design rationale)

Every other StickShift HTML tool so far operates on the *current state as a
noun* — render it, edit it, save it. ModelBuilder is the first tool that
**manufactures a new artifact** rather than re-presenting an existing one. It
sits in the same role HTML always plays in the stack: the deterministic
factory. The division of labor is the AHA pattern applied to production:

- A human (or DHSChat, later) decides *which files* and *what should connect to
  what* — judgment.
- The HTML layer guarantees *form* — a schema-valid workbook with correct
  formula references, every time — deterministic.

The reason this one is "relatively easy" is that it needs **zero stochastic
authorship**. Wiring sheets together is a structural computation, not a
language task. That makes it the cleanest first proof that the HTML layer is a
document factory and not just a viewer.

### The one real design decision

The manufacturing is trivial. The only genuine design question is **what
"wire in" means** — i.e., what the input sheet actually does. See §4. Pin that
down before building; everything else is mechanical.

---

## 2. Hard constraints (read before scoping)

These come from SheetJS Community edition and from Excel's reference model.
They define what v1 can honestly promise.

1. **SheetJS Community reads/writes values, formulas, and basic cell structure.
   It does NOT round-trip rich formatting, charts, pivot tables, conditional
   formatting, or macros.** A source workbook full of styling and charts will
   come out as data + formulas only. v1 promise is *data fidelity*, not
   *presentation fidelity*. Demo on plain data sheets, not on a formatted
   dashboard.

2. **Cross-workbook external references are fragile.** A formula like
   `='C:\Users\...\[SourceA.xlsx]Sheet1'!B2` breaks the moment a path changes
   or the source moves — exactly the failure mode federal file shuffling
   produces. **Therefore: do not wire across workbooks. Consolidate everything
   into ONE output workbook and wire with intra-workbook references
   (`=SheetName!B2`).** Self-contained is the whole point.

3. **Formulas inside copied source sheets must still resolve.** If `SourceA`
   has a sheet whose formula reads `=OtherTab!A1`, that reference has to keep
   pointing at the right tab after consolidation. This is the one place naming
   collisions bite (see §3, step 3).

4. **`.xlsx` is binary; it does not survive the clipboard→VBA text path.** So
   the output is a **browser download**, which is terminal for a deliverable.
   If a future version needs the workbook to land in OneDrive state, that's a
   separate write-gate route, explicitly out of scope here.

---

## 3. Pipeline (what the tool does, in order)

### Step 1 — Ingest
- Drag-drop zone + file picker. Accept multiple `.xlsx`.
- For each file: `XLSX.read(arrayBuffer, { type: 'array', cellFormula: true })`.
- Hold all parsed workbooks in session memory (same discipline as the file
  explorer — in-memory working set, dies with the tab, nothing persisted until
  the user acts).

### Step 2 — Inventory (the legibility step)
Render a panel listing, per source file:
- filename
- each sheet name
- used range (e.g., `A1:F240`)
- detected header row (first non-empty row), shown as chips

This is the demystify move carried over from the explorer: before the tool
builds anything, the human sees exactly what got loaded. Inventory is also
where the human picks targets in §4.

### Step 3 — Consolidate
- Create one output workbook.
- Copy every source sheet into it as its own tab.
- **Collision handling:** prefix each tab with a slug of its source filename,
  e.g. `SourceA__Sheet1`, `SourceB__Sheet1`. Keep a map of
  `original → renamed`.
- **Reference rewrite:** when copying a sheet that contains intra-source
  formulas, rewrite tab references in those formulas using the rename map so
  they still resolve. (v1 can scope this to simple `Sheet!Cell` references and
  *flag* anything it can't safely rewrite rather than silently breaking it —
  fail friendly, like BriefingBuilder.)

### Step 4 — Generate the input sheet
Insert a new first tab, `INPUT`, built per the §4 wiring mode.

### Step 5 — Emit
- `XLSX.write(wb, { bookType: 'xlsx', type: 'array' })` → `Blob` → download.
- Filename default: `model_<timestamp>.xlsx`. Carry the grab-time timestamp
  pattern from the explorer so the artifact is dated at creation.

---

## 4. Wiring modes (pick the default; this is the real decision)

"Wire the source sheets in via formulas" can mean three different things. They
are not equally hard and they serve different jobs. Recommend shipping
**Mode A** as v1 default and leaving B/C as flags.

### Mode A — Pass-through aggregation (recommended v1)
The `INPUT` sheet surfaces each source sheet's key range via live formulas, so
the consolidated workbook has one place to read from and edits propagate.

Example — three source files, each a small table. `INPUT` gets one labeled
block per source:

```
A1: SourceA            B1: (header row mirrored)   C1: ...
A2: =SourceA__Sheet1!A2  B2: =SourceA__Sheet1!B2     ...
A3: =SourceA__Sheet1!A3  ...
...
A12: SourceB
A13: =SourceB__Sheet1!A2 ...
```

Because these are formulas, not pasted values, changing a source tab updates
`INPUT` automatically. The output is a single readable rollup wired live to
its sources. This is the most generally useful and the easiest to reason
about.

### Mode B — Named-input mapping
The human, in the inventory panel, tags specific cells as named inputs
(`tax_rate`, `base_volume`). `INPUT` becomes a two-column
`name | =Source!cell` table, and the tool optionally writes Excel **defined
names** so downstream formulas can use `=tax_rate` instead of cell coordinates.
More powerful, needs a tagging UI. Flag, not v1.

### Mode C — Stack/union
When sources are the *same schema* (e.g., monthly extracts), `INPUT` stacks
them into one long table with a `source_file` column appended, via formulas or
values. This is closer to "append queries." Useful but a different tool in
spirit. Flag.

**Decision needed from you:** confirm Mode A is the v1 target. If your actual
mental picture is B (a real input panel feeding a model) or C (combining like
extracts), say so — it changes the inventory UI, not the plumbing.

---

## 5. UI (two-pane, consistent with the explorer)

- **Left:** drop zone + source inventory tree (file → sheets → range).
- **Right:** preview of the `INPUT` sheet that *will* be generated, rendered as
  an HTML table, with formula cells shown as their formula (e.g.
  `=SourceA__Sheet1!B2`) so the human sees the wiring before downloading.
- **Top:** grab-time timestamp, loud and plain ("Built from N files · 14:32").
- **Action:** single **Build & Download** button.
- **Footer:** `⚙ Part of StickShift ▸` discovery affordance.

Preview-before-emit is the trust step: the human verifies the wiring in
readable form, then commits to the download. Same surfaces-and-verifies logic
as the rest of the stack.

---

## 6. Dependencies
- SheetJS Community (`xlsx`), single bundled JS file, no CDN, no network.
- Everything else is vanilla DOM. One file, fully offline, opens from
  `file://`.

---

## 7. Explicitly out of scope for v1
- Formatting / chart / pivot / macro preservation.
- Cross-workbook external references.
- Write-gate / OneDrive landing (download is terminal).
- LLM involvement (this is the pure-manufacturing tool; DHSChat adds nothing
  to a structural wiring task).
- Mode B tagging UI and Mode C stacking.

---

## 8. Acceptance test (the demo that proves it works)
1. Make three trivial `.xlsx` files, each a 5-row table, plain data.
2. Drop all three in. Inventory shows three files, three sheets, ranges.
3. Right pane previews an `INPUT` sheet with three formula-wired blocks.
4. Build & Download. Open in Excel.
5. Edit a value in `SourceB__Sheet1`. Confirm the matching `INPUT` cell
   updates live. That live update is the whole proof — it shows the output is
   *wired*, not *pasted*.
