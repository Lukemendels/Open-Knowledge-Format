---
okf_version: "0.1"
type: Note
title: OKF Roadmap / Deferred Ideas
description: Ideas considered and intentionally deferred (YAGNI). Each entry records the idea, why it was deferred, and the concrete trigger that should make us revisit it.
---

# OKF Roadmap / Deferred Ideas

Things we thought through, understand the options for, and chose **not** to build yet.
The point of this file is to capture the reasoning so a future decision is informed, not re-derived from scratch. An entry graduates to a real task only when its **revisit trigger** actually fires.

## User-editable index grouping axes

**Idea:** let an operator change how a directory's `index.md` is grouped — beyond the built-in `status`-or-flat behavior — without editing VBA. The natural home would be a new reserved config surface (e.g. `_foundation/okf-config.md`) carrying an ordered list of grouping fields the generator may use.

**Shipped instead (Task 7):** grouping is *inferred* from the concepts' own frontmatter. The generator holds an ordered `GROUP_BY_CANDIDATES` list (currently one entry, `"status"`); the first candidate present in a folder's concepts wins; none present → flat alphabetical. This fixed the real problem in front of us (`skills/` filing under `(unset)`) with zero config and no prompt change.

**Why deferred:** the configurability is genuinely useful for **OKF-as-a-published-standard** (adopters with folder genres we haven't imagined), but today there are zero users who need it. Designing the user-configuration ergonomics now means guessing which world we're in (solo operator vs. multi-adopter) before we have to.

**Two known extension paths, in rising effort:**
1. **Widen the in-code list** — append a field name (e.g. `"domain"`) to `GROUP_BY_CANDIDATES`. One edit, one place. Fine for a solo operator. Still requires touching VBA.
2. **External config surface** — `_foundation/okf-config.md` declares the axis list as data; generator reads it. Gets operators fully out of the code, but adds: a reserved file, a parser, missing/malformed handling, and a mention in the operator guide / system map.

**Known hard boundary either way:** if a *single* folder's concepts carry *two* recognized axes (e.g. both `status` and `domain`), "first present wins" needs an explicit per-folder override to choose — and an override reintroduces declaration + the machine-owned-`index.md` ownership question. Pure inference cannot resolve this; it's the true edge of the current design.

**Revisit trigger:** a real non-`skills/` folder needs grouping by a field other than `status`, **or** an OKF adopter asks for configurable grouping. At that point decide between path 1 and path 2 based on how many adopters actually exist — a decision we can only make well once it's concrete.

**Seam already in place:** `GROUP_BY_CANDIDATES` in `OKFIndexGenerator.bas`, marked with an extension-seam comment.

## Whole-folder retrieval (`mode: folder` bundler addition)

**Idea:** a first-class `mode: folder` in `<CONTEXT_REQUEST>` that pulls every concept in a named folder in one hop, so a sharded document (now a folder of section files) loads whole without the operator knowing the internal layout.

**Shipped instead:** nothing new in the bundler. Whole-folder pull already works by seeding the folder's generated `index.md` at `depth: 1` — the hub lists its sections as outbound links, so a one-layer BFS returns `index.md` plus every section in a single assemble. The doc-sharding spec documents this as the canonical "load it whole" move.

**Why deferred:** `mode: folder` would be a second code path that produces the same result the index-at-depth-1 trick already produces. Building it now means owning a new mode branch, its parser line, and a test matrix, to save the operator from typing one seed line. YAGNI.

**Known rough edge (the one that may fire the trigger):** the index-seed trick only returns current sections if `index.md` was regenerated after the last shard/rename — which it is, on every VBA write — and it depends on the operator (or the model) knowing to seed the *index*, not the stub. That second-order knowledge is exactly the friction the **viewer "Load folder -> DHSChat" button** removes: it derives the folder from the file in view and emits the index-at-depth-1 request for you. So the practical fix shipped in the HTML layer, not the bundler.

**Revisit trigger:** the viewer button proves insufficient (e.g. a headless/agentic caller with no viewer needs whole-folder pull and hand-seeding the index is error-prone), **or** an OKF adopter without the viewer asks for it. Until then the trick + the button cover it.

**Seam already in place:** the `mode` switch in `StickShiftContextBundle.bas` (the `If mode = "index" ... ElseIf mode = "all"` ladder) is where a `folder` branch would slot.

## Section-patch / append write mode

**Idea:** a write envelope that rewrites *part* of a file (a section, or an append) instead of replacing the whole file, to cut tokens-per-second on edits to long concepts.

**Shipped instead:** document sharding. Splitting a long concept into a folder of small section files means an edit re-emits one small file, getting the same tokens/sec win structurally — without a second envelope grammar, merge semantics, or a new failure surface. It also preserves the load-bearing invariant that *every write restates a whole, schema-valid file*.

**Why deferred:** patch-mode is a much larger change than sharding and weakens an invariant the system leans on. Sharding makes files small enough that whole-file replacement is already cheap.

**Revisit trigger:** a single *section* file routinely grows long enough that re-emitting it whole is itself the friction — i.e. sharding alone stops being enough. At that point patch-mode (or finer recursive sharding) earns its complexity.

## Edit-diff capture as a skill-tuning flywheel

**Idea:** gate the VBA write so that every human edit committed through the viewer stores a before/after snapshot plus a one-line "why" (git-commit style) in a diff file. The accumulated diffs are labeled data — "the bundle said X, I changed it to Y" — that surface which parts of a context bundle get hand-fixed repeatedly, telling us what to bake into a skill or system-prompt so the fix stops being manual.

**Status: not built, not yet spec'd.** Designed in the *StickShift development quirks* session (2026-06-27) but no diff/reason/review code exists in any module or the viewer yet. The crystallized concept at the end of that session — the HTML write-review surface (agent proposes -> PENDING -> rendered for review -> APPROVED -> committed, where approval both commits the write and fires the capture) — was never written to a committed spec.

**Design decisions already made (carry these into the spec):**
1. **Gate at the VBA write button** — the commit point, the one moment "what it was" (on disk) and "what it's becoming" (in the packet) both exist in the same hand. Instrument the write, not the writer.
2. **Key the diff on a stable `file_id`, not the path** — so history survives sharding/rename. Same lesson as the Day-1 state-drift bug: do not orphan a record from the thing it points at.
3. **Store fat now, compute lean later** — full before/after snapshots per write, not deltas. Disk is free at this scale; you can derive deltas from snapshots but never recover snapshots from deltas.
4. **Capture the reason at write time** — a half-sentence prompted on save, stored in the diff file's OKF front matter. The diff records *that* an edit happened; the reason records *why*, while it is still in the operator's head.
5. **Write order (crash-safe):** write the AFTER file to the diff file, append the BEFORE content under an explicit header, then overwrite the live file location.

**Open question to resolve in the spec:** per-skill drift signal vs. global signal — per-skill tells you "this skill drifts in this direction"; global surfaces taste-level patterns ("I always tighten verbosity"). Decide before wiring the gate.

**Revisit trigger:** ready to build now — this is the next spec to write, not a deferred idea. Listed here so the design decisions are not re-derived from scratch.
