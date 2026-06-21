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
