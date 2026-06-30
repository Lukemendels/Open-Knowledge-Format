# Build prompt (verification loop): Edit-Diff Capture

You are implementing `_meta/spec-edit-capture.md` in this repository. Read that spec
first; it is the source of truth for *what* to build. This prompt defines *how* to
work: a closed loop where you implement, run verification, read the failures, fix, and
repeat until verification is green — then stop and report. Do not ask me questions
mid-loop unless you are truly blocked; prefer the spec's stated decisions.

---

## GOAL (definition of done)

All of the following are simultaneously true:

1. `pytest tests/ -q` passes with **zero** failures and zero skips/xfails that you added.
2. The **ASCII guard** passes: `grep -rIP "[^\x00-\x7F]" builds/*.bas` prints nothing.
3. The **ChrW guard** passes: `grep -rIn "ChrW" builds/*.bas` prints nothing.
4. New twin tests exist and pass for the three pure-logic behaviors in spec §9:
   reason-line detection, `EnsureConceptId`, and diff-record body format
   (edit -> AFTER+BEFORE; new -> AFTER only).
5. **Every pre-existing test and golden vector still passes, unmodified.** The
   `### FILE:` parse/compute contract is unchanged once the optional `reason:` line is
   stripped.
6. An **untracked** write (a `<VBA_WRITE>` with no `reason:` line) produces byte-identical
   results to the current code: same files written, same `log.md` lines, no diff file.

You are done only when 1–6 hold at the same time on a clean run.

---

## VERIFICATION (run this exact block every iteration)

```bash
pip install pytest -q
echo "== pytest ==" && pytest tests/ -q
echo "== ascii guard (want: no output) ==" && grep -rIP "[^\x00-\x7F]" builds/*.bas || echo "ASCII OK"
echo "== chrw guard (want: no output) ==" && grep -rIn "ChrW" builds/*.bas || echo "ChrW OK"
```

Treat the run as green only if pytest reports all-passed AND both guards print their OK
line (no matches). Run the **whole** block each time — not just your new test — so a
regression in an existing golden vector cannot hide.

---

## THE LOOP

1. Make the smallest meaningful change toward the goal (one helper, one behavior).
2. Run the full VERIFICATION block.
3. If green on all of 1–6, go to COMPLETION REPORT.
4. If red: read the actual failure output. Diagnose the real cause. Fix the
   implementation. Return to step 2.
5. Repeat. If you are not green after ~8 focused iterations, **stop** and report what is
   blocking you (see "If you get stuck"). Do not keep thrashing and do not hack.

---

## RULES (these are what make the loop trustworthy — do not break them)

- **Never edit, weaken, skip, `xfail`, or delete a pre-existing test or golden vector to
  get green.** If an existing test fails, your change is wrong — fix your code, not the
  test. The existing suite is the contract.
- **Never weaken or remove the ASCII or ChrW guards**, and never satisfy them by deleting
  `.bas` content the spec requires. Write the VBA in pure ASCII with no `ChrW`, as the
  spec already constrains.
- **The twin must stay a faithful mirror of the VBA.** The Python twin (in the test file)
  and the VBA in `builds/*.bas` must describe the *same* behavior. If you change one,
  change the other to match — never make them agree by making the twin wrong or by
  loosening an assertion. A green twin that does not reflect the VBA is a failed build,
  not a passed one.
- **Do not reduce scope to pass.** All three implementation items below must be present.
- **Do not invent behavior the spec does not specify.** Where the spec made a decision
  (global capture, lazy `id` in frontmatter, per-batch reason, diff-first/live-last
  order, `reason:` line as the capture trigger), implement that decision as written.

---

## SCOPE

**In scope** (spec §10 items 1–3):
1. `StickShiftConfig.bas`: `EditStoreDir()` — copy `DistDir()`, swap `-dist` -> `-edits`.
2. `StickShiftWriteApply.bas`: reason-line detection (§6.1), gated capture in the per-file
   loop (§6.2), and the private helpers `CaptureEditDiff` (§6.3), `EnsureConceptId` (§4),
   `FrontmatterValue`, `SanitizeId` (§6.4). Capture runs *before* the live `WriteUtf8`;
   `WriteUtf8` writes the id-injected `afterContent`. Public contract
   (`ApplyWriteEnvelopeText` / `ApplyStickShiftWrite`), reserved-file guard, log format,
   and auto-reindex are unchanged.
3. The oracle `tests/test_okf_edit_capture.py` **ALREADY EXISTS in the repo and is
   verified green** — do NOT recreate, overwrite, or weaken it. Make the VBA satisfy it.
   If you need further twin coverage, ADD assertions/functions to that file without
   altering or loosening the existing ones (it mirrors the inline-twin style of
   `tests/test_okf_write_apply.py`).

**Out of scope — do NOT build:** the viewer reason-prompt modal (spec §7), the
PENDING/APPROVED review surface, and any diff-mining / skill-tuning tool (spec §8). Leave
them alone.

---

## WHAT THIS LOOP CANNOT VERIFY (hand these back to me)

The loop closes around pure logic and the static guards. It runs on Linux with no Excel,
so it **cannot** execute the VBA or touch the real filesystem. These remain human-verified
on Windows — green here means "logic verified, ready for the Excel smoke test," not "done":

- That `CaptureEditDiff` actually runs **before** the live `WriteUtf8` (crash-safe order),
  proven by stepping the macro and confirming the live file still holds BEFORE at the
  breakpoint after capture.
- That the diff file lands in the real `-edits` sibling, UTF-8, with AFTER then BEFORE.
- That `EnsureConceptId` injects the `id` into a real on-disk file's frontmatter and the
  live file ends byte-consistent with the diff's AFTER section.
- That an untracked write on Windows is unchanged.

Produce these as a short manual checklist in your completion report so I can run them in
Excel.

---

## IF YOU GET STUCK

Stop after ~8 iterations without green. Report: the last VERIFICATION output, which of
goals 1–6 are failing, your best diagnosis of why, and the options you see. Do not modify
existing tests or guards to escape. A clear blocker is a better outcome than a green
build that cheated.

---

## COMPLETION REPORT (only when 1–6 are green)

Output:
1. The final VERIFICATION block output, showing pytest all-passed and both guards OK.
2. A file-by-file summary of what changed and why (tie each to a spec section).
3. A 2–4 line **twin/VBA correspondence note**: state that the Python twin and the VBA
   implement the same behavior, and name the one or two places that are easiest to drift
   so I know where to look during the Excel smoke test.
4. The **manual checklist** from the section above, ready for me to run on Windows.
