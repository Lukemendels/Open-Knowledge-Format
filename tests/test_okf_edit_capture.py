"""
Python twin of the edit-capture additions to StickShiftWriteApply.bas.

These are the ORACLE for the verification loop: the VBA in builds/*.bas must implement
the same behavior these twins describe. Treat this file as authoritative over any prose
sketch in spec-edit-capture.md sec.6 -- where the spec's pseudocode and these tests
disagree, the tests are correct (they encode two subtleties the prose glossed: a leading
newline before the `reason:` line, and the fallback-key argument to ensure_concept_id).

Behaviors covered:
  - detect_reason:      a leading `reason:` line flips capture on and is stripped; its
                        absence leaves the body byte-identical (the untracked contract).
  - ensure_concept_id:  existing id preserved; missing id injected once after the opening
                        ---; no-frontmatter file falls back to a path-keyed id.
  - frontmatter_value:  read one scalar from the leading frontmatter block.
  - build_diff_record:  edit -> AFTER + BEFORE; new -> AFTER only; carries id/type/reason.
"""

import pytest


# -- Twin logic (mirror of the VBA) ---------------------------------------------------

def detect_reason(body):
    """
    Mirror of sec.6.1. Returns (capture_on, reason_text, body_for_file_parser).

    The envelope body (text between <VBA_WRITE> and </VBA_WRITE>) normally begins with a
    newline, so the `reason:` line -- when present -- is the first NON-EMPTY line. When no
    reason line is present, the original body is returned untouched so an untracked write
    is byte-identical to today.
    """
    norm = body.replace("\r\n", "\n").replace("\r", "\n")
    leading = norm.lstrip("\n")
    nl = leading.find("\n")
    first_line = (leading if nl == -1 else leading[:nl]).strip()
    if first_line[:7].lower() == "reason:":
        reason = first_line[7:].strip()
        rest = "" if nl == -1 else leading[nl + 1:]
        return True, reason, rest
    return False, "", body


def frontmatter_value(content, key):
    """Mirror of FrontmatterValue: one scalar from the leading --- block, else ''."""
    lines = content.replace("\r\n", "\n").split("\n")
    if not lines or lines[0].strip() != "---":
        return ""
    for ln in lines[1:]:
        if ln.strip() == "---":
            break
        s = ln.strip()
        if s.lower().startswith(key.lower() + ":"):
            return s[len(key) + 1:].strip()
    return ""


def ensure_concept_id(content, seq, now_str, fallback_key):
    """
    Mirror of sec.4 EnsureConceptId. Returns (concept_id, possibly_mutated_content, source).

    - frontmatter with an id:   return it, content unchanged.
    - frontmatter without an id: generate c-<now>-<seq>, insert after the opening ---.
    - no frontmatter at all:     fall back to "path:<fallback_key>", content unchanged.
    """
    norm = content.replace("\r\n", "\n")
    lines = norm.split("\n")
    if lines and lines[0].strip() == "---":
        for ln in lines[1:]:
            if ln.strip() == "---":
                break
            s = ln.strip()
            if s.lower().startswith("id:"):
                return s[3:].strip(), content, "frontmatter-existing"
        new_id = "c-{}-{}".format(now_str, seq)
        lines.insert(1, "id: " + new_id)
        return new_id, "\n".join(lines), "frontmatter-injected"
    return "path:" + fallback_key, content, "path"


def build_diff_record(cid, rel_path, after, before, action, reason, concept_type, captured):
    """Mirror of CaptureEditDiff's record body (sec.6.3)."""
    out = "---\n"
    out += 'okf_version: "0.1"\n'
    out += "type: EditDiff\n"
    out += "id: " + cid + "\n"
    out += "path_at_write: " + rel_path + "\n"
    out += "concept_type: " + concept_type + "\n"
    out += "action: " + action + "\n"
    out += "reason: " + reason + "\n"
    out += "captured: " + captured + "\n"
    out += "---\n\n"
    out += "## AFTER\n\n" + after + "\n\n"
    if action == "edit":
        out += "## BEFORE\n\n" + before + "\n"
    return out


# -- Golden vectors -------------------------------------------------------------------

UNTRACKED_BODY = "\n### FILE: builds/a.md\nhello\n### END FILE\n"
TRACKED_BODY = "\nreason: tightened wage citation\n### FILE: builds/a.md\nhello\n### END FILE\n"

FM_WITH_ID = '---\nokf_version: "0.1"\ntype: Build\nid: c-existing-1\ntitle: A\n---\n\nbody\n'
FM_NO_ID = '---\nokf_version: "0.1"\ntype: Skill\ntitle: A\n---\n\nbody\n'
NO_FRONTMATTER = "# just a stub\n\npointer to builds/a/index.md\n"


def test_reason_detected_and_stripped():
    on, reason, rest = detect_reason(TRACKED_BODY)
    assert on is True
    assert reason == "tightened wage citation"
    # the stripped body is exactly the untracked body (reason line is the only difference)
    assert rest == UNTRACKED_BODY.lstrip("\n")


def test_no_reason_is_byte_identical():
    on, reason, rest = detect_reason(UNTRACKED_BODY)
    assert on is False
    assert reason == ""
    assert rest == UNTRACKED_BODY  # untouched -> untracked write unchanged


def test_reason_requires_first_nonempty_line():
    # a `reason:` that appears AFTER a FILE block is not a capture trigger
    body = "\n### FILE: builds/a.md\nreason: not me\n### END FILE\n"
    on, _, rest = detect_reason(body)
    assert on is False
    assert rest == body


def test_ensure_id_preserves_existing():
    cid, content, source = ensure_concept_id(FM_WITH_ID, seq=0, now_str="20260630090000",
                                             fallback_key="builds/a.md")
    assert cid == "c-existing-1"
    assert source == "frontmatter-existing"
    assert content == FM_WITH_ID  # unchanged


def test_ensure_id_injects_when_missing():
    cid, content, source = ensure_concept_id(FM_NO_ID, seq=2, now_str="20260630090000",
                                             fallback_key="builds/a.md")
    assert cid == "c-20260630090000-2"
    assert source == "frontmatter-injected"
    # injected exactly once, immediately after the opening ---
    assert content.split("\n")[1] == "id: c-20260630090000-2"
    assert content.count("id: c-20260630090000-2") == 1
    # and it is now readable as a frontmatter scalar
    assert frontmatter_value(content, "id") == "c-20260630090000-2"


def test_ensure_id_path_fallback_without_frontmatter():
    cid, content, source = ensure_concept_id(NO_FRONTMATTER, seq=0, now_str="x",
                                             fallback_key="builds/a.md")
    assert cid == "path:builds/a.md"
    assert source == "path"
    assert content == NO_FRONTMATTER  # never invents a frontmatter block


def test_frontmatter_value_reads_type_and_missing():
    assert frontmatter_value(FM_NO_ID, "type") == "Skill"
    assert frontmatter_value(NO_FRONTMATTER, "type") == ""


def test_diff_record_edit_has_after_and_before():
    rec = build_diff_record(
        cid="c-1", rel_path="builds/a.md", after="NEW", before="OLD",
        action="edit", reason="why", concept_type="Build",
        captured="2026-06-30 09:00:00",
    )
    assert "## AFTER\n\nNEW" in rec
    assert "## BEFORE\n\nOLD" in rec
    assert "id: c-1" in rec
    assert "concept_type: Build" in rec
    assert "reason: why" in rec


def test_diff_record_new_has_after_only():
    rec = build_diff_record(
        cid="c-2", rel_path="builds/b.md", after="NEW", before="",
        action="new", reason="created", concept_type="Build",
        captured="2026-06-30 09:00:00",
    )
    assert "## AFTER\n\nNEW" in rec
    assert "## BEFORE" not in rec  # no before section for a new file


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
