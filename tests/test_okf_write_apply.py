"""
Python twin of OKFWriteApply.bas — write envelope parser and gate logic.

Keeps the same behavioural contract as the VBA macro:
  - parse_write_envelope: extract (path, contents) pairs from a <VBA_WRITE> block.
  - compute_operations:   apply the gate — new file → write, existing → .proposed.
  - Never produces a delete operation.

Golden vectors cover: single new file, single existing file, mixed batch,
empty envelope, and the no-block error path.
"""

import re
import pytest


# ── Twin logic ──────────────────────────────────────────────────────────────────────────

def parse_write_envelope(text: str) -> list[tuple[str, str]]:
    """
    Extract files from a <VBA_WRITE> envelope.

    Returns a list of (path, contents) tuples in the order they appear.
    Raises ValueError if no <VBA_WRITE> block is present.
    """
    match = re.search(r"<VBA_WRITE>(.*?)</VBA_WRITE>", text, re.DOTALL)
    if not match:
        raise ValueError("No <VBA_WRITE> block found in input")

    body = match.group(1)
    # Normalise CRLF so the regex below works regardless of clipboard line-endings.
    body = body.replace("\r\n", "\n").replace("\r", "\n")

    result = []
    # Each block: ### FILE: <path>\n<contents>\n### END FILE
    block_re = re.compile(r"### FILE:\s*(.+?)\n(.*?)### END FILE", re.DOTALL)
    for m in block_re.finditer(body):
        path = m.group(1).strip()
        contents = m.group(2)
        # Strip the single separator newline that follows ### FILE: <path>
        if contents.startswith("\n"):
            contents = contents[1:]
        # Strip the trailing newline that precedes ### END FILE
        contents = contents.rstrip("\n")
        result.append((path, contents))

    return result


def compute_operations(
    parsed_files: list[tuple[str, str]],
    existing_files: set[str],
) -> list[tuple[str, str, str]]:
    """
    Determine file operations from a parsed envelope.

    For each (path, contents):
      path NOT in existing_files  →  ('write',  path,              contents)
      path IN  existing_files     →  ('staged', path + '.proposed', contents)

    No delete operation is ever produced.
    """
    ops = []
    for path, contents in parsed_files:
        if path in existing_files:
            ops.append(("staged", path + ".proposed", contents))
        else:
            ops.append(("write", path, contents))
    return ops


# ── Shared fixture content ────────────────────────────────────────────────────────────────────

_CONTENTS_A = """\
---
type: Build
title: Build A
description: First fixture build.
status: idea
effort: S
impact: high
domain: TSA
timestamp: 2026-06-18T00:00:00Z
last_touched: 2026-06-18
---

# What it is
Fixture build A.

# Next action
Do something.

# Dependencies
(none)

# Notes
(none)"""

_CONTENTS_B = """\
---
type: Build
title: Build B
description: Second fixture build.
status: production
effort: S
impact: high
domain: TSA
timestamp: 2026-06-18T00:00:00Z
last_touched: 2026-06-18
---

# What it is
Fixture build B.

# Next action
Monitor.

# Dependencies
(none)

# Notes
(none)"""


def _envelope(*path_content_pairs: tuple[str, str]) -> str:
    parts = []
    for path, contents in path_content_pairs:
        parts.append(f"### FILE: {path}\n{contents}\n### END FILE")
    return "<VBA_WRITE>\n" + "\n".join(parts) + "\n</VBA_WRITE>"


# ── Golden vectors ─────────────────────────────────────────────────────────────────────────────

GOLDEN_VECTORS = [
    {
        "name": "single_new_file",
        "envelope": _envelope(("builds/new-build.md", _CONTENTS_A)),
        "existing": set(),
        "expected_ops": [("write", "builds/new-build.md", _CONTENTS_A)],
    },
    {
        "name": "single_existing_file_stages_as_proposed",
        "envelope": _envelope(("builds/existing-build.md", _CONTENTS_A)),
        "existing": {"builds/existing-build.md"},
        "expected_ops": [("staged", "builds/existing-build.md.proposed", _CONTENTS_A)],
    },
    {
        "name": "mixed_new_and_existing",
        "envelope": _envelope(
            ("builds/new-build.md", _CONTENTS_A),
            ("builds/existing-build.md", _CONTENTS_B),
        ),
        "existing": {"builds/existing-build.md"},
        "expected_ops": [
            ("write", "builds/new-build.md", _CONTENTS_A),
            ("staged", "builds/existing-build.md.proposed", _CONTENTS_B),
        ],
    },
    {
        "name": "empty_envelope_yields_no_ops",
        "envelope": "<VBA_WRITE>\n</VBA_WRITE>",
        "existing": set(),
        "expected_ops": [],
    },
    {
        "name": "two_new_files",
        "envelope": _envelope(
            ("builds/alpha.md", _CONTENTS_A),
            ("builds/beta.md", _CONTENTS_B),
        ),
        "existing": set(),
        "expected_ops": [
            ("write", "builds/alpha.md", _CONTENTS_A),
            ("write", "builds/beta.md", _CONTENTS_B),
        ],
    },
    {
        "name": "two_existing_files_both_staged",
        "envelope": _envelope(
            ("builds/alpha.md", _CONTENTS_A),
            ("builds/beta.md", _CONTENTS_B),
        ),
        "existing": {"builds/alpha.md", "builds/beta.md"},
        "expected_ops": [
            ("staged", "builds/alpha.md.proposed", _CONTENTS_A),
            ("staged", "builds/beta.md.proposed", _CONTENTS_B),
        ],
    },
]


# ── Tests ─────────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("vec", GOLDEN_VECTORS, ids=[v["name"] for v in GOLDEN_VECTORS])
def test_golden_vector(vec: dict) -> None:
    parsed = parse_write_envelope(vec["envelope"])
    ops = compute_operations(parsed, vec["existing"])
    assert ops == vec["expected_ops"]


def test_no_vba_write_block_raises_value_error() -> None:
    with pytest.raises(ValueError, match="No <VBA_WRITE>"):
        parse_write_envelope("just some text with no envelope")


def test_preamble_and_postamble_ignored() -> None:
    """Prose before and after <VBA_WRITE> is silently discarded."""
    text = (
        "Here is the output you requested:\n\n"
        + _envelope(("builds/new.md", _CONTENTS_A))
        + "\n\nLet me know if you need changes."
    )
    parsed = parse_write_envelope(text)
    assert len(parsed) == 1
    assert parsed[0][0] == "builds/new.md"


def test_proposed_path_has_dot_proposed_suffix() -> None:
    env = _envelope(("builds/foo.md", _CONTENTS_A))
    parsed = parse_write_envelope(env)
    ops = compute_operations(parsed, {"builds/foo.md"})
    assert ops[0][1] == "builds/foo.md.proposed"


def test_original_path_unchanged_for_new_file() -> None:
    env = _envelope(("builds/bar.md", _CONTENTS_A))
    parsed = parse_write_envelope(env)
    ops = compute_operations(parsed, set())
    assert ops[0][1] == "builds/bar.md"


def test_no_delete_operations_ever() -> None:
    env = _envelope(
        ("builds/new.md", _CONTENTS_A),
        ("builds/old.md", _CONTENTS_B),
    )
    parsed = parse_write_envelope(env)
    ops = compute_operations(parsed, {"builds/old.md"})
    for op_type, _, _ in ops:
        assert op_type in ("write", "staged"), f"unexpected op type: {op_type}"


def test_contents_preserved_exactly() -> None:
    env = _envelope(("builds/precise.md", _CONTENTS_A))
    parsed = parse_write_envelope(env)
    assert parsed[0][1] == _CONTENTS_A


def test_crlf_envelope_normalised() -> None:
    """Windows CRLF line endings in the envelope are handled correctly."""
    env = _envelope(("builds/win.md", _CONTENTS_A)).replace("\n", "\r\n")
    parsed = parse_write_envelope(env)
    assert len(parsed) == 1
    assert parsed[0][0] == "builds/win.md"
