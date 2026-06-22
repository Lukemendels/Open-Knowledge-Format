"""
Guard test: every xl*/mso* name *used* in builds/*.bas must also appear as a
Private Const declaration in the same file.

This catches the class of bug where a late-binding refactor drops a type-library
reference but leaves the enum constant names in the code.  VBA reports these as
undefined variables under Option Explicit, one at a time, so the defect looks
like an endless "fixed it, still won't compile" loop.  This test makes the whole
class CI-catchable without a VBA engine.
"""

import re
from pathlib import Path

import pytest

BUILDS_DIR = Path(__file__).parent.parent / "builds"

# Pattern that matches xl* or mso* identifiers.
_IDENT_RE = re.compile(r'\b((?:xl|mso)[A-Za-z]+)\b')

# Pattern that matches a Private Const declaration of an xl*/mso* name.
# Accepts both single and double spaces between keywords.
_CONST_RE = re.compile(r'^\s*Private\s+Const\s+((?:xl|mso)[A-Za-z]+)\b', re.MULTILINE)


def _used_names(src: str) -> set[str]:
    """All xl*/mso* identifiers that appear anywhere in *src*."""
    return set(_IDENT_RE.findall(src))


def _declared_names(src: str) -> set[str]:
    """All xl*/mso* names declared as Private Const in *src*."""
    return set(_CONST_RE.findall(src))


@pytest.mark.parametrize("bas_file", sorted(BUILDS_DIR.glob("*.bas")))
def test_all_used_constants_are_declared(bas_file: Path) -> None:
    src = bas_file.read_text(encoding="utf-8")
    used = _used_names(src)
    declared = _declared_names(src)
    undeclared = used - declared
    assert not undeclared, (
        f"{bas_file.name}: xl*/mso* names used but not declared as Private Const: "
        + ", ".join(sorted(undeclared))
    )
