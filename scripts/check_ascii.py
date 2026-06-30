#!/usr/bin/env python3
"""
StickShift stack guard: the single source of truth for the encoding rules that
both the local pre-commit hook and CI enforce.

Rules (mirrors what breaks CI today):
  - builds/*.bas                 must be pure ASCII (VBE import mangles non-ASCII)
  - builds/*.bas                 must not contain ChrW (use plain ASCII literals)
  - builds/html-tools/**/*.html  must be pure ASCII (air-gapped tool hygiene;
                                 the *_ascii_only compliance tests enforce the same)

Exit 0 if clean, 1 if any violation. Run from anywhere -- paths resolve against the
repo root (this file's grandparent), so the hook and CI agree regardless of CWD.

    python scripts/check_ascii.py
"""

import sys
import glob
import unicodedata
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent


def _ascii_violations(path: Path):
    """Yield (line, col, byte_index, char) once per non-ASCII character."""
    text = path.read_text(encoding="utf-8", errors="replace")
    line = 1
    col = 1
    byte_index = 0
    for ch in text:
        if ch == "\n":
            line += 1
            col = 1
            byte_index += 1
            continue
        if ord(ch) > 0x7F:
            yield (line, col, byte_index, ch)
        byte_index += len(ch.encode("utf-8"))
        col += 1


def check_ascii(rel_glob, recursive=False):
    findings = []
    for f in sorted(glob.glob(str(REPO / rel_glob), recursive=recursive)):
        p = Path(f)
        for line, col, idx, ch in _ascii_violations(p):
            name = unicodedata.name(ch, "UNKNOWN")
            rel = p.relative_to(REPO).as_posix()
            findings.append(
                f"  {rel}:{line}:{col}  non-ASCII U+{ord(ch):04X} {name} ({ch!r}) at byte {idx}"
            )
    return findings


def check_chrw(rel_glob):
    findings = []
    for f in sorted(glob.glob(str(REPO / rel_glob))):
        p = Path(f)
        text = p.read_text(encoding="utf-8", errors="replace")
        for n, ln in enumerate(text.splitlines(), start=1):
            if "ChrW" in ln:
                rel = p.relative_to(REPO).as_posix()
                findings.append(f"  {rel}:{n}  ChrW found -- use a plain ASCII literal")
    return findings


def main():
    problems = []

    bas_ascii = check_ascii("builds/*.bas")
    if bas_ascii:
        problems.append(("Non-ASCII in builds/*.bas", bas_ascii))

    bas_chrw = check_chrw("builds/*.bas")
    if bas_chrw:
        problems.append(("ChrW in builds/*.bas", bas_chrw))

    html_ascii = check_ascii("builds/html-tools/**/*.html", recursive=True)
    if html_ascii:
        problems.append(("Non-ASCII in builds/html-tools/**/*.html", html_ascii))

    if not problems:
        print("stack guard passed: .bas and html-tools .html are pure ASCII, no ChrW.")
        return 0

    print("STACK GUARD FAILED -- fix before committing (CI will reject these):\n")
    for title, findings in problems:
        print(title + ":")
        for line in findings:
            print(line)
        print()
    print("Tip: smart quotes and em dashes are the usual culprits "
          "-- replace with \" ' and - .")
    return 1


if __name__ == "__main__":
    sys.exit(main())
