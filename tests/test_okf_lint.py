"""
Python twin of OKFLint.bas — bundle integrity checker.

Implements the same seven checks as the VBA macro:
  1. Missing / empty `type` field
  2. Missing required fields: status, effort, impact
  3. Broken cross-links (.md links whose target file is absent)
  4. WIP violation (> 1 build at status:working)
  5. Stalls (working builds with missing or old last_touched)
  6. Pending .proposed files awaiting review
  7. Active-to-archived links

Golden vectors use two fixture bundles:
  fixtures/clean_bundle   — no findings expected
  fixtures/broken_bundle  — deliberately broken, specific findings expected
"""

import re
from dataclasses import dataclass
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

import pytest


FIXTURES = Path(__file__).parent / "fixtures"
CLEAN_BUNDLE = FIXTURES / "clean_bundle"
BROKEN_BUNDLE = FIXTURES / "broken_bundle"

# Fixed reference date for deterministic tests.
REF_DATE = date(2026, 6, 18)


# ── Data types ────────────────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Finding:
    severity: str   # 'error' | 'warning'
    file: str       # basename of the offending file, or '(portfolio)'
    message: str    # one-line description


# ── Twin logic ──────────────────────────────────────────────────────────────────────────

def _parse_frontmatter(content: str) -> dict[str, str]:
    """
    Extract top-level scalar frontmatter fields from a YAML front-matter block.
    Returns an empty dict if no frontmatter is present.
    """
    content = content.replace("\r\n", "\n").replace("\r", "\n")
    lines = content.split("\n")
    if not lines or lines[0].strip() != "---":
        return {}

    fm: dict[str, str] = {}
    for line in lines[1:]:
        if line.strip() == "---":
            break
        if ":" in line:
            key, _, val = line.partition(":")
            key = key.strip().lower()
            val = val.strip().strip("\"'")
            fm[key] = val
    return fm


def _extract_links(content: str) -> list[str]:
    """
    Return all local .md link URLs found in a markdown document.
    Strips anchors; skips http/https links.
    """
    links = []
    for m in re.finditer(r"\]\(([^)]+)\)", content):
        url = m.group(1).strip()
        if url.lower().startswith(("http://", "https://")):
            continue
        # Strip anchor fragment before checking extension.
        base = url.split("#")[0]
        if base.lower().endswith(".md"):
            links.append(url)
    return links


def _resolve_link(link: str, from_file: Path, bundle_root: Path) -> Optional[Path]:
    """
    Resolve a markdown link URL to an absolute Path.
    Returns None for external links, empty links, or non-.md links.
    """
    # Strip anchor.
    base = link.split("#")[0].strip()
    if not base:
        return None
    if base.lower().startswith(("http://", "https://")):
        return None
    if not base.lower().endswith(".md"):
        return None

    if base.startswith("/"):
        # Root-relative: /builds/foo.md → bundle_root/builds/foo.md
        return bundle_root / base.lstrip("/")
    else:
        # Relative to the file's own directory.
        return from_file.parent / base


_SEVERITY_RANK = {"error": 0, "warning": 1}


def lint_bundle(
    bundle_root: Path,
    stale_days: int = 14,
    reference_date: Optional[date] = None,
) -> list[Finding]:
    """
    Scan a bundle directory and return all integrity findings, sorted by
    (severity rank, file name).  Errors precede warnings; within each severity
    findings are sorted alphabetically by file name for deterministic output.
    """
    if reference_date is None:
        reference_date = date.today()

    findings: list[Finding] = []
    builds_dir = bundle_root / "builds"

    if not builds_dir.exists():
        return [Finding("error", "(bundle)", f"builds/ directory not found: {builds_dir}")]

    # Concept files: .md files that are not index.md or log.md.
    concept_files = sorted(
        f for f in builds_dir.iterdir()
        if f.is_file()
        and f.suffix == ".md"
        and f.name not in ("index.md", "log.md")
    )

    fm_by_path: dict[str, dict[str, str]] = {}
    links_by_path: dict[str, list[str]] = {}
    working_builds: list[tuple[str, Path]] = []  # (last_touched, path)

    # ── Phase 1: per-file checks ────────────────────────────────────────────────────
    for f in concept_files:
        content = f.read_text(encoding="utf-8")
        fm = _parse_frontmatter(content)
        fm_by_path[str(f)] = fm

        # Check 1: missing / empty type
        if not fm.get("type"):
            findings.append(Finding("error", f.name, "missing or empty `type` field"))

        # Check 2: missing required fields
        for field in ("status", "effort", "impact"):
            if not fm.get(field):
                findings.append(Finding("error", f.name, f"missing required field `{field}`"))

        # Collect links for cross-file checks.
        links_by_path[str(f)] = _extract_links(content)

        # Accumulate working builds for WIP + stall.
        if fm.get("status", "").lower() == "working":
            working_builds.append((fm.get("last_touched", ""), f))

    # ── Phase 2: cross-file checks ────────────────────────────────────────────────────

    # Check 3: broken links
    for file_path_str, links in links_by_path.items():
        f = Path(file_path_str)
        for link in links:
            resolved = _resolve_link(link, f, bundle_root)
            if resolved is not None and not resolved.exists():
                findings.append(Finding("error", f.name, f"broken link: {link}"))

    # Check 4: WIP violation
    if len(working_builds) > 1:
        names = ", ".join(p.name for _, p in working_builds)
        findings.append(
            Finding(
                "error",
                "(portfolio)",
                f"WIP violation: {len(working_builds)} builds at status:working ({names})",
            )
        )

    # Check 5: stalls — flag working builds with missing or old last_touched,
    #          report oldest first.
    stale_threshold = reference_date - timedelta(days=stale_days)
    stale_findings: list[tuple[str, Finding]] = []  # (sort_key, Finding)

    for lt_str, f in working_builds:
        if not lt_str:
            sort_key = "0000-00-00"
            msg = "working build has missing last_touched"
        else:
            try:
                lt_date = date.fromisoformat(lt_str)
            except ValueError:
                sort_key = "0000-00-00"
                msg = f"working build has invalid last_touched: {lt_str!r}"
            else:
                if lt_date <= stale_threshold:
                    age = (reference_date - lt_date).days
                    sort_key = lt_str
                    msg = f"stale working build: last_touched {lt_str} ({age} days ago)"
                else:
                    continue  # fresh — not stale

        stale_findings.append((sort_key, Finding("warning", f.name, msg)))

    # Sort stale findings oldest-first, then add in that order.
    stale_findings.sort(key=lambda t: t[0])
    for _, finding in stale_findings:
        findings.append(finding)

    # Check 6: pending .proposed files (recursive scan of whole bundle).
    for proposed in sorted(bundle_root.rglob("*.proposed")):
        rel = proposed.relative_to(bundle_root)
        findings.append(
            Finding("warning", proposed.name, f"pending .proposed awaiting review: {rel}")
        )

    # Check 7: active-to-archived links.
    archived_paths = {
        path_str
        for path_str, fm in fm_by_path.items()
        if fm.get("status", "").lower() == "archived"
    }

    for file_path_str, links in links_by_path.items():
        f = Path(file_path_str)
        if fm_by_path.get(file_path_str, {}).get("status", "").lower() == "archived":
            continue  # archived→archived is fine

        for link in links:
            resolved = _resolve_link(link, f, bundle_root)
            if resolved is not None and str(resolved) in archived_paths:
                findings.append(
                    Finding("warning", f.name, f"active build links to archived build: {link}")
                )

    # Sort by (severity rank, file name) for deterministic output.
    findings.sort(key=lambda x: (_SEVERITY_RANK.get(x.severity, 99), x.file))
    return findings


# ── Tests ─────────────────────────────────────────────────────────────────────────────────

def test_clean_bundle_produces_no_findings() -> None:
    findings = lint_bundle(CLEAN_BUNDLE, reference_date=REF_DATE)
    assert findings == [], f"Expected no findings, got: {findings}"


class TestBrokenBundle:
    """All checks against the deliberately broken fixture bundle."""

    def setup_method(self) -> None:
        self.findings = lint_bundle(BROKEN_BUNDLE, stale_days=14, reference_date=REF_DATE)
        self.errors = [f for f in self.findings if f.severity == "error"]
        self.warnings = [f for f in self.findings if f.severity == "warning"]

    # ── Check 1: missing type ───────────────────────────────────────────────────────────────────

    def test_missing_type_is_flagged_as_error(self) -> None:
        matches = [f for f in self.errors if f.file == "no-type-build.md"]
        assert len(matches) == 1
        assert "type" in matches[0].message

    # ── Check 3: broken link ────────────────────────────────────────────────────────────────────

    def test_broken_link_is_flagged_as_error(self) -> None:
        matches = [f for f in self.errors if f.file == "broken-link-build.md"]
        assert len(matches) >= 1
        assert any("broken link" in f.message and "nonexistent.md" in f.message for f in matches)

    # ── Check 4: WIP violation ────────────────────────────────────────────────────────────────────

    def test_wip_violation_is_flagged_as_error(self) -> None:
        wip = [f for f in self.errors if "(portfolio)" in f.file]
        assert len(wip) == 1
        assert "WIP violation" in wip[0].message
        assert "2 builds" in wip[0].message

    def test_wip_violation_names_both_working_builds(self) -> None:
        wip = [f for f in self.errors if "(portfolio)" in f.file][0]
        assert "working-fresh.md" in wip.message
        assert "working-stale.md" in wip.message

    # ── Check 5: stalls ─────────────────────────────────────────────────────────────────────────────

    def test_stale_working_build_is_flagged_as_warning(self) -> None:
        stale = [f for f in self.warnings if f.file == "working-stale.md"]
        assert len(stale) == 1
        assert "stale" in stale[0].message
        assert "2020-01-01" in stale[0].message

    def test_fresh_working_build_not_flagged_as_stale(self) -> None:
        """working-fresh.md is only 3 days old — must not trigger a stale warning."""
        stale_fresh = [
            f for f in self.warnings
            if f.file == "working-fresh.md" and "stale" in f.message
        ]
        assert stale_fresh == []

    # ── Check 6: pending .proposed ───────────────────────────────────────────────────────────────────

    def test_pending_proposed_file_is_flagged_as_warning(self) -> None:
        proposed = [f for f in self.warnings if ".proposed" in f.file]
        assert len(proposed) == 1
        assert "pending" in proposed[0].message
        assert "working-fresh.md.proposed" in proposed[0].message

    # ── Check 7: active-to-archived link ─────────────────────────────────────────────────────────────

    def test_active_to_archived_link_is_flagged_as_warning(self) -> None:
        arch_links = [
            f for f in self.warnings if f.file == "active-links-archived.md"
        ]
        assert len(arch_links) == 1
        assert "archived" in arch_links[0].message
        assert "archived-build.md" in arch_links[0].message

    def test_archived_build_itself_not_flagged_for_its_own_links(self) -> None:
        """archived-build.md links to nothing — must not generate any finding."""
        from_archived = [f for f in self.findings if f.file == "archived-build.md"]
        assert from_archived == []

    # ── Good build has no findings ───────────────────────────────────────────────────────────────────

    def test_good_build_produces_no_findings(self) -> None:
        from_good = [f for f in self.findings if f.file == "good-build.md"]
        assert from_good == []

    # ── Output is sorted ─────────────────────────────────────────────────────────────────────────────

    def test_errors_precede_warnings_in_output(self) -> None:
        severities = [f.severity for f in self.findings]
        # Find first warning index and last error index.
        error_indices = [i for i, s in enumerate(severities) if s == "error"]
        warning_indices = [i for i, s in enumerate(severities) if s == "warning"]
        if error_indices and warning_indices:
            assert max(error_indices) < min(warning_indices)

    def test_findings_within_severity_sorted_by_file(self) -> None:
        error_files = [f.file for f in self.errors]
        assert error_files == sorted(error_files)
        warning_files = [f.file for f in self.warnings]
        assert warning_files == sorted(warning_files)


# ── Edge-case unit tests ────────────────────────────────────────────────────────────────────────────

def test_missing_builds_directory_yields_error(tmp_path: Path) -> None:
    findings = lint_bundle(tmp_path)
    assert len(findings) == 1
    assert findings[0].severity == "error"
    assert "builds" in findings[0].message


def test_link_to_existing_file_not_flagged(tmp_path: Path) -> None:
    builds = tmp_path / "builds"
    builds.mkdir()
    target = builds / "target.md"
    target.write_text(
        "---\ntype: Build\ntitle: T\ndescription: T\n"
        "status: production\neffort: S\nimpact: high\ndomain: TSA\n"
        "timestamp: 2026-01-01T00:00:00Z\nlast_touched: 2026-01-01\n---\n\n"
        "# What it is\nTarget.\n\n# Next action\nNone.\n\n"
        "# Dependencies\n(none)\n\n# Notes\n(none)\n",
        encoding="utf-8",
    )
    source = builds / "source.md"
    source.write_text(
        "---\ntype: Build\ntitle: S\ndescription: S\n"
        "status: production\neffort: S\nimpact: high\ndomain: TSA\n"
        "timestamp: 2026-01-01T00:00:00Z\nlast_touched: 2026-01-01\n---\n\n"
        "# What it is\nSource.\n\n# Next action\nNone.\n\n"
        "# Dependencies\n* [Target](/builds/target.md) — ok\n\n# Notes\n(none)\n",
        encoding="utf-8",
    )
    findings = lint_bundle(tmp_path, reference_date=REF_DATE)
    broken = [f for f in findings if "broken link" in f.message]
    assert broken == []


def test_single_working_build_no_wip_violation(tmp_path: Path) -> None:
    builds = tmp_path / "builds"
    builds.mkdir()
    (builds / "solo.md").write_text(
        "---\ntype: Build\ntitle: Solo\ndescription: Only one.\n"
        "status: working\neffort: S\nimpact: high\ndomain: TSA\n"
        "timestamp: 2026-06-18T00:00:00Z\nlast_touched: 2026-06-18\n---\n\n"
        "# What it is\nSolo.\n\n# Next action\nGo.\n\n"
        "# Dependencies\n(none)\n\n# Notes\n(none)\n",
        encoding="utf-8",
    )
    findings = lint_bundle(tmp_path, reference_date=REF_DATE)
    wip = [f for f in findings if "WIP violation" in f.message]
    assert wip == []
