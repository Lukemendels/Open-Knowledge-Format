"""
Python twin of OKFIndexGenerator.bas -- index file producer.

Implements the same logic as the VBA macro:
  - Grouping is inferred from concepts' own frontmatter.
  - First field in GROUP_BY_CANDIDATES present in any concept of a folder
    wins as the grouping axis; none present -> flat alphabetical list,
    no group headings.
  - GROUP_ORDER determines heading order; extras go alphabetical.
  - STALL_GROUP (working) sorts oldest-last_touched first, shows date inline.
  - Reserved files (index.md, log.md) are never listed as concepts.

Tests cover:
  - skills-shaped folder (no status) -> flat list, no headings
  - builds-shaped folder (has status) -> lifecycle board (regression guard)
  - mixed folder (some status, some not) -> grouped, status-less under (unset)
"""

from pathlib import Path
from typing import Optional

import pytest


# ── Constants (mirrors VBA module-level) ─────────────────────────────────────────────────

# EXTENSION SEAM: append a field name (e.g. "domain") to support a new grouping axis.
# See _meta/okf-roadmap.md before adding one.
GROUP_BY_CANDIDATES: list[str] = ["status"]   # extend: ["status", "domain"]
STALL_GROUP = "working"
GROUP_ORDER = ["working", "boilerplate", "spec", "idea", "parked", "production", "archived"]
OKF_VERSION = "0.1"


# ── Twin logic ───────────────────────────────────────────────────────────────────────────

def _is_concept_file(name: str) -> bool:
    n = name.lower()
    return n.endswith(".md") and n not in ("index.md", "log.md")


def _parse_frontmatter(content: str) -> dict[str, str]:
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
            fm[key.strip().lower()] = val.strip().strip("\"'")
    return fm


def _basename(name: str) -> str:
    return name[:-3] if name.lower().endswith(".md") else name


def _pick_group_field(concepts: list[dict]) -> str:
    """Return the first GROUP_BY_CANDIDATES field present in any concept, or ''."""
    for candidate in GROUP_BY_CANDIDATES:
        if any(c.get(candidate, "") for c in concepts):
            return candidate
    return ""


def _sort_entries(entries: list[tuple[str, str]]) -> list[str]:
    """Sort (sort_key, display_line) pairs and return display lines."""
    return [line for _, line in sorted(entries, key=lambda t: t[0].lower())]


def _ordered_group_keys(groups: dict) -> list[str]:
    """GROUP_ORDER first (in that order), then remaining keys alphabetically."""
    seen: set[str] = set()
    result: list[str] = []
    for key in GROUP_ORDER:
        if key in groups and key not in seen:
            result.append(key)
            seen.add(key)
    result += sorted(k for k in groups if k not in seen)
    return result


def generate_dir_index(
    dir_path: Path,
    is_root: bool = False,
) -> Optional[str]:
    """
    Generate the index.md content for one directory.

    Returns None if the directory has no concepts and no subdirectories
    (nothing worth writing).
    """
    concept_files = sorted(
        f for f in dir_path.iterdir()
        if f.is_file() and _is_concept_file(f.name)
    )

    # Collect concept data.
    concepts: list[dict] = []
    for f in concept_files:
        fm = _parse_frontmatter(f.read_text(encoding="utf-8"))
        title = fm.get("title", "") or _basename(f.name)
        concepts.append({
            "name": f.name,
            "title": title,
            "desc": fm.get("description", ""),
            "status": fm.get("status", ""),
            "last_touched": fm.get("last_touched", ""),
            "fm": fm,
        })

    subdirs = sorted(d for d in dir_path.iterdir() if d.is_dir())

    sb = ""

    if is_root:
        sb += f'---\nokf_version: "{OKF_VERSION}"\n---\n\n'

    if concepts:
        group_field = _pick_group_field(concepts)

        if group_field == "":
            # Flat alphabetical list -- no group headings.
            entries: list[tuple[str, str]] = []
            for c in concepts:
                line = f"* [{c['title']}]({c['name']})"
                if c["desc"]:
                    line += f" - {c['desc']}"
                entries.append((c["title"].lower(), line))
            for line in _sort_entries(entries):
                sb += line + "\n"
            sb += "\n"

        else:
            # Grouped under headings (lifecycle-board behavior).
            groups: dict[str, list[tuple[str, str]]] = {}
            for c in concepts:
                grp = c["fm"].get(group_field, "") or "(unset)"

                if grp.lower() == STALL_GROUP.lower():
                    lt = c["last_touched"]
                    line = f"* [{c['title']}]({c['name']})"
                    if c["desc"]:
                        line += f" - {c['desc']}"
                    line += f"  _(last touched {lt if lt else 'never'})_"
                    sort_key = lt if lt else "0000-00-00"
                else:
                    line = f"* [{c['title']}]({c['name']})"
                    if c["desc"]:
                        line += f" - {c['desc']}"
                    sort_key = c["title"].lower()

                groups.setdefault(grp, []).append((sort_key, line))

            for grp in _ordered_group_keys(groups):
                sb += f"# {grp}\n"
                for line in _sort_entries(groups[grp]):
                    sb += line + "\n"
                sb += "\n"

    if subdirs:
        sb += "# Subdirectories\n"
        for d in subdirs:
            sb += f"* [{d.name}]({d.name}/)\n"
        sb += "\n"

    return sb if sb.strip() else None


# ── Fixtures helpers ─────────────────────────────────────────────────────────────────────

def _skill(tmp: Path, name: str, title: str, desc: str) -> None:
    """Write a Skill concept (no status field)."""
    tmp.write_text(
        f"---\ntype: Skill\ntitle: {title}\ndescription: {desc}\n---\n\n# Procedure\n\nDo it.\n",
        encoding="utf-8",
    )


def _build(tmp: Path, name: str, title: str, desc: str, status: str,
           last_touched: str = "2026-06-01") -> None:
    """Write a Build concept (has status field)."""
    tmp.write_text(
        f"---\ntype: Build\ntitle: {title}\ndescription: {desc}\n"
        f"status: {status}\neffort: S\nimpact: high\n"
        f"last_touched: {last_touched}\n---\n\n# Body\n\nContent.\n",
        encoding="utf-8",
    )


# ── Tests: skills-shaped folder (no status) -> flat list ─────────────────────────────────

class TestSkillsFolder:
    """Folder whose concepts carry no status -> flat alphabetical list, no headings."""

    def setup_method(self, tmp_path_factory) -> None:
        pass  # each test uses its own tmp_path

    def test_no_unset_heading(self, tmp_path: Path) -> None:
        _skill(tmp_path / "writing.md", "writing.md", "Writing", "Write well.")
        _skill(tmp_path / "analysis.md", "analysis.md", "Analysis", "Analyse data.")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert "(unset)" not in result

    def test_no_group_headings_at_all(self, tmp_path: Path) -> None:
        _skill(tmp_path / "alpha.md", "alpha.md", "Alpha Skill", "First.")
        _skill(tmp_path / "beta.md", "beta.md", "Beta Skill", "Second.")
        result = generate_dir_index(tmp_path)
        assert result is not None
        # No lines starting with "# " (except potential Subdirectories, which won't exist here)
        heading_lines = [l for l in result.splitlines() if l.startswith("# ")]
        assert heading_lines == []

    def test_entries_sorted_alphabetically_by_title(self, tmp_path: Path) -> None:
        _skill(tmp_path / "zebra.md", "zebra.md", "Zebra", "Last alpha.")
        _skill(tmp_path / "apple.md", "apple.md", "Apple", "First alpha.")
        _skill(tmp_path / "mango.md", "mango.md", "Mango", "Middle.")
        result = generate_dir_index(tmp_path)
        assert result is not None
        pos_apple = result.index("Apple")
        pos_mango = result.index("Mango")
        pos_zebra = result.index("Zebra")
        assert pos_apple < pos_mango < pos_zebra

    def test_entries_include_description(self, tmp_path: Path) -> None:
        _skill(tmp_path / "doc-review.md", "doc-review.md", "Document Review",
               "Review a document for quality and completeness.")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert "* [Document Review](doc-review.md) - Review a document" in result

    def test_entry_format_no_link_prefix(self, tmp_path: Path) -> None:
        _skill(tmp_path / "s.md", "s.md", "My Skill", "Does something.")
        result = generate_dir_index(tmp_path)
        assert "* [My Skill](s.md) - Does something." in result


# ── Tests: builds-shaped folder (has status) -> lifecycle board ───────────────────────────

class TestBuildsFolder:
    """Regression guard: concepts with status -> grouped lifecycle board, unchanged."""

    def test_idea_group_heading_present(self, tmp_path: Path) -> None:
        _build(tmp_path / "a.md", "a.md", "Alpha Build", "First.", "idea")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert "# idea\n" in result

    def test_group_order_working_before_idea(self, tmp_path: Path) -> None:
        _build(tmp_path / "a.md", "a.md", "Alpha", "Desc.", "idea")
        _build(tmp_path / "b.md", "b.md", "Beta", "Desc.", "working", "2026-06-15")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert result.index("# working") < result.index("# idea")

    def test_working_entry_shows_last_touched(self, tmp_path: Path) -> None:
        _build(tmp_path / "w.md", "w.md", "Active Build", "Desc.", "working", "2026-06-15")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert "_(last touched 2026-06-15)_" in result

    def test_non_working_entries_no_last_touched(self, tmp_path: Path) -> None:
        _build(tmp_path / "a.md", "a.md", "Alpha", "Desc.", "idea")
        result = generate_dir_index(tmp_path)
        assert "last touched" not in result

    def test_entries_within_group_sorted_by_title(self, tmp_path: Path) -> None:
        _build(tmp_path / "z.md", "z.md", "Zebra Build", "Z.", "idea")
        _build(tmp_path / "a.md", "a.md", "Alpha Build", "A.", "idea")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert result.index("Alpha Build") < result.index("Zebra Build")

    def test_no_flat_list_when_status_present(self, tmp_path: Path) -> None:
        _build(tmp_path / "a.md", "a.md", "Alpha", "Desc.", "production")
        result = generate_dir_index(tmp_path)
        # Should have a "# production" heading, not a bare bullet
        assert "# production\n" in result


# ── Tests: mixed folder (some status, some not) -> grouped, (unset) for missing ──────────

class TestMixedFolder:
    """Inference surfaces the inconsistency: status-less files land under (unset)."""

    def test_mixed_has_group_headings(self, tmp_path: Path) -> None:
        _build(tmp_path / "b.md", "b.md", "Build One", "Has status.", "idea")
        _skill(tmp_path / "s.md", "s.md", "Skill One", "No status.")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert "# idea\n" in result

    def test_status_less_concept_under_unset(self, tmp_path: Path) -> None:
        _build(tmp_path / "b.md", "b.md", "Build One", "Has status.", "idea")
        _skill(tmp_path / "s.md", "s.md", "Skill One", "No status.")
        result = generate_dir_index(tmp_path)
        assert result is not None
        assert "# (unset)\n" in result
        unset_pos = result.index("# (unset)")
        assert result.index("Skill One") > unset_pos

    def test_no_flat_list_when_any_status_present(self, tmp_path: Path) -> None:
        _build(tmp_path / "b.md", "b.md", "Build", "Has status.", "production")
        _skill(tmp_path / "s.md", "s.md", "Skill", "No status.")
        result = generate_dir_index(tmp_path)
        # At least one concept has status, so we get grouped output
        assert "# production\n" in result or "# (unset)\n" in result


# ── Tests: root frontmatter and empty dir ─────────────────────────────────────────────────

def test_root_index_has_okf_version_frontmatter(tmp_path: Path) -> None:
    _build(tmp_path / "a.md", "a.md", "Alpha", "Desc.", "idea")
    result = generate_dir_index(tmp_path, is_root=True)
    assert result is not None
    assert result.startswith('---\nokf_version: "0.1"\n---\n\n')


def test_non_root_has_no_frontmatter(tmp_path: Path) -> None:
    _build(tmp_path / "a.md", "a.md", "Alpha", "Desc.", "idea")
    result = generate_dir_index(tmp_path, is_root=False)
    assert result is not None
    assert not result.startswith("---")


def test_empty_dir_returns_none(tmp_path: Path) -> None:
    result = generate_dir_index(tmp_path)
    assert result is None


def test_subdirectories_section(tmp_path: Path) -> None:
    sub = tmp_path / "sub"
    sub.mkdir()
    _build(tmp_path / "a.md", "a.md", "Alpha", "Desc.", "idea")
    result = generate_dir_index(tmp_path)
    assert result is not None
    assert "# Subdirectories\n" in result
    assert "* [sub](sub/)\n" in result


def test_index_and_log_not_listed_as_concepts(tmp_path: Path) -> None:
    (tmp_path / "index.md").write_text("# index\n", encoding="utf-8")
    (tmp_path / "log.md").write_text("# log\n", encoding="utf-8")
    _skill(tmp_path / "real.md", "real.md", "Real Skill", "Counts.")
    result = generate_dir_index(tmp_path)
    assert result is not None
    assert "index.md" not in result
    assert "log.md" not in result
    assert "Real Skill" in result
