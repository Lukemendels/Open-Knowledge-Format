"""
Python twin of OKFContextBundle.bas — context bundle assembler.

Implements the same logic as the VBA macro:
  - parse_context_request: extract mode, depth, direction, via, seeds from a
    <CONTEXT_REQUEST> block (no block → index defaults).
  - assemble_index: Hop-1 opener (foundation + /index.md + /skills/index.md).
  - assemble_bundle: BFS expansion from seeds at configurable depth/direction/via.
  - build_header: the <!-- OKF-CONTEXT-BUNDLE … --> header block.
  - build_context_bundle: full entry-point returning the complete file content.
  - derive_dist_dir: the -dist sibling folder of the bundle root.

Golden vectors use the fixture bundle at tests/fixtures/context_bundle/:
  _foundation/01-profile.md, _foundation/02-standards.md
  index.md, skills/index.md, skills/document-review.md
  builds/alpha.md  (→ beta via Dependencies, → deep via Notes)
  builds/beta.md   (→ alpha cycle, → epsilon depth-2)
  builds/gamma.md  (→ alpha, used for inbound test)
  builds/deep.md   (no outbound links)
  builds/epsilon.md (no outbound links, reachable at depth 2 via beta)
"""

import re
from pathlib import Path
from typing import Optional

import pytest


FIXTURES = Path(__file__).parent / "fixtures"
CB = FIXTURES / "context_bundle"      # the fixture bundle root

FIXED_TS = "2026-06-20T00:00:00Z"


# ── Twin logic ──────────────────────────────────────────────────────────────────────────

def parse_context_request(text: str) -> dict:
    """
    Parse a <CONTEXT_REQUEST> block from text.

    Returns a dict with keys: mode, depth, direction, via, seeds.
    Returns index-mode defaults when no block is present.
    """
    defaults: dict = {
        "mode": "index",
        "depth": 1,
        "direction": "outbound",
        "via": "",
        "seeds": [],
    }

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    start = text.find("<CONTEXT_REQUEST>")
    end   = text.find("</CONTEXT_REQUEST>")

    if start == -1 or end == -1 or end <= start:
        return defaults.copy()

    block = text[start + len("<CONTEXT_REQUEST>"):end]
    result = defaults.copy()
    seeds: list[str] = []
    in_include = False

    for line in block.split("\n"):
        s = line.strip()
        if s.startswith("mode: "):
            result["mode"] = s[6:].strip()
            in_include = False
        elif s.startswith("depth: "):
            try:
                result["depth"] = int(s[7:].strip())
            except ValueError:
                pass
            in_include = False
        elif s.startswith("direction: "):
            result["direction"] = s[11:].strip()
            in_include = False
        elif s.startswith("via: "):
            result["via"] = s[5:].strip()
            in_include = False
        elif s == "include:":
            in_include = True
        elif in_include and s.startswith("- "):
            seeds.append(s[2:].strip())
        elif s and not s.startswith("- "):
            in_include = False

    result["seeds"] = seeds
    return result


def _is_concept_file(name: str) -> bool:
    n = name.lower()
    return n.endswith(".md") and n not in ("index.md", "log.md")


def _get_layer(rel_path: str) -> str:
    rel = rel_path.replace("\\", "/")
    if rel.startswith("_foundation/"):
        return "foundation"
    parts = rel.split("/")
    if parts[-1] == "index.md":
        return "map"
    return "selected"


def _make_anchor(rel_path: str, content: str) -> str:
    layer = _get_layer(rel_path)
    rel   = rel_path.replace("\\", "/")
    out   = f"<!-- OKF:BEGIN concept={rel} layer={layer} -->\n"
    out  += content
    if not content.endswith("\n"):
        out += "\n"
    out  += f"<!-- OKF:END concept={rel} -->\n\n"
    return out


def _heading_level(line: str) -> int:
    s = line.lstrip()
    if not s.startswith("#"):
        return 0
    i = 0
    while i < len(s) and s[i] == "#":
        i += 1
    return i


def _heading_text(line: str) -> str:
    s = line.lstrip()
    i = 0
    while i < len(s) and s[i] == "#":
        i += 1
    return s[i:].strip()


def _extract_section(content: str, heading_name: str) -> str:
    """Return the body content under a named heading until the next same/higher heading."""
    content = content.replace("\r\n", "\n").replace("\r", "\n")
    lines = content.split("\n")

    start_idx = -1
    h_level = 0
    for i, line in enumerate(lines):
        lvl = _heading_level(line)
        if lvl > 0 and _heading_text(line).lower() == heading_name.lower():
            start_idx = i
            h_level = lvl
            break

    if start_idx == -1:
        return ""

    section_lines = []
    for line in lines[start_idx + 1:]:
        lvl = _heading_level(line)
        if lvl > 0 and lvl <= h_level:
            break
        section_lines.append(line)

    return "\n".join(section_lines)


def _extract_links_scoped(content: str, via: str = "") -> list[str]:
    """Extract .md links, optionally restricted to content under a named heading."""
    content = content.replace("\r\n", "\n").replace("\r", "\n")
    scan = _extract_section(content, via) if via else content

    links: list[str] = []
    seen: set[str] = set()
    for m in re.finditer(r"\]\(([^)]+)\)", scan):
        url = m.group(1).strip()
        if url.lower().startswith(("http://", "https://")):
            continue
        base = url.split("#")[0]
        if base.lower().endswith(".md") and url not in seen:
            links.append(url)
            seen.add(url)
    return links


def _resolve_link(link: str, from_rel: str, bundle_root: Path) -> Optional[str]:
    """
    Resolve a markdown link to a bundle-root-relative forward-slash path.
    Returns None if external, non-.md, or the target file doesn't exist.
    """
    base = link.split("#")[0].strip()
    if not base or base.lower().startswith(("http://", "https://")):
        return None
    if not base.lower().endswith(".md"):
        return None

    if base.startswith("/"):
        abs_path = bundle_root / base.lstrip("/")
    else:
        from_dir = Path(from_rel).parent
        abs_path = bundle_root / from_dir / base

    try:
        normalized = abs_path.resolve()
    except Exception:
        return None

    if not normalized.exists():
        return None

    try:
        rel = normalized.relative_to(bundle_root.resolve())
    except ValueError:
        return None

    return str(rel).replace("\\", "/")


def _collect_all_concepts(bundle_root: Path) -> list[str]:
    """Return all concept files as bundle-root-relative forward-slash paths."""
    result = []
    for f in sorted(bundle_root.rglob("*.md")):
        if _is_concept_file(f.name):
            rel = str(f.relative_to(bundle_root)).replace("\\", "/")
            result.append(rel)
    return result


def assemble_index(bundle_root: Path) -> tuple[str, int, int, int]:
    """
    Assemble Hop-1 index bundle.
    Returns (content, foundation_count, map_count, selected_count).
    """
    parts: list[str] = []
    found_c = map_c = sel_c = 0

    # 1. Foundation files: all under _foundation/, sorted by filename ascending.
    foundation_dir = bundle_root / "_foundation"
    if foundation_dir.exists():
        foundation_files = sorted(
            (f for f in foundation_dir.rglob("*.md") if _is_concept_file(f.name)),
            key=lambda f: f.name,
        )
        for f in foundation_files:
            rel = str(f.relative_to(bundle_root)).replace("\\", "/")
            parts.append(_make_anchor(rel, f.read_text(encoding="utf-8")))
            found_c += 1

    # 2. Root index.md.
    root_idx = bundle_root / "index.md"
    if root_idx.exists():
        parts.append(_make_anchor("index.md", root_idx.read_text(encoding="utf-8")))
        map_c += 1

    # 3. skills/index.md.
    skills_idx = bundle_root / "skills" / "index.md"
    if skills_idx.exists():
        parts.append(_make_anchor("skills/index.md", skills_idx.read_text(encoding="utf-8")))
        map_c += 1

    return "".join(parts), found_c, map_c, sel_c


def assemble_bundle(
    bundle_root: Path,
    seeds: list[str],
    depth: int = 1,
    direction: str = "outbound",
    via: str = "",
) -> tuple[str, int, int, int]:
    """
    Assemble a bundle via BFS graph traversal.
    Returns (content, foundation_count, map_count, selected_count).
    """
    # Normalise seeds: forward slashes, no leading /.
    normalized = [s.replace("\\", "/").lstrip("/") for s in seeds]

    # visited preserves insertion order (Python 3.7+).
    visited: dict[str, bool] = {}
    frontier: list[str] = []

    for seed in normalized:
        if seed not in visited and (bundle_root / seed).exists():
            visited[seed] = True
            frontier.append(seed)
        # Absent seeds skipped silently per OKF §9.

    for _ in range(depth):
        if not frontier:
            break
        next_frontier: list[str] = []

        if direction.lower() == "outbound":
            for from_rel in frontier:
                abs_path = bundle_root / from_rel
                if not abs_path.exists():
                    continue
                content = abs_path.read_text(encoding="utf-8")
                for link in _extract_links_scoped(content, via):
                    resolved = _resolve_link(link, from_rel, bundle_root)
                    if resolved is not None and resolved not in visited:
                        visited[resolved] = True
                        next_frontier.append(resolved)

        elif direction.lower() == "inbound":
            frontier_set = set(frontier)
            for concept_rel in _collect_all_concepts(bundle_root):
                if concept_rel in visited:
                    continue
                abs_path = bundle_root / concept_rel
                if not abs_path.exists():
                    continue
                content = abs_path.read_text(encoding="utf-8")
                for link in _extract_links_scoped(content, ""):  # via ignored for inbound
                    resolved = _resolve_link(link, concept_rel, bundle_root)
                    if resolved is not None and resolved in frontier_set:
                        visited[concept_rel] = True
                        next_frontier.append(concept_rel)
                        break

        frontier = next_frontier

    # Assemble in BFS insertion order.
    parts: list[str] = []
    found_c = map_c = sel_c = 0
    for rel_path in visited:
        abs_path = bundle_root / rel_path
        if not abs_path.exists():
            continue
        content = abs_path.read_text(encoding="utf-8")
        layer = _get_layer(rel_path)
        parts.append(_make_anchor(rel_path, content))
        if layer == "foundation":
            found_c += 1
        elif layer == "map":
            map_c += 1
        else:
            sel_c += 1

    return "".join(parts), found_c, map_c, sel_c


def build_header(
    mode: str,
    foundation_count: int,
    map_count: int,
    selected_count: int,
    body: str,
    timestamp: str = FIXED_TS,
) -> str:
    """Return the <!-- OKF-CONTEXT-BUNDLE … --> header string (with trailing blank line)."""
    total = foundation_count + map_count + selected_count
    approx_tokens = len(body) // 4
    return (
        f"<!-- OKF-CONTEXT-BUNDLE\n"
        f"mode: {mode}\n"
        f"okf_version: 0.1\n"
        f"assembled: {timestamp}\n"
        f"concepts: {total} ({foundation_count} foundation, {map_count} map, {selected_count} selected)\n"
        f"approx_tokens: {approx_tokens}\n"
        f"-->\n\n"
    )


def build_context_bundle(
    bundle_root: Path,
    request_text: str,
    timestamp: str = FIXED_TS,
) -> str:
    """
    Main entry: parse request, assemble bundle, return full OKF-context.md content.
    Raises ValueError if bundle root does not exist.
    """
    if not bundle_root.exists():
        raise ValueError(f"Bundle root not found: {bundle_root}")

    req = parse_context_request(request_text)
    mode = req["mode"]

    if mode == "index":
        body, found_c, map_c, sel_c = assemble_index(bundle_root)
    else:
        if not req["seeds"]:
            raise ValueError("No include: paths in bundle request")
        body, found_c, map_c, sel_c = assemble_bundle(
            bundle_root,
            req["seeds"],
            req["depth"],
            req["direction"],
            req["via"],
        )

    header = build_header(mode, found_c, map_c, sel_c, body, timestamp)
    return header + body


def derive_dist_dir(bundle_root: Path) -> Path:
    """
    Return the -dist sibling folder of bundle_root.
    Strips any trailing path separator before appending -dist.
    """
    stripped = str(bundle_root).rstrip("/\\")
    return Path(stripped + "-dist")


# ── Tests — §11.2: request parsing ────────────────────────────────────────────────────

class TestParseContextRequest:

    def test_no_block_returns_index_defaults(self) -> None:
        result = parse_context_request("some text with no envelope")
        assert result["mode"] == "index"
        assert result["depth"] == 1
        assert result["direction"] == "outbound"
        assert result["via"] == ""
        assert result["seeds"] == []

    def test_index_block(self) -> None:
        text = "<CONTEXT_REQUEST>\nmode: index\n</CONTEXT_REQUEST>"
        result = parse_context_request(text)
        assert result["mode"] == "index"
        assert result["seeds"] == []

    def test_full_bundle_block(self) -> None:
        text = (
            "<CONTEXT_REQUEST>\n"
            "mode: bundle\n"
            "depth: 2\n"
            "direction: inbound\n"
            "via: Dependencies\n"
            "include:\n"
            "- builds/alpha.md\n"
            "- builds/beta.md\n"
            "</CONTEXT_REQUEST>"
        )
        result = parse_context_request(text)
        assert result["mode"] == "bundle"
        assert result["depth"] == 2
        assert result["direction"] == "inbound"
        assert result["via"] == "Dependencies"
        assert result["seeds"] == ["builds/alpha.md", "builds/beta.md"]

    def test_depth_zero(self) -> None:
        text = "<CONTEXT_REQUEST>\nmode: bundle\ndepth: 0\ninclude:\n- builds/alpha.md\n</CONTEXT_REQUEST>"
        assert parse_context_request(text)["depth"] == 0

    def test_preamble_ignored(self) -> None:
        text = "intent line\n<CONTEXT_REQUEST>\nmode: index\n</CONTEXT_REQUEST>"
        assert parse_context_request(text)["mode"] == "index"

    def test_crlf_normalised(self) -> None:
        text = "<CONTEXT_REQUEST>\r\nmode: bundle\r\ninclude:\r\n- builds/alpha.md\r\n</CONTEXT_REQUEST>"
        result = parse_context_request(text)
        assert result["mode"] == "bundle"
        assert result["seeds"] == ["builds/alpha.md"]


# ── Tests — §11.1: index mode ─────────────────────────────────────────────────────────

class TestAssembleIndex:

    def setup_method(self) -> None:
        self.body, self.found_c, self.map_c, self.sel_c = assemble_index(CB)

    def test_foundation_count(self) -> None:
        assert self.found_c == 2

    def test_map_count(self) -> None:
        assert self.map_c == 2  # index.md + skills/index.md

    def test_selected_count(self) -> None:
        assert self.sel_c == 0  # no build files in index mode

    def test_foundation_files_in_order(self) -> None:
        pos1 = self.body.find("concept=_foundation/01-profile.md")
        pos2 = self.body.find("concept=_foundation/02-standards.md")
        assert pos1 != -1 and pos2 != -1
        assert pos1 < pos2

    def test_root_index_included(self) -> None:
        assert "concept=index.md layer=map" in self.body

    def test_skills_index_included(self) -> None:
        assert "concept=skills/index.md layer=map" in self.body

    def test_skill_body_not_included(self) -> None:
        # skills/document-review.md is a concept file, not in index mode
        assert "concept=skills/document-review.md" not in self.body

    def test_foundation_layer_tag(self) -> None:
        assert "concept=_foundation/01-profile.md layer=foundation" in self.body

    def test_map_order_foundation_before_index(self) -> None:
        pos_found = self.body.find("layer=foundation")
        pos_map   = self.body.find("layer=map")
        assert pos_found < pos_map

    def test_anchor_begin_end_format(self) -> None:
        assert "<!-- OKF:BEGIN concept=_foundation/01-profile.md layer=foundation -->" in self.body
        assert "<!-- OKF:END concept=_foundation/01-profile.md -->" in self.body

    def test_no_request_in_build_bundle_gives_index(self) -> None:
        content = build_context_bundle(CB, "no envelope here", FIXED_TS)
        assert "mode: index" in content
        assert "okf_version: 0.1" in content


# ── Tests — §11.3: bundle outbound ───────────────────────────────────────────────────

class TestAssembleBundleOutbound:

    def test_depth_1_from_alpha_includes_seeds_and_direct_links(self) -> None:
        body, _, _, sel_c = assemble_bundle(CB, ["builds/alpha.md"], depth=1)
        assert "concept=builds/alpha.md" in body
        assert "concept=builds/beta.md" in body    # alpha → beta (Dependencies)
        assert "concept=builds/deep.md" in body    # alpha → deep (Notes)
        assert sel_c == 3

    def test_depth_2_from_alpha_reaches_epsilon(self) -> None:
        body, _, _, sel_c = assemble_bundle(CB, ["builds/alpha.md"], depth=2)
        assert "concept=builds/epsilon.md" in body  # alpha→beta→epsilon
        assert sel_c == 4  # alpha, beta, deep, epsilon

    def test_depth_2_alpha_not_duplicated(self) -> None:
        # beta links back to alpha (A→B→A cycle); alpha must appear exactly once.
        body, _, _, _ = assemble_bundle(CB, ["builds/alpha.md"], depth=2)
        assert body.count("<!-- OKF:BEGIN concept=builds/alpha.md") == 1

    def test_depth_2_order_seed_first(self) -> None:
        body, _, _, _ = assemble_bundle(CB, ["builds/alpha.md"], depth=2)
        pos_alpha = body.find("<!-- OKF:BEGIN concept=builds/alpha.md")
        pos_beta  = body.find("<!-- OKF:BEGIN concept=builds/beta.md")
        assert pos_alpha < pos_beta


# ── Tests — §11.4: via scoping and inbound ────────────────────────────────────────────

class TestViaAndInbound:

    def test_via_dependencies_excludes_notes_links(self) -> None:
        # alpha→beta in Dependencies; alpha→deep in Notes
        body, _, _, sel_c = assemble_bundle(
            CB, ["builds/alpha.md"], depth=1, via="Dependencies"
        )
        assert "concept=builds/beta.md" in body
        assert "concept=builds/deep.md" not in body
        assert sel_c == 2   # alpha + beta

    def test_no_via_includes_all_links(self) -> None:
        body, _, _, sel_c = assemble_bundle(
            CB, ["builds/alpha.md"], depth=1, via=""
        )
        assert "concept=builds/beta.md" in body
        assert "concept=builds/deep.md" in body
        assert sel_c == 3

    def test_inbound_from_alpha_collects_backlinks(self) -> None:
        # gamma → alpha and beta → alpha both make those two concepts inbound.
        body, _, _, sel_c = assemble_bundle(
            CB, ["builds/alpha.md"], depth=1, direction="inbound"
        )
        assert "concept=builds/alpha.md" in body
        assert "concept=builds/beta.md" in body
        assert "concept=builds/gamma.md" in body
        assert sel_c == 3

    def test_inbound_via_is_ignored(self) -> None:
        # Passing via= for inbound must not change the result
        body_no_via, _, _, _ = assemble_bundle(
            CB, ["builds/alpha.md"], depth=1, direction="inbound", via=""
        )
        body_via, _, _, _ = assemble_bundle(
            CB, ["builds/alpha.md"], depth=1, direction="inbound", via="Dependencies"
        )
        assert body_no_via == body_via


# ── Tests — §11.5: depth 0 (seeds only) ──────────────────────────────────────────────

def test_depth_zero_returns_seeds_only() -> None:
    body, _, _, sel_c = assemble_bundle(
        CB, ["skills/document-review.md"], depth=0
    )
    assert "concept=skills/document-review.md" in body
    assert "concept=builds/alpha.md" not in body
    assert sel_c == 1


# ── Tests — §11.6: de-dup and cycle break ────────────────────────────────────────────

def test_cycle_does_not_duplicate_nodes() -> None:
    # alpha→beta→alpha is a cycle; both should appear exactly once.
    body, _, _, _ = assemble_bundle(CB, ["builds/alpha.md"], depth=3)
    assert body.count("<!-- OKF:BEGIN concept=builds/alpha.md") == 1
    assert body.count("<!-- OKF:BEGIN concept=builds/beta.md") == 1


# ── Tests — §11.7: absent link ───────────────────────────────────────────────────────

def test_absent_link_skipped_assembly_succeeds(tmp_path: Path) -> None:
    builds = tmp_path / "builds"
    builds.mkdir()
    concept = builds / "source.md"
    concept.write_text(
        "---\ntype: Build\ntitle: Source\ndescription: has a broken link.\n"
        "status: idea\neffort: S\nimpact: low\n---\n\n"
        "# Dependencies\n* [Missing](/builds/nonexistent.md) — gone\n",
        encoding="utf-8",
    )
    body, _, _, sel_c = assemble_bundle(
        tmp_path, ["builds/source.md"], depth=1
    )
    assert "concept=builds/source.md" in body
    # The absent link should not have generated a concept anchor of its own.
    assert "concept=builds/nonexistent.md" not in body
    assert sel_c == 1


# ── Tests — §11.8: anchor and header exactness ────────────────────────────────────────

class TestAnchorAndHeaderFormat:

    def test_anchor_begin_format(self) -> None:
        result = _make_anchor("builds/alpha.md", "content\n")
        assert result.startswith("<!-- OKF:BEGIN concept=builds/alpha.md layer=selected -->")

    def test_anchor_end_format(self) -> None:
        result = _make_anchor("builds/alpha.md", "content\n")
        assert "<!-- OKF:END concept=builds/alpha.md -->" in result

    def test_anchor_trailing_blank_line(self) -> None:
        result = _make_anchor("builds/alpha.md", "content\n")
        assert result.endswith("\n\n")

    def test_anchor_content_appended_newline_if_missing(self) -> None:
        result = _make_anchor("builds/alpha.md", "content without newline")
        lines = result.split("\n")
        # Content should be followed by the END comment on its own line.
        end_line_idx = next(i for i, l in enumerate(lines) if "OKF:END" in l)
        assert lines[end_line_idx - 1] in ("content without newline", "")

    def test_header_mode_field(self) -> None:
        h = build_header("bundle", 0, 0, 3, "x" * 400, FIXED_TS)
        assert "mode: bundle\n" in h

    def test_header_okf_version(self) -> None:
        h = build_header("index", 2, 2, 0, "", FIXED_TS)
        assert "okf_version: 0.1\n" in h

    def test_header_assembled_timestamp(self) -> None:
        h = build_header("index", 2, 2, 0, "", FIXED_TS)
        assert f"assembled: {FIXED_TS}\n" in h

    def test_header_concept_counts(self) -> None:
        h = build_header("bundle", 1, 2, 3, "", FIXED_TS)
        assert "concepts: 6 (1 foundation, 2 map, 3 selected)\n" in h

    def test_header_approx_token_math(self) -> None:
        body = "x" * 400
        h = build_header("bundle", 0, 0, 1, body, FIXED_TS)
        assert "approx_tokens: 100\n" in h   # 400 chars // 4 = 100

    def test_header_starts_with_comment_open(self) -> None:
        h = build_header("index", 0, 0, 0, "", FIXED_TS)
        assert h.startswith("<!-- OKF-CONTEXT-BUNDLE\n")

    def test_header_ends_with_comment_close_and_blank_line(self) -> None:
        h = build_header("index", 0, 0, 0, "", FIXED_TS)
        assert h.endswith("-->\n\n")

    def test_full_output_header_then_body(self) -> None:
        content = build_context_bundle(
            CB,
            "<CONTEXT_REQUEST>\nmode: bundle\ndepth: 0\ninclude:\n- builds/alpha.md\n</CONTEXT_REQUEST>",
            FIXED_TS,
        )
        header_end = content.index("-->\n\n") + len("-->\n\n")
        after_header = content[header_end:]
        assert after_header.startswith("<!-- OKF:BEGIN concept=builds/alpha.md")


# ── Tests — §11.9: output path derivation ─────────────────────────────────────────────

class TestDeriveDistDir:

    def test_basic_derivation(self, tmp_path: Path) -> None:
        root = tmp_path / "build-portfolio"
        dist = derive_dist_dir(root)
        assert dist == tmp_path / "build-portfolio-dist"

    def test_not_under_bundle_root(self, tmp_path: Path) -> None:
        root = tmp_path / "my-bundle"
        dist = derive_dist_dir(root)
        # dist must NOT be inside root
        assert not str(dist).startswith(str(root) + "/") and dist != root

    def test_output_filename_constant(self, tmp_path: Path) -> None:
        root = tmp_path / "bundle"
        dist = derive_dist_dir(root)
        out  = dist / "OKF-context.md"
        assert out.name == "OKF-context.md"
        assert out.parent == dist

    def test_strips_trailing_slash(self, tmp_path: Path) -> None:
        root_with_slash = Path(str(tmp_path / "bundle") + "/")
        root_plain      = tmp_path / "bundle"
        assert derive_dist_dir(root_with_slash) == derive_dist_dir(root_plain)

    def test_dist_sibling_of_root(self, tmp_path: Path) -> None:
        root = tmp_path / "my-okf-bundle"
        dist = derive_dist_dir(root)
        # Same parent directory, name ends with -dist
        assert dist.parent == root.parent
        assert dist.name == "my-okf-bundle-dist"


# ── Tests — integration: build_context_bundle end-to-end ──────────────────────────────

def test_integration_index_mode(tmp_path: Path) -> None:
    found = tmp_path / "_foundation"
    found.mkdir()
    (found / "schema.md").write_text(
        "---\ntype: Foundation\ntitle: Schema\ndescription: Schema.\n---\n\nSchema content.\n",
        encoding="utf-8",
    )
    (tmp_path / "index.md").write_text(
        "---\nokf_version: \"0.1\"\n---\n\n# all\n\n* [Schema](_foundation/schema.md)\n",
        encoding="utf-8",
    )

    content = build_context_bundle(tmp_path, "", FIXED_TS)

    assert content.startswith("<!-- OKF-CONTEXT-BUNDLE\n")
    assert "mode: index\n" in content
    assert "okf_version: 0.1\n" in content
    assert "layer=foundation" in content
    assert "layer=map" in content


def test_integration_bundle_mode(tmp_path: Path) -> None:
    builds = tmp_path / "builds"
    builds.mkdir()
    (builds / "a.md").write_text(
        "---\ntype: Build\ntitle: A\ndescription: A.\nstatus: idea\neffort: S\nimpact: low\n---\n\n"
        "# Dependencies\n* [B](/builds/b.md)\n",
        encoding="utf-8",
    )
    (builds / "b.md").write_text(
        "---\ntype: Build\ntitle: B\ndescription: B.\nstatus: idea\neffort: S\nimpact: low\n---\n\n"
        "Body of B.\n",
        encoding="utf-8",
    )

    request = (
        "<CONTEXT_REQUEST>\n"
        "mode: bundle\n"
        "depth: 1\n"
        "include:\n"
        "- builds/a.md\n"
        "</CONTEXT_REQUEST>"
    )
    content = build_context_bundle(tmp_path, request, FIXED_TS)

    assert "mode: bundle\n" in content
    assert "concept=builds/a.md layer=selected" in content
    assert "concept=builds/b.md layer=selected" in content
    assert "concepts: 2 (0 foundation, 0 map, 2 selected)\n" in content
