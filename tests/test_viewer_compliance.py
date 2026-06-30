"""
Unit tests for the StickShift Viewer HTML tool.
Validates ASCII encoding, HTML Tool Compliance Standard requirements,
and the explorer tree construction logic.
"""

from pathlib import Path
import re
import pytest

# Path to the HTML tool
TOOL_PATH = Path(__file__).parent.parent / "builds" / "html-tools" / "stickshift-viewer" / "stickshift-viewer.html"

def test_viewer_exists() -> None:
    """Ensure the viewer HTML tool exists in the expected location."""
    assert TOOL_PATH.exists(), f"Viewer tool not found at {TOOL_PATH}"
    assert TOOL_PATH.is_file()


def test_viewer_ascii_only() -> None:
    """Ensure the viewer HTML file contains strictly ASCII characters for VBA compatibility."""
    content = TOOL_PATH.read_bytes()
    try:
        content.decode("ascii")
    except UnicodeDecodeError as e:
        # Show context of decode error
        start = max(0, e.start - 40)
        end = min(len(content), e.end + 40)
        snippet = content[start:end]
        pytest.fail(
            f"Viewer tool contains non-ASCII characters at index {e.start} "
            f"(byte value 0x{content[e.start]:02x}). Context: {snippet!r}"
        )


def test_viewer_compliance_identity() -> None:
    """Ensure the window.STICKSHIFT_TOOL identity declaration is present and correct."""
    html_text = TOOL_PATH.read_text(encoding="utf-8")
    
    # Check for the window.STICKSHIFT_TOOL object structure
    # Match id, file, skillSlug, and title
    id_match = re.search(r'id:\s*["\']stickshift-viewer["\']', html_text)
    file_match = re.search(r'file:\s*["\']stickshift-viewer\.html["\']', html_text)
    slug_match = re.search(r'skillSlug:\s*["\']stickshift-viewer["\']', html_text)
    title_match = re.search(r'title:\s*["\']StickShift Viewer["\']', html_text)
    
    assert id_match is not None, "window.STICKSHIFT_TOOL is missing ID or has incorrect ID"
    assert file_match is not None, "window.STICKSHIFT_TOOL is missing file or has incorrect file"
    assert slug_match is not None, "window.STICKSHIFT_TOOL is missing skillSlug or has incorrect skillSlug"
    assert title_match is not None, "window.STICKSHIFT_TOOL is missing title or has incorrect title"


def test_viewer_compliance_embedded_skill() -> None:
    """Ensure the embedded companion skill block is present, correct, and matching."""
    html_text = TOOL_PATH.read_text(encoding="utf-8")
    
    # 1. Check for the script tag structure
    script_match = re.search(
        r'<script\s+type=["\']text/markdown["\']\s+id=["\']stickshift-skill["\']\s+data-skill-slug=["\']stickshift-viewer["\']>',
        html_text
    )
    assert script_match is not None, "Missing or incorrect #stickshift-skill script tag"
    
    # 2. Extract skill content
    skill_block_pattern = r'<script\s+type=["\']text/markdown["\']\s+id=["\']stickshift-skill["\']\s+data-skill-slug=["\']stickshift-viewer["\']>(.*?)</script>'
    skill_content_match = re.search(skill_block_pattern, html_text, re.DOTALL)
    assert skill_content_match is not None, "Failed to extract content from #stickshift-skill tag"
    
    skill_content = skill_content_match.group(1).strip()
    
    # 3. Check for required elements in the skill markdown
    assert "type: Skill" in skill_content, "Skill frontmatter missing type: Skill"
    assert "title: StickShift Viewer" in skill_content, "Skill frontmatter title is incorrect"
    
    # Verify it instructs to open stickshift-viewer.html
    assert "<HTML_OPEN>" in skill_content, "Skill is missing <HTML_OPEN> block"
    assert "tool: stickshift-viewer.html" in skill_content, "Skill HTML_OPEN references incorrect tool file"
    assert "</HTML_OPEN>" in skill_content, "Skill is missing </HTML_OPEN> block"


def test_viewer_compliance_onboarding_elements() -> None:
    """Ensure onboarding and setup elements exist in the DOM as required."""
    html_text = TOOL_PATH.read_text(encoding="utf-8")
    
    # Verify presence of the onboarding/setup panel and buttons
    assert 'id="ssPanel"' in html_text, "Missing StickShift onboarding/setup panel container"
    assert 'id="ssYes"' in html_text, "Missing onboarding Yes button"
    assert 'id="ssNo"' in html_text, "Missing onboarding No button"
    assert 'id="ssCopySkill"' in html_text, "Missing onboarding Copy Skill button"
    assert 'id="ssFootBtn"' in html_text, "Missing footer button to toggle onboarding panel"
    assert 'id="loadAllBtn"' in html_text, "Missing Load all button"


def build_tree_python(files_list: list[dict]) -> dict:
    """Python twin of the JavaScript buildTree() function in stickshift-viewer.html."""
    root = {"name": "Root", "children": {}, "files": []}
    for f in files_list:
        parts = f["path"].split("/")
        current = root
        for part in parts[:-1]:
            if part not in current["children"]:
                current["children"][part] = {"name": part, "children": {}, "files": []}
            current = current["children"][part]
        filename = parts[-1]
        current["files"].append({"name": filename, "file": f})
    return root


def test_sidebar_tree_construction() -> None:
    """Verify that tree view building logic matches explorer hierarchy expectations."""
    # Mimic a set of parsed files from a context bundle
    dummy_files = [
        {"path": "index.md", "text": "Index content"},
        {"path": "projects/air-cargo-ai/board.md", "text": "Board content"},
        {"path": "projects/air-cargo-ai/tasks/T-0001.md", "text": "T-0001 content"},
        {"path": "projects/air-cargo-ai/tasks/T-0002.md", "text": "T-0002 content"},
        {"path": "skills/stickshift-viewer.md", "text": "Viewer skill content"},
    ]
    
    tree = build_tree_python(dummy_files)
    
    # Root level files and folders
    assert len(tree["files"]) == 1
    assert tree["files"][0]["name"] == "index.md"
    assert "projects" in tree["children"]
    assert "skills" in tree["children"]
    
    # Under projects
    projects = tree["children"]["projects"]
    assert "air-cargo-ai" in projects["children"]
    assert len(projects["files"]) == 0
    
    # Under projects/air-cargo-ai
    air_cargo = projects["children"]["air-cargo-ai"]
    assert len(air_cargo["files"]) == 1
    assert air_cargo["files"][0]["name"] == "board.md"
    assert "tasks" in air_cargo["children"]
    
    # Under projects/air-cargo-ai/tasks
    tasks = air_cargo["children"]["tasks"]
    assert len(tasks["files"]) == 2
    task_names = {f["name"] for f in tasks["files"]}
    assert task_names == {"T-0001.md", "T-0002.md"}
    assert len(tasks["children"]) == 0
    
    # Under skills
    skills = tree["children"]["skills"]
    assert len(skills["files"]) == 1
    assert skills["files"][0]["name"] == "stickshift-viewer.md"


def resolve_relative_path_python(base_path: str, relative_path: str) -> str:
    """Python twin of the JavaScript resolveRelativePath() function."""
    if "://" in relative_path or relative_path.startswith("/") or relative_path.startswith("\\"):
        clean = re.sub(r"^[a-zA-Z]+://", "", relative_path)
        return clean
    
    parts = base_path.split("/")
    parts.pop() # remove current filename
    
    rel_parts = relative_path.split("/")
    for p in rel_parts:
        if p == "." or p == "":
            continue
        if p == "..":
            if parts:
                parts.pop()
        else:
            parts.append(p)
    return "/".join(parts)


def test_relative_path_resolver() -> None:
    """Test relative path resolver logic."""
    base = "projects/air-cargo-ai/tasks/T-0001.md"
    
    # Sibling file
    assert resolve_relative_path_python(base, "T-0002.md") == "projects/air-cargo-ai/tasks/T-0002.md"
    
    # Parent folder file
    assert resolve_relative_path_python(base, "../board.md") == "projects/air-cargo-ai/board.md"
    
    # Multi-level relative paths
    assert resolve_relative_path_python(base, "../../index.md") == "projects/index.md"
    
    # Absolute file URL
    assert resolve_relative_path_python(base, "file:///C:/StickShift/projects/index.md") == "/C:/StickShift/projects/index.md"


def test_markdown_link_regex() -> None:
    """Test parsing markdown links [text](url) into HTML links."""
    link_pattern = re.compile(r"\[([^\]\n]+)\]\(([^)\n]+)\)")
    
    text = "Please check the [Kanban Board](../board.md) for updates."
    replaced = link_pattern.sub(r'<a href="\2">\1</a>', text)
    assert replaced == 'Please check the <a href="../board.md">Kanban Board</a> for updates.'
    
    # Bold links
    bold_text = "See [**T-0001**](tasks/T-0001.md)"
    replaced_bold = link_pattern.sub(r'<a href="\2">\1</a>', bold_text)
    assert replaced_bold == 'See <a href="tasks/T-0001.md">**T-0001**</a>'

