"""
Unit tests for the StickShift Network HTML tool.
Validates ASCII encoding, HTML Tool Compliance Standard requirements,
and specific onboarding elements.
"""

from pathlib import Path
import re
import pytest

# Path to the HTML tool
TOOL_PATH = Path(__file__).parent.parent / "builds" / "html-tools" / "stickshift-network" / "stickshift-network.html"

def test_network_exists() -> None:
    """Ensure the network HTML tool exists in the expected location."""
    assert TOOL_PATH.exists(), f"Network tool not found at {TOOL_PATH}"
    assert TOOL_PATH.is_file()


def test_network_ascii_only() -> None:
    """Ensure the network HTML file contains strictly ASCII characters for VBA compatibility."""
    content = TOOL_PATH.read_bytes()
    try:
        content.decode("ascii")
    except UnicodeDecodeError as e:
        # Show context of decode error
        start = max(0, e.start - 40)
        end = min(len(content), e.end + 40)
        snippet = content[start:end]
        pytest.fail(
            f"Network tool contains non-ASCII characters at index {e.start} "
            f"(byte value 0x{content[e.start]:02x}). Context: {snippet!r}"
        )


def test_network_compliance_identity() -> None:
    """Ensure the window.STICKSHIFT_TOOL identity declaration is present and correct."""
    html_text = TOOL_PATH.read_text(encoding="utf-8")
    
    # Check for the window.STICKSHIFT_TOOL object structure
    # Match id, file, skillSlug, and title
    id_match = re.search(r'id:\s*["\']stickshift-network["\']', html_text)
    file_match = re.search(r'file:\s*["\']stickshift-network\.html["\']', html_text)
    slug_match = re.search(r'skillSlug:\s*["\']stickshift-network["\']', html_text)
    title_match = re.search(r'title:\s*["\']StickShift Network["\']', html_text)
    
    assert id_match is not None, "window.STICKSHIFT_TOOL is missing ID or has incorrect ID"
    assert file_match is not None, "window.STICKSHIFT_TOOL is missing file or has incorrect file"
    assert slug_match is not None, "window.STICKSHIFT_TOOL is missing skillSlug or has incorrect skillSlug"
    assert title_match is not None, "window.STICKSHIFT_TOOL is missing title or has incorrect title"


def test_network_compliance_embedded_skill() -> None:
    """Ensure the embedded companion skill block is present, correct, and matching."""
    html_text = TOOL_PATH.read_text(encoding="utf-8")
    
    # 1. Check for the script tag structure
    script_match = re.search(
        r'<script\s+type=["\']text/markdown["\']\s+id=["\']stickshift-skill["\']\s+data-skill-slug=["\']stickshift-network["\']>',
        html_text
    )
    assert script_match is not None, "Missing or incorrect #stickshift-skill script tag"
    
    # 2. Extract skill content
    skill_block_pattern = r'<script\s+type=["\']text/markdown["\']\s+id=["\']stickshift-skill["\']\s+data-skill-slug=["\']stickshift-network["\']>(.*?)</script>'
    skill_content_match = re.search(skill_block_pattern, html_text, re.DOTALL)
    assert skill_content_match is not None, "Failed to extract content from #stickshift-skill tag"
    
    skill_content = skill_content_match.group(1).strip()
    
    # 3. Check for required elements in the skill markdown
    assert "type: Skill" in skill_content, "Skill frontmatter missing type: Skill"
    assert "title: StickShift Network" in skill_content, "Skill frontmatter title is incorrect"
    
    # Verify it instructs to open stickshift-network.html
    assert "<HTML_OPEN>" in skill_content, "Skill is missing <HTML_OPEN> block"
    assert "tool: stickshift-network.html" in skill_content, "Skill HTML_OPEN references incorrect tool file"
    assert "</HTML_OPEN>" in skill_content, "Skill is missing </HTML_OPEN> block"


def test_network_compliance_onboarding_elements() -> None:
    """Ensure onboarding and setup elements exist in the DOM as required."""
    html_text = TOOL_PATH.read_text(encoding="utf-8")
    
    # Verify presence of the onboarding/setup panel and buttons
    assert 'id="ssPanel"' in html_text, "Missing StickShift onboarding/setup panel container"
    assert 'id="ssYes"' in html_text, "Missing onboarding Yes button"
    assert 'id="ssNo"' in html_text, "Missing onboarding No button"
    assert 'id="ssCopySkill"' in html_text, "Missing onboarding Copy Skill button"
    assert 'id="ssFootBtn"' in html_text, "Missing footer button to toggle onboarding panel"
    assert 'id="loadAllBtn"' in html_text, "Missing Load all button"
