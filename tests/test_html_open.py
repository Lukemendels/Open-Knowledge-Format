"""
Python twin of the HTML_OPEN block parser in StickShiftContextBundle.bas.
"""

import pytest
import re

def parse_html_open(text: str) -> tuple[str, int, list[str]]:
    """
    Parses an <HTML_OPEN> block and returns (tool, depth, seeds).
    
    If no <HTML_OPEN> block is found, raises ValueError.
    If 'tool:' is missing, returns an empty tool string.
    """
    match = re.search(r"<HTML_OPEN>(.*?)</HTML_OPEN>", text, re.DOTALL)
    if not match:
        raise ValueError("No <HTML_OPEN> block found")

    body = match.group(1).replace("\r\n", "\n").replace("\r", "\n")
    
    tool = ""
    depth = 0
    seeds = []
    
    lines = body.split("\n")
    in_include = False
    for line in lines:
        ln = line.strip()
        if ln.startswith("tool:"):
            tool = ln[5:].strip()
            in_include = False
        elif ln.startswith("depth:"):
            try:
                depth = int(ln[6:].strip())
            except ValueError:
                pass
            in_include = False
        elif ln == "include:":
            in_include = True
        elif in_include and ln.startswith("- "):
            seeds.append(ln[2:].strip())
        elif ln != "" and not ln.startswith("-"):
            in_include = False
            
    return tool, depth, seeds


# --- Unit Tests ---

def test_html_open_tool_only() -> None:
    text = (
        "<HTML_OPEN>\n"
        "tool: mytool.html\n"
        "</HTML_OPEN>"
    )
    tool, depth, seeds = parse_html_open(text)
    assert tool == "mytool.html"
    assert depth == 0
    assert seeds == []

def test_html_open_tool_and_include() -> None:
    text = (
        "<HTML_OPEN>\n"
        "tool: othertool.html\n"
        "include:\n"
        "- skills/my-skill.md\n"
        "- builds/build-a.md\n"
        "</HTML_OPEN>"
    )
    tool, depth, seeds = parse_html_open(text)
    assert tool == "othertool.html"
    assert depth == 0
    assert seeds == ["skills/my-skill.md", "builds/build-a.md"]

def test_html_open_tool_depth_and_include() -> None:
    text = (
        "<HTML_OPEN>\n"
        "tool: complex.html\n"
        "depth: 2\n"
        "include:\n"
        "- skills/my-skill.md\n"
        "</HTML_OPEN>"
    )
    tool, depth, seeds = parse_html_open(text)
    assert tool == "complex.html"
    assert depth == 2
    assert seeds == ["skills/my-skill.md"]

def test_html_open_ignore_after_non_list() -> None:
    text = (
        "<HTML_OPEN>\n"
        "tool: check.html\n"
        "include:\n"
        "- skills/my-skill.md\n"
        "depth: 1\n"
        "- builds/ignored.md\n"
        "</HTML_OPEN>"
    )
    tool, depth, seeds = parse_html_open(text)
    assert tool == "check.html"
    # "depth: 1" terminates "include:" list, so builds/ignored.md is not collected.
    assert depth == 1
    assert seeds == ["skills/my-skill.md"]

def test_html_open_no_tool_is_empty() -> None:
    text = (
        "<HTML_OPEN>\n"
        "depth: 1\n"
        "</HTML_OPEN>"
    )
    tool, depth, seeds = parse_html_open(text)
    assert tool == ""
    assert depth == 1
    assert seeds == []

def test_no_html_open_block_raises_value_error() -> None:
    with pytest.raises(ValueError, match="No <HTML_OPEN> block"):
        parse_html_open("some text without the envelope")
