---
okf_version: "0.1"
type: Skill
title: AutoReviewer (Tracked-Changes Tool)
description: Produces a Word document with real tracked changes from a review, using a local browser tool that writes OOXML markup directly (no Word, no COM). Use this skill when the operator wants tracked changes or redlines in a .docx, to mark up a document, or to turn review feedback into suggested edits.
tags: [skill, html-tool, document-review]
---

# autoreviewer-tool

## Purpose
Turn review feedback into a .docx with genuine tracked changes, via this local browser tool.
The tool runs on the operator's machine, takes their document, and downloads a marked-up copy.

## When to use this skill
- The operator wants tracked changes / redlines in a Word document.
- The operator has review feedback to apply as suggestions rather than as prose.

## How to open the tool
A local HTML tool cannot be opened from a chat hyperlink - the operator launches it from
StickShift. Give the operator this block verbatim (these exact lines, no code fence), then one
instruction line, and nothing else:

<HTML_OPEN>
tool: autoreviewer.html
include:
- skills/autoreviewer-tool.md
</HTML_OPEN>

Instruction line: "Copy the block above and click Open HTML Tool in StickShift."

## Then walk them through it
Once open: drop the .docx in, paste the review block when prompted, download the marked-up copy.
Proceed with the review content normally.
