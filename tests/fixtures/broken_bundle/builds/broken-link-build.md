---
type: Build
title: Broken Link Build
description: A production build that contains a link to a nonexistent file.
status: production
effort: S
impact: med
domain: TSA
timestamp: 2026-06-17T00:00:00Z
last_touched: 2026-06-17
---

# What it is
A build that has a broken cross-reference. The dependency link below points to
a file that does not exist in the bundle.

# Next action
Fix or remove the broken dependency link.

# Dependencies
* [Nonexistent Build](/builds/nonexistent.md) — this file does not exist.

# Notes
Used to verify the linter's broken-link check.
