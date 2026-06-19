---
type: Build
title: Active Links Archived
description: A spec-stage build that mistakenly depends on an archived build.
status: spec
effort: M
impact: med
domain: TSA
timestamp: 2026-06-17T00:00:00Z
last_touched: 2026-06-17
---

# What it is
A build in spec that lists an archived build as a dependency. The link should
be retargeted to an active replacement before this build can proceed.

# Next action
Retarget the dependency link to a non-archived build.

# Dependencies
* [Archived Build](/builds/archived-build.md) — this dependency is archived.

# Notes
Used to verify the active-to-archived link check. The linter should emit a
warning because the source (spec) is not archived but the target is.
