---
okf_version: "0.1"
type: Skill
title: Skill Retrieval Optimizer
description: Rewrites a skill's description and trigger surface so the LLM actually selects it at the right moment, because the description is the only thing visible before a skill fires. Use this skill when a skill is not firing when it should, when the operator says a skill "is not being picked" or "should have triggered," when the operator wants to make a skill more discoverable or tune its description, or when a newly authored skill needs its retrieval surface checked before rollout.
tags: [skill, meta, retrieval, discoverability]
---

# skill-retrieval-optimizer

## Purpose

Make a skill get selected when it should, and stay quiet when it should not. You are not
changing what a skill does - you are rewriting the surface the model reads when it decides
whether to load it.

The thing being optimized is almost always the `description` field. Everything else in a
skill file (the body, the procedure, the examples) is invisible at selection time. So if a
skill is not firing, the body is rarely the cause and rarely the fix.

## The core principle

A skill is selected in two hops. Hop one: the model sees a manifest - just the `title` and
`description` of every available skill - and picks one. Hop two: only then is the chosen
skill's body loaded. The body cannot influence a decision that was already made before it
was read.

This has one hard consequence that drives everything below: **all retrieval signal must
live in the description.** A perfect procedure with a vague description never runs. The
description is not a summary of the skill; it is the skill's bid to be chosen.

## When to use this skill

- The operator reports a skill that did not fire when it obviously should have ("I asked
  for an email draft and the email composer never triggered").
- The operator wants a skill to be more discoverable, or to tune/rewrite its description.
- A new skill is about to ship and its retrieval surface has not been pressure-tested.
- Two skills keep getting confused for each other (a collision - see below).

Do not use this skill to change what a skill does, fix a bug in a skill's procedure, or
write a new skill from scratch. This is surface work only.

## What you need from the operator

Ask for, at minimum, the target skill's current `title` and `description`, and its body if
diagnosing a miss (you may need to see what it actually does to name its triggers well).

Ask also, if you can get it, for the `title` + `description` of the neighboring skills it
might be confused with. Retrieval is competitive: a description is good or bad only relative
to its manifest siblings. You cannot optimize distinguishability against skills you cannot
see, so if siblings are withheld, optimize in isolation and flag collision risk explicitly.

## Why a skill fails to fire (diagnose before rewriting)

1. **Abstraction gap.** The description names the concept the author thinks in
   ("communication artifact generation") instead of the words the operator types ("draft an
   email," "write a reply"). The model matches the operator's words, not the author's.
2. **Missing surface forms.** The description covers one phrasing but not its synonyms and
   inflections (email / message / reply / note; draft / write / compose / send).
3. **No trigger clause at all.** The description says what the skill does but never says
   when to use it, so the model has nothing situational to match against.
4. **Collision.** Two descriptions claim overlapping trigger space, so the model picks the
   wrong one or hesitates. Fixed by sharpening the boundary in both descriptions, not by
   making one louder.
5. **Over-broad.** The description fires on everything, crowding out more specific skills.
   The fix is to add negative space (what it is NOT for).

## Procedure

1. **Name the real triggers.** List the actual phrasings an operator would type to want this
   skill - terse, verbose, and oblique. These are concrete utterances, not topics. For an
   email skill: "draft an email to X," "reply to this," "write a note declining," "send them
   a follow-up." Pull the recurring verbs and nouns out of that list.
2. **Collect surface forms.** For each trigger noun and verb, add the synonyms and
   inflections operators actually use. Put them in the description even if they feel
   redundant - redundancy in the discovery surface is a feature, since you are matching
   unpredictable phrasings.
3. **Mark the boundary.** Write one clause of negative space: the nearest thing this skill
   is NOT for, especially the sibling it collides with. This is what keeps it from firing on
   a neighbor's territory.
4. **Rewrite the description** to the template below.
5. **Mirror, do not duplicate, in the body.** Update the body's "When to use" section so a
   human reading the file sees the same triggers and the same boundary. The body does not
   drive selection, but a drifted body confuses the next author.
6. **Build the validation set** (next section) and hand it to the operator.

## Description template

    <One sentence: what the skill produces or does, in plain operator words.>
    Use this skill when <primary trigger phrasing>, when <situational trigger>, or when
    <alternate phrasing / synonym cluster>. Do not use it for <nearest neighbor / boundary>.

Keep it to two or three sentences. Lead the trigger clause with the single most common
phrasing - it carries the most matching weight. Spend words on triggers, not on describing
internal mechanics the model cannot use to decide.

## Validation (the operator runs this on the DHSChat side)

You cannot observe whether your rewrite worked - only the operator can, by watching what
fires in a live session. So your deliverable is not just a new description; it is a new
description plus a test the operator can run.

Produce two labeled lists:

- **Should-fire (positives):** 6-10 phrasings that ought to select this skill. Vary them
  hard - one terse, one wordy, one using a synonym, one oblique that names the goal not the
  tool ("let them know I cannot make the meeting" for an email skill).
- **Should-not-fire (negatives):** 3-5 near-miss phrasings that should select a sibling or
  nothing - the boundary cases. These catch over-broad descriptions.

The operator pastes each phrasing into DHSChat and records which skill fired. The target is:
every positive fires this skill, every negative does not. If a positive misses, that
phrasing's words are not in the description - add them. If a negative fires, the description
is too broad - sharpen the boundary clause. Re-run until both lists pass. This is the loop;
the operator's live session is the oracle.

## Example

Operator: "I asked DHSChat to draft an email and the email composer skill never fired."

Current description (the miss):

    Generates professional communication artifacts from operator intent.

Diagnosis: abstraction gap (#1) and no surface forms (#2). The operator typed "draft an
email"; the description says "communication artifacts." None of the operator's words are
present, so there was nothing to match.

Rewritten description:

    Drafts an email, message, or reply from a short instruction, ready to copy into the
    operator's mail client. Use this skill when the operator asks to draft, write, compose,
    or reply to an email or message, when they say "send them a note" or "let them know," or
    when they describe what to communicate and to whom. Do not use it for editing an existing
    document's prose - that is the editor skill.

Validation set handed to the operator:

    Should-fire:
      - "draft an email to Heather about the detail extension"
      - "reply to Paul saying the install worked"
      - "write a short note declining the meeting"
      - "compose a follow-up to the vendor"
      - "let them know I cannot make Thursday"      (oblique - names goal, not tool)
      - "send Mya a thank-you message"
    Should-not-fire:
      - "tighten the wording in this paragraph"      (-> editor skill)
      - "summarize this email thread"                (-> summarizer, not composer)
      - "shard builds/foo.md"                         (-> doc-sharding)

The operator runs these on DHSChat; if "let them know I cannot make Thursday" misses, add
"let them know / tell them" to the trigger clause and re-run.

## Edge cases

- **Two skills must share territory.** If the boundary is genuinely fuzzy, write the boundary
  clause into BOTH descriptions so each points at the other ("for X use this; for Y use the
  other skill"). Sharpening one side alone just moves the collision.
- **The operator's phrasing is idiosyncratic.** Optimize for how THIS operator actually
  talks, observed from their real requests, over generic phrasings. The manifest only has to
  serve its actual users.
- **You were given no siblings.** Optimize the description in isolation, but tell the operator
  the rewrite cannot account for collisions, and recommend re-running validation against the
  full live manifest where neighbors compete.
