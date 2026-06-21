OKF — READ ME FIRST
A manual-activation knowledge agent. It runs entirely inside this workbook.

------------------------------------------------------------
WHAT THIS IS
------------------------------------------------------------
This connects your AI chat to a folder of markdown notes (the "bundle").
The AI asks for context, or proposes file edits. You copy what it gives you,
click a button here, and the workbook does the work.

Nothing runs on its own. Nothing happens until you click. You stay in the loop.

------------------------------------------------------------
SET UP — ONE TIME
------------------------------------------------------------
1. Enable macros: when the file opens, click the yellow bar at the top
   ("Enable Content" / "Enable Editing").
   (If you ever see a red "macros have been blocked" bar instead: close Excel,
   right-click the file -> Properties -> tick "Unblock" at the bottom -> OK,
   then reopen.)

2. Make a folder and add the starter files: create a folder on your C: drive —
   for example C:\OKF — and copy in the starter bundle (Luke will give you these;
   it includes a "skills" folder). The folder name doesn't matter; the tool
   remembers the path, not the name.

3. Point the tool at it: click button 5, "Set Bundle Root", and pick the folder
   you just made. Stored per machine — you only do this once.

4. Build the listings: click button 2, "Regenerate Index", once. This makes the
   bundle's index.md listings current so the AI can see what's in there.

------------------------------------------------------------
THE LOOP — THIS IS THE WHOLE THING
------------------------------------------------------------
   copy from chat  ->  click a button here  ->  get the result  ->  back to chat

Two buttons do the real work:

   Button 4 — BUILD CONTEXT BUNDLE
     Use when the AI replies with a <CONTEXT_REQUEST> block.
     Copy the whole block, click button 4, and the assembled context opens in
     the -dist folder as OKF-context.md. Paste that back into the chat.

   Button 1 — APPLY WRITE ENVELOPE
     Use when the AI replies with a <VBA_WRITE> block (file edits).
     Copy it, click button 1, and the files are written straight into your
     bundle. Every write is recorded in log.md, so you can always see what
     changed and when.

------------------------------------------------------------
TRY THIS FIRST — your first round-trips
------------------------------------------------------------
Your bundle comes seeded with a "skill" — a reusable procedure the AI can pick
up and follow. This one happens to teach the AI how to write more skills. So
your first session is: see it find the skill, then use it to make your own.

ROUND-TRIP 1 — see retrieval work:
1. In your AI chat, ask: "What skills are in my bundle?"
2. The AI replies with a <CONTEXT_REQUEST> block — that's it asking to look in
   the bundle. Copy the whole block.
3. Click button 4, "Build Context Bundle". The -dist folder opens with
   OKF-context.md. Paste its contents back into the chat.
4. The AI now lists your skills — including "Skill MD Authoring".

ROUND-TRIP 2 — make something of your own:
5. Ask: "Use the skill-authoring skill to help me write a new skill for
   <a task you actually do>."
6. The AI may ask for the full procedure first (another <CONTEXT_REQUEST> — same
   button 4, same steps). Then it walks you through it and hands you a
   <VBA_WRITE> block.
7. Copy that block, click button 1, "Apply Write Envelope".
8. Open your skills folder — your new skill is there. Open log.md — the write is
   recorded, with a timestamp.

That is the whole agent: it found context for you, helped you build something,
wrote it where it belongs, and kept a record — and nothing happened until you
clicked. You were in the loop the entire time.

------------------------------------------------------------
THE OTHER BUTTONS — not needed on day one
------------------------------------------------------------
   Button 2 — Regenerate Index : rebuilds the index.md listings. This runs
              automatically after every write; button 2 is just the manual version.
   Button 3 — Run Linter       : checks the bundle for problems (broken links,
              stalls, etc.). Findings appear in the "OKF Lint Report" sheet.

------------------------------------------------------------
IF YOU GET STUCK
------------------------------------------------------------
Note exactly where you hesitated — that spot is genuinely useful, it tells us
what to make smoother for the next person.

Want to see how it works under the hood? Ask Luke to walk you through the
modules — but do it AFTER your first run, not before. Feel the loop first.
