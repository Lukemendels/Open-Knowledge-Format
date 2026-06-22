StickShift — READ ME FIRST
Everything automatic AI promises — except your hand's on the StickShift.

A manual-activation knowledge agent. It runs entirely inside this workbook.

------------------------------------------------------------
WHAT THIS IS
------------------------------------------------------------
This connects your AI chat to a folder of markdown notes (a "context").
The AI asks for context, or proposes file edits. You copy what it gives you,
click a button here, and the workbook does the work.

Nothing runs on its own. Nothing happens until you click. You stay in the loop.

------------------------------------------------------------
YOUR CONTEXT  (the bar at the top)
------------------------------------------------------------
The top of the dashboard shows your current context - the folder StickShift is
pointed at right now. The "Switch Context" button changes it.

You can keep more than one: a personal context, a shared team context, a
per-project context. Switch between them anytime; StickShift remembers the last
one per machine.

------------------------------------------------------------
SET UP - TWO CLICKS
------------------------------------------------------------
1. Enable macros: when the file opens, click the yellow bar at the top
   ("Enable Content").
   (If you see a red "macros have been blocked" bar instead: close Excel,
   right-click the file -> Properties -> tick "Unblock" -> OK, then reopen.)

2. Set your context: click "Switch Context" in the top bar and pick (or make) a
   folder - for example C:\StickShift. The name doesn't matter; the tool
   remembers the path.

3. Click button 1, "Initialize Context". This seeds the starter files -
   including a skill that teaches the AI how to write more skills - and builds
   the listings the AI reads. That's setup done.

   (Switching to a context someone already set up? Skip step 3 - just start
   using it.)

------------------------------------------------------------
THE LOOP - THIS IS THE WHOLE THING
------------------------------------------------------------
   copy from chat  ->  click a button here  ->  get the result  ->  back to chat

Two buttons do the real work:

   Button 2 - BUILD CONTEXT BUNDLE
     Use when the AI replies with a <CONTEXT_REQUEST> block.
     Copy the whole block, click button 2, and the assembled context opens as
     StickShift-context.md. Paste that back into the chat.

   Button 3 - APPLY WRITE ENVELOPE
     Use when the AI replies with a <VBA_WRITE> block (file edits).
     Copy it, click button 3, and the files are written into your context.
     Every write is recorded in log.md, so you can always see what changed.

------------------------------------------------------------
TRY THIS FIRST - your first round-trips
------------------------------------------------------------
Your context comes seeded with a "skill" - a reusable procedure the AI can pick
up and follow. This one teaches the AI how to write more skills. So your first
session is: watch it find the skill, then use it to make your own.

ROUND-TRIP 1 - see retrieval work:
1. In your AI chat, ask: "What skills are in my context?"
2. The AI replies with a <CONTEXT_REQUEST> block - that's it asking to look in
   the folder. Copy the whole block.
3. Click button 2, "Build Context Bundle". StickShift-context.md opens. Paste
   its contents back into the chat.
4. The AI now lists your skills - including "Skill MD Authoring".

ROUND-TRIP 2 - make something of your own:
5. Ask: "Use the skill-authoring skill to help me write a new skill for
   <a task you actually do>."
6. The AI may ask for the full procedure first (another <CONTEXT_REQUEST> - same
   button 2). Then it walks you through it and hands you a <VBA_WRITE> block.
7. Copy that block, click button 3, "Apply Write Envelope".
8. Open your skills folder - your new skill is there. Open log.md - the write is
   recorded, with a timestamp.

That is the whole agent: it found context for you, helped you build something,
wrote it where it belongs, and kept a record - and nothing happened until you
clicked. You were in the loop the entire time.

------------------------------------------------------------
THE OTHER BUTTONS - not needed on day one
------------------------------------------------------------
   Button 4 - Regenerate Index : rebuilds the listings. Runs automatically after
              every write; button 4 is the manual version.
   Button 5 - Run Linter       : checks the context for problems (broken links,
              stalls, etc.). Findings appear in the "StickShift Lint Report" sheet.

------------------------------------------------------------
IF YOU GET STUCK
------------------------------------------------------------
Note exactly where you hesitated - that spot is genuinely useful; it tells us
what to make smoother for the next person.
