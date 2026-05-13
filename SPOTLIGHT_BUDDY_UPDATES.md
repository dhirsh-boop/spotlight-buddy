# Spotlight Buddy — Progress Sheet

**Current version:** V1.8.2
**Last updated:** May 13, 2026

---

## What is Spotlight Buddy?

Spotlight Buddy is an internal tool that turns a finished video recording into a **ready-to-edit Adobe Premiere Pro project** with one click. Instead of a senior editor manually placing every speaker name, quote card, and slide on the timeline, Spotlight Buddy reads the transcript and a few supporting documents, then generates a script that builds the entire graphics layer automatically.

**Who uses it:** Producers and editors on WebMD video programs.
**What it saves:** Hours of manual timeline work per program.
**How it's used:** Upload the transcript + program PDF (and optionally slide markers + slide images), click Generate, download a zip, and run it in Premiere.

---

## Progress Timeline

*Listed oldest to newest. Each entry describes what got better and why it matters for the team.*

---

### Phase 1 — Foundation (versions 1.0 through 1.6)

**What was built:**
- The core idea was proven: an AI reads the video transcript and decides where graphics should appear.
- Integrated Google's Gemini AI to read transcripts and choose appropriate graphics at appropriate moments.
- Built the bridge to Premiere Pro using ExtendScript (the language Adobe uses for automation).

**Why it matters:**
- Established that this kind of automation is actually feasible inside our existing Premiere workflow — no new software for editors to learn.

---

### Phase 2 — The Graphics Library (versions 1.6 → 1.7)

**What was built:**
- Wired in all six WebMD-standard graphic types: Program Title, Speaker Intro, Banner Quote, Short Quote, Full-Screen Quote, and Bullet-Point Lists.
- Added an optional PDF upload so the AI can pull the **exact** program title, speaker credentials, and learning objectives — no more typos or guesses.
- Built smart property-matching so the tool works even when MOGRT field names vary between templates.
- Added a Manager Overview CSV alongside the Premiere script so reviewers can see what was generated without opening Premiere.

**Why it matters:**
- The tool now handles every graphic type the team uses on a typical program.
- Producers get a human-readable summary of every graphic placed — easy quality check before sending to the editor.

---

### Phase 3 — Reliability Hardening (version 1.7 → 1.8.0)

**What was built:**
- Hardened the AI-to-Premiere handoff so long transcripts don't get cut off mid-output.
- Added detailed logging — every run writes a debug report to the Desktop so we can troubleshoot quickly if something looks off.
- Locked the prompt and ExtendScript engine as the "stable core" that future features must build *around*, not on top of.

**Why it matters:**
- The tool is now dependable enough to use on real deliverables, not just experiments.
- New features can be added without risking the parts editors already rely on.

---

### Phase 4 — Automated Slide Syncing (version 1.8.2 — current)

*Built during the week of May 4, following the April 30 kickoff with Marina.*

**What was built:**
- The tool can now **automatically place full-screen slides** on the Premiere timeline — not just MOGRT graphics.
- The Medical Education team marks slides directly in Descript using simple labels like `slide 1`, `slide 2`. Spotlight Buddy reads the Descript export and matches each marker to the corresponding slide image.
- The producer downloads the slide deck from PowerPoint Online as a zip of JPEGs (PowerPoint does this natively) and uploads it. No new export tool needed.
- Slides land on their own dedicated track in Premiere, separate from the graphics track — editors can adjust each layer independently.
- Built-in **safeguards**:
  - Slides never appear while a graphic is on screen (no visual collisions to clean up).
  - Each slide stays on screen for a maximum of **12 seconds**, so the speaker video gets breathing room between slides.
  - Only slides the MED team explicitly marked will appear — nothing automatic or guessed.
- The slide-syncing feature is **optional**. If a program doesn't have slide markers, the tool still produces a graphics-only output exactly as before.
- Added a one-click launcher (`Launch Spotlight Buddy.command`) so the producer doesn't have to type any commands — just double-click and the app opens in their browser.

**Why it matters:**
- This was the big ask from Laura's feedback: complex slides that need on-screen time now get handled automatically instead of by hand.
- The MED team's workflow stays inside Descript — no new tools, no new training.
- The producer's job stays simple: drop in a zip of slides, click Generate.
- Editors get a clean, organized Premiere timeline with slides and graphics on separate tracks — easy to polish.

---

## What's Next (proposed)

- **MED "Important Slide" flag** — let the MED team mark certain slides as needing extra hold time (longer than the 12-second default). Discussed with Marina; deferred until MED defines the annotation format.
- **Migrate to the newer Google AI SDK** — the current `google.generativeai` library has been deprecated by Google. Low risk, but worth doing before Google stops shipping security patches.
- **Add release tagging discipline** (done as of v1.8.2) — every future change is now tracked in git history, so this update sheet writes itself going forward.
