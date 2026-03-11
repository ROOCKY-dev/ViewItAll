# ViewItAll-md — Agent Memo
# Long-term memory shared across all agents and sessions.
# Each agent documents their experience and observations about the engineer.
# READ ALL PRIOR SECTIONS before starting work in a new session.

---
---

# #1 | Copilot-S46 | 2026-03-11 :

## Engineer Profile
Ahmed is a hands-on product engineer who thinks in sprints and iterates fast. He gives high-level intent ("make it native", "no plugin vibe") and trusts the agent to fill in the technical detail. He will approve or redirect — he doesn't micromanage implementation.

## Communication Style
- **Shorthand heavy.** Ahmed writes quickly and informally — typos and abbreviations are normal. Parse intent, not literal words. "verifiy the Pdf color picker" means "verify + test the PDF colour picker feature". "acheved" = "achieved".
- **Brief approvals.** "great" / "approved" / "great job" mean: what was done is good, proceed to the next logical step.
- **Direct redirects.** If he doesn't like something he'll say so simply — no need to over-explain or apologise.

## Preferences Observed

### Code
- No hardcoded values anywhere — Ahmed specifically called this out when asking for modularisation (Sprint 4). If a value should be user-configurable, put it in settings.
- Modular, clean file structure. He approved the Sprint 4 modularisation explicitly.
- TypeScript strict; he expects the build to be clean before anything is called done.

### UI/UX
- **"Zero plugin vibe"** is his exact phrase — the gold standard for this project. He wants the plugin to feel invisible, as if Obsidian natively handles these file types.
- He cares deeply about visual polish. He asked for a note redesign unprompted just because it "didn't look right".
- Native Obsidian patterns (Lucide icons, CSS vars, `setIcon`, `setTooltip`, `addAction`) satisfy him. Emojis and hardcoded colours do not.

### Process
- Sprint-based delivery. Ahmed thinks in sprints; structure proposals accordingly.
- He asks for plans before implementation (use `[[PLAN]]` prefix sessions).
- He reviews and approves before saying "start" or "go ahead".
- He prefers one commit per sprint or logical batch — not many micro-commits.

### Git
- **Critical:** NEVER use `Co-authored-by:` trailer. He explicitly corrected this. Use "Assisted by GitHub Copilot" in the commit body only.
- Conventional commit format: `type(scope): description (vX.Y.Z)`

### Keyboard
- Ahmed chose Alt as the snap modifier key. He wanted it as a default.
- He thinks in "hold to activate + tap to cycle" patterns for modifier-based interactions.
- He expects shortcuts to be configurable in settings.

## What Ahmed Values Most (in order)
1. **Native feel** — UI that matches Obsidian's own quality bar
2. **No hardcoded config** — everything user-configurable
3. **Clean builds** — TypeScript errors are not acceptable to ship
4. **Feature completeness** — finish what is started before moving on
5. **Good defaults** — settings should work out of the box without tuning

## Implicit Requirements I Noticed
- Ahmed never explicitly said "the note overlay looks bad" but asked for a redesign after seeing the first version. Keep an eye on new UI components — they likely need a visual pass before they're truly done.
- He mentioned "erase" and "highlight" alongside pen snapping before I listed them — he thinks about tool parity. When adding a feature to one tool, check if it applies to the others.
- He asked about the `.toon` format — he's curious about agent tooling and LLM-specific formats. He may want to explore proper TOON format for requirements/constraints files in a future session.

## Open Questions for Next Agent
- Ahmed mentioned "we can even make our own file handling library or method only if necessary" for future file types. Clarify with him whether he wants a unified file handler abstraction (a `ViewFactory` pattern) before Sprint 6 begins.
- The `.toon` files currently use a TOML-like custom format, not the real TOON spec. Consider asking Ahmed if he wants to adopt proper TOON format for machine-readable files in a future session.
