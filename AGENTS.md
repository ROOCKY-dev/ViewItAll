# ViewItAll-md — Agent Context
> Read this file at the start of **every** session. ~2 min read. Then `npm run build` to confirm baseline.

---

## 🔴 Critical Rules — Never Violate

1. **TypeScript strict. Zero `any`.** Use `unknown` + type guards or define a local interface.
2. **CSS: Obsidian variables only.** No hardcoded hex, named colours, or px font-sizes. Only `rgba(0,0,0,N)` for shadows.
3. **Icons: `setIcon()` + Lucide only.** No emoji as labels. `setTooltip()` not `.title =`.
4. **No React.** Native DOM — `createEl`, `setIcon`, `setTooltip`, `addAction`.
5. **No hardcoded config.** Every key, colour, width, flag reads from `PluginSettings`.
6. **Build must pass before commit.** `npm run build` — tsc + esbuild, zero errors.
7. **No `console.log()` in commits.** Use `new Notice()` for user messages.
8. **Heavy libs: dynamic `import()` only.** Never top-level import for pptx/epub/three.js etc.
9. **Commit format** — see Quick Reference below. **Never** add `Co-authored-by:` trailer.
10. **Docs first.** Read `docs/` before using any Obsidian API. Never guess at signatures.

**Rule priority when rules conflict:**
1. Build must pass → 2. No data loss → 3. Native feel → 4. Performance → 5. Code cleanliness

---

## 🟡 Engineer Profile

**Ahmed** — hands-on product engineer. Thinks in sprints. Gives high-level intent, trusts agents on technical detail. Approves or redirects; doesn't micromanage implementation.

| Trait | What it means |
|-------|---------------|
| Short approvals ("great", "approved") | What was done is good — proceed to next logical step |
| Informal writing / typos | Parse intent, not literal words |
| Direct redirects | No need to apologise — just adjust and continue |
| "Zero plugin vibe" | His gold standard. The plugin should feel invisible. |
| Cares about visual polish | Unsatisfied UI will be called out — pre-empt it |
| Sprint-based | Structure proposals as sprints with clear deliverables |
| Plans before code | Confirm understanding before implementing |

**Non-negotiables Ahmed has enforced:**
- No `Co-authored-by:` (corrected once, never forget)
- Modular code with all config in settings (Sprint 4 driving motivation)
- Native Obsidian patterns — enforced every sprint since Sprint 5

---

## 🔵 Patterns & Conventions

### File Structure
```
src/
  main.ts          # Lifecycle only — load, unload, register
  settings.ts      # PluginSettings interface, defaults, settings tab
  types.ts         # VIEW_TYPE_* constants, shared types
  views/           # FileView subclasses (one file per type, max 600 lines)
  views/pdf/       # PDF-specific controllers + pdfTypes.ts
  utils/           # Pure functions — no Obsidian state
styles.css         # All CSS (single file)
docs/              # Obsidian API docs — read before using any API
.Agent/            # This governance directory
scripts/           # Automation scripts (CSS checker, etc.)
```

### Obsidian Native Patterns
```typescript
// Icon button
const btn = el.createEl('div', { cls: 'clickable-icon' });
setIcon(btn, 'lucide-icon-name');
setTooltip(btn, 'Action (Shortcut)');
btn.addEventListener('click', handler);

// Tool group pill
const group = el.createEl('div', { cls: 'via-tool-group' });
// active state: btn.classList.toggle('is-active', condition)

// Secondary actions (infrequent) → view header, NOT toolbar
this.addAction('icon', 'Tooltip', callback); // in onload(), persists for view lifetime

// Modal
class M extends Modal {
  onOpen() {
    this.setTitle('Confirm');                                   // not createEl('h3')
    const row = this.contentEl.createEl('div', { cls: 'modal-button-container' });
    row.createEl('button', { text: 'Cancel' }).addEventListener('click', ...);
    row.createEl('button', { text: 'OK', cls: 'mod-cta' }).addEventListener('click', ...);
  }
}
```

### CSS Class Naming
All plugin classes: `via-` prefix (e.g. `.via-pdf-toolbar`, `.via-tool-group`)

### Keyboard Shortcuts
- All shortcuts configurable in PluginSettings
- Document-level handlers guard with: `this.app.workspace.getActiveViewOfType(MyView) === this`
- Display configured key in tooltip: `` `Action (${s.keyName.toUpperCase()})` ``

### Annotation Storage
- Sidecar file: `<basename>.annotations.json` alongside the PDF
- Use `vault.modify()` / `vault.modifyBinary()` — not `adapter.write*` — for proper Obsidian events

### Cleanup
Every `onUnloadFile()` must: destroy pdfDoc, disconnect observers, remove dynamic DOM, null out all element refs, call `hideColorPopover()`.

---

## 📋 Checklists

### New File Type
- [ ] `VIEW_TYPE_XXX` constant in `src/types.ts`
- [ ] `src/views/XxxView.ts` extending `FileView`
- [ ] Register in `src/main.ts` `onload()`
- [ ] `enableXxx: boolean` in `PluginSettings` (default `true`)
- [ ] Settings toggle in tab under "File Types" section
- [ ] `canAcceptExtension()` checks the toggle at runtime
- [ ] Lazy-import heavy library inside `onLoadFile()` with `await import()`
- [ ] Full cleanup in `onUnloadFile()`
- [ ] Check bundle: `ls -lh main.js` — must stay under 2MB
- [ ] Update `.Agent/STATUS.md` file type table
- [ ] Update `.Agent/requirements.yaml` status → `done`
- [ ] Bump minor version

### New Setting
- [ ] Typed field on `PluginSettings` interface
- [ ] Sensible default in `DEFAULT_SETTINGS`
- [ ] UI control in correct settings tab section
- [ ] Read from `this.plugin.settings` at runtime — never hardcode
- [ ] `this.plugin.saveSettings()` after any change

### Session End
- [ ] `npm run build` passes cleanly
- [ ] `npm run lint` passes (or new violations are intentional and noted)
- [ ] Update `.Agent/STATUS.md`
- [ ] Append to `.Agent/DECISIONS.md` if any non-obvious choices were made
- [ ] Commit follows format in Quick Reference

### Rollback
```bash
# Undo last commit (keep changes staged)
git reset --soft HEAD~1

# Discard all uncommitted changes
git checkout -- .

# Restore a specific file to last commit
git checkout HEAD -- src/views/PdfView.ts

# Stash work in progress
git stash && git stash pop   # restore later
```

---

## ⚡ Quick Reference

```bash
npm run build              # tsc + esbuild (must pass before commit)
npm run dev                # watch mode
npm run lint               # ESLint all src/ files
npm run lint:css           # check styles.css for hardcoded values
npm version patch --no-git-tag-version   # bump + sync manifest + versions.json
npm version minor --no-git-tag-version
```

**Commit format:**
```
type(scope): short description (vX.Y.Z)

- bullet detail

Assisted by GitHub Copilot
```
Types: `feat` | `fix` | `refactor` | `style` | `docs` | `chore` | `perf`
**No** `Co-authored-by:` trailer — ever.
