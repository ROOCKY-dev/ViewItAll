# ViewItAll-md — UI/UX Design Rules & Requirements

> **One-line mandate:** Every UI element must feel like it was built by the Obsidian team — not added by a plugin.

---

## 1. Core Philosophy

### The "Zero Plugin Vibe" Rule
A user who doesn't know ViewItAll-md is installed should not be able to tell a plugin is rendering their files. This means:
- No custom colour palettes that clash with the active theme
- No bespoke icon sets — use Lucide (what Obsidian uses)
- No floating UI that looks "app-like" outside Obsidian's chrome
- Toolbar density and spacing must match Obsidian's own toolbars

### Native First
Before building any custom component, ask: *Does Obsidian already provide this?*
- Modal → use `Modal` class with `setTitle()`, `modal-button-container`, `.mod-cta`
- Icon button → `.clickable-icon` + `setIcon()`
- Tooltip → `setTooltip()` (never `element.title = ...`)
- View header action → `addAction()` (not a toolbar button)
- Notice → `new Notice(message)` (not a custom toast)

---

## 2. CSS Token Rules

### Mandatory Tokens
All visual properties must reference Obsidian CSS variables:

| Property | Use |
|----------|-----|
| Background | `--background-primary`, `--background-secondary`, `--background-modifier-hover` |
| Text | `--text-normal`, `--text-muted`, `--text-faint`, `--text-accent`, `--text-error` |
| Borders | `--background-modifier-border`, `--background-modifier-border-hover` |
| Accent / interactive | `--interactive-accent`, `--interactive-accent-hover` |
| Radius | `--radius-s`, `--radius-m`, `--radius-l` |
| Spacing | `--size-4-1` through `--size-4-8` |
| Font size | `--font-ui-small`, `--font-ui-medium`, `--font-ui-large` |
| Font family | `--font-ui`, `--font-text`, `--font-monospace` |
| Shadow | `0 Xpx Ypx rgba(0,0,0,N)` — only rgba with 0-alpha is permitted as a non-token value |

### Banned in CSS
```
❌  #3a5aef           (hardcoded hex)
❌  color: blue        (named colour)
❌  font-size: 13px    (hardcoded size)
❌  background: white  (hardcoded colour)
```

---

## 3. Component Guidelines

### Toolbar
- Position: top by default, user-configurable to bottom
- Height: minimum 36px, align-items: center
- Separator: `.via-toolbar-sep` — 1px line using `--background-modifier-border`
- Overflow: `overflow-x: auto; scrollbar-width: none` — never wrap or clip buttons
- Spacer: `.via-toolbar-spacer` (flex: 1) to push secondary actions right
- **Primary tools** (used constantly) → in toolbar
- **Secondary actions** (infrequent) → view header via `addAction()`

### Tool Group Pill
- Wrap related toggle buttons in `.via-tool-group`
- Background: `--background-modifier-hover` (pill background)
- Active button: adds `is-active` class → white card lift effect
- Use `setIcon()` + `setTooltip()` on each button — no text labels

### Color Picker Pattern
- **Never** inline swatches/sliders in the toolbar
- Use a single coloured dot button (`.via-color-dot-btn`) 
- Click opens a floating `.via-color-popover` panel positioned below the dot
- Popover contains: preset swatches → separator → size slider → opacity slider (if applicable)
- Dismiss: click outside (`mousedown` on `document`, capture phase)
- Destroy popover on tool switch or file unload

### Icon Buttons
```typescript
// Correct pattern
const btn = el.createEl('div', { cls: 'clickable-icon' });
setIcon(btn, 'lucide-icon-name');
setTooltip(btn, 'Human-readable action name');
btn.addEventListener('click', handler);
```

### Modals
```typescript
// Correct pattern
class MyModal extends Modal {
  onOpen() {
    this.setTitle('Confirm Action');  // NOT createEl('h3')
    this.contentEl.createEl('p', { text: 'Description text' });
    const btnRow = this.contentEl.createEl('div', { cls: 'modal-button-container' });
    btnRow.createEl('button', { text: 'Cancel' }).addEventListener('click', ...);
    btnRow.createEl('button', { text: 'Confirm', cls: 'mod-cta' }).addEventListener('click', ...);
  }
}
```

### Text Notes (PDF overlays)
- **Background**: `--background-primary` (neutral card — never the user's chosen colour as fill)
- **Colour identity**: 4px `border-top` in `var(--note-color)` + glow ring on hover/focus
- **Header**: visible only on hover (`opacity: 0 → 1` transition) — clean default state
- **Resize**: `resize: both`, `min-width: 140px`, `max-width: 400px`
- **Text area**: `--text-normal`, `--font-text`, italic `--text-faint` placeholder

### Snap Button
- Icon + short text label (H / V / 45°) — not icon alone
- Border: `1px solid --background-modifier-border`
- Active (snapping in progress): `--interactive-accent` background, `--text-on-accent`

---

## 4. Interaction Patterns

### Keyboard Shortcuts
- All shortcuts must be configurable in the Settings tab — no hardcoded keys
- Display the configured key in `setTooltip()` text: `Pen (${key.toUpperCase()})`
- Document-level key handlers must guard with `getActiveViewOfType(ViewClass) === this`

### Loading States
- Show `.via-pdf-loading` spinner while parsing/rendering
- Use `new Notice('message')` for operations longer than ~2s
- Never block the UI thread — use `async/await` with `requestAnimationFrame` for canvas redraws

### Error States
- `.via-error` class for in-view error messages
- `.via-warning` for non-fatal warnings (e.g., mammoth.js conversion warnings)
- `new Notice('Error: ' + message)` for critical failures

---

## 5. Per-View UX Goals

### PDF Viewer
- Goal: feel like Obsidian's built-in PDF viewer but with annotation superpowers
- The annotation toolbar should disappear when in "View" mode (cursor tool active)
- Zoom label is the only text button — everything else is icon-only
- Page indicator is subtle; page numbers appear below each page canvas

### DOCX Viewer
- Goal: feel like a lightweight Google Docs embedded in Obsidian
- View mode: read-only, no visible toolbar controls except the edit toggle
- Edit mode: clean `outline` on the content area using `--interactive-accent`
- Dirty state: small yellow dot (`.via-docx-dirty-dot`) — never text like "Unsaved"

### Future Viewers (Spreadsheet, Presentation, eBook)
- Every new viewer must go through the "zero plugin vibe" checklist before shipping
- Spreadsheets: table must use Obsidian's table border tokens, not browser defaults
- Presentations: slide navigation should feel like a minimal, focused reading mode
- eBooks: reading mode with font size preference stored in settings

---

## 6. Quality Checklist for New UI Components

Before shipping any new UI element, verify:
- [ ] All colours come from CSS variables — no hardcoded values
- [ ] Icons use `setIcon()` with a Lucide icon name
- [ ] Tooltips use `setTooltip()` — not `title` attribute
- [ ] Interactive elements have `.clickable-icon` or `via-btn` with proper hover states
- [ ] The component looks correct in both light and dark Obsidian themes
- [ ] The component does not interfere with keyboard navigation
- [ ] Error states are handled and communicated to the user
- [ ] Any dynamically created elements are cleaned up in `onUnloadFile()` or equivalent
