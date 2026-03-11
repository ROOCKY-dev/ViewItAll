# ViewItAll-md — Decision Log
> Append-only. Add entries when a non-obvious choice is made. One entry per logical decision.
> Format: `## YYYY-MM-DD | SprintN | Topic`

---

## 2026-03-11 | Sprint 1 | Sidecar `.annotations.json`
**Decision:** Store annotations alongside the PDF as `<basename>.annotations.json`, not embedded in the PDF.  
**Why:** Non-destructive. Original PDF is never modified. Vault-sync friendly. Embeddable annotations would require `pdf-lib` write on every stroke (heavy) and would lose the diff-ability of the sidecar approach.

## 2026-03-11 | Sprint 1 | 3 Canvas Layers Per PDF Page
**Decision:** Stack `pdfCanvas` (z-index 1) / `annotCanvas` (z-index 2) / `searchCanvas` (z-index 3) per page.  
**Why:** Decouples rendering concerns. Annotation canvas can be cleared/redrawn on tool change without re-rendering the PDF page. Search highlights don't contaminate the annotation layer.

## 2026-03-11 | Sprint 1 | IntersectionObserver for Lazy Rendering
**Decision:** Only render PDF pages that are in the viewport; `unrender` off-screen pages after threshold.  
**Why:** Large PDFs (100+ pages) would OOM the tab if all pages were rendered upfront. Canvas memory is the constraint, not DOM nodes.

## 2026-03-11 | Sprint 1 | Annotation Coordinates as Normalised 0–1 Fractions
**Decision:** Store all annotation points as fractions of page dimensions, not pixel values.  
**Why:** Scale-independent. A user zooming from 100% to 200% or printing the page gets correctly positioned annotations. Avoids re-normalising saved data on load.

## 2026-03-11 | Sprint 1 | DOCX Edit via contentEditable HTML (mammoth → html-to-docx)
**Decision:** mammoth.js converts `.docx` → HTML for display and editing; html-to-docx converts back to `.docx` on save.  
**Why:** Avoids a full Office XML parser/writer. mammoth is read-only and loses some complex formatting, but that is an acceptable tradeoff for MVP. Round-trip fidelity is best-effort.

## 2026-03-11 | Sprint 2 | Bug — DOCX Save Used Buffer Not ArrayBuffer
**Decision (fix):** Use `buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)` to extract ArrayBuffer from Node Buffer.  
**Why:** Obsidian's `vault.modifyBinary()` requires an ArrayBuffer. `html-to-docx` returns a Node.js `Buffer` in the Electron environment, not a Web `Blob`. Slicing `.buffer` with correct byte offsets is the standard cross-environment fix.

## 2026-03-11 | Sprint 3b | Alt Key for Snap Modifier
**Decision:** Default snap activation key = `Alt`.  
**Why:** Ahmed specified Alt. It is a modifier key (not captured by typing), is not commonly bound by Obsidian for PDF views, and has good ergonomics for hold-to-activate interaction.

## 2026-03-11 | Sprint 3b | Alt + S to Cycle Snap Direction
**Decision:** One key (configurable, default `s`) pressed while modifier is held cycles H → V → 45° → H.  
**Why:** Ahmed's original concept. Single button concept for discoverability. Document-level keydown listener on `containerEl` with focus guard so user doesn't need to click a specific button first.

## 2026-03-11 | Sprint 4 | All Config in PluginSettings
**Decision:** No hardcoded values for keys, sizes, colours, feature flags.  
**Why:** Ahmed explicitly required this. Modular, user-overridable defaults. Avoids "magic number archaeology" when changing a threshold.

## 2026-03-11 | Sprint 5 | No React
**Decision:** Native DOM only — `createEl`, `setIcon`, `setTooltip`, `addAction`. No React, no virtual DOM.  
**Why:** React adds ~40KB to the bundle. Obsidian's API is already DOM-first. Native DOM code with Obsidian helpers achieves the "zero plugin vibe" goal without the overhead. React would also break Obsidian theming if components rendered their own style sheets.

## 2026-03-11 | Sprint 5 | addAction() for Infrequent Secondary Actions
**Decision:** Search, TOC, and Export PDF are view header actions (`addAction()`), not toolbar buttons.  
**Why:** Keeps the annotation toolbar focused on annotation tools. Secondary actions (used once per session vs. once per stroke) should not compete visually. `addAction()` places them in the standard Obsidian view header location — exactly where a native feature would put them.

## 2026-03-11 | Sprint 5 | Color Popover Built Fresh on Open
**Decision:** `showColorPopover()` creates DOM fresh; `hideColorPopover()` removes and nulls it.  
**Why:** No stale state across tool switches or file changes. Simpler than a hidden/shown panel. Popover is appended to `document.body` and positioned absolutely — avoids overflow clipping from ancestor containers.

## 2026-03-11 | Sprint 5 | CSS Custom Property `--note-color` on Note Element
**Decision:** `el.style.setProperty('--note-color', color)` — a CSS variable set on each note element.  
**Why:** Single-property theming without JS re-renders. All visual states (default border, hover glow, dot indicator) derive from this one variable via CSS. Changing colour = change one CSS variable, not four element styles.

## 2026-03-11 | .Agent/ Restructure | 3 Files > 7 Files
**Decision:** Consolidate governance from 7 files → `AGENT.md` + `STATUS.md` + `DECISIONS.md`. Add automated enforcement via ESLint + CSS checker.  
**Why:** 7 files = ~6,000 boot tokens with massive redundancy (commit format appeared in 4 files). Consolidated = ~2,500 tokens, single source of truth per concern, automated rules that can't be "forgotten" by an agent. Maintenance burden drops from 7-file sync to 2 files that need updating (STATUS.md + DECISIONS.md).

## 2026-03-11 | Sprint 7 | PPTX Model-First Parser Split
**Decision:** Split PPTX parsing into dedicated modules under `src/views/pptx/` and expand the parsed model to include shape IDs, bounds, z-index, and style tokens (fill/stroke) while keeping current HTML rendering path intact.  
**Why:** Current text-snapshot parsing blocks move/resize/format editing. A model-first parser enables incremental editor features without a risky full renderer rewrite and keeps `PptxView` focused on view lifecycle/UI concerns.

## 2026-03-11 | Sprint 7 | PPTX Draft Persistence via Sidecar
**Decision:** Persist PPTX edit-session transforms and style overrides in a companion sidecar file `<file>.pptx.edits.json`, loaded on open and saved via explicit toolbar action.  
**Why:** Non-destructive draft persistence protects original `.pptx` during early editor phases, enables recovery across reloads, and keeps apply-to-PPTX writing as a separate, explicit action.

## 2026-03-11 | Sprint 7 | Explicit Apply-to-PPTX Writeback
**Decision:** Add a dedicated “Apply draft edits to PPTX” action that rewrites slide OOXML (`p:spPr/a:xfrm`, fill, stroke tokens), saves the binary `.pptx`, then clears draft sidecar edits.  
**Why:** Users need both safe drafts and a deterministic commit step to modify source files. Explicit apply prevents accidental source mutation during iterative editing and avoids repeated double-application of transform deltas.
