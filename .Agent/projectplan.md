# ViewItAll-md — Project Plan

> Living document. **Must be updated after every commit, patch, or push.**

---

## Overview

**Plugin:** ViewItAll-md  
**Purpose:** Native Obsidian viewer and editor for non-markdown file types — starting with `.pdf` and `.docx`, with a roadmap to spreadsheets, presentations, eBooks, and more.  
**Architecture:** TypeScript → esbuild → `main.js`. Entry: `src/main.ts`. Views extend Obsidian's `FileView`.  
**Current Version:** `1.4.1`  
**Repository:** Local vault at `.obsidian/plugins/ViewItAll-md`

---

## Sprint History

### Sprint 1 — Core Viewers (v1.0.0)
**Goal:** Minimal viable PDF and DOCX viewers inside Obsidian.

**Delivered:**
- PDF viewer using `pdf.js` with lazy-loading via `IntersectionObserver`
- 3-canvas layer stack per page: pdf (z1) / annotation (z2) / search (z3)
- Freehand pen annotation with colour presets and custom colour input
- Highlighter annotation with opacity setting
- Eraser tool
- Annotation persistence to `.annotations.json` sidecar files
- DOCX viewer using `mammoth.js` → contentEditable HTML
- DOCX save back to `.docx` using `html-to-docx`
- Confirm-on-save modal for DOCX
- Configurable toolbar position (top/bottom)
- Zoom with Ctrl+wheel and step buttons; zoom level persistence
- Versioning workflow: `npm version` + `version-bump.mjs`

---

### Sprint 2 — PDF Enhancements (v1.1.0)
**Goal:** Search, history, and usability improvements for PDF.

**Delivered:**
- Full-text search with match highlighting on a third canvas layer
- Previous/next match navigation
- Undo/redo for pen strokes (per-page stacks)
- Configurable default zoom level in settings
- Page jump input (click indicator → type page number → Enter)
- Configurable toolbar position per view type

**Bug fix (v1.1.1):** DOCX save was failing due to `HtmlToDocx` returning a Node.js `Buffer` not a `Blob` in Electron. Fixed with `buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)`.

---

### Sprint 3 — PDF Extra Features (v1.2.0)
**Goal:** Table of contents, text notes, and PDF export.

**Delivered:**
- Table of contents sidebar from PDF outline (`pdfDoc.getOutline()`)
- Collapsible tree with scroll-to-page on click
- Sticky text notes overlaid on PDF pages (draggable, colour-customisable)
- Text notes stored in annotations sidecar alongside strokes
- Export annotated PDF with strokes embedded via `pdf-lib`

---

### Sprint 3b — Constrained Stroke Snapping (v1.2.4)
**Goal:** Snap pen/highlighter/eraser strokes to H/V/45° directions.

**Delivered:**
- Snap activated by holding a configurable modifier key (default: Alt)
- Three snap directions: horizontal, vertical, 45° slope
- Cycle direction with modifier key + S (single button, configurable)
- Snap logic extracted to pure `snapPoint()` function in `src/utils/pdfSnap.ts`
- Snap direction button in toolbar showing current mode

---

### Sprint 4 — Modularisation & Settings Expansion (v1.3.0)
**Goal:** No hardcoded values; full settings coverage; clean module separation.

**Delivered:**
- `PluginSettings` interface expanded with all configurable values
- Full settings tab: DOCX section, PDF General, Annotation Tools, Snap, Keyboard Shortcuts
- All keyboard shortcuts and defaults read from settings at runtime
- `PdfSearchController` extracted to `src/views/pdf/PdfSearchController.ts`
- `snapPoint()` extracted to `src/utils/pdfSnap.ts`
- `exportAnnotatedPdf()` extracted to `src/utils/pdfExport.ts`
- `PageCtx`, `SearchMatch`, `PageRenderState` types moved to `src/views/pdf/pdfTypes.ts`

---

### Sprint 5 — Native Obsidian UI/UX Overhaul (v1.4.0)
**Goal:** Zero plugin vibe — everything looks and feels like a native Obsidian feature.

**Delivered:**
- All emoji buttons replaced with Lucide icons via `setIcon()` + `setTooltip()`
- Tool buttons in `.via-tool-group` pill with `.clickable-icon` + `is-active` state
- Color picker moved from inline toolbar to a floating `.via-color-popover` panel
  - Swatch row + custom colour input + size slider + opacity slider (highlighter)
  - Dismissed on outside click; destroyed on tool switch or file close
- Secondary actions (Search, TOC, Export) moved to view header via `addAction()`
- Snap button rebuilt with `setIcon()` + text label (H / V / 45°)
- Note drag handle → `grip-vertical`; delete → `x`; TOC close → `x`
- DocxView: `pencil`/`eye` toggle icon, `undo-2`, `redo-2`, `save`; yellow dirty dot
- ConfirmModal: `setTitle()` + `modal-button-container` + `mod-cta`
- `styles.css` complete rewrite — only Obsidian CSS variables, no hardcoded values

---

### Sprint 5 (patch) — Note Visual Redesign (v1.4.1)
**Goal:** Make PDF text notes look polished and theme-native.

**Delivered:**
- Note body is now a neutral card (`--background-primary`) for readable text in all themes
- Coloured accent: `border-top: 4px solid var(--note-color)` + matching glow ring on hover
- Header controls fade in on hover (opacity 0 → 1) — clean at rest, interactive on demand
- Color dot indicator in header centre for at-a-glance note colour identification
- Note is user-resizable (`resize: both`, 140px–400px)
- Delete button uses `--text-error` on hover for clear destructive affordance
- `--note-color` CSS variable replaces inline `background` on note element

---

## Current State (v1.4.1)

| File type | View | Edit | Annotate | Export | Search | Status |
|-----------|------|------|----------|--------|--------|--------|
| `.pdf`    | ✅   | —    | ✅ (pen/hl/erase/notes) | ✅ | ✅ | Shipped |
| `.docx`   | ✅   | ✅   | —        | —      | —      | Shipped |
| `.xlsx`   | —    | —    | —        | —      | —      | Planned (Sprint 6) |
| `.csv`    | —    | —    | —        | —      | —      | Planned (Sprint 6) |
| `.pptx`   | —    | —    | —        | —      | —      | Planned (Sprint 7) |
| `.epub`   | —    | —    | —        | —      | —      | Planned (Sprint 8) |
| `.mp4/mp3`| —    | —    | —        | —      | —      | Planned (Sprint 6) |

### Module Structure
```
src/
  main.ts                          # Plugin lifecycle, view registration
  settings.ts                      # PluginSettings + defaults + tab UI
  types.ts                         # VIEW_TYPE_*, shared types
  views/
    PdfView.ts                     # PDF viewer + annotation engine
    DocxView.ts                    # DOCX viewer + editor
    pdf/
      PdfSearchController.ts       # Search state + UI
      pdfTypes.ts                  # PageCtx, SearchMatch, PageRenderState
  utils/
    pdfSnap.ts                     # snapPoint() pure function
    pdfExport.ts                   # exportAnnotatedPdf() async
    docxUtils.ts                   # readDocxAsHtml, saveHtmlAsDocx
```

---

## Roadmap

### Sprint 6 — New File Types I (v1.5.0)
- `.xlsx` / `.csv` spreadsheet viewer (SheetJS)
- `.mp4` / `.mp3` media player (native HTML5)
- Settings toggle per file type (`enableXlsx`, `enableCsv`, `enableMedia`)

### Sprint 7 — New File Types II (v1.6.0)
- `.pptx` presentation viewer (Phase 1: text extraction; Phase 2: visual slides)
- Settings toggle: `enablePptx`

### Sprint 8 — New File Types III (v1.7.0)
- `.epub` eBook reader (epub.js)
- Font size preference, reading position memory
- Settings toggle: `enableEpub`

### Sprint 9 — Polish & Performance (v1.8.0)
- `.zip` archive inspector (JSZip)
- Performance audit: bundle size, memory on large PDFs, annotation render time
- Keyboard navigation accessibility audit

### Future (v2.x)
- 3D model viewer `.glb`/`.gltf` (Three.js, lazy-loaded)
- `.odt` / `.rtf` support
- PDF annotation collaboration (shared sidecar via vault sync)

---

## Key Technical Decisions Log

| Decision | Rationale | Sprint |
|----------|-----------|--------|
| Sidecar `.annotations.json` for PDF annotations | Non-destructive; original PDF unchanged | 1 |
| `vault.modifyBinary` not `adapter.writeBinary` | Fires proper Obsidian file events | 1 |
| `document`-level key handler for Alt snap key | `containerEl` handlers need focus; Alt is a modifier | 3b |
| No React | Native DOM + Obsidian APIs = native feel, no overhead | 5 |
| `addAction()` for search/TOC/export | Keeps annotation toolbar focused; these are infrequent | 5 |
| CSS variable `--note-color` on note element | Enables single-property colour theming without JS rerenders | 5p |
| Dynamic `import()` required for heavy future libs | Keeps initial bundle size under 2MB | — |
