# ViewItAll-md — Current Status
> Updated after every commit. Single source of truth for project state.

---

## Version: 1.4.1

## File Type Support

| Extension | View | Edit | Annotate | Export | Search | Status |
|-----------|------|------|----------|--------|--------|--------|
| `.pdf`    | ✅   | —    | ✅ pen/hl/erase/notes | ✅ | ✅ | Shipped |
| `.docx`   | ✅   | ✅   | —        | —      | —      | Shipped |
| `.xlsx`   | —    | —    | —        | —      | —      | Sprint 6 |
| `.csv`    | —    | —    | —        | —      | —      | Sprint 6 |
| `.mp4/.mp3` | —  | —    | —        | —      | —      | Sprint 6 |
| `.pptx`   | —    | —    | —        | —      | —      | Sprint 7 |
| `.epub`   | —    | —    | —        | —      | —      | Sprint 8 |
| `.zip`    | —    | —    | —        | —      | —      | Sprint 9 |

## Module Structure

```
src/
  main.ts
  settings.ts
  types.ts
  views/
    PdfView.ts                 # ~600 lines — PDF viewer + annotation engine
    DocxView.ts                # ~200 lines — DOCX viewer + editor
    pdf/
      PdfSearchController.ts   # Search state + bar UI
      pdfTypes.ts              # PageCtx, SearchMatch, PageRenderState
  utils/
    pdfSnap.ts                 # snapPoint() pure function
    pdfExport.ts               # exportAnnotatedPdf() async
    docxUtils.ts               # readDocxAsHtml, saveHtmlAsDocx
scripts/
  check-css.js                 # Automated CSS variable enforcement
```

## Active: Sprint 7 (v1.6.0)

In progress. Targets:
- `.pptx` model-first parser + shape metadata extraction
- Phase 1 editor groundwork (shape ids, bounds, style tokens)
- Phase 2 next: move/resize + text editing + fill/stroke formatting

## Roadmap

| Sprint | Version | Targets |
|--------|---------|---------|
| 6 | 1.5.0 | xlsx/csv viewer, media player, file-type toggles |
| 7 | 1.6.0 | pptx viewer (Phase 1: text; Phase 2: visual) |
| 8 | 1.7.0 | epub reader (epub.js) |
| 9 | 1.8.0 | zip inspector, performance audit |
| Future | 2.x | 3D models, odt/rtf, annotation sharing |

## Recent Changes

| Version | Date | Summary |
|---------|------|---------|
| 1.4.1 | 2026-03-11 | PDF note visual redesign — card + coloured accent, resizable, hover controls |
| 1.4.0 | 2026-03-11 | Sprint 5 UI/UX overhaul — Lucide icons, CSS vars, color popover, view header actions |
| 1.3.0 | 2026-03-11 | Sprint 4 modularisation — full settings, extracted controllers, no hardcoded values |
| 1.2.4 | 2026-03-11 | Sprint 3b — constrained stroke snapping H/V/45° with Alt modifier |
| 1.2.0 | 2026-03-11 | Sprint 3 — TOC, text notes, PDF export |
| 1.1.1 | 2026-03-11 | Fix DOCX save (Buffer→ArrayBuffer for Electron) |
| 1.1.0 | 2026-03-11 | Sprint 2 — search, undo/redo, page jump, toolbar position |
| 1.0.0 | 2026-03-11 | Sprint 1 — PDF + DOCX viewers, pen/highlight/erase annotations |

## Requirements Status (summary)

Full registry → `.Agent/requirements.yaml`

- CORE: 5 done, 1 planned (CORE-06 file-type toggles)
- PDF: 18 done
- DOCX: 5 done
- UX: 4 done
- Future file types: FT-01 through FT-08 planned
