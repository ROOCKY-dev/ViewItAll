# ViewItAll — Backlog

> Agile feature list for **beta 1.0.0**. Tackled one feature at a time, highest priority first.  
> See `UI-UX_constraints.md` for design standards every feature must meet.

---

## ✅ MVP Done (v0.1.x)

- [x] Plugin scaffold, esbuild config (punycode fix, pdf-worker inline)
- [x] DOCX viewer + editor (mammoth → contentEditable → html-to-docx)
- [x] PDF viewer (pdfjs-dist v3, blob-URL worker, flex layout)
- [x] PDF loading spinner & correct page sizing
- [x] PDF freehand annotations (pen, highlighter, eraser)
- [x] Annotation sidecar save (.annotations.json)
- [x] PDF zoom controls (step buttons, 9 snap levels, normalised annotation coords)
- [x] Settings tab
- [x] GitHub repo (ROOCKY-dev/ViewItAll-obsidian)
- [x] UI/UX constraints doc

---

## 🏃 Sprint 1 — Core UX Polish

| ID | Feature | Notes |
|----|---------|-------|
| `pdf-scroll-zoom` | **PDF: Ctrl+scroll zoom** | Smooth zoom centred on pointer; Ctrl+0 reset; Ctrl+=/- keyboard |
| `pdf-page-numbers` | **PDF: Page numbers + total** | "N / Total" indicator + label below each page canvas |
| `pdf-page-jump` | **PDF: Jump-to-page** | Click indicator → inline input → scrollIntoView *(needs page-numbers first)* |
| `pdf-keyboard-tools` | **PDF: Tool keyboard shortcuts** | V P H E keys; show hints in button titles |
| `pdf-color-picker` | **PDF: Per-stroke colour picker** | ✅ 6 swatches + custom slot; persist per tool type |

---

## 🏃 Sprint 2 — Tools Quality

| ID | Feature | Notes |
|----|---------|-------|
| `pdf-width-slider` | **PDF: Width & opacity sliders** | ✅ Inline slider; persist per tool |
| `pdf-text-search` | **PDF: Full-text search** | ✅ Ctrl+F; pdfjs getTextContent; highlight + prev/next |
| `docx-undo-redo` | **DOCX: Undo/redo** | ✅ execCommand undo/redo; toolbar buttons; dirty-state indicator |
| `toolbar-position` | **Both: Configurable toolbar position** | ✅ top/bottom; per file-type setting |

---

## 🏃 Sprint 3 — Advanced Features

| ID | Feature | Notes |
|----|---------|-------|
| `pdf-toc` | **PDF: TOC / outline sidebar** | pdfjs getOutline(); collapsible panel; click to scroll |
| `pdf-export` | **PDF: Export with embedded annotations** | pdf-lib vector draw; export as .annotated.pdf |
| `pdf-snap` | **PDF: Shape snap** | Near-rect/circle/line detection; snap grid overlay (§3.6 constraints) |
| `pdf-text-note` | **PDF: Text note anchored to page** | Click-to-place text box; normalised coords |

---

## 🔭 Post- release (Future)

| ID | Feature |
|----|---------|
| `xlsx-viewer` | Excel viewer (SheetJS, read-only, sheet tabs) |
| `pptx-viewer` | PowerPoint viewer (slide strip + main view) |

---

## Known Limitations

- **DOCX round-trip is lossy** — custom styles, macros, tracked changes not preserved on save.
- **PDF annotations are sidecar-only** — not embedded in the PDF file (use `pdf-export` feature for that).
- **Desktop only** (`isDesktopOnly: true`) — requires Electron renderer.
- **Annotation coords pre-0.1.4** — absolute px coords (old sidecar files will render incorrectly at new zoom levels).
