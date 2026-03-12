# ViewItAll

> Open **PDF**, **Word (.docx)**, **Excel (.xlsx)**, **CSV**, and **PowerPoint (.pptx)** files directly inside [Obsidian](https://obsidian.md) — no external apps, no context switching.

![Version](https://img.shields.io/badge/version-1.5.0-blue)
![Obsidian](https://img.shields.io/badge/Obsidian-0.15%2B-purple)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Supported Formats

| Format | Extensions | Capabilities |
|--------|-----------|--------------|
| PDF | `.pdf` | View + Annotate + Export |
| Word | `.docx` | View + Edit + Save |
| Excel | `.xlsx` | View + Edit + Save |
| CSV | `.csv` | View + Edit + Save |
| PowerPoint | `.pptx` | View |

Each format can be individually enabled or disabled in **Settings → ViewItAll**.

---

## Features

### PDF Viewer & Annotator
- Lazy page rendering — only pages near the viewport are loaded, keeping memory low on large documents
- **Zoom** — steps from 25% to 400%, Ctrl+scroll, Ctrl+`=`/`-`/`0`
- **Annotation tools** — Pen, Highlighter, Eraser, and sticky Notes, each with configurable colours, widths, and opacity
- **Snap-to-line** — hold a modifier key to constrain strokes to horizontal, vertical, or 45-degree angles
- **Full-text search** — Ctrl+F search across all pages with match highlighting and navigation
- **Table of contents** — extracted from PDF bookmarks/outline, toggle sidebar from the toolbar
- **Export annotated PDF** — bakes pen and highlighter strokes into a new `.annotated.pdf` file via pdf-lib
- Annotations saved as a companion `.pdf.annotations.json` sidecar (normalised coordinates)
- Configurable keyboard shortcuts for every tool and action

### Word (.docx) Viewer / Editor
- Renders documents as clean HTML via [mammoth](https://github.com/mwilliamson/mammoth.js)
- Toggle **Edit mode** with Undo / Redo support
- Save back to `.docx` via [html-to-docx](https://github.com/privateOmega/html-to-docx)
- Optional confirmation dialog warns about lossy round-trip (custom styles, tracked changes)
- Configurable: open in edit mode by default, toolbar position (top/bottom), open target (tab/sidebar)

### Spreadsheet (.xlsx / .csv) Viewer / Editor
- Parses workbooks with [SheetJS](https://sheetjs.com/) (lazy-loaded on first use)
- **Multi-sheet tabs** for `.xlsx` files
- **Inline cell editing** — double-click or type in the formula bar; Enter, Tab, and Escape navigation
- **Formula engine** — evaluates `SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF`, `AND`, `OR`, `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `ROUND`, `TODAY`, `NOW`, and more with cell/range references
- **Add rows & columns**, or right-click headers for Insert / Delete context menus
- Save back as `.xlsx` or `.csv`; revert to last save with Undo
- Dirty-state indicator and optional confirm-before-save dialog
- Ctrl/Cmd+S to save

### PowerPoint (.pptx) Viewer
- Canvas-based slide rendering via pptxviewjs
- Previous / Next slide navigation with slide counter
- Zoom in / out (25%–400%) and **Fit to container**
- Reads real theme colours from the PPTX file for accurate rendering
- Detects actual slide dimensions from metadata

---

## Installation

### Manual (development)
1. Clone this repo into your vault's plugin folder:
   ```
   .obsidian/plugins/ViewItAll-md/
   ```
2. Install dependencies:
   ```bash
   npm install
   ```
3. Build:
   ```bash
   npm run build
   ```
4. Enable the plugin in **Settings → Community plugins**.

---

## Development

```bash
npm install       # install dependencies
npm run dev       # watch mode — rebuilds on save
npm run build     # production build
```

Requires **Node 18+** and **npm**.

---

## Architecture

```
src/
  main.ts                    # Plugin lifecycle, view registration
  types.ts                   # Shared interfaces & view-type constants
  settings.ts                # Settings UI + defaults
  utils/
    fileUtils.ts             # Vault path helpers
    docxUtils.ts             # mammoth read + html-to-docx write
    annotations.ts           # PDF annotation sidecar load/save
    formulaEval.ts           # Spreadsheet formula engine
    pdfExport.ts             # Annotated PDF export (pdf-lib)
    pdfSnap.ts               # Snap-to-line geometry
  views/
    DocxView.ts              # DOCX FileView
    PdfView.ts               # PDF FileView (pdfjs-dist)
    PptxView.ts              # PPTX FileView (pptxviewjs)
    SpreadsheetView.ts       # XLSX / CSV FileView (SheetJS)
    pdf/
      PdfSearchController.ts # Full-text PDF search
      pdfTypes.ts            # PDF-specific types
styles.css                   # All scoped via-* CSS classes
```

---

## Known Limitations

- **DOCX round-trip is lossy** — custom paragraph styles, macros, tracked changes, and complex tables may not survive save.
- **PDF annotation coordinates are normalised** (0–1); annotations from versions < 0.1.4 (which used absolute px) will render incorrectly.
- **Spreadsheet formulas** cover common functions but not the full Excel formula language.
- **PPTX viewer is read-only** — slide editing is not supported.
- Obsidian's built-in PDF viewer is replaced while the plugin is active; it is restored on unload.

---

## License

MIT © [ROOCKY.dev](https://github.com/ROOCKY-dev)
