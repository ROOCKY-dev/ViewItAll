# ViewItAll

> View and edit **Word documents (.docx)** natively inside [Obsidian](https://obsidian.md) — no external apps, no context switching.

![Version](https://img.shields.io/badge/version-2.0.1-blue)
![Obsidian](https://img.shields.io/badge/Obsidian-0.15%2B-purple)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

### Word (.docx) viewer and editor
- **Native OOXML rendering** — parses .docx files directly into a typed document model and renders to DOM. No third-party conversion libraries.
- **Edit mode** — toggle between view and edit mode. Type directly into paragraphs with full contentEditable support.
- **Formatting toolbar** — bold, italic, underline, strikethrough, font family, font size, text color, highlight color, paragraph alignment, and clear formatting.
- **Keyboard shortcuts** — Ctrl+B, Ctrl+I, Ctrl+U for formatting. Ctrl+Z / Ctrl+Shift+Z for undo/redo. Ctrl+S to save.
- **Image support** — embedded images render inline with correct sizing. Insert new images via toolbar. Drag resize handles to scale images.
- **Table support** — renders tables with cell shading, column spans, and vertical alignment. Edit cell content with Tab/Shift+Tab navigation.
- **Round-trip save** — saves back to .docx by regenerating `word/document.xml` while preserving all other ZIP entries (styles, numbering, relationships, media, theme) unchanged.
- **Undo/Redo** — snapshot-based history with up to 100 levels.
- **Auto-save** — optionally save unsaved changes when closing a file.
- Configurable toolbar position (top/bottom), default zoom, and edit mode defaults.

---

## Installation

### From community plugins
1. Open **Settings → Community plugins**.
2. Search for **View It All**.
3. Click **Install**, then **Enable**.

### Manual
1. Clone this repo into your vault's plugin folder:
   ```
   .obsidian/plugins/ViewItAll-md/
   ```
2. Install dependencies and build:
   ```bash
   npm install
   npm run build
   ```
3. Enable the plugin in **Settings → Community plugins**.

---

## Development

```bash
npm install       # install dependencies
npm run dev       # watch mode — rebuilds on save
npm run build     # production build
npm run lint      # run ESLint
```

Requires **Node 18+** and **npm**.

---

## Architecture

```
src/
  main.ts              # Plugin lifecycle, view registration
  types.ts             # Shared interfaces and view-type constants
  settings.ts          # Settings UI and defaults
  docx/
    model.ts           # Pure TypeScript document model interfaces
    parser.ts          # OOXML ZIP → document model (JSZip + DOMParser)
    renderer.ts        # Document model → DOM (createEl, no innerHTML)
    serializer.ts      # Document model → OOXML XML → ZIP round-trip
    editing.ts         # ContentEditable ↔ model bridge, formatting, undo
    selection.ts       # Browser Selection ↔ model coordinate mapping
    toolbar.ts         # Formatting toolbar (bold, italic, color, etc.)
    history.ts         # Snapshot-based undo/redo stack
    styles.ts          # OOXML style resolution and inheritance
    numbering.ts       # List numbering definition parser
    relationships.ts   # Relationship type constants
  utils/
    xml.ts             # XML namespace constants, parseXml helper
    units.ts           # px↔EMU, pt↔half-points, twips conversions
  views/
    DocxView.ts        # DOCX FileView (view/edit/save lifecycle)
styles.css             # All scoped via-* CSS classes, Obsidian variables
```

---

## Known limitations

- **Complex OOXML features** — headers/footers, shapes, SmartArt, embedded OLE objects, and tracked changes are not rendered.
- **Style fidelity** — the plugin resolves paragraph and run styles but does not implement the full OOXML style inheritance chain (e.g., document defaults, theme fonts).
- **Desktop only** — requires Electron APIs available in the desktop version of Obsidian.

---

## License

MIT © [ROOCKY.dev](https://github.com/ROOCKY-dev)
