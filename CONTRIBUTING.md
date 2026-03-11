# Contributing to ViewItAll-md

First off, thank you for considering contributing to ViewItAll-md! It's people like you that make this plugin a great tool for the Obsidian community.

Before you start writing code, please read through these critical guidelines. They are enforced rigorously to maintain performance, stability, and the native "zero plugin vibe" of this plugin.

## 🔴 Critical Development Rules

1. **TypeScript strict. Zero `any`.** Use `unknown` + type guards or define a local interface.
2. **CSS: Obsidian variables only.** No hardcoded hex, named colours, or px font-sizes. Only `rgba(0,0,0,N)` for shadows. All custom classes must use the `via-` prefix. OCD stuff
3. **Icons: `setIcon()` + Lucide only.** No emoji as labels. Use `setTooltip()` for hover text, not `.title =`. i adore consistency
4. **No React.** Native DOM only. Use Obsidian's `createEl`, `setIcon`, `setTooltip`, and `addAction`.
5. **No hardcoded config.** Every key, colour, width, flag must read from `PluginSettings`. i mean yah 
6. **Build must pass before commit.** Run `npm run build` (tsc + esbuild). It must pass with zero errors. test test test 
7. **No `console.log()` in commits.** Use `new Notice()` for user messages.
8. **Heavy libs: dynamic `import()` only.** Never use top-level imports for heavy libraries like `pptx`, `epub`, or `three.js`.

## 💻 Development Setup

1. **Clone the repository** into your local Obsidian vault's plugins folder:
   `<VaultFolder>/.obsidian/plugins/ViewItAll-md`
2. **Install dependencies**:
   `npm install`
3. **Run the development watch process**:
   `npm run dev`
4. **Lint your code**:
   `npm run lint`       *(ESLint all src/ files)*
   `npm run lint:css`   *(Check styles.css for hardcoded values)*

## 🏗️ Folder Structure

- `src/main.ts`: Lifecycle only — load, unload, register.
- `src/settings.ts`: PluginSettings interface, defaults, settings tab.
- `src/types.ts`: `VIEW_TYPE_*` constants and shared types.
- `src/views/`: FileView subclasses (keep strictly to one file per type, max 600 lines).
- `src/utils/`: Pure functions — no Obsidian state.
- `styles.css`: All CSS (single file).

## 📝 Committing Code

When you are ready to commit, please adhere to our strict commit format.

**Format:**
```
type(scope): short description (vX.Y.Z)

- bullet detail

"just small changes" 
```
Types: `feat` | `fix` | `refactor` | `style` | `docs` | `chore` | `perf`
