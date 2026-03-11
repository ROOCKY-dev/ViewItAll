import { FileView, TFile, WorkspaceLeaf, Notice } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import type { RefProxy as PdfRefProxy } from 'pdfjs-dist/types/src/display/api';
import { VIEW_TYPE_PDF } from '../types';
import type { PageAnnotations, AnnotationPath, AnnotationFile, TextNote } from '../types';
import {
loadAnnotations,
saveAnnotations,
getPageAnnotations,
setPageAnnotations,
} from '../utils/pdfAnnotations';
import { snapPoint }          from '../utils/pdfSnap';
import { exportAnnotatedPdf } from '../utils/pdfExport';
import { PdfSearchController } from './pdf/PdfSearchController';
import type { PageCtx, PageRenderState } from './pdf/pdfTypes';
import type ViewItAllPlugin from '../main';
import type { SnapDirection } from '../settings';

// eslint-disable-next-line @typescript-eslint/no-var-requires
const _pdfWorkerSrc: string = require('pdfjs-worker-src');
let _workerBlobUrl: string | null = null;
function getPdfWorkerUrl(): string {
if (!_workerBlobUrl) {
const blob = new Blob([_pdfWorkerSrc], { type: 'application/javascript' });
_workerBlobUrl = URL.createObjectURL(blob);
}
return _workerBlobUrl;
}

type AnnotTool = 'none' | 'pen' | 'highlighter' | 'eraser' | 'note';
// PageRenderState imported from pdfTypes

export class PdfView extends FileView {
private plugin: ViewItAllPlugin;
private pdfDoc: pdfjsLib.PDFDocumentProxy | null = null;
private annotData: AnnotationFile = { version: 1, pages: [] };
private pages: PageCtx[] = [];
private currentTool: AnnotTool = 'none';
private isDrawing = false;
private currentPath: AnnotationPath | null = null;
private currentFile: TFile | null = null;

// Zoom
private currentScale = 1.0;
private readonly ZOOM_STEPS = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0, 3.0, 4.0];
private scrollAreaEl: HTMLElement | null = null;
private zoomLabelEl: HTMLElement | null = null;
private _zoomDebounceTimer: ReturnType<typeof setTimeout> | null = null;

// Lazy rendering
private pageIndicatorEl: HTMLElement | null = null;
private pageObserver: IntersectionObserver | null = null;
private renderObserver: IntersectionObserver | null = null;
private _renderGen = 0;

// Color picker
private colorSectionEl: HTMLElement | null = null;
private colorSepEl: HTMLElement | null = null;
private colorSwatchEls: HTMLButtonElement[] = [];
private colorCustomInputEl: HTMLInputElement | null = null;

private readonly PEN_PRESETS       = ['#e03131', '#1971c2', '#2f9e44', '#212529', '#e8590c', '#7048e8'];
private readonly HIGHLIGHT_PRESETS = ['#ffd43b', '#22b8cf', '#f783ac', '#69db7c', '#ffa94d', '#da77f2'];

// Width / opacity sliders
private widthSectionEl: HTMLElement | null = null;
private widthSepEl: HTMLElement | null = null;
private widthLabelEl: HTMLElement | null = null;
private widthSliderEl: HTMLInputElement | null = null;
private opacityRowEl: HTMLElement | null = null;
private opacityLabelEl: HTMLElement | null = null;
private opacitySliderEl: HTMLInputElement | null = null;

// Snapping
private snapDirection: SnapDirection = 'horizontal';
private snapDirBtnEl: HTMLButtonElement | null = null;

// Layout containers
private wrapperEl: HTMLElement | null = null;
private bodyEl: HTMLElement | null = null;
private tocSidebarEl: HTMLElement | null = null;
private tocVisible = false;

// Text notes overlay (keyed by note id)
private noteEls = new Map<string, HTMLElement>();

// Search controller
private search = new PdfSearchController();

constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
super(leaf);
this.plugin = plugin;
pdfjsLib.GlobalWorkerOptions.workerSrc = getPdfWorkerUrl();
}

onload(): void {
super.onload();

const isActiveLeaf = () => this.app.workspace.getActiveViewOfType(PdfView) === this;

// ── Document-level snap key handler ──────────────────────────────────
// Registered at document level so it fires regardless of focus, guarded
// by the active-leaf check.
this.registerDomEvent(document as unknown as HTMLElement, 'keydown', (e: KeyboardEvent) => {
if (!isActiveLeaf()) return;
const s = this.plugin.settings;
const snapKey     = s.snapActivateKey; // 'Alt' | 'Shift'
const snapPressed = snapKey === 'Alt' ? e.altKey : e.shiftKey;

if (e.key === snapKey) {
e.preventDefault();
const drawing = this.currentTool === 'pen'
             || this.currentTool === 'highlighter'
             || this.currentTool === 'eraser';
if (drawing) this.snapDirBtnEl?.classList.add('via-btn-snap-active');
} else if (snapPressed && e.key.toLowerCase() === s.keySnapCycle) {
// Snap modifier + cycle key → cycle snap direction
e.preventDefault();
const dirs: SnapDirection[] = ['horizontal', 'vertical', 'slope'];
this.snapDirection = dirs[(dirs.indexOf(this.snapDirection) + 1) % dirs.length]!;
this.updateSnapDirBtn();
}
});

this.registerDomEvent(document as unknown as HTMLElement, 'keyup', (e: KeyboardEvent) => {
if (!isActiveLeaf()) return;
if (e.key === this.plugin.settings.snapActivateKey) {
this.snapDirBtnEl?.classList.remove('via-btn-snap-active');
}
});

// ── Container-level shortcuts (zoom, tools, search) ──────────────────
this.registerDomEvent(this.containerEl as HTMLElement, 'keydown', (e: KeyboardEvent) => {
const s = this.plugin.settings;

// Ctrl/Cmd shortcuts
if (e.ctrlKey || e.metaKey) {
if (e.key === '0') {
e.preventDefault();
this.setZoom(s.pdfDefaultZoom, this.viewportCenterFrac());
} else if (e.key === '=' || e.key === '+') {
e.preventDefault(); this.stepZoom(+1);
} else if (e.key === '-') {
e.preventDefault(); this.stepZoom(-1);
} else if (e.key.toLowerCase() === s.keySearch) {
e.preventDefault(); this.search.open();
}
return;
}

// Skip snap modifier combos — handled at document level
if (e.key === s.snapActivateKey
 || (s.snapActivateKey === 'Alt'   && e.altKey)
 || (s.snapActivateKey === 'Shift' && e.shiftKey)) return;

// Tool shortcuts — skip if an input element has focus
const target = e.target as HTMLElement;
if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable) return;

const toolMap: Record<string, AnnotTool> = {
[s.keyToolView]:      'none',
[s.keyToolPen]:       'pen',
[s.keyToolHighlight]: 'highlighter',
[s.keyToolErase]:     'eraser',
[s.keyToolNote]:      'note',
};
const tool = toolMap[e.key.toLowerCase()];
if (tool !== undefined) { e.preventDefault(); this.setTool(tool); }
});
}

getViewType(): string { return VIEW_TYPE_PDF; }
getDisplayText(): string { return this.file?.basename ?? 'PDF'; }
getIcon(): string { return 'file'; }
canAcceptExtension(extension: string): boolean { return extension === 'pdf'; }

async onLoadFile(file: TFile): Promise<void> {
this.currentFile  = file;
this.currentTool  = this.plugin.settings.pdfDefaultTool;
this.currentScale = this.plugin.settings.pdfDefaultZoom;
this.snapDirection = this.plugin.settings.snapDefaultDirection;
this.annotData    = await loadAnnotations(this.app, file);
await this.renderPdf(file);
}

async onUnloadFile(_file: TFile): Promise<void> {
this._renderGen++;
if (this._zoomDebounceTimer !== null) {
clearTimeout(this._zoomDebounceTimer); this._zoomDebounceTimer = null;
}
this.renderObserver?.disconnect(); this.renderObserver = null;
this.pageObserver?.disconnect();   this.pageObserver   = null;
if (this.pdfDoc) { this.pdfDoc.destroy(); this.pdfDoc = null; }
this.pages = [];
this.search.destroy();
this.noteEls.clear();
this.tocSidebarEl = null;
this.tocVisible   = false;
this.bodyEl       = null;
this.snapDirBtnEl = null;
this.contentEl.empty();
}

private async renderPdf(file: TFile): Promise<void> {
this._renderGen++;
this.contentEl.empty();
this.pages = [];
this.search.destroy();
this.noteEls.clear();
this.tocSidebarEl = null;
this.tocVisible   = false;
this.renderObserver?.disconnect(); this.renderObserver = null;
this.pageObserver?.disconnect();   this.pageObserver   = null;

const isBottom = this.plugin.settings.pdfToolbarPosition === 'bottom';
const wrapper = this.contentEl.createEl('div', { cls: 'via-pdf-wrapper' });
if (isBottom) wrapper.classList.add('via-pdf-wrapper--toolbar-bottom');
this.wrapperEl = wrapper;

const toolbar = this.buildToolbar();
wrapper.appendChild(toolbar);

const bodyEl = wrapper.createEl('div', { cls: 'via-pdf-body' });
this.bodyEl = bodyEl;

const scrollArea = bodyEl.createEl('div', { cls: 'via-pdf-scroll' });
this.scrollAreaEl = scrollArea;
scrollArea.addEventListener('wheel', (e: WheelEvent) => this.handleWheelZoom(e), { passive: false });

const loadingEl = scrollArea.createEl('div', { cls: 'via-pdf-loading' });
loadingEl.createEl('div', { cls: 'via-pdf-loading-spinner' });
loadingEl.createEl('span', { text: 'Loading PDF…' });

let buffer: ArrayBuffer;
try {
buffer = await this.app.vault.adapter.readBinary(file.path);
} catch (err) {
loadingEl.remove();
scrollArea.createEl('p', { cls: 'via-error', text: `Cannot read file: ${String(err)}` });
return;
}

this.pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(buffer) }).promise;
const numPages = this.pdfDoc.numPages;

const sizes = await Promise.all(
Array.from({ length: numPages }, async (_, i) => {
const page = await this.pdfDoc!.getPage(i + 1);
const vp   = page.getViewport({ scale: this.currentScale });
return { w: Math.ceil(vp.width), h: Math.ceil(vp.height) };
})
);

loadingEl.remove();

for (let i = 0; i < numPages; i++) {
const { w, h } = sizes[i]!;
const container = scrollArea.createEl('div', { cls: 'via-pdf-page' });
container.style.cssText = `width:${w}px;height:${h}px;min-width:${w}px;min-height:${h}px`;
scrollArea.createEl('div', { cls: 'via-pdf-page-label', text: `${i + 1} / ${numPages}` });
this.pages.push({
pageNum: i + 1, state: 'placeholder' as PageRenderState,
container, pdfCanvas: null, annotCanvas: null, searchCanvas: null, w, h,
});
}

if (this.pageIndicatorEl) this.pageIndicatorEl.textContent = `1 / ${numPages}`;

// Wire search controller to current document
this.search.setContext(this.pdfDoc, this.pages, wrapper, bodyEl);

this.attachRenderObserver();
this.attachPageObserver();

// Render existing text notes
for (const note of (this.annotData.notes ?? [])) this.renderNoteEl(note);

// Load TOC; auto-open if configured
this.loadToc().then(() => {
if (this.plugin.settings.showTocOnOpen && (this as any)._outline?.length > 0) {
this.toggleToc();
}
}).catch(console.error);
}

// Lazy rendering ---------------------------------------------------------

private attachRenderObserver(): void {
this.renderObserver?.disconnect();
if (!this.scrollAreaEl || this.pages.length === 0) return;

const pageMap = new Map<Element, PageCtx>(this.pages.map(p => [p.container, p]));

this.renderObserver = new IntersectionObserver(
(entries) => {
for (const entry of entries) {
const ctx = pageMap.get(entry.target);
if (!ctx) continue;
if (entry.isIntersecting) this.renderPageCanvas(ctx).catch(console.error);
else                      this.unloadPageCanvas(ctx);
}
},
{ root: this.scrollAreaEl, rootMargin: '200% 0px' }
);

for (const ctx of this.pages) this.renderObserver.observe(ctx.container);
}

private async renderPageCanvas(ctx: PageCtx): Promise<void> {
if (ctx.state !== 'placeholder') return;
ctx.state = 'rendering';
const gen = this._renderGen;

const page     = await this.pdfDoc!.getPage(ctx.pageNum);
const viewport = page.getViewport({ scale: this.currentScale });

if (gen !== this._renderGen) { ctx.state = 'placeholder'; return; }

const pdfCanvas = ctx.container.createEl('canvas', { cls: 'via-pdf-canvas' });
pdfCanvas.width = ctx.w; pdfCanvas.height = ctx.h;

const annotCanvas = ctx.container.createEl('canvas', { cls: 'via-pdf-annot-canvas' });
annotCanvas.width = ctx.w; annotCanvas.height = ctx.h;

const searchCanvas = ctx.container.createEl('canvas', { cls: 'via-pdf-search-canvas' });
searchCanvas.width = ctx.w; searchCanvas.height = ctx.h;

await page.render({ canvasContext: pdfCanvas.getContext('2d')!, viewport }).promise;

if (gen !== this._renderGen) {
pdfCanvas.remove(); annotCanvas.remove(); searchCanvas.remove();
ctx.state = 'placeholder';
return;
}

ctx.pdfCanvas    = pdfCanvas;
ctx.annotCanvas  = annotCanvas;
ctx.searchCanvas = searchCanvas;
ctx.state        = 'rendered';

this.redrawAnnotations(ctx);
if (this.search.hasMatches) this.search.drawHighlightsForPage(ctx);
this.attachDrawListeners(ctx);
this.updateCanvasInteraction();
}

private unloadPageCanvas(ctx: PageCtx): void {
if (ctx.state !== 'rendered') return;
ctx.pdfCanvas?.remove();    ctx.pdfCanvas    = null;
ctx.annotCanvas?.remove();  ctx.annotCanvas  = null;
ctx.searchCanvas?.remove(); ctx.searchCanvas = null;
ctx.state = 'placeholder';
}

// Toolbar ----------------------------------------------------------------

private buildToolbar(): HTMLElement {
const bar = document.createElement('div');
bar.className = 'via-pdf-toolbar';
const s = this.plugin.settings;

// TOC toggle
const tocBtn = bar.createEl('button', { cls: 'via-btn', text: '📑 TOC' });
tocBtn.title = 'Toggle table of contents';
tocBtn.addEventListener('click', () => this.toggleToc());

bar.createEl('div', { cls: 'via-toolbar-sep' });

// Tool buttons — titles show the configured key
const tools: { id: AnnotTool; label: string; key: string }[] = [
{ id: 'none',        label: '👁 View',      key: s.keyToolView.toUpperCase() },
{ id: 'pen',         label: '✏️ Pen',        key: s.keyToolPen.toUpperCase() },
{ id: 'highlighter', label: '🖊 Highlight',  key: s.keyToolHighlight.toUpperCase() },
{ id: 'eraser',      label: '⬜ Erase',      key: s.keyToolErase.toUpperCase() },
{ id: 'note',        label: '📝 Note',       key: s.keyToolNote.toUpperCase() },
];

for (const t of tools) {
const btn = bar.createEl('button', { cls: 'via-btn', text: t.label });
btn.dataset.tool = t.id;
btn.title = `${t.label} (${t.key})`;
if (t.id === this.currentTool) btn.classList.add('via-btn-active');
btn.addEventListener('click', () => this.setTool(t.id));
}

bar.createEl('div', { cls: 'via-toolbar-sep' });

// Color picker section
this.colorSwatchEls = [];
this.colorSectionEl = bar.createEl('div', { cls: 'via-pdf-color-section' });
const initPresets = this.currentTool === 'highlighter' ? this.HIGHLIGHT_PRESETS : this.PEN_PRESETS;
for (const color of initPresets) {
const swatch = this.colorSectionEl.createEl('button', { cls: 'via-color-swatch' });
swatch.style.background = color;
swatch.dataset.color    = color;
swatch.title            = color;
swatch.addEventListener('click', () => this.applyColor(swatch.dataset.color!));
this.colorSwatchEls.push(swatch);
}
const customLabel = this.colorSectionEl.createEl('label', { cls: 'via-color-swatch via-color-custom', title: 'Custom colour' });
const customInput = customLabel.createEl('input');
customInput.type      = 'color';
customInput.className = 'via-color-custom-input';
this.colorCustomInputEl = customInput;

customInput.addEventListener('input', () => {
const tool = this.currentTool;
if (tool !== 'pen' && tool !== 'highlighter') return;
const presets = tool === 'pen' ? this.PEN_PRESETS : this.HIGHLIGHT_PRESETS;
const color   = customInput.value.toLowerCase();
for (const sw of this.colorSwatchEls)
sw.classList.toggle('via-color-swatch-active', sw.dataset.color?.toLowerCase() === color);
customInput.parentElement?.classList.toggle('via-color-swatch-active', !presets.some(c => c.toLowerCase() === color));
});
customInput.addEventListener('change', () => this.applyColor(customInput.value));

this.colorSepEl = bar.createEl('div', { cls: 'via-toolbar-sep' });

const showColors = this.currentTool === 'pen' || this.currentTool === 'highlighter';
this.colorSectionEl.style.display = showColors ? 'flex' : 'none';
this.colorSepEl.style.display     = showColors ? '' : 'none';
if (showColors) this.syncColorPicker(this.currentTool as 'pen' | 'highlighter');

// Width / opacity sliders
this.widthSectionEl = bar.createEl('div', { cls: 'via-pdf-width-section' });
const widthRow = this.widthSectionEl.createEl('div', { cls: 'via-pdf-width-row' });
widthRow.createEl('span', { cls: 'via-pdf-slider-label', text: 'Size' });
this.widthSliderEl = widthRow.createEl('input');
this.widthSliderEl.type      = 'range';
this.widthSliderEl.className = 'via-pdf-slider';
this.widthLabelEl = widthRow.createEl('span', { cls: 'via-pdf-slider-value' });

this.widthSliderEl.addEventListener('input',  () => {
if (this.widthLabelEl) this.widthLabelEl.textContent = `${this.widthSliderEl!.value}px`;
});
this.widthSliderEl.addEventListener('change', () => this.applyWidth(Number(this.widthSliderEl!.value)));

this.opacityRowEl = this.widthSectionEl.createEl('div', { cls: 'via-pdf-width-row' });
this.opacityRowEl.createEl('span', { cls: 'via-pdf-slider-label', text: 'Opacity' });
this.opacitySliderEl = this.opacityRowEl.createEl('input');
this.opacitySliderEl.type      = 'range';
this.opacitySliderEl.min       = '0.1';
this.opacitySliderEl.max       = '1.0';
this.opacitySliderEl.step      = '0.05';
this.opacitySliderEl.className = 'via-pdf-slider';
this.opacityLabelEl = this.opacityRowEl.createEl('span', { cls: 'via-pdf-slider-value' });

this.opacitySliderEl.addEventListener('input',  () => {
const v = Number(this.opacitySliderEl!.value);
if (this.opacityLabelEl) this.opacityLabelEl.textContent = `${Math.round(v * 100)}%`;
});
this.opacitySliderEl.addEventListener('change', () => this.applyOpacity(Number(this.opacitySliderEl!.value)));

this.widthSepEl = bar.createEl('div', { cls: 'via-toolbar-sep' });
this.widthSectionEl.style.display = showColors ? 'flex' : 'none';
this.widthSepEl.style.display     = showColors ? '' : 'none';
if (showColors) this.syncWidthSlider(this.currentTool as 'pen' | 'highlighter');

// Snap direction button
this.snapDirBtnEl = bar.createEl('button', { cls: 'via-btn via-pdf-snap-btn' }) as HTMLButtonElement;
this.updateSnapDirBtn();
this.snapDirBtnEl.addEventListener('click', () => {
const dirs: SnapDirection[] = ['horizontal', 'vertical', 'slope'];
this.snapDirection = dirs[(dirs.indexOf(this.snapDirection) + 1) % dirs.length]!;
this.updateSnapDirBtn();
});

bar.createEl('div', { cls: 'via-toolbar-sep' });

const zoomOut = bar.createEl('button', { cls: 'via-btn via-btn-zoom', text: '\u2212' });
zoomOut.title = 'Zoom out (Ctrl+\u2212)';
zoomOut.addEventListener('click', () => this.stepZoom(-1));

this.zoomLabelEl = bar.createEl('button', {
cls:  'via-btn via-btn-zoom-label',
text: `${Math.round(this.currentScale * 100)}%`,
});
this.zoomLabelEl.title = `Reset zoom (Ctrl+0)`;
this.zoomLabelEl.addEventListener('click', () => this.setZoom(s.pdfDefaultZoom, this.viewportCenterFrac()));

const zoomIn = bar.createEl('button', { cls: 'via-btn via-btn-zoom', text: '+' });
zoomIn.title = 'Zoom in (Ctrl+=)';
zoomIn.addEventListener('click', () => this.stepZoom(+1));

bar.createEl('div', { cls: 'via-toolbar-sep' });

this.pageIndicatorEl = bar.createEl('button', { cls: 'via-pdf-page-indicator', text: '\u2014 / \u2014' });
this.pageIndicatorEl.title = 'Click to jump to page';
this.pageIndicatorEl.addEventListener('click', () => this.openPageJumpInput());

bar.createEl('div', { cls: 'via-toolbar-sep' });

const clearBtn = bar.createEl('button', { cls: 'via-btn via-btn-danger', text: '🗑 Clear page' });
clearBtn.addEventListener('click', () => this.clearCurrentPageAnnotations());

const saveBtn = bar.createEl('button', { cls: 'via-btn via-btn-save', text: '💾 Save annotations' });
saveBtn.addEventListener('click', () => this.persistAnnotations());

const exportBtn = bar.createEl('button', { cls: 'via-btn via-btn-export', text: '📤 Export PDF' });
exportBtn.title = 'Export PDF with annotations embedded';
exportBtn.addEventListener('click', () => {
if (this.currentFile && this.pdfDoc) {
exportAnnotatedPdf(this.app, this.currentFile, this.pdfDoc, this.annotData);
}
});

return bar;
}

// Zoom -------------------------------------------------------------------

private stepZoom(direction: -1 | 1): void {
const idx  = this.ZOOM_STEPS.findIndex(s => Math.abs(s - this.currentScale) < 0.01);
const next = this.ZOOM_STEPS[Math.max(0, Math.min(this.ZOOM_STEPS.length - 1, idx + direction))];
if (next !== undefined) this.setZoom(next, this.viewportCenterFrac());
}

private async setZoom(scale: number, frac?: { x: number; y: number; pX: number; pY: number }): Promise<void> {
if (Math.abs(scale - this.currentScale) < 0.001) return;
this.currentScale = scale;
if (this.zoomLabelEl) this.zoomLabelEl.textContent = `${Math.round(scale * 100)}%`;
await this.reRenderPages(frac);
}

private handleWheelZoom(e: WheelEvent): void {
if (!e.ctrlKey && !e.metaKey) return;
e.preventDefault();
const scrollEl = this.scrollAreaEl;
if (!scrollEl) return;

const rect = scrollEl.getBoundingClientRect();
const pX   = e.clientX - rect.left;
const pY   = e.clientY - rect.top;
const frac = {
x:  (scrollEl.scrollLeft + pX) / (scrollEl.scrollWidth  || 1),
y:  (scrollEl.scrollTop  + pY) / (scrollEl.scrollHeight || 1),
pX, pY,
};

const idx  = this.ZOOM_STEPS.findIndex(s => Math.abs(s - this.currentScale) < 0.01);
const next = this.ZOOM_STEPS[Math.max(0, Math.min(this.ZOOM_STEPS.length - 1, idx + (e.deltaY < 0 ? 1 : -1)))];
if (next === undefined || Math.abs(next - this.currentScale) < 0.001) return;

this.currentScale = next;
if (this.zoomLabelEl) this.zoomLabelEl.textContent = `${Math.round(next * 100)}%`;

if (this._zoomDebounceTimer !== null) clearTimeout(this._zoomDebounceTimer);
this._zoomDebounceTimer = setTimeout(() => {
this._zoomDebounceTimer = null;
this.reRenderPages(frac).catch(console.error);
}, 250);
}

private viewportCenterFrac(): { x: number; y: number; pX: number; pY: number } | undefined {
const el = this.scrollAreaEl;
if (!el) return undefined;
const pX = el.clientWidth / 2, pY = el.clientHeight / 2;
return {
x: (el.scrollLeft + pX) / (el.scrollWidth  || 1),
y: (el.scrollTop  + pY) / (el.scrollHeight || 1),
pX, pY,
};
}

private async reRenderPages(frac?: { x: number; y: number; pX: number; pY: number }): Promise<void> {
if (!this.pdfDoc || !this.scrollAreaEl) return;
this._renderGen++;
const scrollEl = this.scrollAreaEl;

this.renderObserver?.disconnect(); this.renderObserver = null;
this.pageObserver?.disconnect();   this.pageObserver   = null;

const sizes = await Promise.all(
this.pages.map(async (ctx) => {
const page = await this.pdfDoc!.getPage(ctx.pageNum);
const vp   = page.getViewport({ scale: this.currentScale });
return { w: Math.ceil(vp.width), h: Math.ceil(vp.height) };
})
);

for (let i = 0; i < this.pages.length; i++) {
const ctx     = this.pages[i]!;
const { w, h } = sizes[i]!;
ctx.w = w; ctx.h = h;
ctx.container.style.cssText = `width:${w}px;height:${h}px;min-width:${w}px;min-height:${h}px`;
this.unloadPageCanvas(ctx);
}

if (frac) {
scrollEl.scrollLeft = frac.x * scrollEl.scrollWidth  - frac.pX;
scrollEl.scrollTop  = frac.y * scrollEl.scrollHeight - frac.pY;
}

this.attachRenderObserver();
this.attachPageObserver();
}

// Page indicator ---------------------------------------------------------

private attachPageObserver(): void {
this.pageObserver?.disconnect();
if (!this.scrollAreaEl || this.pages.length === 0) return;

const total   = this.pdfDoc!.numPages;
const pageMap = new Map<Element, number>(this.pages.map(p => [p.container, p.pageNum]));
const visibleRatio = new Map<number, number>();

this.pageObserver = new IntersectionObserver(
(entries) => {
for (const entry of entries) {
const num = pageMap.get(entry.target);
if (num !== undefined) visibleRatio.set(num, entry.intersectionRatio);
}
let bestPage = 1, bestRatio = -1;
for (const [num, ratio] of visibleRatio) {
if (ratio > bestRatio) { bestRatio = ratio; bestPage = num; }
}
if (this.pageIndicatorEl) this.pageIndicatorEl.textContent = `${bestPage} / ${total}`;
},
{ root: this.scrollAreaEl, threshold: Array.from({ length: 11 }, (_, i) => i / 10) }
);

for (const ctx of this.pages) this.pageObserver.observe(ctx.container);
if (this.pageIndicatorEl) this.pageIndicatorEl.textContent = `1 / ${total}`;
}

private updateCanvasInteraction(): void {
for (const ctx of this.pages) {
if (!ctx.annotCanvas) continue;
const drawing = this.currentTool !== 'none' && this.currentTool !== 'note';
ctx.annotCanvas.style.pointerEvents = drawing ? 'auto' : 'none';
ctx.annotCanvas.style.cursor        = drawing ? 'crosshair' : 'default';
ctx.container.style.cursor          = this.currentTool === 'note' ? 'text' : '';
}
}

private setTool(tool: AnnotTool): void {
this.currentTool = tool;
this.containerEl.querySelectorAll('.via-pdf-toolbar [data-tool]').forEach(b =>
b.classList.toggle('via-btn-active', (b as HTMLElement).dataset.tool === tool)
);
this.updateCanvasInteraction();
const showColors = tool === 'pen' || tool === 'highlighter';
if (this.colorSectionEl) this.colorSectionEl.style.display = showColors ? 'flex' : 'none';
if (this.colorSepEl)     this.colorSepEl.style.display     = showColors ? '' : 'none';
if (this.widthSectionEl) this.widthSectionEl.style.display = showColors ? 'flex' : 'none';
if (this.widthSepEl)     this.widthSepEl.style.display     = showColors ? '' : 'none';
if (showColors) {
this.syncColorPicker(tool);
this.syncWidthSlider(tool);
}
}

// Color picker ------------------------------------------------------------

private syncColorPicker(tool: 'pen' | 'highlighter'): void {
const presets     = tool === 'pen' ? this.PEN_PRESETS : this.HIGHLIGHT_PRESETS;
const activeColor = (tool === 'pen'
? this.plugin.settings.penColor
: this.plugin.settings.highlighterColor
).toLowerCase();

for (let i = 0; i < this.colorSwatchEls.length; i++) {
const swatch = this.colorSwatchEls[i];
if (!swatch) continue;
const color = presets[i] ?? '';
swatch.style.background = color;
swatch.dataset.color    = color;
swatch.title            = color;
swatch.classList.toggle('via-color-swatch-active', color.toLowerCase() === activeColor);
}

if (this.colorCustomInputEl) {
this.colorCustomInputEl.value = activeColor;
const isCustom = !presets.some(c => c.toLowerCase() === activeColor);
this.colorCustomInputEl.parentElement?.classList.toggle('via-color-swatch-active', isCustom);
}
}

private applyColor(color: string): void {
const tool = this.currentTool;
if (tool !== 'pen' && tool !== 'highlighter') return;
if (tool === 'pen') this.plugin.settings.penColor = color;
else                this.plugin.settings.highlighterColor = color;
this.plugin.saveSettings();
this.syncColorPicker(tool);
}

// Width / opacity slider -------------------------------------------------

private syncWidthSlider(tool: 'pen' | 'highlighter'): void {
if (!this.widthSliderEl || !this.widthLabelEl) return;

if (tool === 'pen') {
this.widthSliderEl.min   = '1';
this.widthSliderEl.max   = '20';
this.widthSliderEl.step  = '1';
const w = this.plugin.settings.penWidth;
this.widthSliderEl.value  = String(w);
this.widthLabelEl.textContent = `${w}px`;
} else {
this.widthSliderEl.min   = '10';
this.widthSliderEl.max   = '40';
this.widthSliderEl.step  = '2';
const w = this.plugin.settings.highlighterWidth;
this.widthSliderEl.value  = String(w);
this.widthLabelEl.textContent = `${w}px`;
}

const showOpacity = tool === 'highlighter';
if (this.opacityRowEl) this.opacityRowEl.style.display = showOpacity ? 'flex' : 'none';
if (showOpacity && this.opacitySliderEl && this.opacityLabelEl) {
const op = this.plugin.settings.highlighterOpacity;
this.opacitySliderEl.value        = String(op);
this.opacityLabelEl.textContent   = `${Math.round(op * 100)}%`;
}
}

private applyWidth(value: number): void {
const tool = this.currentTool;
if (tool !== 'pen' && tool !== 'highlighter') return;
if (tool === 'pen') this.plugin.settings.penWidth          = value;
else                this.plugin.settings.highlighterWidth  = value;
this.plugin.saveSettings();
}

private applyOpacity(value: number): void {
this.plugin.settings.highlighterOpacity = value;
this.plugin.saveSettings();
}

// Snap -------------------------------------------------------------------

private updateSnapDirBtn(): void {
if (!this.snapDirBtnEl) return;
const labels = { horizontal: '⟷ H', vertical: '↕ V', slope: '↗ 45°' };
const { snapActivateKey, keySnapCycle } = this.plugin.settings;
this.snapDirBtnEl.textContent = labels[this.snapDirection];
this.snapDirBtnEl.title =
`Snap: ${this.snapDirection} — hold ${snapActivateKey} while drawing to activate · ${snapActivateKey}+${keySnapCycle.toUpperCase()} to cycle`;
}

// Page jump --------------------------------------------------------------

private openPageJumpInput(): void {
if (!this.pdfDoc || !this.pageIndicatorEl) return;
const total       = this.pdfDoc.numPages;
const currentPage = this.getVisiblePageNum();
const indicator   = this.pageIndicatorEl;

const input = document.createElement('input');
input.type      = 'number';
input.min       = '1';
input.max       = String(total);
input.value     = String(currentPage);
input.className = 'via-pdf-page-jump-input';

indicator.parentElement!.insertBefore(input, indicator);
indicator.style.display = 'none';
input.focus();
input.select();

const cleanup = () => { input.remove(); indicator.style.display = ''; };
const commit  = () => {
const val = parseInt(input.value, 10);
if (!isNaN(val)) this.scrollToPage(Math.max(1, Math.min(total, val)));
cleanup();
};

input.addEventListener('keydown', (e) => {
e.stopPropagation();
if (e.key === 'Enter')  { e.preventDefault(); commit();  }
if (e.key === 'Escape') { e.preventDefault(); cleanup(); }
});
input.addEventListener('blur', cleanup);
}

private scrollToPage(pageNum: number): void {
const ctx = this.pages.find(p => p.pageNum === pageNum);
if (ctx) ctx.container.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// Drawing ----------------------------------------------------------------

private attachDrawListeners(ctx: PageCtx): void {
const annotCanvas = ctx.annotCanvas;
if (!annotCanvas) return;
const { pageNum } = ctx;

const getPos = (e: MouseEvent | PointerEvent) => {
const rect = annotCanvas.getBoundingClientRect();
return { x: (e.clientX - rect.left) / rect.width, y: (e.clientY - rect.top) / rect.height };
};

const isSnapActive = (e: PointerEvent) => {
const key = this.plugin.settings.snapActivateKey;
return key === 'Alt' ? e.altKey : e.shiftKey;
};

annotCanvas.addEventListener('pointerdown', e => {
if (this.currentTool === 'none') return;
annotCanvas.setPointerCapture(e.pointerId);
this.isDrawing = true;
this.currentPath = {
tool: this.currentTool === 'pen' ? 'pen' : this.currentTool === 'eraser' ? 'eraser' : 'highlighter',
color: this.currentTool === 'pen'
? this.plugin.settings.penColor
: this.currentTool === 'highlighter'
? this.plugin.settings.highlighterColor
: '#ffffff',
width: this.currentTool === 'pen'
? this.plugin.settings.penWidth
: this.currentTool === 'highlighter'
? this.plugin.settings.highlighterWidth
: this.plugin.settings.eraserWidth,
opacity: this.currentTool === 'highlighter' ? this.plugin.settings.highlighterOpacity : 1,
points: [getPos(e)],
};
});

annotCanvas.addEventListener('pointermove', e => {
if (!this.isDrawing || !this.currentPath) return;
const raw = getPos(e);
if (isSnapActive(e) && this.currentPath.points.length >= 1) {
const origin  = this.currentPath.points[0]!;
const snapped = snapPoint(origin, raw, this.snapDirection);
// Replace trailing point to keep a clean constrained stroke
if (this.currentPath.points.length > 1) {
this.currentPath.points[this.currentPath.points.length - 1] = snapped;
} else {
this.currentPath.points.push(snapped);
}
this.snapDirBtnEl?.classList.add('via-btn-snap-active');
} else {
this.currentPath.points.push(raw);
if (!isSnapActive(e)) this.snapDirBtnEl?.classList.remove('via-btn-snap-active');
}
this.redrawAnnotations(ctx, this.currentPath);
});

const finishDraw = () => {
if (!this.isDrawing || !this.currentPath) return;
this.isDrawing = false;
this.snapDirBtnEl?.classList.remove('via-btn-snap-active');
let pa = getPageAnnotations(this.annotData, pageNum);
pa = { ...pa, paths: [...pa.paths, this.currentPath!] };
this.annotData = setPageAnnotations(this.annotData, pa);
this.currentPath = null;
this.redrawAnnotations(ctx);
};

annotCanvas.addEventListener('pointerup',     finishDraw);
annotCanvas.addEventListener('pointercancel', finishDraw);

// Note tool: click on page container to place a note
ctx.container.addEventListener('click', (e: MouseEvent) => {
if (this.currentTool !== 'note') return;
if ((e.target as HTMLElement).closest('.via-pdf-note')) return;
const rect = ctx.container.getBoundingClientRect();
const x = (e.clientX - rect.left) / rect.width;
const y = (e.clientY - rect.top)  / rect.height;
this.createNote(pageNum, x, y);
});
}

// Annotations ------------------------------------------------------------

private redrawAnnotations(ctx: PageCtx, inProgressPath?: AnnotationPath): void {
if (!ctx.annotCanvas) return;
const canvas = ctx.annotCanvas;
const c = canvas.getContext('2d')!;
c.clearRect(0, 0, canvas.width, canvas.height);

const drawPath = (path: AnnotationPath) => {
if (path.points.length < 2) return;
c.save();
if (path.tool === 'highlighter') {
c.globalAlpha = path.opacity ?? this.plugin.settings.highlighterOpacity;
c.globalCompositeOperation = 'multiply';
} else if (path.tool === 'eraser') {
c.globalCompositeOperation = 'destination-out'; c.globalAlpha = 1;
} else {
c.globalAlpha = 1; c.globalCompositeOperation = 'source-over';
}
c.strokeStyle = path.color;
c.lineWidth   = path.width * this.currentScale;
c.lineCap     = 'round';
c.lineJoin    = 'round';
c.beginPath();
c.moveTo(path.points[0]!.x * canvas.width, path.points[0]!.y * canvas.height);
for (let i = 1; i < path.points.length; i++) {
c.lineTo(path.points[i]!.x * canvas.width, path.points[i]!.y * canvas.height);
}
c.stroke();
c.restore();
};

const pa: PageAnnotations = getPageAnnotations(this.annotData, ctx.pageNum);
for (const path of pa.paths) drawPath(path);
if (inProgressPath) drawPath(inProgressPath);
}

// Persistence ------------------------------------------------------------

private clearCurrentPageAnnotations(): void {
const visiblePage = this.getVisiblePageNum();
this.annotData = setPageAnnotations(this.annotData, { page: visiblePage, paths: [] });
const ctx = this.pages.find(p => p.pageNum === visiblePage);
if (ctx) this.redrawAnnotations(ctx);
}

private getVisiblePageNum(): number {
let best = 1, bestVisible = -Infinity;
for (const ctx of this.pages) {
const rect    = ctx.container.getBoundingClientRect();
const visible = Math.min(rect.bottom, window.innerHeight) - Math.max(rect.top, 0);
if (visible > bestVisible) { bestVisible = visible; best = ctx.pageNum; }
}
return best;
}

private async persistAnnotations(): Promise<void> {
if (!this.currentFile) return;
try {
await saveAnnotations(this.app, this.currentFile, this.annotData);
new Notice('\u2705 Annotations saved');
} catch (err) {
new Notice(`\u274c Failed to save annotations: ${String(err)}`);
}
}

// TOC / Outline ----------------------------------------------------------

private async loadToc(): Promise<void> {
if (!this.pdfDoc) return;
try {
const outline = await this.pdfDoc.getOutline();
if (!outline || outline.length === 0) return;
(this as any)._outline = outline;
} catch {
// Some PDFs throw on getOutline — ignore
}
}

private toggleToc(): void {
if (!this.bodyEl) return;
this.tocVisible = !this.tocVisible;

if (!this.tocVisible) {
this.tocSidebarEl?.remove();
this.tocSidebarEl = null;
return;
}

const sidebar = this.bodyEl.createEl('div', { cls: 'via-pdf-toc' });
this.tocSidebarEl = sidebar;
this.bodyEl.insertBefore(sidebar, this.scrollAreaEl);

const header   = sidebar.createEl('div', { cls: 'via-pdf-toc-header' });
header.createEl('span', { text: 'Contents' });
const closeBtn = header.createEl('button', { cls: 'via-btn', text: '✕' });
closeBtn.addEventListener('click', () => {
this.tocVisible = false;
sidebar.remove();
this.tocSidebarEl = null;
});

const list    = sidebar.createEl('div', { cls: 'via-pdf-toc-list' });
const outline = (this as any)._outline as any[];

if (!outline || outline.length === 0) {
list.createEl('p', { cls: 'via-pdf-toc-empty', text: 'No outline available for this PDF.' });
return;
}

const renderItems = (items: typeof outline, parentEl: HTMLElement, depth = 0) => {
for (const item of items) {
const row = parentEl.createEl('div', { cls: 'via-pdf-toc-item' });
row.style.paddingLeft = `${8 + depth * 14}px`;

if (item.items && item.items.length > 0) {
const toggle    = row.createEl('span', { cls: 'via-pdf-toc-toggle', text: '▾' });
let   collapsed = false;
const childList = parentEl.createEl('div');
renderItems(item.items, childList, depth + 1);

toggle.addEventListener('click', (e) => {
e.stopPropagation();
collapsed = !collapsed;
childList.style.display = collapsed ? 'none' : '';
toggle.textContent      = collapsed ? '▸' : '▾';
});
}

const label = row.createEl('span', { cls: 'via-pdf-toc-label', text: item.title ?? '(untitled)' });
label.addEventListener('click', async () => {
if (!this.pdfDoc) return;
try {
let dest = item.dest;
if (typeof dest === 'string') dest = await this.pdfDoc.getDestination(dest);
if (!Array.isArray(dest) || dest.length === 0) return;
const pageIdx = await this.pdfDoc.getPageIndex(dest[0] as PdfRefProxy);
this.scrollToPage(pageIdx + 1);
} catch {
// Destination lookup failed — silently ignore
}
});
}
};

renderItems(outline, list);
}

// Text notes -------------------------------------------------------------

private createNote(pageNum: number, x: number, y: number): void {
const note: TextNote = {
id:    `note-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
page:  pageNum,
x, y,
text:  '',
color: this.plugin.settings.noteDefaultColor,
};
this.annotData = {
...this.annotData,
notes: [...(this.annotData.notes ?? []), note],
};
this.renderNoteEl(note, true);
}

private renderNoteEl(note: TextNote, focusImmediately = false): void {
const ctx = this.pages.find(p => p.pageNum === note.page);
if (!ctx) return;

const el = ctx.container.createEl('div', { cls: 'via-pdf-note' });
el.style.cssText   = `left:${note.x * 100}%;top:${note.y * 100}%;background:${note.color ?? this.plugin.settings.noteDefaultColor}`;
el.dataset.noteId  = note.id;

const header = el.createEl('div', { cls: 'via-pdf-note-header' });

// Drag handle
header.createEl('span', { cls: 'via-pdf-note-drag', text: '⠿' }).addEventListener('mousedown', (e) => {
e.preventDefault();
const startX = e.clientX, startY = e.clientY;
const origLeft = note.x, origTop = note.y;
const rect = ctx.container.getBoundingClientRect();

const onMove = (me: MouseEvent) => {
const dx = (me.clientX - startX) / rect.width;
const dy = (me.clientY - startY) / rect.height;
note.x = Math.max(0, Math.min(0.9,  origLeft + dx));
note.y = Math.max(0, Math.min(0.95, origTop  + dy));
el.style.left = `${note.x * 100}%`;
el.style.top  = `${note.y * 100}%`;
};
const onUp = () => {
window.removeEventListener('mousemove', onMove);
window.removeEventListener('mouseup',   onUp);
};
window.addEventListener('mousemove', onMove);
window.addEventListener('mouseup',   onUp);
});

const deleteBtn = header.createEl('button', { cls: 'via-pdf-note-delete', text: '✕' });
deleteBtn.addEventListener('click', () => {
this.annotData = {
...this.annotData,
notes: (this.annotData.notes ?? []).filter(n => n.id !== note.id),
};
el.remove();
this.noteEls.delete(note.id);
});

const textarea = el.createEl('textarea', { cls: 'via-pdf-note-text' });
textarea.value       = note.text;
textarea.placeholder = 'Note…';
textarea.addEventListener('input',   () => { note.text = textarea.value; });
textarea.addEventListener('keydown', e => e.stopPropagation());

this.noteEls.set(note.id, el);
if (focusImmediately) textarea.focus();
}
}
