import { FileView, TFile, WorkspaceLeaf, Notice } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import type { TextItem as PdfTextItem, RefProxy as PdfRefProxy } from 'pdfjs-dist/types/src/display/api';
import { PDFDocument, rgb, LineCapStyle } from 'pdf-lib';
import { VIEW_TYPE_PDF } from '../types';
import type { PageAnnotations, AnnotationPath, AnnotationFile, TextNote } from '../types';
import {
	loadAnnotations,
	saveAnnotations,
	getPageAnnotations,
	setPageAnnotations,
} from '../utils/pdfAnnotations';
import type ViewItAllPlugin from '../main';

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
type PageRenderState = 'placeholder' | 'rendering' | 'rendered';

interface PageCtx {
	pageNum: number;
	state: PageRenderState;
	container: HTMLElement;
	pdfCanvas: HTMLCanvasElement | null;
	annotCanvas: HTMLCanvasElement | null;
	searchCanvas: HTMLCanvasElement | null;
	w: number;
	h: number;
}

interface SearchMatch {
	pageNum: number;
	x: number; y: number; w: number; h: number; // normalised 0-1
}

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

	// Lazy rendering — renderObserver triggers canvas creation as pages enter viewport;
	// pageObserver drives the toolbar page indicator.
	private pageIndicatorEl: HTMLElement | null = null;
	private pageObserver: IntersectionObserver | null = null;
	private renderObserver: IntersectionObserver | null = null;
	// Incremented on reload/zoom to cancel in-flight async renders
	private _renderGen = 0;

	// Color picker
	private colorSectionEl: HTMLElement | null = null;
	private colorSepEl: HTMLElement | null = null;
	private colorSwatchEls: HTMLButtonElement[] = [];
	private colorCustomInputEl: HTMLInputElement | null = null;

	private readonly PEN_PRESETS    = ['#e03131', '#1971c2', '#2f9e44', '#212529', '#e8590c', '#7048e8'];
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
	private snapDirection: 'horizontal' | 'vertical' | 'slope' = 'horizontal';
	private snapDirBtnEl: HTMLButtonElement | null = null;

	// Text search
	private wrapperEl: HTMLElement | null = null;
	private bodyEl: HTMLElement | null = null;          // flex-row container: toc + scroll
	private tocSidebarEl: HTMLElement | null = null;
	private tocVisible = false;
	private searchBarEl: HTMLElement | null = null;
	private searchInputEl: HTMLInputElement | null = null;
	private searchMatchCountEl: HTMLElement | null = null;
	private searchMatches: SearchMatch[] = [];
	private searchCurrentIdx = -1;
	private _textCache = new Map<number, PdfTextItem[]>();

	// Text notes overlay elements (keyed by note id)
	private noteEls = new Map<string, HTMLElement>();

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
		pdfjsLib.GlobalWorkerOptions.workerSrc = getPdfWorkerUrl();
	}

	onload(): void {
		super.onload();

		// ── Document-level Alt / Alt+S handler ──────────────────────────────
		// Registered on document so it fires regardless of which element has
		// focus — as long as this leaf is the active one.
		const isActiveLeaf = () => this.app.workspace.getActiveViewOfType(PdfView) === this;

		this.registerDomEvent(document as unknown as HTMLElement, 'keydown', (e: KeyboardEvent) => {
			if (!isActiveLeaf()) return;
			if (e.key === 'Alt') {
				// Prevent browser default (usually opens menu bar on Windows)
				e.preventDefault();
				const drawingTool = this.currentTool === 'pen' || this.currentTool === 'highlighter' || this.currentTool === 'eraser';
				if (drawingTool) this.snapDirBtnEl?.classList.add('via-btn-snap-active');
			} else if (e.altKey && e.key.toLowerCase() === 's') {
				// Alt+S — cycle snap direction
				e.preventDefault();
				const dirs: Array<'horizontal' | 'vertical' | 'slope'> = ['horizontal', 'vertical', 'slope'];
				this.snapDirection = dirs[(dirs.indexOf(this.snapDirection) + 1) % dirs.length]!;
				this.updateSnapDirBtn();
			}
		});
		this.registerDomEvent(document as unknown as HTMLElement, 'keyup', (e: KeyboardEvent) => {
			if (!isActiveLeaf()) return;
			if (e.key === 'Alt') this.snapDirBtnEl?.classList.remove('via-btn-snap-active');
		});

		// ── Container-level shortcuts (zoom, tools, search) ─────────────────
		this.registerDomEvent(this.containerEl as HTMLElement, 'keydown', (e: KeyboardEvent) => {
			// Ctrl/Cmd zoom shortcuts
			if (e.ctrlKey || e.metaKey) {
				if (e.key === '0') { e.preventDefault(); this.setZoom(1.0, this.viewportCenterFrac()); }
				else if (e.key === '=' || e.key === '+') { e.preventDefault(); this.stepZoom(+1); }
				else if (e.key === '-') { e.preventDefault(); this.stepZoom(-1); }
				else if (e.key === 'f' || e.key === 'F') { e.preventDefault(); this.openSearchBar(); }
				return;
			}
			// Skip Alt combos — handled above
			if (e.altKey) return;
			// Tool shortcuts — guard: don't fire when an input/textarea is focused
			const target = e.target as HTMLElement;
			if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable) return;
			const toolMap: Record<string, AnnotTool> = { v: 'none', p: 'pen', h: 'highlighter', e: 'eraser', n: 'note' };
			const tool = toolMap[e.key.toLowerCase()];
			if (tool !== undefined) { e.preventDefault(); this.setTool(tool); }
		});
	}

	getViewType(): string { return VIEW_TYPE_PDF; }
	getDisplayText(): string { return this.file?.basename ?? 'PDF'; }
	getIcon(): string { return 'file'; }
	canAcceptExtension(extension: string): boolean { return extension === 'pdf'; }

	async onLoadFile(file: TFile): Promise<void> {
		this.currentFile = file;
		this.currentTool = this.plugin.settings.pdfDefaultTool;
		this.annotData = await loadAnnotations(this.app, file);
		await this.renderPdf(file);
	}

	async onUnloadFile(_file: TFile): Promise<void> {
		this._renderGen++;
		if (this._zoomDebounceTimer !== null) { clearTimeout(this._zoomDebounceTimer); this._zoomDebounceTimer = null; }
		this.renderObserver?.disconnect(); this.renderObserver = null;
		this.pageObserver?.disconnect(); this.pageObserver = null;
		if (this.pdfDoc) { this.pdfDoc.destroy(); this.pdfDoc = null; }
		this.pages = [];
		this._textCache.clear();
		this.searchMatches = [];
		this.searchCurrentIdx = -1;
		this.noteEls.clear();
		this.tocSidebarEl = null;
		this.tocVisible = false;
		this.bodyEl = null;
		this.snapDirBtnEl = null;
		this.contentEl.empty();
	}

	private async renderPdf(file: TFile): Promise<void> {
		this._renderGen++;
		this.contentEl.empty();
		this.pages = [];
		this._textCache.clear();
		this.searchMatches = [];
		this.searchCurrentIdx = -1;
		this.noteEls.clear();
		this.tocSidebarEl = null;
		this.tocVisible = false;
		this.renderObserver?.disconnect(); this.renderObserver = null;
		this.pageObserver?.disconnect(); this.pageObserver = null;

		const isBottom = this.plugin.settings.pdfToolbarPosition === 'bottom';
		const wrapper = this.contentEl.createEl('div', { cls: 'via-pdf-wrapper' });
		if (isBottom) wrapper.classList.add('via-pdf-wrapper--toolbar-bottom');
		this.wrapperEl = wrapper;

		const toolbar = this.buildToolbar();
		wrapper.appendChild(toolbar);

		// Body area: flex-row holds optional TOC sidebar + scroll area
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

		// Pre-calculate ALL page sizes in parallel (viewport math only, no canvas ops)
		const sizes = await Promise.all(
			Array.from({ length: numPages }, async (_, i) => {
				const page = await this.pdfDoc!.getPage(i + 1);
				const vp = page.getViewport({ scale: this.currentScale });
				return { w: Math.ceil(vp.width), h: Math.ceil(vp.height) };
			})
		);

		loadingEl.remove();

		// Create all placeholder divs + labels atomically
		for (let i = 0; i < numPages; i++) {
			const { w, h } = sizes[i]!;
			const container = scrollArea.createEl('div', { cls: 'via-pdf-page' });
			container.style.cssText = `width:${w}px;height:${h}px;min-width:${w}px;min-height:${h}px`;
			scrollArea.createEl('div', { cls: 'via-pdf-page-label', text: `${i + 1} / ${numPages}` });
			this.pages.push({ pageNum: i + 1, state: 'placeholder', container, pdfCanvas: null, annotCanvas: null, searchCanvas: null, w, h });
		}

		if (this.pageIndicatorEl) this.pageIndicatorEl.textContent = `1 / ${numPages}`;
		this.attachRenderObserver();
		this.attachPageObserver();

		// Render existing text notes
		for (const note of (this.annotData.notes ?? [])) {
			this.renderNoteEl(note);
		}

		// Load TOC in background (non-blocking)
		this.loadToc().catch(console.error);
	}

	// Lazy rendering ---------------------------------------------------------

	// rootMargin "200% 0px": render pages within 2x viewport height above/below.
	// Pages leaving that zone are unloaded to free GPU memory.
	private attachRenderObserver(): void {
		this.renderObserver?.disconnect();
		if (!this.scrollAreaEl || this.pages.length === 0) return;

		const pageMap = new Map<Element, PageCtx>(this.pages.map(p => [p.container, p]));

		this.renderObserver = new IntersectionObserver(
			(entries) => {
				for (const entry of entries) {
					const ctx = pageMap.get(entry.target);
					if (!ctx) continue;
					if (entry.isIntersecting) {
						this.renderPageCanvas(ctx).catch(console.error);
					} else {
						this.unloadPageCanvas(ctx);
					}
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

		const page = await this.pdfDoc!.getPage(ctx.pageNum);
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

		ctx.pdfCanvas = pdfCanvas;
		ctx.annotCanvas = annotCanvas;
		ctx.searchCanvas = searchCanvas;
		ctx.state = 'rendered';

		this.redrawAnnotations(ctx);
		if (this.searchMatches.length > 0) this.drawSearchHighlightsForPage(ctx);
		this.attachDrawListeners(ctx);
		this.updateCanvasInteraction();
	}

	private unloadPageCanvas(ctx: PageCtx): void {
		if (ctx.state !== 'rendered') return;
		ctx.pdfCanvas?.remove(); ctx.pdfCanvas = null;
		ctx.annotCanvas?.remove(); ctx.annotCanvas = null;
		ctx.searchCanvas?.remove(); ctx.searchCanvas = null;
		ctx.state = 'placeholder';
	}

	// Toolbar ----------------------------------------------------------------

	private buildToolbar(): HTMLElement {
		const bar = document.createElement('div');
		bar.className = 'via-pdf-toolbar';

		// TOC toggle
		const tocBtn = bar.createEl('button', { cls: 'via-btn', text: '📑 TOC' });
		tocBtn.title = 'Toggle table of contents';
		tocBtn.addEventListener('click', () => this.toggleToc());

		bar.createEl('div', { cls: 'via-toolbar-sep' });

		const tools: { id: AnnotTool; label: string; key: string }[] = [
			{ id: 'none',        label: '👁 View',      key: 'V' },
			{ id: 'pen',         label: '✏️ Pen',        key: 'P' },
			{ id: 'highlighter', label: '🖊 Highlight',  key: 'H' },
			{ id: 'eraser',      label: '⬜ Erase',      key: 'E' },
			{ id: 'note',        label: '📝 Note',       key: 'N' },
		];

		for (const t of tools) {
			const btn = bar.createEl('button', { cls: 'via-btn', text: t.label });
			btn.dataset.tool = t.id;
			btn.title = `${t.label} (${t.key})`;
			if (t.id === this.currentTool) btn.classList.add('via-btn-active');
			btn.addEventListener('click', () => this.setTool(t.id));
		}

		bar.createEl('div', { cls: 'via-toolbar-sep' });

		// Color picker section — shown only when pen / highlighter is active
		this.colorSwatchEls = [];
		this.colorSectionEl = bar.createEl('div', { cls: 'via-pdf-color-section' });
		const initPresets = this.currentTool === 'highlighter' ? this.HIGHLIGHT_PRESETS : this.PEN_PRESETS;
		for (const color of initPresets) {
			const swatch = this.colorSectionEl.createEl('button', { cls: 'via-color-swatch' });
			swatch.style.background = color;
			swatch.dataset.color = color;
			swatch.title = color;
			swatch.addEventListener('click', () => this.applyColor(swatch.dataset.color!));
			this.colorSwatchEls.push(swatch);
		}
		const customLabel = this.colorSectionEl.createEl('label', { cls: 'via-color-swatch via-color-custom', title: 'Custom colour' });
		const customInput = customLabel.createEl('input');
		customInput.type = 'color';
		customInput.className = 'via-color-custom-input';
		this.colorCustomInputEl = customInput;
		customInput.addEventListener('input', () => {
			// Live preview: update swatch active ring without persisting
			const tool = this.currentTool;
			if (tool !== 'pen' && tool !== 'highlighter') return;
			const presets = tool === 'pen' ? this.PEN_PRESETS : this.HIGHLIGHT_PRESETS;
			const color   = customInput.value.toLowerCase();
			for (const s of this.colorSwatchEls) s.classList.toggle('via-color-swatch-active', s.dataset.color?.toLowerCase() === color);
			const isCustom = !presets.some(c => c.toLowerCase() === color);
			customInput.parentElement?.classList.toggle('via-color-swatch-active', isCustom);
		});
		customInput.addEventListener('change', () => this.applyColor(customInput.value));

		this.colorSepEl = bar.createEl('div', { cls: 'via-toolbar-sep' });

		const showColors = this.currentTool === 'pen' || this.currentTool === 'highlighter';
		this.colorSectionEl.style.display = showColors ? 'flex' : 'none';
		this.colorSepEl.style.display     = showColors ? '' : 'none';
		if (showColors) this.syncColorPicker(this.currentTool as 'pen' | 'highlighter');

		// Width / opacity slider section
		this.widthSectionEl = bar.createEl('div', { cls: 'via-pdf-width-section' });

		const widthRow = this.widthSectionEl.createEl('div', { cls: 'via-pdf-width-row' });
		widthRow.createEl('span', { cls: 'via-pdf-slider-label', text: 'Size' });
		this.widthSliderEl = widthRow.createEl('input');
		this.widthSliderEl.type = 'range';
		this.widthSliderEl.className = 'via-pdf-slider';
		this.widthLabelEl = widthRow.createEl('span', { cls: 'via-pdf-slider-value' });

		this.widthSliderEl.addEventListener('input', () => {
			const v = Number(this.widthSliderEl!.value);
			if (this.widthLabelEl) this.widthLabelEl.textContent = `${v}px`;
		});
		this.widthSliderEl.addEventListener('change', () => this.applyWidth(Number(this.widthSliderEl!.value)));

		this.opacityRowEl = this.widthSectionEl.createEl('div', { cls: 'via-pdf-width-row' });
		this.opacityRowEl.createEl('span', { cls: 'via-pdf-slider-label', text: 'Opacity' });
		this.opacitySliderEl = this.opacityRowEl.createEl('input');
		this.opacitySliderEl.type = 'range';
		this.opacitySliderEl.min = '0.1'; this.opacitySliderEl.max = '1.0'; this.opacitySliderEl.step = '0.05';
		this.opacitySliderEl.className = 'via-pdf-slider';
		this.opacityLabelEl = this.opacityRowEl.createEl('span', { cls: 'via-pdf-slider-value' });

		this.opacitySliderEl.addEventListener('input', () => {
			const v = Number(this.opacitySliderEl!.value);
			if (this.opacityLabelEl) this.opacityLabelEl.textContent = `${Math.round(v * 100)}%`;
		});
		this.opacitySliderEl.addEventListener('change', () => this.applyOpacity(Number(this.opacitySliderEl!.value)));

		this.widthSepEl = bar.createEl('div', { cls: 'via-toolbar-sep' });

		this.widthSectionEl.style.display = showColors ? 'flex' : 'none';
		this.widthSepEl.style.display     = showColors ? '' : 'none';
		if (showColors) this.syncWidthSlider(this.currentTool as 'pen' | 'highlighter');

		// Snap direction button — always visible, Shift to activate
		this.snapDirBtnEl = bar.createEl('button', { cls: 'via-btn via-pdf-snap-btn' }) as HTMLButtonElement;
		this.updateSnapDirBtn();
		this.snapDirBtnEl.addEventListener('click', () => {
			const dirs: Array<'horizontal' | 'vertical' | 'slope'> = ['horizontal', 'vertical', 'slope'];
			const idx = dirs.indexOf(this.snapDirection);
			this.snapDirection = dirs[(idx + 1) % dirs.length]!;
			this.updateSnapDirBtn();
		});

		bar.createEl('div', { cls: 'via-toolbar-sep' });

		const zoomOut = bar.createEl('button', { cls: 'via-btn via-btn-zoom', text: '\u2212' });
		zoomOut.title = 'Zoom out (Ctrl+\u2212)';
		zoomOut.addEventListener('click', () => this.stepZoom(-1));

		this.zoomLabelEl = bar.createEl('button', {
			cls: 'via-btn via-btn-zoom-label',
			text: `${Math.round(this.currentScale * 100)}%`,
		});
		this.zoomLabelEl.title = 'Reset to 100% (Ctrl+0)';
		this.zoomLabelEl.addEventListener('click', () => this.setZoom(1.0, this.viewportCenterFrac()));

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
		exportBtn.addEventListener('click', () => this.exportAnnotatedPdf());

		return bar;
	}

	// Zoom -------------------------------------------------------------------

	private stepZoom(direction: -1 | 1): void {
		const idx = this.ZOOM_STEPS.findIndex(s => Math.abs(s - this.currentScale) < 0.01);
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
		const pX = e.clientX - rect.left;
		const pY = e.clientY - rect.top;
		const frac = {
			x: (scrollEl.scrollLeft + pX) / (scrollEl.scrollWidth  || 1),
			y: (scrollEl.scrollTop  + pY) / (scrollEl.scrollHeight || 1),
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

	// On zoom: cancel in-flight renders, resize all placeholders in parallel,
	// then re-observe. Only visible pages get re-rendered — O(visible) not O(total).
	private async reRenderPages(frac?: { x: number; y: number; pX: number; pY: number }): Promise<void> {
		if (!this.pdfDoc || !this.scrollAreaEl) return;
		this._renderGen++;
		const scrollEl = this.scrollAreaEl;

		this.renderObserver?.disconnect(); this.renderObserver = null;
		this.pageObserver?.disconnect(); this.pageObserver = null;

		// Recalculate sizes in parallel
		const sizes = await Promise.all(
			this.pages.map(async (ctx) => {
				const page = await this.pdfDoc!.getPage(ctx.pageNum);
				const vp = page.getViewport({ scale: this.currentScale });
				return { w: Math.ceil(vp.width), h: Math.ceil(vp.height) };
			})
		);

		for (let i = 0; i < this.pages.length; i++) {
			const ctx = this.pages[i]!;
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

		const total = this.pdfDoc!.numPages;
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
			ctx.annotCanvas.style.cursor = drawing ? 'crosshair' : 'default';
			// Note tool uses the page container itself (below canvas layer)
			ctx.container.style.cursor = this.currentTool === 'note' ? 'text' : '';
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
			this.widthSliderEl.min = '1'; this.widthSliderEl.max = '20'; this.widthSliderEl.step = '1';
			const w = this.plugin.settings.penWidth;
			this.widthSliderEl.value = String(w);
			this.widthLabelEl.textContent = `${w}px`;
		} else {
			this.widthSliderEl.min = '10'; this.widthSliderEl.max = '40'; this.widthSliderEl.step = '2';
			const w = this.plugin.settings.highlighterWidth;
			this.widthSliderEl.value = String(w);
			this.widthLabelEl.textContent = `${w}px`;
		}

		const showOpacity = tool === 'highlighter';
		if (this.opacityRowEl) this.opacityRowEl.style.display = showOpacity ? 'flex' : 'none';
		if (showOpacity && this.opacitySliderEl && this.opacityLabelEl) {
			const op = this.plugin.settings.highlighterOpacity;
			this.opacitySliderEl.value = String(op);
			this.opacityLabelEl.textContent = `${Math.round(op * 100)}%`;
		}
	}

	private applyWidth(value: number): void {
		const tool = this.currentTool;
		if (tool !== 'pen' && tool !== 'highlighter') return;
		if (tool === 'pen') this.plugin.settings.penWidth = value;
		else                this.plugin.settings.highlighterWidth = value;
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
		this.snapDirBtnEl.textContent = labels[this.snapDirection];
		this.snapDirBtnEl.title = `Snap: ${this.snapDirection} — Alt+S to cycle · hold Alt while drawing to activate`;
	}

	/** Constrain `raw` to the current snap direction from `origin`. */
	private snapPoint(
		origin: { x: number; y: number },
		raw: { x: number; y: number }
	): { x: number; y: number } {
		const dx = raw.x - origin.x;
		const dy = raw.y - origin.y;
		switch (this.snapDirection) {
			case 'horizontal':
				return { x: raw.x, y: origin.y };
			case 'vertical':
				return { x: origin.x, y: raw.y };
			case 'slope': {
				// Snap to nearest 45° increment (8 directions)
				const angle = Math.atan2(dy, dx);
				const snapped = Math.round(angle / (Math.PI / 4)) * (Math.PI / 4);
				const dist = Math.sqrt(dx * dx + dy * dy);
				return {
					x: origin.x + dist * Math.cos(snapped),
					y: origin.y + dist * Math.sin(snapped),
				};
			}
		}
	}

	// Page jump --------------------------------------------------------------

	private openPageJumpInput(): void {
		if (!this.pdfDoc || !this.pageIndicatorEl) return;
		const total = this.pdfDoc.numPages;
		const currentPage = this.getVisiblePageNum();
		const indicator = this.pageIndicatorEl;

		const input = document.createElement('input');
		input.type = 'number';
		input.min = '1';
		input.max = String(total);
		input.value = String(currentPage);
		input.className = 'via-pdf-page-jump-input';

		indicator.parentElement!.insertBefore(input, indicator);
		indicator.style.display = 'none';
		input.focus();
		input.select();

		const cleanup = () => {
			input.remove();
			indicator.style.display = '';
		};

		const commit = () => {
			const val = parseInt(input.value, 10);
			if (!isNaN(val)) this.scrollToPage(Math.max(1, Math.min(total, val)));
			cleanup();
		};

		input.addEventListener('keydown', (e) => {
			e.stopPropagation();
			if (e.key === 'Enter') { e.preventDefault(); commit(); }
			else if (e.key === 'Escape') { e.preventDefault(); cleanup(); }
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
			if (e.altKey && this.currentPath.points.length >= 1) {
				const origin = this.currentPath.points[0]!;
				const snapped = this.snapPoint(origin, raw);
				// Replace last point to keep a clean constrained stroke
				if (this.currentPath.points.length > 1) {
					this.currentPath.points[this.currentPath.points.length - 1] = snapped;
				} else {
					this.currentPath.points.push(snapped);
				}
				this.snapDirBtnEl?.classList.add('via-btn-snap-active');
			} else {
				this.currentPath.points.push(raw);
				if (!e.altKey) this.snapDirBtnEl?.classList.remove('via-btn-snap-active');
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

		annotCanvas.addEventListener('pointerup', finishDraw);
		annotCanvas.addEventListener('pointercancel', finishDraw);

		// Note tool: click on page container (not canvas) to place a note
		ctx.container.addEventListener('click', (e: MouseEvent) => {
			if (this.currentTool !== 'note') return;
			// Ignore clicks on existing note overlays
			if ((e.target as HTMLElement).closest('.via-pdf-note')) return;
			const rect = ctx.container.getBoundingClientRect();
			const x = (e.clientX - rect.left) / rect.width;
			const y = (e.clientY - rect.top) / rect.height;
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
			c.lineWidth = path.width * this.currentScale;
			c.lineCap = 'round'; c.lineJoin = 'round';
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
			const rect = ctx.container.getBoundingClientRect();
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

	// Text search ------------------------------------------------------------

	private openSearchBar(): void {
		if (this.searchBarEl) { this.searchInputEl?.focus(); return; }
		if (!this.wrapperEl) return;

		const bar = this.wrapperEl.createEl('div', { cls: 'via-pdf-search-bar' });
		this.searchBarEl = bar;

		const searchIcon = bar.createEl('span', { cls: 'via-pdf-search-icon', text: '🔍' });
		searchIcon.style.flexShrink = '0';

		this.searchInputEl = bar.createEl('input');
		this.searchInputEl.type = 'text';
		this.searchInputEl.placeholder = 'Search…';
		this.searchInputEl.className = 'via-pdf-search-input';

		this.searchMatchCountEl = bar.createEl('span', { cls: 'via-pdf-search-count', text: '' });

		const prevBtn = bar.createEl('button', { cls: 'via-btn', text: '↑' });
		prevBtn.title = 'Previous match (Shift+Enter)';
		prevBtn.addEventListener('click', () => this.goToMatch(this.searchCurrentIdx - 1));

		const nextBtn = bar.createEl('button', { cls: 'via-btn', text: '↓' });
		nextBtn.title = 'Next match (Enter)';
		nextBtn.addEventListener('click', () => this.goToMatch(this.searchCurrentIdx + 1));

		const closeBtn = bar.createEl('button', { cls: 'via-btn', text: '✕' });
		closeBtn.title = 'Close search (Escape)';
		closeBtn.addEventListener('click', () => this.closeSearchBar());

		this.searchInputEl.addEventListener('keydown', (e) => {
			e.stopPropagation();
			if (e.key === 'Enter') {
				e.preventDefault();
				e.shiftKey ? this.goToMatch(this.searchCurrentIdx - 1) : this.goToMatch(this.searchCurrentIdx + 1);
			} else if (e.key === 'Escape') {
				e.preventDefault();
				this.closeSearchBar();
			}
		});

		let debounce: ReturnType<typeof setTimeout> | null = null;
		this.searchInputEl.addEventListener('input', () => {
			if (debounce) clearTimeout(debounce);
			debounce = setTimeout(() => this.performSearch(this.searchInputEl!.value), 300);
		});

		// Insert search bar: after toolbar (first child of wrapper), before body area
		const bodyEl = this.bodyEl;
		if (bodyEl) this.wrapperEl!.insertBefore(bar, bodyEl);

		this.searchInputEl.focus();
	}

	private closeSearchBar(): void {
		this.searchMatches = [];
		this.searchCurrentIdx = -1;
		this.clearAllSearchHighlights();
		this.updateSearchCount();
		this.searchBarEl?.remove();
		this.searchBarEl = null;
		this.searchInputEl = null;
		this.searchMatchCountEl = null;
	}

	private async performSearch(query: string): Promise<void> {
		this.searchMatches = [];
		this.searchCurrentIdx = -1;
		this.clearAllSearchHighlights();

		if (!query.trim() || !this.pdfDoc) {
			this.updateSearchCount();
			return;
		}

		const q = query.toLowerCase();
		const total = this.pdfDoc.numPages;

		for (let pn = 1; pn <= total; pn++) {
			const items = await this.getPageTextItems(pn);
			const page = await this.pdfDoc!.getPage(pn);
			const vp = page.getViewport({ scale: 1.0 });

			for (const item of items) {
				const str = item.str.toLowerCase();
				let idx = 0;
				while ((idx = str.indexOf(q, idx)) !== -1) {
					const tx = item.transform[4] ?? 0;
					const ty = item.transform[5] ?? 0;
					const fontSize = Math.abs(item.transform[3] ?? 12);
					const charWidth = (item.width || 0) / (item.str.length || 1);
					const matchX = tx + charWidth * idx;
					const matchW = charWidth * query.length;

					const [cx, cy]   = vp.convertToViewportPoint(matchX, ty);
					const [cx2, cy2] = vp.convertToViewportPoint(matchX + matchW, ty - fontSize);

					this.searchMatches.push({
						pageNum: pn,
						x: Math.min(cx, cx2) / vp.width,
						y: Math.min(cy, cy2) / vp.height,
						w: Math.abs(cx2 - cx) / vp.width,
						h: Math.abs(cy2 - cy) / vp.height,
					});
					idx += q.length;
				}
			}
		}

		if (this.searchMatches.length > 0) {
			this.searchCurrentIdx = 0;
			this.scrollToMatch(0);
		}

		this.updateSearchCount();
		this.redrawAllSearchHighlights();
	}

	private goToMatch(idx: number): void {
		if (this.searchMatches.length === 0) return;
		this.searchCurrentIdx = ((idx % this.searchMatches.length) + this.searchMatches.length) % this.searchMatches.length;
		this.scrollToMatch(this.searchCurrentIdx);
		this.updateSearchCount();
		this.redrawAllSearchHighlights();
	}

	private scrollToMatch(idx: number): void {
		const m = this.searchMatches[idx];
		if (!m) return;
		const ctx = this.pages.find(p => p.pageNum === m.pageNum);
		if (ctx) ctx.container.scrollIntoView({ behavior: 'smooth', block: 'center' });
	}

	private updateSearchCount(): void {
		if (!this.searchMatchCountEl) return;
		if (this.searchMatches.length === 0) {
			this.searchMatchCountEl.textContent = this.searchInputEl?.value.trim() ? 'No matches' : '';
		} else {
			this.searchMatchCountEl.textContent = `${this.searchCurrentIdx + 1} / ${this.searchMatches.length}`;
		}
	}

	private drawSearchHighlightsForPage(ctx: PageCtx): void {
		if (!ctx.searchCanvas) return;
		const canvas = ctx.searchCanvas;
		const c = canvas.getContext('2d')!;
		c.clearRect(0, 0, canvas.width, canvas.height);

		for (let i = 0; i < this.searchMatches.length; i++) {
			const m = this.searchMatches[i]!;
			if (m.pageNum !== ctx.pageNum) continue;
			c.fillStyle = i === this.searchCurrentIdx
				? 'rgba(255, 140, 0, 0.55)'
				: 'rgba(255, 220, 0, 0.38)';
			c.fillRect(m.x * canvas.width, m.y * canvas.height, m.w * canvas.width, m.h * canvas.height);
		}
	}

	private redrawAllSearchHighlights(): void {
		for (const ctx of this.pages) {
			if (ctx.state === 'rendered') this.drawSearchHighlightsForPage(ctx);
		}
	}

	private clearAllSearchHighlights(): void {
		for (const ctx of this.pages) {
			if (!ctx.searchCanvas) continue;
			const c = ctx.searchCanvas.getContext('2d')!;
			c.clearRect(0, 0, ctx.searchCanvas.width, ctx.searchCanvas.height);
		}
	}

	private async getPageTextItems(pageNum: number): Promise<PdfTextItem[]> {
		if (this._textCache.has(pageNum)) return this._textCache.get(pageNum)!;
		const page = await this.pdfDoc!.getPage(pageNum);
		const tc   = await page.getTextContent();
		const items = tc.items.filter((it): it is PdfTextItem => 'str' in it);
		this._textCache.set(pageNum, items);
		return items;
	}

	// TOC / Outline ----------------------------------------------------------

	private async loadToc(): Promise<void> {
		if (!this.pdfDoc) return;
		try {
			const outline = await this.pdfDoc.getOutline();
			if (!outline || outline.length === 0) return; // no bookmarks — keep TOC button visible but sidebar empty is handled
			// Store outline so toggleToc can build the panel
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
		// Insert before the scroll area
		this.bodyEl.insertBefore(sidebar, this.scrollAreaEl);

		const header = sidebar.createEl('div', { cls: 'via-pdf-toc-header' });
		header.createEl('span', { text: 'Contents' });
		const closeBtn = header.createEl('button', { cls: 'via-btn', text: '✕' });
		closeBtn.addEventListener('click', () => { this.tocVisible = false; sidebar.remove(); this.tocSidebarEl = null; });

		const list = sidebar.createEl('div', { cls: 'via-pdf-toc-list' });
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
					const toggle = row.createEl('span', { cls: 'via-pdf-toc-toggle', text: '▾' });
					let collapsed = false;
					const childList = parentEl.createEl('div');
					renderItems(item.items, childList, depth + 1);

					toggle.addEventListener('click', (e) => {
						e.stopPropagation();
						collapsed = !collapsed;
						childList.style.display = collapsed ? 'none' : '';
						toggle.textContent = collapsed ? '▸' : '▾';
					});
				}

				const label = row.createEl('span', { cls: 'via-pdf-toc-label', text: item.title ?? '(untitled)' });
				label.addEventListener('click', async () => {
					if (!this.pdfDoc) return;
					try {
						let dest = item.dest;
						if (typeof dest === 'string') dest = await this.pdfDoc.getDestination(dest);
						if (!Array.isArray(dest) || dest.length === 0) return;
						const pageRef = dest[0];
						const pageIdx = await this.pdfDoc.getPageIndex(pageRef as PdfRefProxy);
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
			id: `note-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
			page: pageNum,
			x, y,
			text: '',
			color: '#ffd43b',
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
		el.style.cssText = `left:${note.x * 100}%;top:${note.y * 100}%;background:${note.color ?? '#ffd43b'}`;
		el.dataset.noteId = note.id;

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
				note.x = Math.max(0, Math.min(0.9, origLeft + dx));
				note.y = Math.max(0, Math.min(0.95, origTop + dy));
				el.style.left = `${note.x * 100}%`;
				el.style.top  = `${note.y * 100}%`;
			};
			const onUp = () => {
				window.removeEventListener('mousemove', onMove);
				window.removeEventListener('mouseup', onUp);
			};
			window.addEventListener('mousemove', onMove);
			window.addEventListener('mouseup', onUp);
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
		textarea.value = note.text;
		textarea.placeholder = 'Note…';
		textarea.addEventListener('input', () => { note.text = textarea.value; });
		textarea.addEventListener('keydown', e => e.stopPropagation());

		this.noteEls.set(note.id, el);
		if (focusImmediately) textarea.focus();
	}

	// Export -----------------------------------------------------------------

	private async exportAnnotatedPdf(): Promise<void> {
		if (!this.currentFile || !this.pdfDoc) return;
		new Notice('Exporting PDF with annotations…');

		try {
			const srcBuffer = await this.app.vault.adapter.readBinary(this.currentFile.path);
			const pdfLibDoc = await PDFDocument.load(srcBuffer);
			const pages = pdfLibDoc.getPages();

			for (const pa of this.annotData.pages) {
				const pageIdx = pa.page - 1;
				const libPage = pages[pageIdx];
				if (!libPage || pa.paths.length === 0) continue;

				const pjsPage = await this.pdfDoc.getPage(pa.page);
				const vp = pjsPage.getViewport({ scale: 1.0 });
				const pdfW = libPage.getWidth();
				const pdfH = libPage.getHeight();
				const scaleX = pdfW / vp.width;
				const scaleY = pdfH / vp.height;

				for (const path of pa.paths) {
					if (path.tool === 'eraser') continue;
					if (path.points.length < 2) continue;

					const pts = path.points.map(p => ({
						px: p.x * vp.width  * scaleX,
						py: pdfH - p.y * vp.height * scaleY,
					}));

					const hexToRgb = (hex: string) => {
						const n = parseInt(hex.replace('#', ''), 16);
						return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
					};

					const lineWidth = (path.width * scaleX + path.width * scaleY) / 2;
					const opacity = path.tool === 'highlighter' ? (path.opacity ?? 0.35) : 1;

					// Build SVG path string
					let d = `M ${pts[0]!.px.toFixed(2)} ${pts[0]!.py.toFixed(2)}`;
					for (let i = 1; i < pts.length; i++) {
						d += ` L ${pts[i]!.px.toFixed(2)} ${pts[i]!.py.toFixed(2)}`;
					}

					libPage.drawSvgPath(d, {
						borderColor: hexToRgb(path.color),
						borderWidth: lineWidth,
						borderOpacity: opacity,
						borderLineCap: LineCapStyle.Round,
						opacity: 0,
					});
				}
			}

			const exportBytes = await pdfLibDoc.save();
			const exportBuffer = exportBytes.buffer as ArrayBuffer;

			// Save as <basename>.annotated.pdf next to original
			const dir = this.currentFile.parent?.path ?? '';
			const base = this.currentFile.basename;
			const exportPath = dir ? `${dir}/${base}.annotated.pdf` : `${base}.annotated.pdf`;

			const existing = this.app.vault.getAbstractFileByPath(exportPath);
			if (existing && existing instanceof TFile) {
				await this.app.vault.modifyBinary(existing, exportBuffer);
			} else {
				await this.app.vault.createBinary(exportPath, exportBuffer);
			}

			new Notice(`✅ Exported to "${base}.annotated.pdf"`);
		} catch (err) {
			new Notice(`❌ Export failed: ${String(err)}`);
		}
	}
}
