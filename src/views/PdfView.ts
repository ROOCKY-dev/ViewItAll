import { FileView, TFile, WorkspaceLeaf, Notice } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import { VIEW_TYPE_PDF } from '../types';
import type { PageAnnotations, AnnotationPath, AnnotationFile } from '../types';
import {
	loadAnnotations,
	saveAnnotations,
	getPageAnnotations,
	setPageAnnotations,
} from '../utils/pdfAnnotations';
import type ViewItAllPlugin from '../main';

// The worker source is inlined by the esbuild `pdf-worker-inline` plugin.
// We create a Blob URL once so Electron's renderer can spawn the worker thread.
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

type AnnotTool = 'none' | 'pen' | 'highlighter' | 'eraser';

interface PageRenderCtx {
	pageNum: number;
	pdfCanvas: HTMLCanvasElement;
	annotCanvas: HTMLCanvasElement;
	container: HTMLElement;
}

export class PdfView extends FileView {
	private plugin: ViewItAllPlugin;
	private pdfDoc: pdfjsLib.PDFDocumentProxy | null = null;
	private annotData: AnnotationFile = { version: 1, pages: [] };
	private pages: PageRenderCtx[] = [];
	private currentTool: AnnotTool = 'none';
	private isDrawing = false;
	private currentPath: AnnotationPath | null = null;
	private currentFile: TFile | null = null;

	// Zoom — stored as a scale factor (1.0 = 100%)
	private currentScale = 1.0;
	private readonly ZOOM_STEPS = [0.25, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0, 3.0, 4.0];
	private scrollAreaEl: HTMLElement | null = null;
	private zoomLabelEl: HTMLElement | null = null;

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
		pdfjsLib.GlobalWorkerOptions.workerSrc = getPdfWorkerUrl();
	}

	getViewType(): string { return VIEW_TYPE_PDF; }
	getDisplayText(): string { return this.file?.basename ?? 'PDF'; }
	getIcon(): string { return 'file'; }

	canAcceptExtension(extension: string): boolean {
		return extension === 'pdf';
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.currentFile = file;
		this.currentTool = this.plugin.settings.pdfDefaultTool;
		this.annotData = await loadAnnotations(this.app, file);
		await this.renderPdf(file);
	}

	async onUnloadFile(_file: TFile): Promise<void> {
		if (this.pdfDoc) { this.pdfDoc.destroy(); this.pdfDoc = null; }
		this.pages = [];
		this.contentEl.empty();
	}

	// ── Render ──────────────────────────────────────────────────────────────

	private async renderPdf(file: TFile): Promise<void> {
		this.contentEl.empty();
		this.pages = [];
		this.scrollAreaEl = null;
		this.zoomLabelEl = null;

		// Flex wrapper fills the full leaf height
		const wrapper = this.contentEl.createEl('div', { cls: 'via-pdf-wrapper' });

		const toolbar = this.buildToolbar();
		wrapper.appendChild(toolbar);

		const scrollArea = wrapper.createEl('div', { cls: 'via-pdf-scroll' });
		this.scrollAreaEl = scrollArea;

		// Loading indicator — shown while PDF is parsed and pages are rendered
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

		const loadingTask = pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
		this.pdfDoc = await loadingTask.promise;

		loadingEl.remove();

		for (let i = 1; i <= this.pdfDoc.numPages; i++) {
			const ctx = await this.renderPage(i, scrollArea);
			this.pages.push(ctx);
			this.redrawAnnotations(ctx);
			this.attachDrawListeners(ctx);
		}
	}

	private async renderPage(pageNum: number, container: HTMLElement): Promise<PageRenderCtx> {
		const page = await this.pdfDoc!.getPage(pageNum);
		const viewport = page.getViewport({ scale: this.currentScale });

		const w = Math.ceil(viewport.width);
		const h = Math.ceil(viewport.height);

		// Explicit block sizing so flex doesn't collapse the page
		const pageWrap = container.createEl('div', { cls: 'via-pdf-page' });
		pageWrap.style.width  = `${w}px`;
		pageWrap.style.height = `${h}px`;
		pageWrap.style.minWidth = `${w}px`;
		pageWrap.style.minHeight = `${h}px`;

		const pdfCanvas = pageWrap.createEl('canvas', { cls: 'via-pdf-canvas' });
		pdfCanvas.width  = w;
		pdfCanvas.height = h;

		const annotCanvas = pageWrap.createEl('canvas', { cls: 'via-pdf-annot-canvas' });
		annotCanvas.width  = w;
		annotCanvas.height = h;

		const renderCtx = {
			canvasContext: pdfCanvas.getContext('2d')!,
			viewport,
		};
		await page.render(renderCtx).promise;

		return { pageNum, pdfCanvas, annotCanvas, container: pageWrap };
	}

	// ── Toolbar ─────────────────────────────────────────────────────────────

	private buildToolbar(): HTMLElement {
		const bar = document.createElement('div');
		bar.className = 'via-pdf-toolbar';

		// Annotation tools
		const tools: { id: AnnotTool; label: string }[] = [
			{ id: 'none', label: '👁 View' },
			{ id: 'pen', label: '✏️ Pen' },
			{ id: 'highlighter', label: '🖊 Highlight' },
			{ id: 'eraser', label: '⬜ Erase' },
		];

		for (const t of tools) {
			const btn = bar.createEl('button', { cls: 'via-btn', text: t.label });
			btn.dataset.tool = t.id;
			if (t.id === this.currentTool) btn.classList.add('via-btn-active');
			btn.addEventListener('click', () => {
				this.currentTool = t.id;
				bar.querySelectorAll('[data-tool]').forEach(b =>
					b.classList.toggle('via-btn-active', (b as HTMLElement).dataset.tool === t.id)
				);
				this.updateCanvasInteraction();
			});
		}

		bar.createEl('div', { cls: 'via-toolbar-sep' });

		// Zoom controls
		const zoomOut = bar.createEl('button', { cls: 'via-btn via-btn-zoom', text: '−' });
		zoomOut.title = 'Zoom out';
		zoomOut.addEventListener('click', () => this.stepZoom(-1));

		this.zoomLabelEl = bar.createEl('button', {
			cls: 'via-btn via-btn-zoom-label',
			text: `${Math.round(this.currentScale * 100)}%`,
		});
		this.zoomLabelEl.title = 'Reset to 100%';
		this.zoomLabelEl.addEventListener('click', () => this.setZoom(1.0));

		const zoomIn = bar.createEl('button', { cls: 'via-btn via-btn-zoom', text: '+' });
		zoomIn.title = 'Zoom in';
		zoomIn.addEventListener('click', () => this.stepZoom(+1));

		bar.createEl('div', { cls: 'via-toolbar-sep' });

		const clearBtn = bar.createEl('button', { cls: 'via-btn via-btn-danger', text: '🗑 Clear page' });
		clearBtn.addEventListener('click', () => this.clearCurrentPageAnnotations());

		const saveBtn = bar.createEl('button', { cls: 'via-btn via-btn-save', text: '💾 Save annotations' });
		saveBtn.addEventListener('click', () => this.persistAnnotations());

		return bar;
	}

	// ── Zoom ────────────────────────────────────────────────────────────────

	private stepZoom(direction: -1 | 1): void {
		const idx = this.ZOOM_STEPS.findIndex(s => Math.abs(s - this.currentScale) < 0.01);
		const next = this.ZOOM_STEPS[Math.max(0, Math.min(this.ZOOM_STEPS.length - 1, idx + direction))];
		if (next !== undefined) this.setZoom(next);
	}

	private setZoom(scale: number): void {
		this.currentScale = scale;
		if (this.zoomLabelEl) this.zoomLabelEl.textContent = `${Math.round(scale * 100)}%`;
		this.reRenderPages();
	}

	private async reRenderPages(): Promise<void> {
		if (!this.pdfDoc || !this.scrollAreaEl) return;
		for (const ctx of this.pages) ctx.container.remove();
		this.pages = [];
		for (let i = 1; i <= this.pdfDoc.numPages; i++) {
			const ctx = await this.renderPage(i, this.scrollAreaEl);
			this.pages.push(ctx);
			this.redrawAnnotations(ctx);
			this.attachDrawListeners(ctx);
		}
		this.updateCanvasInteraction();
	}

	private updateCanvasInteraction(): void {
		for (const ctx of this.pages) {
			const drawing = this.currentTool !== 'none';
			ctx.annotCanvas.style.pointerEvents = drawing ? 'auto' : 'none';
			ctx.annotCanvas.style.cursor = drawing ? 'crosshair' : 'default';
		}
	}

	// ── Drawing ─────────────────────────────────────────────────────────────

	private attachDrawListeners(ctx: PageRenderCtx): void {
		const { annotCanvas, pageNum } = ctx;

		// Coordinates are stored **normalized** (0–1 relative to canvas size)
		// so they remain valid when the canvas is re-rendered at a different zoom level.
		const getPos = (e: MouseEvent | PointerEvent) => {
			const rect = annotCanvas.getBoundingClientRect();
			return {
				x: (e.clientX - rect.left) / rect.width,
				y: (e.clientY - rect.top) / rect.height,
			};
		};

		annotCanvas.addEventListener('pointerdown', e => {
			if (this.currentTool === 'none') return;
			annotCanvas.setPointerCapture(e.pointerId);
			this.isDrawing = true;
			const pos = getPos(e);
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
						: 20,
				points: [pos],
			};
		});

		annotCanvas.addEventListener('pointermove', e => {
			if (!this.isDrawing || !this.currentPath) return;
			const pos = getPos(e);
			this.currentPath.points.push(pos);
			this.redrawAnnotations(ctx, this.currentPath);
		});

		const finishDraw = () => {
			if (!this.isDrawing || !this.currentPath) return;
			this.isDrawing = false;
			let pa = getPageAnnotations(this.annotData, pageNum);
			pa = { ...pa, paths: [...pa.paths, this.currentPath!] };
			this.annotData = setPageAnnotations(this.annotData, pa);
			this.currentPath = null;
			this.redrawAnnotations(ctx);
		};

		annotCanvas.addEventListener('pointerup', finishDraw);
		annotCanvas.addEventListener('pointercancel', finishDraw);
	}

	// ── Annotation rendering ─────────────────────────────────────────────────

	private redrawAnnotations(ctx: PageRenderCtx, inProgressPath?: AnnotationPath): void {
		const canvas = ctx.annotCanvas;
		const c = canvas.getContext('2d')!;
		c.clearRect(0, 0, canvas.width, canvas.height);

		const pa: PageAnnotations = getPageAnnotations(this.annotData, ctx.pageNum);

		const drawPath = (path: AnnotationPath) => {
			if (path.points.length < 2) return;
			const w = canvas.width;
			const h = canvas.height;
			c.save();
			if (path.tool === 'highlighter') {
				c.globalAlpha = 0.35;
				c.globalCompositeOperation = 'multiply';
			} else if (path.tool === 'eraser') {
				c.globalCompositeOperation = 'destination-out';
				c.globalAlpha = 1;
			} else {
				c.globalAlpha = 1;
				c.globalCompositeOperation = 'source-over';
			}
			c.strokeStyle = path.color;
			// Line width also needs to scale with the canvas (stored relative to 100% zoom)
			c.lineWidth = path.width * this.currentScale;
			c.lineCap = 'round';
			c.lineJoin = 'round';
			c.beginPath();
			const p0 = path.points[0]!;
			c.moveTo(p0.x * w, p0.y * h);
			for (let i = 1; i < path.points.length; i++) {
				const pi = path.points[i]!;
				c.lineTo(pi.x * w, pi.y * h);
			}
			c.stroke();
			c.restore();
		};

		for (const path of pa.paths) drawPath(path);
		if (inProgressPath) drawPath(inProgressPath);
	}

	// ── Persistence ──────────────────────────────────────────────────────────

	private clearCurrentPageAnnotations(): void {
		// Clears the page currently visible (page 1 as proxy for now; improve with intersection observer)
		const visiblePage = this.getVisiblePageNum();
		const pa: PageAnnotations = { page: visiblePage, paths: [] };
		this.annotData = setPageAnnotations(this.annotData, pa);
		const ctx = this.pages.find(p => p.pageNum === visiblePage);
		if (ctx) this.redrawAnnotations(ctx);
	}

	private getVisiblePageNum(): number {
		// Find the page whose container is most visible in the scroll area
		let best = 1;
		let bestVisible = -Infinity;
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
			new Notice('✅ Annotations saved');
		} catch (err) {
			new Notice(`❌ Failed to save annotations: ${String(err)}`);
		}
	}
}
