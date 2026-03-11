import { FileView, Notice, TFile, WorkspaceLeaf, setIcon, setTooltip } from 'obsidian';
import { VIEW_TYPE_PPTX } from '../types';
import type ViewItAllPlugin from '../main';
import { parseSlidesFromZip } from './pptx/pptxParser';
import type { JSZipConstructor, RunData, SlideData } from './pptx/pptxTypes';
import { loadPptxEdits, savePptxEdits } from '../utils/pptxEdits';
import { applyEditsToPptxZip } from './pptx/pptxWriter';

type DragMode = 'move' | 'resize';
type ResizeHandle = 'nw' | 'ne' | 'sw' | 'se';

interface ShapeEditState {
	translateX: number;
	translateY: number;
	widthPx: number | null;
	heightPx: number | null;
}

interface ShapeFormatState {
	fillToken: string | null;
	lineToken: string | null;
}

interface DragState {
	shapeKey: string;
	mode: DragMode;
	startX: number;
	startY: number;
	startTranslateX: number;
	startTranslateY: number;
	startWidth: number;
	startHeight: number;
	handle: ResizeHandle | null;
}

export class PptxView extends FileView {
	private plugin: ViewItAllPlugin;
	private currentFile: TFile | null = null;
	private slides: SlideData[] = [];
	private activeSlide = 0;

	// Zoom & panel state
	private zoomLevel = 1.0;
	private stripVisible = true;
	private isFullscreen = false;
	private readonly ZOOM_STEP = 0.2;
	private readonly ZOOM_MIN = 0.3;
	private readonly ZOOM_MAX = 2.0;

	// Editing state
	private editMode = true;
	private selectedShapeKey: string | null = null;
	private shapeEdits = new Map<string, ShapeEditState>();
	private shapeFormats = new Map<string, ShapeFormatState>();
	private shapeTextOverrides = new Map<string, string>();
	private activeDrag: DragState | null = null;
	private fillSelectEl: HTMLSelectElement | null = null;
	private strokeSelectEl: HTMLSelectElement | null = null;

	// DOM refs
	private wrapper: HTMLElement | null = null;
	private slideContainer: HTMLElement | null = null;
	private slideStrip: HTMLElement | null = null;
	private slideCounter: HTMLElement | null = null;

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
	}

	getViewType(): string { return VIEW_TYPE_PPTX; }
	getDisplayText(): string { return this.file?.basename ?? 'Presentation'; }
	getIcon(): string { return 'presentation'; }

	canAcceptExtension(extension: string): boolean {
		return extension === 'pptx' && this.plugin.settings.enablePptx;
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.currentFile = file;
		this.activeSlide = 0;
		this.slides = [];
		await this.loadDraftEdits(file);
		await this.renderFile(file);
	}

	async onUnloadFile(_file: TFile): Promise<void> {
		this.contentEl.empty();
		this.slides = [];
		this.wrapper = null;
		this.slideContainer = null;
		this.slideStrip = null;
		this.slideCounter = null;
		this.currentFile = null;
		this.zoomLevel = 1.0;
		this.stripVisible = true;
		this.isFullscreen = false;
		this.selectedShapeKey = null;
		this.activeDrag = null;
		this.fillSelectEl = null;
		this.strokeSelectEl = null;
		this.shapeEdits.clear();
		this.shapeFormats.clear();
		this.shapeTextOverrides.clear();
	}

	private async renderFile(file: TFile): Promise<void> {
		this.contentEl.empty();

		const isBottom = this.plugin.settings.pptxToolbarPosition === 'bottom';

		// Read file
		let data: ArrayBuffer;
		try {
			data = await this.app.vault.adapter.readBinary(file.path);
		} catch (err) {
			this.contentEl.createEl('p', {
				cls: 'via-error',
				text: `Failed to read file: ${String(err)}`,
			});
			return;
		}

		// Show loading
		const loading = this.contentEl.createEl('div', { cls: 'via-pdf-loading' });
		loading.createEl('div', { cls: 'via-pdf-loading-spinner' });
		loading.createEl('span', { text: 'Parsing presentation...' });

		// Lazy-load JSZip + parse
		try {
			const JSZip = ((await import('jszip')) as unknown as { default: JSZipConstructor }).default;
			const zip = await JSZip.loadAsync(data);
			this.slides = await parseSlidesFromZip(zip);
		} catch (err) {
			loading.remove();
			this.contentEl.createEl('p', {
				cls: 'via-error',
				text: `Failed to parse PPTX: ${String(err)}`,
			});
			return;
		}

		loading.remove();

		if (this.slides.length === 0) {
			this.contentEl.createEl('p', {
				cls: 'via-sheet-empty',
				text: 'This presentation has no slides.',
			});
			return;
		}

		// Wrapper
		this.wrapper = this.contentEl.createEl('div', { cls: 'via-pptx-wrapper' });
		if (isBottom) this.wrapper.classList.add('via-pptx-wrapper--toolbar-bottom');

		// Toolbar
		const toolbar = this.wrapper.createEl('div', { cls: 'via-pptx-toolbar' });

		// Strip panel toggle
		const stripToggle = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(stripToggle, 'panel-left');
		setTooltip(stripToggle, 'Toggle slide panel');
		if (this.stripVisible) stripToggle.classList.add('is-active');
		stripToggle.addEventListener('click', () => {
			this.stripVisible = !this.stripVisible;
			this.slideStrip?.classList.toggle('is-hidden', !this.stripVisible);
			stripToggle.classList.toggle('is-active', this.stripVisible);
		});

		toolbar.createEl('div', { cls: 'via-toolbar-sep' });

		// First slide
		const firstBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(firstBtn, 'chevrons-left');
		setTooltip(firstBtn, 'First slide');
		firstBtn.addEventListener('click', () => this.goToSlide(0));

		// Prev
		const prevBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(prevBtn, 'chevron-left');
		setTooltip(prevBtn, 'Previous slide');
		prevBtn.addEventListener('click', () => this.goToSlide(this.activeSlide - 1));

		// Slide counter
		this.slideCounter = toolbar.createEl('div', { cls: 'via-pptx-counter' });
		this.updateCounter();

		// Next
		const nextBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(nextBtn, 'chevron-right');
		setTooltip(nextBtn, 'Next slide');
		nextBtn.addEventListener('click', () => this.goToSlide(this.activeSlide + 1));

		// Last slide
		const lastBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(lastBtn, 'chevrons-right');
		setTooltip(lastBtn, 'Last slide');
		lastBtn.addEventListener('click', () => this.goToSlide(this.slides.length - 1));

		toolbar.createEl('div', { cls: 'via-toolbar-sep' });

		// File info
		const fileLabel = toolbar.createEl('div', { cls: 'via-pptx-file-label' });
		const fileIcon = fileLabel.createEl('div', { cls: 'clickable-icon' });
		setIcon(fileIcon, 'presentation');
		fileLabel.createEl('span', {
			text: file.basename,
			cls: 'via-pptx-file-name',
		});

		toolbar.createEl('div', { cls: 'via-toolbar-spacer' });

		// Slide count info
		toolbar.createEl('div', {
			cls: 'via-pptx-info',
			text: `${this.slides.length} slides`,
		});

		toolbar.createEl('div', { cls: 'via-toolbar-sep' });

		// Fullscreen toggle
		const fullscreenBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(fullscreenBtn, 'expand');
		setTooltip(fullscreenBtn, 'Fullscreen slide');
		fullscreenBtn.addEventListener('click', () => {
			this.isFullscreen = !this.isFullscreen;
			this.wrapper?.classList.toggle('via-pptx-fullscreen', this.isFullscreen);
			this.slideStrip?.classList.toggle('is-hidden', this.isFullscreen || !this.stripVisible);
			setIcon(fullscreenBtn, this.isFullscreen ? 'shrink' : 'expand');
			setTooltip(fullscreenBtn, this.isFullscreen ? 'Exit fullscreen' : 'Fullscreen slide');
			fullscreenBtn.classList.toggle('is-active', this.isFullscreen);
		});

			toolbar.createEl('div', { cls: 'via-toolbar-sep' });

			const editToggleBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
			setIcon(editToggleBtn, 'mouse-pointer-2');
			setTooltip(editToggleBtn, this.editMode ? 'Disable shape edit mode' : 'Enable shape edit mode');
			editToggleBtn.classList.toggle('is-active', this.editMode);
			editToggleBtn.addEventListener('click', () => {
				this.editMode = !this.editMode;
				if (!this.editMode) {
					this.selectedShapeKey = null;
					this.activeDrag = null;
				}
				editToggleBtn.classList.toggle('is-active', this.editMode);
				setTooltip(editToggleBtn, this.editMode ? 'Disable shape edit mode' : 'Enable shape edit mode');
				this.renderSlide();
				this.updateFormatControls();
			});

			toolbar.createEl('div', { cls: 'via-toolbar-sep' });

			const fillWrap = toolbar.createEl('div', { cls: 'via-pptx-format-control' });
			fillWrap.createEl('span', { cls: 'via-pptx-format-label', text: 'Fill' });
			this.fillSelectEl = fillWrap.createEl('select', { cls: 'via-pptx-format-select' });
			this.fillSelectEl.createEl('option', { value: '', text: 'From file' });
			this.fillSelectEl.createEl('option', { value: 'accent', text: 'Accent' });
			this.fillSelectEl.createEl('option', { value: 'muted', text: 'Muted' });
			this.fillSelectEl.createEl('option', { value: 'none', text: 'None' });
			this.fillSelectEl.addEventListener('change', () => this.applySelectedShapeFormatting());

			const strokeWrap = toolbar.createEl('div', { cls: 'via-pptx-format-control' });
			strokeWrap.createEl('span', { cls: 'via-pptx-format-label', text: 'Stroke' });
			this.strokeSelectEl = strokeWrap.createEl('select', { cls: 'via-pptx-format-select' });
			this.strokeSelectEl.createEl('option', { value: '', text: 'From file' });
			this.strokeSelectEl.createEl('option', { value: 'accent', text: 'Accent' });
			this.strokeSelectEl.createEl('option', { value: 'muted', text: 'Muted' });
			this.strokeSelectEl.createEl('option', { value: 'normal', text: 'Normal' });
			this.strokeSelectEl.createEl('option', { value: 'none', text: 'None' });
			this.strokeSelectEl.addEventListener('change', () => this.applySelectedShapeFormatting());

			const boldBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
			boldBtn.setText('B');
			setTooltip(boldBtn, 'Bold selected text');
			boldBtn.addEventListener('click', () => document.execCommand('bold'));

			const italicBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
			italicBtn.setText('I');
			setTooltip(italicBtn, 'Italic selected text');
			italicBtn.addEventListener('click', () => document.execCommand('italic'));

			const underlineBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
			underlineBtn.setText('U');
			setTooltip(underlineBtn, 'Underline selected text');
			underlineBtn.addEventListener('click', () => document.execCommand('underline'));

			const saveDraftBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
			setIcon(saveDraftBtn, 'save');
			setTooltip(saveDraftBtn, 'Save draft edits');
			saveDraftBtn.addEventListener('click', async () => {
				if (!this.currentFile) return;
				await this.saveDraftEdits(this.currentFile);
			});

			const applyBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
			setIcon(applyBtn, 'file-check');
			setTooltip(applyBtn, 'Apply draft edits to PPTX');
			applyBtn.addEventListener('click', async () => {
				if (!this.currentFile) return;
				await this.applyDraftEditsToPptx(this.currentFile);
			});

		toolbar.createEl('div', { cls: 'via-toolbar-sep' });

		// Zoom out
		const zoomOutBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(zoomOutBtn, 'zoom-out');
		setTooltip(zoomOutBtn, 'Zoom out');
		zoomOutBtn.addEventListener('click', () => {
			this.zoomLevel = Math.max(this.ZOOM_MIN, this.zoomLevel - this.ZOOM_STEP);
			this.applyZoom();
		});

		// Zoom reset
		const zoomResetBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(zoomResetBtn, 'maximize-2');
		setTooltip(zoomResetBtn, 'Reset zoom');
		zoomResetBtn.addEventListener('click', () => {
			this.zoomLevel = 1.0;
			this.applyZoom();
		});

		// Zoom in
		const zoomInBtn = toolbar.createEl('div', { cls: 'clickable-icon' });
		setIcon(zoomInBtn, 'zoom-in');
		setTooltip(zoomInBtn, 'Zoom in');
		zoomInBtn.addEventListener('click', () => {
			this.zoomLevel = Math.min(this.ZOOM_MAX, this.zoomLevel + this.ZOOM_STEP);
			this.applyZoom();
		});

		// Body: slide strip + main slide
		const body = this.wrapper.createEl('div', { cls: 'via-pptx-body' });

		// Slide strip (thumbnail sidebar)
		this.slideStrip = body.createEl('div', { cls: 'via-pptx-strip' });
		this.renderStrip();

		// Main slide area
		const scrollEl = body.createEl('div', { cls: 'via-pptx-scroll' });
		this.slideContainer = scrollEl.createEl('div', { cls: 'via-pptx-slide' });
		this.renderSlide();

		this.registerDomEvent(window, 'pointermove', (e: PointerEvent) => this.onPointerMove(e));
		this.registerDomEvent(window, 'pointerup', () => this.onPointerUp());

		// Keyboard navigation
		this.registerDomEvent(this.contentEl, 'keydown', (e: KeyboardEvent) => {
			if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
				e.preventDefault();
				this.goToSlide(this.activeSlide + 1);
			} else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
				e.preventDefault();
				this.goToSlide(this.activeSlide - 1);
			} else if (e.key === 'Home') {
				e.preventDefault();
				this.goToSlide(0);
			} else if (e.key === 'End') {
				e.preventDefault();
				this.goToSlide(this.slides.length - 1);
			}
		});
		// Make focusable for keyboard events
		this.contentEl.tabIndex = 0;
	}

	private applyZoom(): void {
		if (!this.slideContainer) return;
		this.slideContainer.style.transform = `scale(${this.zoomLevel})`;
	}

	private goToSlide(index: number): void {
		if (index < 0 || index >= this.slides.length) return;
		this.activeSlide = index;
		this.renderSlide();
		this.updateCounter();
		this.updateStripActive();
		// Scroll strip thumbnail into view
		const thumb = this.slideStrip?.children[index] as HTMLElement | undefined;
		if (thumb) thumb.scrollIntoView({ block: 'nearest' });
	}

	private updateCounter(): void {
		if (!this.slideCounter) return;
		this.slideCounter.textContent = `${this.activeSlide + 1} / ${this.slides.length}`;
	}

	private renderStrip(): void {
		if (!this.slideStrip) return;
		this.slideStrip.empty();

		for (let i = 0; i < this.slides.length; i++) {
			const thumb = this.slideStrip.createEl('div', { cls: 'via-pptx-thumb' });
			if (i === this.activeSlide) thumb.classList.add('is-active');

			const numEl = thumb.createEl('div', { cls: 'via-pptx-thumb-num' });
			numEl.textContent = String(i + 1);

			const preview = thumb.createEl('div', { cls: 'via-pptx-thumb-preview' });
			// Show first meaningful text as mini preview
			const slide = this.slides[i];
			const firstText = this.getSlidePreviewText(slide);
			preview.createEl('span', {
				text: firstText.slice(0, 60) + (firstText.length > 60 ? '…' : ''),
			});

			thumb.addEventListener('click', () => this.goToSlide(i));
		}
	}

	/** Get a short text preview for the strip thumbnail. */
	private getSlidePreviewText(slide: SlideData | undefined): string {
		if (!slide) return '';
		for (const shape of slide.shapes) {
			for (const para of shape.paragraphs) {
				const text = para.runs.map(r => r.text).join('');
				if (text.trim()) return text;
			}
		}
		return '';
	}

	private updateStripActive(): void {
		if (!this.slideStrip) return;
		const thumbs = this.slideStrip.children;
		for (let i = 0; i < thumbs.length; i++) {
			const el = thumbs[i];
			if (el) el.classList.toggle('is-active', i === this.activeSlide);
		}
	}

	private renderSlide(): void {
		if (!this.slideContainer) return;
		this.slideContainer.empty();

		const slide = this.slides[this.activeSlide];
		if (!slide) return;

		// Images
		for (const dataUrl of slide.imageDataUrls) {
			this.slideContainer.createEl('img', {
				cls: 'via-pptx-slide-img',
				attr: { src: dataUrl },
			});
		}

		// Shapes — ordered by type priority: title/ctrTitle first, then subtitle, then body, then other
		const typeOrder: Record<string, number> = { ctrTitle: 0, title: 1, subTitle: 2, body: 3, other: 4 };
		const sorted = [...slide.shapes].sort(
			(a, b) => (typeOrder[a.type] ?? 4) - (typeOrder[b.type] ?? 4)
		);

		for (const shape of sorted) {
			const hasContent = shape.paragraphs.some(p => p.runs.some(r => r.text.trim()));
			if (!hasContent) continue;

			const shapeKey = this.makeShapeKey(slide.index, shape.id);
			const shapeEl = this.slideContainer.createEl('div', { cls: `via-pptx-shape via-pptx-shape--${shape.type}` });
			if (this.editMode) shapeEl.classList.add('via-pptx-shape--editable');
			shapeEl.setAttribute('data-shape-id', shape.id);
			shapeEl.setAttribute('data-shape-key', shapeKey);
			shapeEl.setAttribute('data-z-index', String(shape.zIndex));
			if (shape.bounds) {
				shapeEl.setAttribute('data-x-emu', String(shape.bounds.xEmu));
				shapeEl.setAttribute('data-y-emu', String(shape.bounds.yEmu));
				shapeEl.setAttribute('data-width-emu', String(shape.bounds.widthEmu));
				shapeEl.setAttribute('data-height-emu', String(shape.bounds.heightEmu));
				shapeEl.setAttribute('data-rotation-deg', String(shape.bounds.rotationDeg));
			}
			if (shape.style.fillToken) shapeEl.setAttribute('data-fill-token', shape.style.fillToken);
			if (shape.style.lineToken) shapeEl.setAttribute('data-line-token', shape.style.lineToken);
			if (shape.style.lineDash) shapeEl.setAttribute('data-line-dash', shape.style.lineDash);
			if (shape.style.lineWidth !== null) shapeEl.setAttribute('data-line-width', String(shape.style.lineWidth));

			this.applyShapeEditStyle(shapeEl, shapeKey);
			this.applyShapeFormatStyle(shapeEl, shapeKey, shape.style.fillToken, shape.style.lineToken);
			shapeEl.classList.toggle('is-selected', this.selectedShapeKey === shapeKey);
			shapeEl.contentEditable = this.editMode && this.selectedShapeKey === shapeKey ? 'true' : 'false';

			shapeEl.addEventListener('input', () => this.captureShapeText(shapeEl, shapeKey));

			if (this.editMode) {
				shapeEl.addEventListener('pointerdown', (event: PointerEvent) => this.onShapePointerDown(event, shapeKey));
			}

			const textOverride = this.shapeTextOverrides.get(shapeKey);
			if (textOverride !== undefined) {
				this.renderTextOverride(shapeEl, textOverride);
				if (this.editMode) this.ensureResizeHandles(shapeEl);
				continue;
			}

			// Check if the entire shape is a bullet list
			const isList = shape.paragraphs.length > 1 && shape.paragraphs.some(p => p.isBullet);

			if (isList) {
				// Render as a list: non-bullet paragraphs as regular text, bullet items as <li>
				let currentList: HTMLElement | null = null;
				for (const para of shape.paragraphs) {
					const text = para.runs.map(r => r.text).join('');
					if (!text.trim()) {
						currentList = null;
						continue;
					}
					if (para.isBullet) {
						if (!currentList) {
							currentList = shapeEl.createEl('ul', { cls: 'via-pptx-list' });
						}
						const li = currentList.createEl('li');
						this.renderRuns(li, para.runs);
					} else {
						currentList = null;
						const p = shapeEl.createEl('p');
						this.renderRuns(p, para.runs);
					}
				}
			} else {
				// Render paragraphs as <p> elements
				for (const para of shape.paragraphs) {
					const text = para.runs.map(r => r.text).join('');
					if (!text.trim()) continue;
					const p = shapeEl.createEl('p');
					this.renderRuns(p, para.runs);
				}
			}

			if (this.editMode) this.ensureResizeHandles(shapeEl);
		}

		if (slide.shapes.length === 0 && slide.imageDataUrls.length === 0) {
			this.slideContainer.createEl('div', {
				cls: 'via-pptx-slide-empty',
				text: '(empty slide)',
			});
		}
	}

	private makeShapeKey(slideIndex: number, shapeId: string): string {
		return `${slideIndex}:${shapeId}`;
	}

	private getShapeEditState(shapeKey: string): ShapeEditState {
		const existing = this.shapeEdits.get(shapeKey);
		if (existing) return existing;
		const initial: ShapeEditState = {
			translateX: 0,
			translateY: 0,
			widthPx: null,
			heightPx: null,
		};
		this.shapeEdits.set(shapeKey, initial);
		return initial;
	}

	private applyShapeEditStyle(shapeEl: HTMLElement, shapeKey: string): void {
		const state = this.getShapeEditState(shapeKey);
		shapeEl.style.transform = `translate(${state.translateX}px, ${state.translateY}px)`;
		if (state.widthPx !== null) {
			shapeEl.style.width = `${Math.max(48, state.widthPx)}px`;
		} else {
			shapeEl.style.removeProperty('width');
		}
		if (state.heightPx !== null) {
			shapeEl.style.minHeight = `${Math.max(24, state.heightPx)}px`;
		} else {
			shapeEl.style.removeProperty('min-height');
		}
	}

	private ensureResizeHandles(shapeEl: HTMLElement): void {
		const handles: ResizeHandle[] = ['nw', 'ne', 'sw', 'se'];
		for (const handle of handles) {
			const handleEl = shapeEl.createEl('div', { cls: `via-pptx-shape-handle via-pptx-shape-handle--${handle}` });
			handleEl.setAttribute('data-resize-handle', handle);
			handleEl.addEventListener('pointerdown', (event: PointerEvent) => {
				event.preventDefault();
				event.stopPropagation();
				const shapeKey = shapeEl.getAttribute('data-shape-key');
				if (!shapeKey) return;
				this.beginResize(shapeEl, shapeKey, handle, event);
			});
		}
	}

	private onShapePointerDown(event: PointerEvent, shapeKey: string): void {
		if (!this.editMode) return;
		const target = event.target;
		if (!(target instanceof HTMLElement)) return;
		if (target.hasAttribute('data-resize-handle')) return;

		const shapeEl = target.closest('[data-shape-key]');
		if (!(shapeEl instanceof HTMLElement)) return;

		this.selectedShapeKey = shapeKey;
		this.updateSelectionClasses();
		this.updateFormatControls();

		if (event.detail > 1) {
			shapeEl.contentEditable = 'true';
			shapeEl.focus();
			return;
		}

		event.preventDefault();
		const state = this.getShapeEditState(shapeKey);
		this.activeDrag = {
			shapeKey,
			mode: 'move',
			startX: event.clientX,
			startY: event.clientY,
			startTranslateX: state.translateX,
			startTranslateY: state.translateY,
			startWidth: shapeEl.getBoundingClientRect().width,
			startHeight: shapeEl.getBoundingClientRect().height,
			handle: null,
		};
	}

	private beginResize(shapeEl: HTMLElement, shapeKey: string, handle: ResizeHandle, event: PointerEvent): void {
		if (!this.editMode) return;
		this.selectedShapeKey = shapeKey;
		this.updateSelectionClasses();
		this.updateFormatControls();

		const state = this.getShapeEditState(shapeKey);
		const rect = shapeEl.getBoundingClientRect();
		this.activeDrag = {
			shapeKey,
			mode: 'resize',
			startX: event.clientX,
			startY: event.clientY,
			startTranslateX: state.translateX,
			startTranslateY: state.translateY,
			startWidth: rect.width,
			startHeight: rect.height,
			handle,
		};
	}

	private onPointerMove(event: PointerEvent): void {
		if (!this.activeDrag) return;
		const state = this.shapeEdits.get(this.activeDrag.shapeKey);
		if (!state) return;

		const dx = event.clientX - this.activeDrag.startX;
		const dy = event.clientY - this.activeDrag.startY;

		if (this.activeDrag.mode === 'move') {
			state.translateX = this.activeDrag.startTranslateX + dx;
			state.translateY = this.activeDrag.startTranslateY + dy;
		}

		if (this.activeDrag.mode === 'resize' && this.activeDrag.handle) {
			const horizontalSign = this.activeDrag.handle.endsWith('w') ? -1 : 1;
			const verticalSign = this.activeDrag.handle.startsWith('n') ? -1 : 1;

			const widthPx = Math.max(48, this.activeDrag.startWidth + dx * horizontalSign);
			const heightPx = Math.max(24, this.activeDrag.startHeight + dy * verticalSign);
			state.widthPx = widthPx;
			state.heightPx = heightPx;
		}

		this.applyActiveShapeState();
	}

	private onPointerUp(): void {
		if (!this.activeDrag) return;
		this.activeDrag = null;
	}

	private applyActiveShapeState(): void {
		if (!this.activeDrag) return;
		const shapeEl = this.slideContainer?.querySelector(`[data-shape-key="${this.activeDrag.shapeKey}"]`);
		if (!(shapeEl instanceof HTMLElement)) return;
		this.applyShapeEditStyle(shapeEl, this.activeDrag.shapeKey);
	}

	private applyShapeFormatStyle(
		shapeEl: HTMLElement,
		shapeKey: string,
		defaultFillToken: string | null,
		defaultLineToken: string | null
	): void {
		const override = this.shapeFormats.get(shapeKey);
		const fillToken = override?.fillToken ?? defaultFillToken;
		const lineToken = override?.lineToken ?? defaultLineToken;

		if (fillToken === 'accent') {
			shapeEl.setAttribute('data-fill-style', 'accent');
		} else if (fillToken === 'muted') {
			shapeEl.setAttribute('data-fill-style', 'muted');
		} else if (fillToken === 'none') {
			shapeEl.setAttribute('data-fill-style', 'none');
		} else {
			shapeEl.removeAttribute('data-fill-style');
		}

		if (lineToken === 'accent') {
			shapeEl.setAttribute('data-stroke-style', 'accent');
		} else if (lineToken === 'muted') {
			shapeEl.setAttribute('data-stroke-style', 'muted');
		} else if (lineToken === 'normal') {
			shapeEl.setAttribute('data-stroke-style', 'normal');
		} else if (lineToken === 'none') {
			shapeEl.setAttribute('data-stroke-style', 'none');
		} else {
			shapeEl.removeAttribute('data-stroke-style');
		}
	}

	private applySelectedShapeFormatting(): void {
		if (!this.selectedShapeKey) return;
		const fillToken = this.fillSelectEl?.value ?? '';
		const lineToken = this.strokeSelectEl?.value ?? '';

		this.shapeFormats.set(this.selectedShapeKey, {
			fillToken: fillToken === '' ? null : fillToken,
			lineToken: lineToken === '' ? null : lineToken,
		});

		const shapeEl = this.slideContainer?.querySelector(`[data-shape-key="${this.selectedShapeKey}"]`);
		if (!(shapeEl instanceof HTMLElement)) return;
		this.applyShapeFormatStyle(shapeEl, this.selectedShapeKey, null, null);
	}

	private updateFormatControls(): void {
		if (!this.fillSelectEl || !this.strokeSelectEl) return;
		if (!this.selectedShapeKey) {
			this.fillSelectEl.value = '';
			this.strokeSelectEl.value = '';
			return;
		}
		const state = this.shapeFormats.get(this.selectedShapeKey);
		this.fillSelectEl.value = state?.fillToken ?? '';
		this.strokeSelectEl.value = state?.lineToken ?? '';
	}

	private captureShapeText(shapeEl: HTMLElement, shapeKey: string): void {
		const textContent = shapeEl.innerText.replace(/\r/g, '');
		this.shapeTextOverrides.set(shapeKey, textContent);
	}

	private renderTextOverride(shapeEl: HTMLElement, text: string): void {
		shapeEl.empty();
		const lines = text.split('\n');
		for (const line of lines) {
			if (line.trim().length === 0) {
				shapeEl.createEl('p', { text: ' ' });
				continue;
			}
			shapeEl.createEl('p', { text: line });
		}
	}

	private async loadDraftEdits(file: TFile): Promise<void> {
		const data = await loadPptxEdits(this.app, file);
		this.shapeEdits.clear();
		this.shapeFormats.clear();

		const entries = Object.entries(data.shapes);
		for (const [shapeKey, entry] of entries) {
			this.shapeEdits.set(shapeKey, {
				translateX: entry.translateX,
				translateY: entry.translateY,
				widthPx: entry.widthPx,
				heightPx: entry.heightPx,
			});
			this.shapeFormats.set(shapeKey, {
				fillToken: entry.fillToken,
				lineToken: entry.lineToken,
			});
			if (entry.textContent !== null) {
				this.shapeTextOverrides.set(shapeKey, entry.textContent);
			}
		}
	}

	private async saveDraftEdits(file: TFile): Promise<void> {
		await savePptxEdits(this.app, file, this.collectDraftPayload());
		new Notice('PPTX draft edits saved');
	}

	private async applyDraftEditsToPptx(file: TFile): Promise<void> {
		const draft = this.collectDraftPayload();
		if (Object.keys(draft.shapes).length === 0) {
			new Notice('No draft edits to apply');
			return;
		}

		const source = await this.app.vault.adapter.readBinary(file.path);
		const JSZip = ((await import('jszip')) as unknown as { default: JSZipConstructor }).default;
		const zip = await JSZip.loadAsync(source);
		const changedSlides = await applyEditsToPptxZip(zip, draft);
		if (changedSlides === 0) {
			new Notice('No matching shapes found in PPTX for current draft edits');
			return;
		}

		const updated = await zip.generateAsync({ type: 'arraybuffer' });
		await this.app.vault.modifyBinary(file, updated);

		this.shapeEdits.clear();
		this.shapeFormats.clear();
		this.selectedShapeKey = null;
		this.activeDrag = null;
		await savePptxEdits(this.app, file, { version: 1, shapes: {} });

		await this.renderFile(file);
		new Notice(`Applied edits to ${changedSlides} slide${changedSlides === 1 ? '' : 's'}`);
	}

	private collectDraftPayload(): {
		version: 1;
		shapes: Record<string, {
			translateX: number;
			translateY: number;
			widthPx: number | null;
			heightPx: number | null;
			fillToken: string | null;
			lineToken: string | null;
			textContent: string | null;
		}>;
	} {
		const shapes: Record<string, {
			translateX: number;
			translateY: number;
			widthPx: number | null;
			heightPx: number | null;
			fillToken: string | null;
			lineToken: string | null;
			textContent: string | null;
		}> = {};
		const allShapeKeys = new Set<string>();
		for (const shapeKey of this.shapeEdits.keys()) allShapeKeys.add(shapeKey);
		for (const shapeKey of this.shapeFormats.keys()) allShapeKeys.add(shapeKey);
		for (const shapeKey of this.shapeTextOverrides.keys()) allShapeKeys.add(shapeKey);

		for (const shapeKey of allShapeKeys) {
			const state = this.shapeEdits.get(shapeKey) ?? {
				translateX: 0,
				translateY: 0,
				widthPx: null,
				heightPx: null,
			};
			const format = this.shapeFormats.get(shapeKey);
			shapes[shapeKey] = {
				translateX: state.translateX,
				translateY: state.translateY,
				widthPx: state.widthPx,
				heightPx: state.heightPx,
				fillToken: format?.fillToken ?? null,
				lineToken: format?.lineToken ?? null,
				textContent: this.shapeTextOverrides.get(shapeKey) ?? null,
			};
		}

		return { version: 1, shapes };
	}

	private updateSelectionClasses(): void {
		const shapes = this.slideContainer?.querySelectorAll('.via-pptx-shape--editable');
		if (!shapes) return;
		for (let i = 0; i < shapes.length; i++) {
			const shape = shapes[i];
			if (!(shape instanceof HTMLElement)) continue;
			const key = shape.getAttribute('data-shape-key');
			shape.classList.toggle('is-selected', key !== null && key === this.selectedShapeKey);
		}
	}

	/** Render formatted runs into a container element. */
	private renderRuns(container: HTMLElement, runs: RunData[]): void {
		for (const run of runs) {
			if (!run.text) continue;
			if (run.bold && run.italic) {
				const b = container.createEl('strong');
				b.createEl('em', { text: run.text });
			} else if (run.bold) {
				container.createEl('strong', { text: run.text });
			} else if (run.italic) {
				container.createEl('em', { text: run.text });
			} else {
				container.appendText(run.text);
			}
		}
	}
}
