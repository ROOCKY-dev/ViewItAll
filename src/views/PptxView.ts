import { FileView, TFile, WorkspaceLeaf } from "obsidian";
import { PPTXViewer } from "pptxviewjs";
import { VIEW_TYPE_PPTX } from "../types";
import ViewItAllPlugin from "../main";

// Internal pptxviewjs types for patching theme colors
interface PptxColor {
	scheme?: string;
	tint?: number;
	shade?: number;
	lumMod?: number;
	lumOff?: number;
}

interface PptxDrawingDocument {
	parseColorToHex: (color: unknown) => string;
	applyColorModifications?: (hex: string, color: unknown) => string;
}

interface PptxSlide {
	theme?: Record<string, unknown>;
	layout?: { master?: { theme?: Record<string, unknown> } };
}

interface PptxRenderer {
	presentation?: {
		theme?: { colors?: Record<string, string> };
	};
	slides?: PptxSlide[];
	drawingDocument?: PptxDrawingDocument;
	getSlideDimensions?: () => { cx: number; cy: number } | undefined;
}

interface PptxProcessor {
	processor?: PptxRenderer;
	getSlideDimensions?: () => { cx: number; cy: number } | undefined;
}

export class PptxView extends FileView {
	private plugin: ViewItAllPlugin;
	private viewer: PPTXViewer | null = null;
	private canvasEl: HTMLCanvasElement | null = null;
	private canvasWrapper: HTMLElement | null = null;
	private statusEl: HTMLElement | null = null;
	private zoomLabel: HTMLElement | null = null;
	private zoom = 1.0;
	private slideWidthPx = 960;
	private slideHeightPx = 720;
	private resizeObserver: ResizeObserver | null = null;
	private rendering = false;

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
	}

	getViewType(): string {
		return VIEW_TYPE_PPTX;
	}

	getDisplayText(): string {
		return this.file ? this.file.basename : "PPTX";
	}

	getIcon(): string {
		return "presentation";
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.contentEl.empty();
		this.buildUI();
		await this.loadPresentation(file);
	}

	private buildUI() {
		this.contentEl.classList.add("via-pptx-wrapper");

		const toolbar = this.contentEl.createDiv("via-pptx-toolbar");

		const prevBtn = toolbar.createEl("button", { text: "◀ Previous" });
		prevBtn.addEventListener("click", () => { void this.prevSlide(); });

		const nextBtn = toolbar.createEl("button", { text: "Next ▶" });
		nextBtn.addEventListener("click", () => { void this.nextSlide(); });

		toolbar.createEl("span", {
			text: "|",
			cls: "via-separator-faint",
		});

		const zoomOutBtn = toolbar.createEl("button", { text: "−" });
		zoomOutBtn.addEventListener("click", () => {
			void this.setZoom(this.zoom - 0.15);
		});

		this.zoomLabel = toolbar.createEl("span", {
			text: "100%",
			cls: "via-pptx-zoom-label",
		});

		const zoomInBtn = toolbar.createEl("button", { text: "+" });
		zoomInBtn.addEventListener("click", () => {
			void this.setZoom(this.zoom + 0.15);
		});

		const fitBtn = toolbar.createEl("button", { text: "Fit" });
		fitBtn.addEventListener("click", () => { void this.fitToContainer(); });

		toolbar.createEl("span", {
			text: "|",
			cls: "via-separator-faint",
		});

		this.statusEl = toolbar.createEl("span", { cls: "via-pptx-status" });
		this.statusEl.setText("Loading...");

		this.canvasWrapper = this.contentEl.createDiv(
			"via-pptx-canvas-wrapper",
		);

		this.canvasEl = this.canvasWrapper.createEl("canvas", {
			cls: "via-pptx-canvas",
		});
	}

	private async loadPresentation(file: TFile) {
		const buffer = await this.app.vault.readBinary(file);

		this.viewer?.destroy();
		this.viewer = new PPTXViewer({
			canvas: this.canvasEl,
			slideSizeMode: "fit",
			autoChartRerenderDelayMs: 400,
		});

		this.viewer.on("slideChanged", () => this.updateStatus());

		try {
			await this.viewer.loadFile(buffer);
			this.patchThemeColors();
			this.readSlideDimensions();
			await this.renderCurrentSlide();
			this.updateStatus();
			this.setupResizeObserver();
		} catch (e) {
			console.error("PptxView: failed to render PPTX", e);
			this.statusEl?.setText("Failed to render presentation");
		}
	}

	/**
	 * Patch pptxviewjs internals to use real theme colors from the PPTX file
	 * instead of hardcoded fallback scheme colors.
	 */
	private patchThemeColors() {
		try {
			const v = this.viewer as unknown as { processor?: PptxProcessor };
			const outerProc = v?.processor;
			const renderer = outerProc?.processor;
			if (!renderer) return;

			const presentation = renderer.presentation;
			if (!presentation) return;
			const themeColors: Record<string, string> | undefined =
				presentation.theme?.colors;
			if (!themeColors) return;

			// 1. Propagate theme onto every slide so bgRef / currentSlide.theme works
			if (Array.isArray(renderer.slides)) {
				for (const slide of renderer.slides) {
					if (!slide.theme) slide.theme = presentation.theme;
					if (slide.layout?.master && !slide.layout.master.theme) {
						slide.layout.master.theme = presentation.theme;
					}
				}
			}

			// 2. Patch the DrawingDocument's parseColorToHex to resolve scheme
			//    colors from the real theme instead of hardcoded defaults.
			const drawDoc = renderer.drawingDocument;
			if (!drawDoc) return;

			const origParse = drawDoc.parseColorToHex.bind(drawDoc);
			drawDoc.parseColorToHex = (color: unknown): string => {
				if (!color) return "#ffffff";
				if (typeof color === "string") {
					return color.startsWith("#") ? color : `#${color}`;
				}
				const c = color as PptxColor;
				if (c.scheme && themeColors[c.scheme]) {
					let hex = themeColors[c.scheme];
					if (typeof hex === "string") {
						if (!hex.startsWith("#")) hex = `#${hex}`;
						// Apply tint / shade / lumMod / lumOff modifications
						if (
							c.tint !== undefined ||
							c.shade !== undefined ||
							c.lumMod !== undefined ||
							c.lumOff !== undefined
						) {
							if (
								typeof drawDoc.applyColorModifications ===
								"function"
							) {
								return drawDoc.applyColorModifications(
									hex,
									color,
								);
							}
						}
						return hex;
					}
				}
				return origParse(color);
			};
		} catch {
			// non-critical — rendering will still work with hardcoded fallbacks
		}
	}

	private readSlideDimensions() {
		try {
			const proc = (this.viewer as unknown as { processor?: PptxProcessor })?.processor;
			const dims =
				proc?.getSlideDimensions?.() ??
				proc?.processor?.getSlideDimensions?.();
			if (dims && dims.cx && dims.cy) {
				this.slideWidthPx = (dims.cx / 914400) * 96;
				this.slideHeightPx = (dims.cy / 914400) * 96;
			}
		} catch {
			// keep defaults
		}
	}

	private setupResizeObserver() {
		this.resizeObserver?.disconnect();
		if (!this.canvasWrapper) return;

		let timeout: ReturnType<typeof setTimeout>;
		this.resizeObserver = new ResizeObserver(() => {
			clearTimeout(timeout);
			timeout = setTimeout(() => { void this.renderCurrentSlide(); }, 150);
		});
		this.resizeObserver.observe(this.canvasWrapper);
	}

	private async renderCurrentSlide() {
		if (
			!this.viewer ||
			!this.canvasEl ||
			!this.canvasWrapper ||
			this.rendering
		)
			return;
		this.rendering = true;

		try {
			const wrapperRect = this.canvasWrapper.getBoundingClientRect();
			const availW = wrapperRect.width - 32;
			const availH = wrapperRect.height - 32;
			if (availW <= 0 || availH <= 0) return;

			const fitScale = Math.min(
				availW / this.slideWidthPx,
				availH / this.slideHeightPx,
			);
			const scale = fitScale * this.zoom;

			const cssW = Math.round(this.slideWidthPx * scale);
			const cssH = Math.round(this.slideHeightPx * scale);

			// Set only CSS dimensions — the library handles canvas.width/height
			// and DPR scaling internally via its init() method
			this.canvasEl.setCssProps({
				"--via-canvas-width": `${cssW}px`,
				"--via-canvas-height": `${cssH}px`,
			});

			await this.viewer.render(this.canvasEl);
		} finally {
			this.rendering = false;
		}
	}

	private async fitToContainer() {
		this.zoom = 1.0;
		this.updateZoomLabel();
		await this.renderCurrentSlide();
	}

	private updateStatus() {
		if (!this.viewer || !this.statusEl) return;
		const current = this.viewer.getCurrentSlideIndex() + 1;
		const total = this.viewer.getSlideCount();
		this.statusEl.setText(`Slide ${current} / ${total}`);
	}

	private updateZoomLabel() {
		if (this.zoomLabel) {
			this.zoomLabel.setText(`${Math.round(this.zoom * 100)}%`);
		}
	}

	private async prevSlide() {
		if (!this.viewer || this.rendering) return;
		await this.viewer.previousSlide(this.canvasEl);
		await this.renderCurrentSlide();
		this.updateStatus();
	}

	private async nextSlide() {
		if (!this.viewer || this.rendering) return;
		await this.viewer.nextSlide(this.canvasEl);
		await this.renderCurrentSlide();
		this.updateStatus();
	}

	private async setZoom(newZoom: number) {
		this.zoom = Math.max(0.25, Math.min(newZoom, 4.0));
		this.updateZoomLabel();
		await this.renderCurrentSlide();
	}

	// No async work needed — returns resolved promise for type compatibility
	protected onClose(): Promise<void> {
		this.resizeObserver?.disconnect();
		this.resizeObserver = null;
		this.viewer?.destroy();
		this.viewer = null;
		this.canvasEl = null;
		this.canvasWrapper = null;
		this.statusEl = null;
		this.zoomLabel = null;
		this.contentEl.empty();
		return Promise.resolve();
	}
}
