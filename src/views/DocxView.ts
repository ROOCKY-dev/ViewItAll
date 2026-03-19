/**
 * ViewItAll — DocxView
 *
 * FileView subclass that renders .docx files natively using the
 * OOXML parser + DOM renderer pipeline.
 */

import { FileView, Notice, TFile, WorkspaceLeaf, setIcon, setTooltip } from "obsidian";
import type ViewItAllPlugin from "../main";
import { VIEW_TYPE_DOCX } from "../types";
import type { DocxDocument } from "../docx/model";
import { parseDocx } from "../docx/parser";
import { renderDocument } from "../docx/renderer";

export class DocxView extends FileView {
	plugin: ViewItAllPlugin;

	private wrapperEl: HTMLElement | null = null;
	private toolbarEl: HTMLElement | null = null;
	private scrollEl: HTMLElement | null = null;
	private contentEl_: HTMLElement | null = null;
	private blobUrls: string[] = [];
	private docModel: DocxDocument | null = null;
	private currentZoom = 1.0;

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
	}

	getViewType(): string {
		return VIEW_TYPE_DOCX;
	}

	getDisplayText(): string {
		return this.file?.basename ?? "Word document";
	}

	getIcon(): string {
		return "file-text";
	}

	async onLoadFile(file: TFile): Promise<void> {
		const s = this.plugin.settings;
		this.currentZoom = s.docxDefaultZoom;

		// Clear previous content
		this.contentEl.empty();

		// Build layout
		const isBottom = s.docxToolbarPosition === "bottom";
		this.wrapperEl = this.contentEl.createEl("div", {
			cls: `via-docx-wrapper${isBottom ? " via-docx-wrapper--toolbar-bottom" : ""}`,
		});

		this.toolbarEl = this.wrapperEl.createEl("div", {
			cls: "via-docx-toolbar",
		});

		this.scrollEl = this.wrapperEl.createEl("div", {
			cls: "via-docx-scroll",
		});

		this.contentEl_ = this.scrollEl.createEl("div", {
			cls: "via-docx-content",
		});

		// Build toolbar
		this.buildToolbar();

		// Apply initial zoom
		this.applyZoom();

		// Parse and render
		try {
			const data = await this.app.vault.readBinary(file);
			this.docModel = await parseDocx(data);
			this.blobUrls = renderDocument(this.docModel, this.contentEl_);
		} catch (err: unknown) {
			const msg = err instanceof Error ? err.message : String(err);
			new Notice(`Failed to open .docx: ${msg}`);
			this.contentEl_.createEl("div", {
				cls: "via-error",
				text: `Error loading document: ${msg}`,
			});
		}
	}

	async onUnloadFile(): Promise<void> {
		// Revoke all blob URLs to prevent memory leaks
		for (const url of this.blobUrls) {
			URL.revokeObjectURL(url);
		}
		this.blobUrls = [];

		// Clear DOM refs
		this.contentEl.empty();
		this.wrapperEl = null;
		this.toolbarEl = null;
		this.scrollEl = null;
		this.contentEl_ = null;
		this.docModel = null;
	}

	canAcceptExtension(extension: string): boolean {
		return extension === "docx" && this.plugin.settings.enableDocx;
	}

	// ── Toolbar ─────────────────────────────────────────────────────────────

	private buildToolbar(): void {
		if (!this.toolbarEl) return;

		// File name label
		const fileLabel = this.toolbarEl.createEl("div", {
			cls: "via-docx-file-label",
		});
		const fileIcon = fileLabel.createEl("div", { cls: "clickable-icon" });
		setIcon(fileIcon, "file-text");
		fileLabel.createEl("span", {
			cls: "via-docx-file-name",
			text: this.file?.name ?? "Untitled",
		});

		// Separator
		this.toolbarEl.createEl("div", { cls: "via-toolbar-sep" });

		// Zoom out
		const zoomOut = this.toolbarEl.createEl("div", { cls: "clickable-icon" });
		setIcon(zoomOut, "minus");
		setTooltip(zoomOut, "Zoom out");
		zoomOut.addEventListener("click", () => this.adjustZoom(-0.25));

		// Zoom label
		const zoomLabel = this.toolbarEl.createEl("button", {
			cls: "via-btn via-btn-zoom-label",
			text: this.formatZoom(),
		});
		setTooltip(zoomLabel, "Reset zoom");
		zoomLabel.addEventListener("click", () => {
			this.currentZoom = 1.0;
			this.applyZoom();
			zoomLabel.textContent = this.formatZoom();
		});

		// Zoom in
		const zoomIn = this.toolbarEl.createEl("div", { cls: "clickable-icon" });
		setIcon(zoomIn, "plus");
		setTooltip(zoomIn, "Zoom in");
		zoomIn.addEventListener("click", () => this.adjustZoom(0.25));

		// Spacer
		this.toolbarEl.createEl("div", { cls: "via-toolbar-spacer" });

		// Store zoom label ref for updates
		this.zoomLabelEl = zoomLabel;
	}

	private zoomLabelEl: HTMLElement | null = null;

	private adjustZoom(delta: number): void {
		const newZoom = Math.max(0.25, Math.min(3.0, this.currentZoom + delta));
		this.currentZoom = newZoom;
		this.applyZoom();
		if (this.zoomLabelEl) {
			this.zoomLabelEl.textContent = this.formatZoom();
		}
	}

	private applyZoom(): void {
		if (this.contentEl_) {
			this.contentEl_.style.transform = `scale(${this.currentZoom})`;
			this.contentEl_.style.transformOrigin = "top center";
			// Adjust width to compensate for scaling
			if (this.currentZoom !== 1) {
				this.contentEl_.style.width = `${100 / this.currentZoom}%`;
			} else {
				this.contentEl_.style.width = "";
			}
		}
	}

	private formatZoom(): string {
		return `${Math.round(this.currentZoom * 100)}%`;
	}
}
