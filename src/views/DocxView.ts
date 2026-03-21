/**
 * ViewItAll — DocxView
 *
 * FileView subclass that renders .docx files natively using the
 * OOXML parser + DOM renderer pipeline. Supports view and edit modes
 * with a formatting toolbar, table editing, and image insertion.
 */

import { FileView, Notice, TFile, WorkspaceLeaf, setIcon, setTooltip } from "obsidian";
import type ViewItAllPlugin from "../main";
import { VIEW_TYPE_DOCX } from "../types";
import type { DocxDocument, DocxRunProperties, DocxParagraphProperties } from "../docx/model";
import { parseDocx, type ParseResult } from "../docx/parser";
import { renderDocument } from "../docx/renderer";
import { createEditingController, type EditingController } from "../docx/editing";
import { serializeDocx } from "../docx/serializer";
import { createFormattingToolbar, type FormattingToolbar, type FormatCommand } from "../docx/toolbar";
import {
	domSelectionToModel,
	getRunPropertiesAtSelection,
	getParagraphPropertiesAtSelection,
} from "../docx/selection";

export class DocxView extends FileView {
	plugin: ViewItAllPlugin;

	private wrapperEl: HTMLElement | null = null;
	private toolbarEl: HTMLElement | null = null;
	private formatToolbarContainer: HTMLElement | null = null;
	private scrollEl: HTMLElement | null = null;
	private contentEl_: HTMLElement | null = null;
	private blobUrls: string[] = [];
	private docModel: DocxDocument | null = null;
	private zip: unknown = null;
	private currentZoom = 1.0;

	// Editing state
	private editCtrl: EditingController | null = null;
	private formatToolbar: FormattingToolbar | null = null;
	private dirty = false;
	private editing = false;
	private editToggleBtn: HTMLElement | null = null;
	private saveBtn: HTMLElement | null = null;
	private styleCache = new Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>();
	private keydownHandler: ((e: KeyboardEvent) => void) | null = null;
	private selectionHandler: (() => void) | null = null;

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
		this.dirty = false;
		this.editing = false;

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

		// Formatting toolbar container (below main toolbar)
		this.formatToolbarContainer = this.wrapperEl.createEl("div", {
			cls: "via-docx-format-toolbar-container via-hidden",
		});

		this.scrollEl = this.wrapperEl.createEl("div", {
			cls: "via-docx-scroll",
		});

		this.contentEl_ = this.scrollEl.createEl("div", {
			cls: "via-docx-content",
		});

		// Build toolbars
		this.buildToolbar();
		this.buildFormattingToolbar();

		// Apply initial zoom
		this.applyZoom();

		// Create editing controller
		this.editCtrl = createEditingController();

		// Register Ctrl+S / Cmd+S handler
		this.keydownHandler = (e: KeyboardEvent) => this.handleGlobalKeydown(e);
		this.contentEl.addEventListener("keydown", this.keydownHandler);

		// Register selection change handler for toolbar state
		this.selectionHandler = () => this.handleSelectionChange();
		document.addEventListener("selectionchange", this.selectionHandler);

		// Parse and render
		try {
			const data = await this.app.vault.readBinary(file);
			const result: ParseResult = await parseDocx(data);
			this.docModel = result.doc;
			this.zip = result.zip;
			this.blobUrls = renderDocument(this.docModel, this.contentEl_);

			// Auto-enter edit mode if setting enabled
			if (s.docxEditMode) {
				this.toggleEditMode(true);
			}
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
		// Auto-save dirty documents if enabled
		const saveTask = (this.dirty && this.plugin.settings.docxAutoSave)
			? this.saveDocument()
			: Promise.resolve();
		await saveTask;

		// Disable editing
		if (this.editCtrl?.isEnabled()) {
			this.editCtrl.disable();
		}
		this.editCtrl = null;

		// Destroy formatting toolbar
		if (this.formatToolbar) {
			this.formatToolbar.destroy();
			this.formatToolbar = null;
		}

		// Remove handlers
		if (this.keydownHandler) {
			this.contentEl.removeEventListener("keydown", this.keydownHandler);
			this.keydownHandler = null;
		}
		if (this.selectionHandler) {
			document.removeEventListener("selectionchange", this.selectionHandler);
			this.selectionHandler = null;
		}

		// Revoke all blob URLs to prevent memory leaks
		for (const url of this.blobUrls) {
			URL.revokeObjectURL(url);
		}
		this.blobUrls = [];

		// Clear DOM refs
		this.contentEl.empty();
		this.wrapperEl = null;
		this.toolbarEl = null;
		this.formatToolbarContainer = null;
		this.scrollEl = null;
		this.contentEl_ = null;
		this.docModel = null;
		this.zip = null;
		this.editToggleBtn = null;
		this.saveBtn = null;
		this.styleCache.clear();
		this.dirty = false;
		this.editing = false;
	}

	canAcceptExtension(extension: string): boolean {
		return extension === "docx" && this.plugin.settings.enableDocx;
	}

	// ── Main Toolbar ────────────────────────────────────────────────────────

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

		// Edit/View toggle
		this.editToggleBtn = this.toolbarEl.createEl("div", { cls: "clickable-icon" });
		setIcon(this.editToggleBtn, "eye");
		setTooltip(this.editToggleBtn, "Toggle edit mode");
		this.editToggleBtn.addEventListener("click", () => this.toggleEditMode());

		// Save button
		this.saveBtn = this.toolbarEl.createEl("div", { cls: "clickable-icon" });
		setIcon(this.saveBtn, "save");
		setTooltip(this.saveBtn, "Save (Ctrl+S)");
		this.saveBtn.addEventListener("click", () => { void this.saveDocument(); });
		this.saveBtn.classList.add("via-hidden");

		// Separator
		this.toolbarEl.createEl("div", { cls: "via-toolbar-sep" });

		// Insert image button (only functional in edit mode)
		const insertImageBtn = this.toolbarEl.createEl("div", { cls: "clickable-icon via-hidden" });
		setIcon(insertImageBtn, "image-plus");
		setTooltip(insertImageBtn, "Insert image");
		insertImageBtn.addEventListener("click", () => this.insertImage());
		insertImageBtn.dataset.editOnly = "true";

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

	// ── Formatting Toolbar ──────────────────────────────────────────────────

	private buildFormattingToolbar(): void {
		if (!this.formatToolbarContainer) return;

		this.formatToolbar = createFormattingToolbar();
		this.formatToolbar.build(
			this.formatToolbarContainer,
			(cmd: FormatCommand) => this.handleFormatCommand(cmd),
		);
	}

	private handleFormatCommand(cmd: FormatCommand): void {
		if (!this.editCtrl?.isEnabled()) return;
		this.editCtrl.applyFormat(cmd);
		this.markDirty();

		// Update toolbar state after formatting
		this.handleSelectionChange();
	}

	private handleSelectionChange(): void {
		if (!this.editing || !this.docModel || !this.contentEl_ || !this.formatToolbar) return;

		const msel = domSelectionToModel(this.contentEl_, this.docModel);
		if (!msel) return;

		const rProps = getRunPropertiesAtSelection(msel, this.docModel, this.styleCache);
		const pProps = getParagraphPropertiesAtSelection(msel, this.docModel);
		this.formatToolbar.updateState(rProps, pProps);
	}

	// ── Edit Mode ───────────────────────────────────────────────────────────

	private toggleEditMode(forceEdit?: boolean): void {
		const shouldEdit = forceEdit !== undefined ? forceEdit : !this.editing;

		if (shouldEdit) {
			this.enterEditMode();
		} else {
			this.exitEditMode();
		}
	}

	private enterEditMode(): void {
		if (!this.editCtrl || !this.contentEl_ || !this.docModel) return;
		if (this.editing) return;

		this.editing = true;
		this.editCtrl.enable(
			this.contentEl_,
			this.docModel,
			this.styleCache,
			() => this.markDirty(),
		);

		// Update toggle button icon
		if (this.editToggleBtn) {
			setIcon(this.editToggleBtn, "pencil");
			setTooltip(this.editToggleBtn, "Switch to view mode");
			this.editToggleBtn.classList.add("is-active");
		}

		// Show save button and formatting toolbar
		if (this.saveBtn) {
			this.saveBtn.classList.remove("via-hidden");
		}
		if (this.formatToolbarContainer) {
			this.formatToolbarContainer.classList.remove("via-hidden");
		}
		this.formatToolbar?.show();

		// Show edit-only toolbar buttons
		if (this.toolbarEl) {
			const editBtns = Array.from(
				this.toolbarEl.querySelectorAll<HTMLElement>("[data-edit-only]"),
			);
			for (const btn of editBtns) {
				btn.classList.remove("via-hidden");
			}
		}
	}

	private exitEditMode(): void {
		if (!this.editCtrl || !this.editing) return;

		this.editing = false;
		this.editCtrl.disable();

		// Update toggle button icon
		if (this.editToggleBtn) {
			setIcon(this.editToggleBtn, "eye");
			setTooltip(this.editToggleBtn, "Switch to edit mode");
			this.editToggleBtn.classList.remove("is-active");
		}

		// Hide formatting toolbar
		if (this.formatToolbarContainer) {
			this.formatToolbarContainer.classList.add("via-hidden");
		}
		this.formatToolbar?.hide();

		// Hide edit-only toolbar buttons
		if (this.toolbarEl) {
			const editBtns = Array.from(
				this.toolbarEl.querySelectorAll<HTMLElement>("[data-edit-only]"),
			);
			for (const btn of editBtns) {
				btn.classList.add("via-hidden");
			}
		}

		// Hide save button when not dirty
		if (this.saveBtn && !this.dirty) {
			this.saveBtn.classList.add("via-hidden");
		}
	}

	private markDirty(): void {
		if (this.dirty) return;
		this.dirty = true;

		if (this.saveBtn) {
			this.saveBtn.classList.add("via-docx-toolbar-save--dirty");
		}
	}

	private clearDirty(): void {
		this.dirty = false;

		if (this.saveBtn) {
			this.saveBtn.classList.remove("via-docx-toolbar-save--dirty");
			if (!this.editing) {
				this.saveBtn.classList.add("via-hidden");
			}
		}
	}

	// ── Save ────────────────────────────────────────────────────────────────

	private async saveDocument(): Promise<void> {
		if (!this.docModel || !this.zip || !this.file) {
			new Notice("Nothing to save");
			return;
		}

		try {
			const buffer = await serializeDocx(this.docModel, this.zip);
			await this.app.vault.modifyBinary(this.file, buffer);
			this.clearDirty();
			new Notice("Document saved");
		} catch (err: unknown) {
			const msg = err instanceof Error ? err.message : String(err);
			new Notice(`Failed to save: ${msg}`);
		}
	}

	private handleGlobalKeydown(e: KeyboardEvent): void {
		// Ctrl+S / Cmd+S to save
		if ((e.ctrlKey || e.metaKey) && e.key === "s") {
			if (this.dirty) {
				e.preventDefault();
				e.stopPropagation();
				void this.saveDocument();
			}
		}
	}

	// ── Image Insertion ─────────────────────────────────────────────────────

	private insertImage(): void {
		if (!this.editing || !this.docModel || !this.contentEl_) return;

		// Create a file picker
		const input = document.createElement("input");
		input.type = "file";
		input.accept = "image/png,image/jpeg,image/gif,image/webp,image/svg+xml";
		input.classList.add("via-hidden");

		input.addEventListener("change", () => { void (async () => {
			const file = input.files?.[0];
			if (!file || !this.docModel || !this.contentEl_) return;

			try {
				const buffer = await file.arrayBuffer();
				const blob = new Blob([buffer], { type: file.type });

				// Generate a unique relationship ID
				const rId = `rId${Date.now()}`;

				// Determine file extension
				const ext = file.name.split(".").pop()?.toLowerCase() ?? "png";
				const target = `media/image_${Date.now()}.${ext}`;

				// Add to model
				this.docModel.images.set(rId, blob);
				this.docModel.relationships.set(rId, {
					type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
					target,
				});

				// Create image element in model at current caret position
				const sel = window.getSelection();
				let blockIdx = this.docModel.body.length - 1;
				if (sel && sel.rangeCount) {
					let node: HTMLElement | null = sel.anchorNode instanceof HTMLElement
						? sel.anchorNode
						: sel.anchorNode?.parentElement ?? null;
					while (node && node !== this.contentEl_) {
						if (node.dataset.blockIdx !== undefined) {
							blockIdx = parseInt(node.dataset.blockIdx, 10);
							break;
						}
						node = node.parentElement;
					}
				}

				const block = this.docModel.body[blockIdx];
				if (block && block.type === "paragraph") {
					// Read image dimensions
					const img = new Image();
					const url = URL.createObjectURL(blob);
					await new Promise<void>((resolve) => {
						img.onload = () => resolve();
						img.onerror = () => resolve();
						img.src = url;
					});

					let width = img.naturalWidth || 300;
					let height = img.naturalHeight || 200;
					URL.revokeObjectURL(url);

					// Scale down to fit within content area (max ~600px wide)
					const maxWidth = 600;
					if (width > maxWidth) {
						const scale = maxWidth / width;
						width = Math.round(width * scale);
						height = Math.round(height * scale);
					}

					block.children.push({
						type: "image",
						relationshipId: rId,
						width,
						height,
						altText: file.name,
					});
				}

				// Re-render
				this.contentEl_.empty();
				this.blobUrls = renderDocument(this.docModel, this.contentEl_);

				// Re-enable editing on new elements
				if (this.editCtrl?.isEnabled()) {
					this.editCtrl.disable();
					this.editCtrl.enable(
						this.contentEl_,
						this.docModel,
						this.styleCache,
						() => this.markDirty(),
					);
				}

				this.markDirty();
				new Notice("Image inserted");
			} catch (err: unknown) {
				const msg = err instanceof Error ? err.message : String(err);
				new Notice(`Failed to insert image: ${msg}`);
			}

			input.remove();
		})(); });

		document.body.appendChild(input);
		input.click();
	}

	// ── Zoom ────────────────────────────────────────────────────────────────

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
			this.contentEl_.classList.add("via-docx-zoom");
			this.contentEl_.setCssStyles({ transform: `scale(${this.currentZoom})` });
			// Adjust width to compensate for scaling
			if (this.currentZoom !== 1) {
				this.contentEl_.setCssStyles({ width: `${100 / this.currentZoom}%` });
			} else {
				this.contentEl_.setCssStyles({ width: "" });
			}
		}
	}

	private formatZoom(): string {
		return `${Math.round(this.currentZoom * 100)}%`;
	}
}
