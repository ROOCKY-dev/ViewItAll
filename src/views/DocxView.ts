import { FileView, TFile, WorkspaceLeaf, Notice, Modal, App } from 'obsidian';
import { VIEW_TYPE_DOCX } from '../types';
import { readDocxAsHtml, saveHtmlAsDocx } from '../utils/docxUtils';
import type ViewItAllPlugin from '../main';

export class DocxView extends FileView {
	private plugin: ViewItAllPlugin;
	private editMode = false;
	private isDirty = false;
	private contentDiv: HTMLElement | null = null;
	private editToggleBtn: HTMLElement | null = null;
	private saveBtn: HTMLElement | null = null;
	private undoBtn: HTMLElement | null = null;
	private redoBtn: HTMLElement | null = null;
	private dirtyIndicator: HTMLElement | null = null;
	private currentFile: TFile | null = null;

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
	}

	getViewType(): string { return VIEW_TYPE_DOCX; }
	getDisplayText(): string { return this.file?.basename ?? 'Word Document'; }
	getIcon(): string { return 'file-text'; }

	canAcceptExtension(extension: string): boolean {
		return extension === 'docx';
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.currentFile = file;
		this.editMode = this.plugin.settings.docxDefaultEditMode;
		this.isDirty = false;
		await this.renderFile(file);
	}

	async onUnloadFile(_file: TFile): Promise<void> {
		this.contentEl.empty();
		this.contentDiv = null;
		this.editToggleBtn = null;
		this.saveBtn = null;
		this.undoBtn = null;
		this.redoBtn = null;
		this.dirtyIndicator = null;
	}

	private async renderFile(file: TFile): Promise<void> {
		this.contentEl.empty();

		const isBottom = this.plugin.settings.docxToolbarPosition === 'bottom';

		// Wrapper — flex column so toolbar can be ordered top or bottom
		const wrapper = this.contentEl.createEl('div', { cls: 'via-docx-wrapper' });
		if (isBottom) wrapper.classList.add('via-docx-wrapper--toolbar-bottom');

		// Scroll container
		const scrollEl = wrapper.createEl('div', { cls: 'via-docx-scroll' });

		// ── Toolbar ────────────────────────────────────────────────────────
		const toolbar = wrapper.createEl('div', { cls: 'via-docx-toolbar' });

		this.editToggleBtn = toolbar.createEl('button', {
			cls: 'via-btn',
			text: this.editMode ? '👁 View' : '✏️ Edit',
		});
		this.editToggleBtn.addEventListener('click', () => this.toggleEdit());

		// Undo / redo (visible in edit mode only)
		this.undoBtn = toolbar.createEl('button', { cls: 'via-btn', text: '↩ Undo' });
		this.undoBtn.title = 'Undo (Ctrl+Z)';
		this.undoBtn.style.display = this.editMode ? '' : 'none';
		this.undoBtn.addEventListener('click', () => document.execCommand('undo'));

		this.redoBtn = toolbar.createEl('button', { cls: 'via-btn', text: '↪ Redo' });
		this.redoBtn.title = 'Redo (Ctrl+Shift+Z)';
		this.redoBtn.style.display = this.editMode ? '' : 'none';
		this.redoBtn.addEventListener('click', () => document.execCommand('redo'));

		this.dirtyIndicator = toolbar.createEl('span', { cls: 'via-docx-dirty', text: '● Unsaved' });
		this.dirtyIndicator.style.display = 'none';

		this.saveBtn = toolbar.createEl('button', {
			cls: 'via-btn via-btn-save',
			text: '💾 Save',
		});
		this.saveBtn.style.display = this.editMode ? '' : 'none';
		this.saveBtn.addEventListener('click', () => this.saveFile());

		// ── Conversion warnings ────────────────────────────────────────────
		let html: string;
		let messages: string[];
		try {
			const buffer = await this.app.vault.adapter.readBinary(file.path);
			({ html, messages } = await readDocxAsHtml(buffer));
		} catch (err) {
			scrollEl.createEl('p', {
				cls: 'via-error',
				text: `Failed to read file: ${String(err)}`,
			});
			return;
		}

		if (messages.length > 0) {
			const warn = scrollEl.createEl('div', { cls: 'via-warning' });
			warn.createEl('strong', { text: '⚠️ Conversion notes: ' });
			warn.createEl('span', { text: messages.join('; ') });
		}

		// ── Content ────────────────────────────────────────────────────────
		this.contentDiv = scrollEl.createEl('div', { cls: 'via-docx-content' });
		this.contentDiv.innerHTML = html;
		this.contentDiv.contentEditable = this.editMode ? 'true' : 'false';
		if (this.editMode) this.contentDiv.classList.add('via-editable');

		// Track dirty state
		this.contentDiv.addEventListener('input', () => this.setDirty(true));
	}

	private toggleEdit(): void {
		this.editMode = !this.editMode;
		if (!this.contentDiv || !this.editToggleBtn || !this.saveBtn) return;
		this.contentDiv.contentEditable = this.editMode ? 'true' : 'false';
		this.contentDiv.classList.toggle('via-editable', this.editMode);
		this.editToggleBtn.textContent = this.editMode ? '👁 View' : '✏️ Edit';
		this.saveBtn.style.display = this.editMode ? '' : 'none';
		if (this.undoBtn) this.undoBtn.style.display = this.editMode ? '' : 'none';
		if (this.redoBtn) this.redoBtn.style.display = this.editMode ? '' : 'none';
		// Hide dirty indicator when leaving edit mode without saving
		if (!this.editMode) this.setDirty(false);
	}

	private setDirty(dirty: boolean): void {
		this.isDirty = dirty;
		if (this.dirtyIndicator) this.dirtyIndicator.style.display = dirty ? '' : 'none';
	}

	private async saveFile(): Promise<void> {
		if (!this.currentFile || !this.contentDiv) return;

		if (this.plugin.settings.confirmOnSave) {
			const confirmed = await confirmModal(
				this.app,
				`Overwrite "${this.currentFile.name}"?`,
				'This will replace the original file. Complex formatting may be simplified.'
			);
			if (!confirmed) return;
		}

		try {
			const buffer = await saveHtmlAsDocx(this.contentDiv.innerHTML);
			await this.app.vault.adapter.writeBinary(this.currentFile.path, buffer);
			this.setDirty(false);
			new Notice(`✅ Saved "${this.currentFile.name}"`);
		} catch (err) {
			new Notice(`❌ Save failed: ${String(err)}`);
		}
	}
}

// ── Simple confirmation modal ──────────────────────────────────────────────
function confirmModal(app: App, title: string, message: string): Promise<boolean> {
	return new Promise(resolve => {
		const modal = new ConfirmModal(app, title, message, resolve);
		modal.open();
	});
}

class ConfirmModal extends Modal {
	constructor(
		app: App,
		private title: string,
		private message: string,
		private resolve: (v: boolean) => void
	) {
		super(app);
	}

	onOpen(): void {
		const { contentEl } = this;
		contentEl.createEl('h3', { text: this.title });
		contentEl.createEl('p', { text: this.message });
		const btnRow = contentEl.createEl('div', { cls: 'via-modal-btns' });
		btnRow.createEl('button', { text: 'Cancel', cls: 'via-btn' })
			.addEventListener('click', () => { this.resolve(false); this.close(); });
		btnRow.createEl('button', { text: 'Overwrite', cls: 'via-btn via-btn-danger' })
			.addEventListener('click', () => { this.resolve(true); this.close(); });
	}

	onClose(): void { this.contentEl.empty(); }
}
