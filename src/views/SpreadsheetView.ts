import {
	FileView,
	Notice,
	TFile,
	WorkspaceLeaf,
	Menu,
	Modal,
	App,
	setIcon,
	setTooltip,
} from "obsidian";
import { VIEW_TYPE_SPREADSHEET } from "../types";
import { evaluateSheet } from "../utils/formulaEval";
import type ViewItAllPlugin from "../main";

// ── SheetJS type surface ──────────────────────────────────────────────────────

type CellValue = string | number | boolean | null | undefined;

interface XLSXCell {
	v?: CellValue; // primitive value
	f?: string; // formula (without leading '=')
	t?: string; // type: 'n' | 's' | 'b' | 'e' | 'z'
	w?: string; // formatted text
}

/**
 * Opaque sheet record: cell refs like "A1" map to XLSXCell,
 * and special keys "!ref", "!cols", etc. map to metadata.
 */
type XLSXSheet = Record<string, unknown>;

interface XLSXWorkBook {
	SheetNames: string[];
	Sheets: Record<string, XLSXSheet>;
}

interface SheetRange {
	s: { r: number; c: number };
	e: { r: number; c: number };
}

interface XLSXModule {
	read(data: Uint8Array | ArrayBuffer, opts: { type: "array" }): XLSXWorkBook;
	write(wb: XLSXWorkBook, opts: { type: string; bookType: string }): unknown;
	utils: {
		encode_cell(addr: { r: number; c: number }): string;
		decode_cell(addr: string): { r: number; c: number };
		encode_range(range: SheetRange): string;
		decode_range(range: string): SheetRange;
	};
}

// ── Sheet helpers ─────────────────────────────────────────────────────────────

function sheetGetRef(sheet: XLSXSheet): string | undefined {
	return sheet["!ref"] as string | undefined;
}

function sheetSetRef(sheet: XLSXSheet, ref: string): void {
	(sheet as Record<string, unknown>)["!ref"] = ref;
}

function sheetGetCell(sheet: XLSXSheet, ref: string): XLSXCell | undefined {
	return sheet[ref] as XLSXCell | undefined;
}

function sheetSetCell(sheet: XLSXSheet, ref: string, cell: XLSXCell): void {
	(sheet as Record<string, unknown>)[ref] = cell;
}

// ── View ──────────────────────────────────────────────────────────────────────

export class SpreadsheetView extends FileView {
	private plugin: ViewItAllPlugin;
	private currentFile: TFile | null = null;
	private workbook: XLSXWorkBook | null = null;
	private xlsx: XLSXModule | null = null;
	private activeSheet = 0;
	private isDirty = false;
	private editMode = false;
	private savedData: ArrayBuffer | null = null;

	// Selection
	private selRow = -1;
	private selCol = -1;

	// DOM refs
	private wrapper: HTMLElement | null = null;
	private tabBar: HTMLElement | null = null;
	private tableContainer: HTMLElement | null = null;
	private formulaRef: HTMLElement | null = null;
	private formulaInput: HTMLInputElement | null = null;
	private saveBtn: HTMLElement | null = null;
	private undoBtn: HTMLElement | null = null;
	private editToggleBtn: HTMLElement | null = null;
	private dirtyIndicator: HTMLElement | null = null;
	private addRowBtn: HTMLElement | null = null;
	private addColBtn: HTMLElement | null = null;
	private infoEl: HTMLElement | null = null;

	constructor(leaf: WorkspaceLeaf, plugin: ViewItAllPlugin) {
		super(leaf);
		this.plugin = plugin;
	}

	getViewType(): string {
		return VIEW_TYPE_SPREADSHEET;
	}
	getDisplayText(): string {
		return this.file?.basename ?? "Spreadsheet";
	}
	getIcon(): string {
		return "table";
	}

	canAcceptExtension(extension: string): boolean {
		if (extension === "xlsx") return this.plugin.settings.enableXlsx;
		if (extension === "csv") return this.plugin.settings.enableCsv;
		return false;
	}

	async onLoadFile(file: TFile): Promise<void> {
		this.currentFile = file;
		this.activeSheet = 0;
		this.isDirty = false;
		this.selRow = -1;
		this.selCol = -1;

		if (!this.xlsx) {
			this.xlsx = (await import("xlsx")) as unknown as XLSXModule;
		}
		await this.renderFile(file);
	}

	// No async work needed — returns resolved promise for type compatibility
	onUnloadFile(_file: TFile): Promise<void> {
		this.contentEl.empty();
		this.workbook = null;
		this.savedData = null;
		this.wrapper = null;
		this.tabBar = null;
		this.tableContainer = null;
		this.formulaRef = null;
		this.formulaInput = null;
		this.saveBtn = null;
		this.undoBtn = null;
		this.editToggleBtn = null;
		this.dirtyIndicator = null;
		this.addRowBtn = null;
		this.addColBtn = null;
		this.infoEl = null;
		this.currentFile = null;
		this.editMode = false;
		return Promise.resolve();
	}

	// ── Render ────────────────────────────────────────────────────────────────

	private async renderFile(file: TFile): Promise<void> {
		this.contentEl.empty();
		if (!this.xlsx) return;

		const isBottom =
			this.plugin.settings.spreadsheetToolbarPosition === "bottom";
		const isCsv = file.extension === "csv";

		// Read raw bytes
		let data: ArrayBuffer;
		try {
			data = await this.app.vault.adapter.readBinary(file.path);
		} catch (err) {
			this.contentEl.createEl("p", {
				cls: "via-error",
				text: `Failed to read file: ${String(err)}`,
			});
			return;
		}

		// Keep a copy for undo/revert
		this.savedData = data.slice(0);

		// Parse workbook
		try {
			this.workbook = this.xlsx.read(new Uint8Array(data), {
				type: "array",
			});
		} catch (err) {
			this.contentEl.createEl("p", {
				cls: "via-error",
				text: `Failed to parse spreadsheet: ${String(err)}`,
			});
			return;
		}

		// Evaluate formulas on initial load
		this.evaluateAllFormulas();

		// Root wrapper
		this.wrapper = this.contentEl.createEl("div", {
			cls: "via-sheet-wrapper",
		});
		if (isBottom)
			this.wrapper.classList.add("via-sheet-wrapper--toolbar-bottom");

		// ── Toolbar ───────────────────────────────────────────────────────────
		const toolbar = this.wrapper.createEl("div", {
			cls: "via-sheet-toolbar",
		});

		// Edit / View toggle
		this.editToggleBtn = toolbar.createEl("div", { cls: "clickable-icon" });
		setIcon(this.editToggleBtn, "pencil");
		setTooltip(this.editToggleBtn, "Switch to edit mode");
		this.editToggleBtn.addEventListener("click", () =>
			this.toggleEditMode(),
		);

		toolbar.createEl("div", { cls: "via-toolbar-sep" });

		// Save button
		this.saveBtn = toolbar.createEl("div", {
			cls: "clickable-icon via-icon-save",
		});
		setIcon(this.saveBtn, "save");
		setTooltip(this.saveBtn, "Save (Ctrl+S)");
		this.saveBtn.classList.add("via-hidden");
		this.saveBtn.addEventListener("click", () => { void this.saveFile(); });

		// Undo (revert to last save)
		this.undoBtn = toolbar.createEl("div", { cls: "clickable-icon" });
		setIcon(this.undoBtn, "undo-2");
		setTooltip(this.undoBtn, "Revert to last save");
		this.undoBtn.classList.add("via-hidden");
		this.undoBtn.addEventListener("click", () => this.revertToSaved());

		// Dirty indicator (yellow dot)
		this.dirtyIndicator = toolbar.createEl("div", {
			cls: "via-docx-dirty-dot via-hidden",
		});
		setTooltip(this.dirtyIndicator, "Unsaved changes");

		toolbar.createEl("div", { cls: "via-toolbar-sep" });

		// Add Row
		this.addRowBtn = toolbar.createEl("div", { cls: "clickable-icon" });
		setIcon(this.addRowBtn, "plus-square");
		setTooltip(this.addRowBtn, "Add row at bottom");
		this.addRowBtn.classList.add("via-hidden");
		this.addRowBtn.addEventListener("click", () => this.addRow());

		// Add Column
		this.addColBtn = toolbar.createEl("div", { cls: "clickable-icon" });
		setIcon(this.addColBtn, "between-vertical-start");
		setTooltip(this.addColBtn, "Add column at right");
		this.addColBtn.classList.add("via-hidden");
		this.addColBtn.addEventListener("click", () => this.addColumn());

		toolbar.createEl("div", { cls: "via-toolbar-sep" });

		// File label
		const fileLabel = toolbar.createEl("div", {
			cls: "via-sheet-file-label",
		});
		setIcon(
			fileLabel.createEl("div", { cls: "clickable-icon" }),
			isCsv ? "file-text" : "table",
		);
		fileLabel.createEl("span", {
			text: file.basename,
			cls: "via-sheet-file-name",
		});

		toolbar.createEl("div", { cls: "via-toolbar-sep" });

		// Sheet tabs (multi-sheet xlsx)
		this.tabBar = toolbar.createEl("div", { cls: "via-sheet-tabs" });
		if (!isCsv && this.workbook.SheetNames.length > 1) {
			this.renderTabs();
		}

		toolbar.createEl("div", { cls: "via-toolbar-spacer" });

		// Row × Col info
		this.infoEl = toolbar.createEl("div", { cls: "via-sheet-info" });
		this.updateInfo();

		// ── Formula bar ───────────────────────────────────────────────────────
		const fbar = this.wrapper.createEl("div", {
			cls: "via-sheet-formula-bar",
		});
		this.formulaRef = fbar.createEl("div", {
			cls: "via-sheet-formula-ref",
		});
		fbar.createEl("div", { cls: "via-sheet-formula-sep" });
		this.formulaInput = fbar.createEl("input", {
			cls: "via-sheet-formula-input",
			type: "text",
			placeholder: "Select a cell to edit…",
		});
		this.formulaInput.addEventListener("keydown", (e: KeyboardEvent) => {
			if (e.key === "Enter") {
				e.preventDefault();
				this.commitFormulaBar();
			}
			if (e.key === "Escape") {
				e.preventDefault();
				this.refreshFormulaBar();
				this.formulaInput?.blur();
			}
		});

		// Ctrl/Cmd+S — save
		this.registerDomEvent(this.contentEl, "keydown", (e: KeyboardEvent) => {
			if ((e.ctrlKey || e.metaKey) && e.key === "s") {
				e.preventDefault();
				void this.saveFile();
			}
		});
		this.contentEl.tabIndex = 0;

		// ── Scroll + table ────────────────────────────────────────────────────
		const scrollEl = this.wrapper.createEl("div", {
			cls: "via-sheet-scroll",
		});
		this.tableContainer = scrollEl.createEl("div", {
			cls: "via-sheet-table-wrap",
		});
		this.renderSheet();
	}

	private renderTabs(): void {
		if (!this.tabBar || !this.workbook) return;
		this.tabBar.empty();
		for (let i = 0; i < this.workbook.SheetNames.length; i++) {
			const name = this.workbook.SheetNames[i] ?? `Sheet ${i + 1}`;
			const tab = this.tabBar.createEl("div", {
				cls: "via-sheet-tab",
				text: name,
			});
			if (i === this.activeSheet) tab.classList.add("is-active");
			setTooltip(tab, name);
			tab.addEventListener("click", () => {
				if (i === this.activeSheet) return;
				this.activeSheet = i;
				this.selRow = -1;
				this.selCol = -1;
				this.renderTabs();
				this.renderSheet();
				this.updateInfo();
				this.refreshFormulaBar();
			});
		}
	}

	private renderSheet(): void {
		if (!this.tableContainer || !this.workbook || !this.xlsx) return;
		this.tableContainer.empty();

		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) {
			this.tableContainer.createEl("p", {
				cls: "via-sheet-empty",
				text: "This sheet is empty.",
			});
			return;
		}

		const ref = sheetGetRef(sheet);
		if (!ref) {
			this.tableContainer.createEl("p", {
				cls: "via-sheet-empty",
				text: "This sheet is empty.",
			});
			return;
		}

		const range = this.xlsx.utils.decode_range(ref);
		const rowCount = range.e.r + 1;
		const colCount = range.e.c + 1;

		const table = this.tableContainer.createEl("table", {
			cls: "via-sheet-table",
		});

		// Column header row
		const thead = table.createEl("thead");
		const headerRow = thead.createEl("tr");
		headerRow.createEl("th", { cls: "via-sheet-row-num", text: "" });
		for (let c = 0; c < colCount; c++) {
			const th = headerRow.createEl("th", { text: colLetter(c) });
			if (c === this.selCol) th.classList.add("is-selected-col");
			// Context menu on column header (edit mode only)
			const colIdx = c;
			th.addEventListener("contextmenu", (e: MouseEvent) => {
				if (!this.editMode) return;
				e.preventDefault();
				this.showColumnContextMenu(e, colIdx, colCount);
			});
		}

		// Data rows
		const tbody = table.createEl("tbody");
		for (let r = 0; r < rowCount; r++) {
			const tr = tbody.createEl("tr");
			const rowNumTd = tr.createEl("td", {
				cls: "via-sheet-row-num",
				text: String(r + 1),
			});
			if (r === this.selRow) rowNumTd.classList.add("is-selected-row");
			// Context menu on row number (edit mode only)
			const rowIdx = r;
			rowNumTd.addEventListener("contextmenu", (e: MouseEvent) => {
				if (!this.editMode) return;
				e.preventDefault();
				this.showRowContextMenu(e, rowIdx, rowCount);
			});

			for (let c = 0; c < colCount; c++) {
				const cellRef = this.xlsx.utils.encode_cell({ r, c });
				const cell = sheetGetCell(sheet, cellRef);
				const td = tr.createEl("td");
				td.textContent = this.getCellDisplay(cell);
				if (r === this.selRow && c === this.selCol)
					td.classList.add("is-selected");

				td.addEventListener("click", () => this.selectCell(r, c));
				td.addEventListener("dblclick", () => {
					if (this.editMode) this.beginInlineEdit(td, r, c);
				});
			}
		}
	}

	private getCellDisplay(cell: XLSXCell | undefined): string {
		if (!cell) return "";
		if (cell.w !== undefined) return cell.w;
		if (cell.v !== undefined) return String(cell.v);
		return "";
	}

	// ── Selection ─────────────────────────────────────────────────────────────

	private selectCell(row: number, col: number): void {
		this.selRow = row;
		this.selCol = col;
		this.updateSelectionHighlight();
		this.refreshFormulaBar();
	}

	private updateSelectionHighlight(): void {
		if (!this.tableContainer) return;
		this.tableContainer
			.querySelectorAll(
				".is-selected, .is-selected-col, .is-selected-row",
			)
			.forEach((el) =>
				el.classList.remove(
					"is-selected",
					"is-selected-col",
					"is-selected-row",
				),
			);

		if (this.selRow < 0 || this.selCol < 0) return;
		const table = this.tableContainer.querySelector(".via-sheet-table");
		if (!table) return;

		// Column header highlight
		const ths = table.querySelectorAll("thead th");
		ths[this.selCol + 1]?.classList.add("is-selected-col");

		// Row header + cell highlight
		const rows = table.querySelectorAll("tbody tr");
		const row = rows[this.selRow];
		if (row) {
			row.querySelector(".via-sheet-row-num")?.classList.add(
				"is-selected-row",
			);
			row.querySelectorAll("td")[this.selCol + 1]?.classList.add(
				"is-selected",
			);
		}
	}

	// ── Formula bar ───────────────────────────────────────────────────────────

	private refreshFormulaBar(): void {
		if (!this.formulaRef || !this.formulaInput || !this.xlsx) return;
		if (this.selRow < 0 || this.selCol < 0) {
			this.formulaRef.textContent = "";
			this.formulaInput.value = "";
			return;
		}
		const cellRef = this.xlsx.utils.encode_cell({
			r: this.selRow,
			c: this.selCol,
		});
		this.formulaRef.textContent = cellRef;
		const sheetName = this.workbook?.SheetNames[this.activeSheet];
		const sheet = sheetName ? this.workbook?.Sheets[sheetName] : undefined;
		const cell = sheet ? sheetGetCell(sheet, cellRef) : undefined;
		this.formulaInput.value = cell?.f
			? "=" + cell.f
			: this.getCellDisplay(cell);
	}

	private commitFormulaBar(): void {
		if (!this.formulaInput || this.selRow < 0 || this.selCol < 0) return;
		this.writeCell(this.selRow, this.selCol, this.formulaInput.value);
		this.formulaInput.blur();
	}

	// ── Inline cell editing ───────────────────────────────────────────────────

	private beginInlineEdit(td: HTMLElement, row: number, col: number): void {
		if (!this.xlsx || !this.editMode) return;
		const sheetName = this.workbook?.SheetNames[this.activeSheet];
		const sheet = sheetName ? this.workbook?.Sheets[sheetName] : undefined;
		const cellRef = this.xlsx.utils.encode_cell({ r: row, c: col });
		const cell = sheet ? sheetGetCell(sheet, cellRef) : undefined;
		const rawValue = cell?.f ? "=" + cell.f : this.getCellDisplay(cell);

		td.empty();
		const input = td.createEl("input", {
			cls: "via-sheet-cell-input",
			type: "text",
		});
		input.value = rawValue;
		input.focus();
		input.select();

		let committed = false;
		const commit = (nextRow?: number, nextCol?: number) => {
			if (committed) return;
			committed = true;
			this.writeCell(row, col, input.value);
			// Navigate to next cell if specified
			if (nextRow !== undefined && nextCol !== undefined) {
				this.selectCell(nextRow, nextCol);
				// Auto-start editing the next cell
				const nextTd = this.getCellTd(nextRow, nextCol);
				if (nextTd) {
					setTimeout(
						() => this.beginInlineEdit(nextTd, nextRow, nextCol),
						0,
					);
				}
			}
		};

		input.addEventListener("keydown", (e: KeyboardEvent) => {
			if (e.key === "Enter") {
				e.preventDefault();
				commit(row + 1, col);
			}
			if (e.key === "Tab") {
				e.preventDefault();
				if (e.shiftKey) {
					commit(row, Math.max(0, col - 1));
				} else {
					commit(row, col + 1);
				}
			}
			if (e.key === "Escape") {
				e.preventDefault();
				committed = true;
				this.renderSheet();
				this.refreshFormulaBar();
			}
		});
		input.addEventListener("blur", () => commit());
	}

	/** Get the <td> element for a given row/col in the rendered table. */
	private getCellTd(row: number, col: number): HTMLElement | null {
		if (!this.tableContainer) return null;
		const rows = this.tableContainer.querySelectorAll("tbody tr");
		const tr = rows[row];
		if (!tr) return null;
		// +1 to skip the row number <td>
		const tds = tr.querySelectorAll("td");
		return (tds[col + 1] as HTMLElement) ?? null;
	}

	private writeCell(row: number, col: number, value: string): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const cellRef = this.xlsx.utils.encode_cell({ r: row, c: col });
		let cell: XLSXCell;
		if (value.startsWith("=")) {
			cell = { f: value.slice(1), t: "n", v: 0 };
		} else if (value.trim() !== "" && !isNaN(Number(value))) {
			cell = { v: Number(value), t: "n", w: value };
		} else {
			cell = { v: value, t: "s", w: value };
		}
		sheetSetCell(sheet, cellRef, cell);

		// Expand sheet !ref if needed
		const ref = sheetGetRef(sheet);
		if (!ref) {
			sheetSetRef(
				sheet,
				this.xlsx.utils.encode_range({
					s: { r: 0, c: 0 },
					e: { r: row, c: col },
				}),
			);
		} else {
			const range = this.xlsx.utils.decode_range(ref);
			if (row > range.e.r) range.e.r = row;
			if (col > range.e.c) range.e.c = col;
			sheetSetRef(sheet, this.xlsx.utils.encode_range(range));
		}

		// Re-evaluate all formulas so dependent cells update live
		this.evaluateAllFormulas();

		this.markDirty();
		this.renderSheet();
		this.refreshFormulaBar();
	}

	// ── Edit Mode Toggle ─────────────────────────────────────────────────────

	private evaluateAllFormulas(): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;
		evaluateSheet(sheet, this.xlsx.utils, sheetGetRef(sheet));
	}

	private toggleEditMode(): void {
		this.editMode = !this.editMode;
		if (!this.editToggleBtn) return;
		setIcon(this.editToggleBtn, this.editMode ? "eye" : "pencil");
		setTooltip(
			this.editToggleBtn,
			this.editMode ? "Switch to view mode" : "Switch to edit mode",
		);
		this.editToggleBtn.classList.toggle("is-active", this.editMode);

		// Show/hide edit-only toolbar buttons
		const hidden = !this.editMode;
		if (this.saveBtn) this.saveBtn.classList.toggle("via-hidden", hidden);
		if (this.undoBtn) this.undoBtn.classList.toggle("via-hidden", hidden);
		if (this.addRowBtn) this.addRowBtn.classList.toggle("via-hidden", hidden);
		if (this.addColBtn) this.addColBtn.classList.toggle("via-hidden", hidden);

		// Hide dirty indicator when leaving edit mode
		if (!this.editMode) {
			this.setDirty(false);
		}
	}

	// ── Add Row / Column ──────────────────────────────────────────────────────

	private addRow(): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const ref = sheetGetRef(sheet) ?? "A1:A1";
		const range = this.xlsx.utils.decode_range(ref);
		range.e.r += 1;
		sheetSetRef(sheet, this.xlsx.utils.encode_range(range));

		this.markDirty();
		this.renderSheet();
		this.updateInfo();
		new Notice("Row added", 1500);
	}

	private addColumn(): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const ref = sheetGetRef(sheet) ?? "A1:A1";
		const range = this.xlsx.utils.decode_range(ref);
		range.e.c += 1;
		sheetSetRef(sheet, this.xlsx.utils.encode_range(range));

		this.markDirty();
		this.renderSheet();
		this.updateInfo();
		new Notice("Column added", 1500);
	}

	// ── Insert / Delete Row ───────────────────────────────────────────────────

	private insertRowAt(index: number): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const ref = sheetGetRef(sheet) ?? "A1:A1";
		const range = this.xlsx.utils.decode_range(ref);
		const colCount = range.e.c + 1;

		// Shift rows down from bottom to insertion point
		for (let r = range.e.r; r >= index; r--) {
			for (let c = 0; c < colCount; c++) {
				const srcRef = this.xlsx.utils.encode_cell({ r, c });
				const dstRef = this.xlsx.utils.encode_cell({ r: r + 1, c });
				const cell = sheetGetCell(sheet, srcRef);
				if (cell) {
					sheetSetCell(sheet, dstRef, cell);
				} else {
					delete (sheet as Record<string, unknown>)[dstRef];
				}
			}
		}

		// Clear the inserted row
		for (let c = 0; c < colCount; c++) {
			const ref = this.xlsx.utils.encode_cell({ r: index, c });
			delete (sheet as Record<string, unknown>)[ref];
		}

		range.e.r += 1;
		sheetSetRef(sheet, this.xlsx.utils.encode_range(range));
		this.markDirty();
		this.renderSheet();
		this.updateInfo();
	}

	private deleteRowAt(index: number): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const ref = sheetGetRef(sheet) ?? "A1:A1";
		const range = this.xlsx.utils.decode_range(ref);
		if (range.e.r <= 0) return; // Don't delete the last row
		const colCount = range.e.c + 1;

		// Shift rows up
		for (let r = index; r < range.e.r; r++) {
			for (let c = 0; c < colCount; c++) {
				const srcRef = this.xlsx.utils.encode_cell({ r: r + 1, c });
				const dstRef = this.xlsx.utils.encode_cell({ r, c });
				const cell = sheetGetCell(sheet, srcRef);
				if (cell) {
					sheetSetCell(sheet, dstRef, cell);
				} else {
					delete (sheet as Record<string, unknown>)[dstRef];
				}
			}
		}

		// Clear last row
		for (let c = 0; c < colCount; c++) {
			const cellRef = this.xlsx.utils.encode_cell({ r: range.e.r, c });
			delete (sheet as Record<string, unknown>)[cellRef];
		}

		range.e.r -= 1;
		sheetSetRef(sheet, this.xlsx.utils.encode_range(range));
		this.selRow = -1;
		this.selCol = -1;
		this.markDirty();
		this.renderSheet();
		this.updateInfo();
		this.refreshFormulaBar();
	}

	// ── Insert / Delete Column ────────────────────────────────────────────────

	private insertColumnAt(index: number): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const ref = sheetGetRef(sheet) ?? "A1:A1";
		const range = this.xlsx.utils.decode_range(ref);
		const rowCount = range.e.r + 1;

		// Shift columns right from rightmost to insertion point
		for (let c = range.e.c; c >= index; c--) {
			for (let r = 0; r < rowCount; r++) {
				const srcRef = this.xlsx.utils.encode_cell({ r, c });
				const dstRef = this.xlsx.utils.encode_cell({ r, c: c + 1 });
				const cell = sheetGetCell(sheet, srcRef);
				if (cell) {
					sheetSetCell(sheet, dstRef, cell);
				} else {
					delete (sheet as Record<string, unknown>)[dstRef];
				}
			}
		}

		// Clear inserted column
		for (let r = 0; r < rowCount; r++) {
			const cellRef = this.xlsx.utils.encode_cell({ r, c: index });
			delete (sheet as Record<string, unknown>)[cellRef];
		}

		range.e.c += 1;
		sheetSetRef(sheet, this.xlsx.utils.encode_range(range));
		this.markDirty();
		this.renderSheet();
		this.updateInfo();
	}

	private deleteColumnAt(index: number): void {
		if (!this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) return;
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) return;

		const ref = sheetGetRef(sheet) ?? "A1:A1";
		const range = this.xlsx.utils.decode_range(ref);
		if (range.e.c <= 0) return; // Don't delete the last column
		const rowCount = range.e.r + 1;

		// Shift columns left
		for (let c = index; c < range.e.c; c++) {
			for (let r = 0; r < rowCount; r++) {
				const srcRef = this.xlsx.utils.encode_cell({ r, c: c + 1 });
				const dstRef = this.xlsx.utils.encode_cell({ r, c });
				const cell = sheetGetCell(sheet, srcRef);
				if (cell) {
					sheetSetCell(sheet, dstRef, cell);
				} else {
					delete (sheet as Record<string, unknown>)[dstRef];
				}
			}
		}

		// Clear last column
		for (let r = 0; r < rowCount; r++) {
			const cellRef = this.xlsx.utils.encode_cell({ r, c: range.e.c });
			delete (sheet as Record<string, unknown>)[cellRef];
		}

		range.e.c -= 1;
		sheetSetRef(sheet, this.xlsx.utils.encode_range(range));
		this.selRow = -1;
		this.selCol = -1;
		this.markDirty();
		this.renderSheet();
		this.updateInfo();
		this.refreshFormulaBar();
	}

	// ── Context Menus ─────────────────────────────────────────────────────────

	private showRowContextMenu(
		e: MouseEvent,
		row: number,
		rowCount: number,
	): void {
		const menu = new Menu();
		menu.addItem((item) => {
			item.setTitle("Insert row above")
				.setIcon("arrow-up")
				.onClick(() => this.insertRowAt(row));
		});
		menu.addItem((item) => {
			item.setTitle("Insert row below")
				.setIcon("arrow-down")
				.onClick(() => this.insertRowAt(row + 1));
		});
		if (rowCount > 1) {
			menu.addSeparator();
			menu.addItem((item) => {
				item.setTitle("Delete row")
					.setIcon("trash-2")
					.onClick(() => this.deleteRowAt(row));
			});
		}
		menu.showAtMouseEvent(e);
	}

	private showColumnContextMenu(
		e: MouseEvent,
		col: number,
		colCount: number,
	): void {
		const menu = new Menu();
		menu.addItem((item) => {
			item.setTitle("Insert column left")
				.setIcon("arrow-left")
				.onClick(() => this.insertColumnAt(col));
		});
		menu.addItem((item) => {
			item.setTitle("Insert column right")
				.setIcon("arrow-right")
				.onClick(() => this.insertColumnAt(col + 1));
		});
		if (colCount > 1) {
			menu.addSeparator();
			menu.addItem((item) => {
				item.setTitle("Delete column")
					.setIcon("trash-2")
					.onClick(() => this.deleteColumnAt(col));
			});
		}
		menu.showAtMouseEvent(e);
	}

	// ── Dirty state & Save ────────────────────────────────────────────────────

	private markDirty(): void {
		this.isDirty = true;
		this.saveBtn?.classList.add("is-dirty");
		this.setDirty(true);
	}

	private setDirty(dirty: boolean): void {
		this.isDirty = dirty;
		if (this.dirtyIndicator)
			this.dirtyIndicator.classList.toggle("via-hidden", !dirty);
		if (!dirty) this.saveBtn?.classList.remove("is-dirty");
	}

	private async saveFile(): Promise<void> {
		if (!this.workbook || !this.xlsx || !this.currentFile) return;

		if (this.plugin.settings.confirmOnSave) {
			const confirmed = await confirmModal(
				this.app,
				`Overwrite "${this.currentFile.name}"?`,
				"This will replace the original file with the current spreadsheet data.",
			);
			if (!confirmed) return;
		}

		const bookType = this.currentFile.extension === "csv" ? "csv" : "xlsx";
		try {
			// Use type:'binary' to avoid readSlice errors in SheetJS CFB writer
			const out = this.xlsx.write(this.workbook, {
				type: "binary",
				bookType,
			}) as string;
			const ab = binaryStringToArrayBuffer(out);
			await this.app.vault.modifyBinary(this.currentFile, ab);
			// Update saved snapshot
			this.savedData = ab.slice(0);
			this.setDirty(false);
			new Notice("Saved", 2000);
		} catch (err) {
			new Notice(`Save failed: ${String(err)}`);
		}
	}

	private revertToSaved(): void {
		if (!this.savedData || !this.xlsx) return;
		try {
			this.workbook = this.xlsx.read(new Uint8Array(this.savedData), {
				type: "array",
			});
			this.setDirty(false);
			this.selRow = -1;
			this.selCol = -1;
			this.renderSheet();
			this.updateInfo();
			this.refreshFormulaBar();
			new Notice("Reverted to last save", 1500);
		} catch (err) {
			new Notice(`Revert failed: ${String(err)}`);
		}
	}

	private updateInfo(): void {
		if (!this.infoEl || !this.workbook || !this.xlsx) return;
		const sheetName = this.workbook.SheetNames[this.activeSheet];
		if (!sheetName) {
			this.infoEl.textContent = "—";
			return;
		}
		const sheet = this.workbook.Sheets[sheetName];
		if (!sheet) {
			this.infoEl.textContent = "—";
			return;
		}
		const ref = sheetGetRef(sheet);
		if (!ref) {
			this.infoEl.textContent = "0 rows";
			return;
		}
		const range = this.xlsx.utils.decode_range(ref);
		this.infoEl.textContent = `${range.e.r + 1} × ${range.e.c + 1}`;
	}
}

// ── Confirm modal ─────────────────────────────────────────────────────────────

function confirmModal(
	app: App,
	title: string,
	message: string,
): Promise<boolean> {
	return new Promise((resolve) => {
		const modal = new ConfirmModal(app, title, message, resolve);
		modal.open();
	});
}

class ConfirmModal extends Modal {
	constructor(
		app: App,
		private titleText: string,
		private message: string,
		private resolve: (v: boolean) => void,
	) {
		super(app);
	}

	onOpen(): void {
		this.setTitle(this.titleText);
		const { contentEl } = this;
		contentEl.createEl("p", { text: this.message });
		const btnRow = contentEl.createEl("div", {
			cls: "modal-button-container",
		});
		btnRow
			.createEl("button", { text: "Cancel" })
			.addEventListener("click", () => {
				this.resolve(false);
				this.close();
			});
		const overwriteBtn = btnRow.createEl("button", {
			text: "Overwrite",
			cls: "mod-cta via-btn-danger",
		});
		overwriteBtn.addEventListener("click", () => {
			this.resolve(true);
			this.close();
		});
	}

	onClose(): void {
		this.contentEl.empty();
	}
}

// ── Helpers ───────────────────────────────────────────────────────────────────

/** Convert 0-based column index to spreadsheet letter (A, B, …, Z, AA, AB…). */
function colLetter(index: number): string {
	let s = "";
	let n = index;
	while (n >= 0) {
		s = String.fromCharCode((n % 26) + 65) + s;
		n = Math.floor(n / 26) - 1;
	}
	return s;
}

/** Convert a SheetJS binary string to an ArrayBuffer (avoids CFB readSlice bug). */
function binaryStringToArrayBuffer(bin: string): ArrayBuffer {
	const buf = new ArrayBuffer(bin.length);
	const view = new Uint8Array(buf);
	for (let i = 0; i < bin.length; i++) view[i] = bin.charCodeAt(i) & 0xff;
	return buf;
}
