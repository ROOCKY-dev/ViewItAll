import { App, PluginSettingTab, Setting } from "obsidian";
import type ViewItAllPlugin from "./main";

export type OpenMode = "tab" | "sidebar-right";
export type ToolbarPosition = "top" | "bottom";
export type SnapModifier = "Alt" | "Shift";
export type SnapDirection = "horizontal" | "vertical" | "slope";
export type PptxRendererMode = "custom" | "library";

export interface PluginSettings {
	// ── File-type toggles ─────────────────────────────────────────────────
	enablePdf: boolean;
	enableDocx: boolean;
	enableXlsx: boolean;
	enableCsv: boolean;
	enablePptx: boolean;

	// ── DOCX ─────────────────────────────────────────────────────────────
	docxOpenMode: OpenMode;
	docxToolbarPosition: ToolbarPosition;
	docxDefaultEditMode: boolean;
	confirmOnSave: boolean;

	// ── Spreadsheet (.xlsx / .csv) ────────────────────────────────────────
	spreadsheetToolbarPosition: ToolbarPosition;

	// ── PowerPoint (.pptx) ────────────────────────────────────────────────
	pptxToolbarPosition: ToolbarPosition;
	pptxRendererMode: PptxRendererMode;

	// ── PDF — general ─────────────────────────────────────────────────────
	pdfOpenMode: OpenMode;
	pdfToolbarPosition: ToolbarPosition;
	pdfDefaultTool: "none" | "pen" | "highlighter";
	pdfDefaultZoom: number;
	showTocOnOpen: boolean;

	// ── PDF — annotation tools ────────────────────────────────────────────
	penColor: string;
	penWidth: number;
	highlighterColor: string;
	highlighterWidth: number;
	highlighterOpacity: number;
	eraserWidth: number;

	// ── PDF — notes ───────────────────────────────────────────────────────
	noteDefaultColor: string;

	// ── PDF — snap ────────────────────────────────────────────────────────
	snapActivateKey: SnapModifier;
	snapDefaultDirection: SnapDirection;

	// ── Keyboard shortcuts (PDF) ──────────────────────────────────────────
	keyToolView: string; // default 'v'
	keyToolPen: string; // default 'p'
	keyToolHighlight: string; // default 'h'
	keyToolErase: string; // default 'e'
	keyToolNote: string; // default 'n'
	keySnapCycle: string; // default 's'  (with snapActivateKey held)
	keySearch: string; // default 'f'  (with Ctrl/Cmd held)
}

export const DEFAULT_SETTINGS: PluginSettings = {
	enablePdf: true,
	enableDocx: true,
	enableXlsx: true,
	enableCsv: true,
	enablePptx: true,

	docxOpenMode: "tab",
	docxToolbarPosition: "top",
	docxDefaultEditMode: false,
	confirmOnSave: true,

	spreadsheetToolbarPosition: "top",

	pptxToolbarPosition: "top",
	pptxRendererMode: "library",

	pdfOpenMode: "tab",
	pdfToolbarPosition: "top",
	pdfDefaultTool: "none",
	pdfDefaultZoom: 1.0,
	showTocOnOpen: false,

	penColor: "#e03131",
	penWidth: 2,
	highlighterColor: "#ffd43b",
	highlighterWidth: 16,
	highlighterOpacity: 0.4,
	eraserWidth: 20,

	noteDefaultColor: "#ffd43b",

	snapActivateKey: "Alt",
	snapDefaultDirection: "horizontal",

	keyToolView: "v",
	keyToolPen: "p",
	keyToolHighlight: "h",
	keyToolErase: "e",
	keyToolNote: "n",
	keySnapCycle: "s",
	keySearch: "f",
};

// ── Helper ────────────────────────────────────────────────────────────────────
/** Validate and normalise a single-character shortcut key entered by the user. */
function normKey(raw: string): string {
	return raw.trim().toLowerCase().slice(0, 1) || "";
}

export class ViewItAllSettingTab extends PluginSettingTab {
	plugin: ViewItAllPlugin;

	constructor(app: App, plugin: ViewItAllPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		// ── File Types ───────────────────────────────────────────────────────
		new Setting(containerEl).setName("File types").setHeading();
		containerEl.createEl("p", {
			text: "Enable or disable support for each file type. Changes take effect after reloading Obsidian.",
			cls: "setting-item-description",
		});

		new Setting(containerEl).setName("PDF (.pdf)").addToggle((t) =>
			t.setValue(this.plugin.settings.enablePdf).onChange(async (v) => {
				this.plugin.settings.enablePdf = v;
				await this.plugin.saveSettings();
			}),
		);

		new Setting(containerEl).setName("Word (.docx)").addToggle((t) =>
			t.setValue(this.plugin.settings.enableDocx).onChange(async (v) => {
				this.plugin.settings.enableDocx = v;
				await this.plugin.saveSettings();
			}),
		);

		new Setting(containerEl).setName("Excel (.xlsx)").addToggle((t) =>
			t.setValue(this.plugin.settings.enableXlsx).onChange(async (v) => {
				this.plugin.settings.enableXlsx = v;
				await this.plugin.saveSettings();
			}),
		);

		new Setting(containerEl).setName("CSV (.csv)").addToggle((t) =>
			t.setValue(this.plugin.settings.enableCsv).onChange(async (v) => {
				this.plugin.settings.enableCsv = v;
				await this.plugin.saveSettings();
			}),
		);

		new Setting(containerEl).setName("PowerPoint (.pptx)").addToggle((t) =>
			t.setValue(this.plugin.settings.enablePptx).onChange(async (v) => {
				this.plugin.settings.enablePptx = v;
				await this.plugin.saveSettings();
			}),
		);

		// ── DOCX ─────────────────────────────────────────────────────────────
		new Setting(containerEl).setName("Word documents (.docx)").setHeading();

		new Setting(containerEl)
			.setName("Open mode")
			.setDesc("Where to open .docx files.")
			.addDropdown((dd) =>
				dd
					.addOption("tab", "New tab")
					.addOption("sidebar-right", "Right sidebar")
					.setValue(this.plugin.settings.docxOpenMode)
					.onChange(async (v) => {
						this.plugin.settings.docxOpenMode = v as OpenMode;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Open in edit mode by default")
			.setDesc("When enabled, .docx files open ready to edit.")
			.addToggle((t) =>
				t
					.setValue(this.plugin.settings.docxDefaultEditMode)
					.onChange(async (v) => {
						this.plugin.settings.docxDefaultEditMode = v;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Toolbar position")
			.setDesc("Where to pin the toolbar.")
			.addDropdown((dd) =>
				dd
					.addOption("top", "Top")
					.addOption("bottom", "Bottom")
					.setValue(this.plugin.settings.docxToolbarPosition)
					.onChange(async (v) => {
						this.plugin.settings.docxToolbarPosition =
							v as ToolbarPosition;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Confirm before saving")
			.setDesc(
				"Show a confirmation dialog before overwriting the original .docx file.",
			)
			.addToggle((t) =>
				t
					.setValue(this.plugin.settings.confirmOnSave)
					.onChange(async (v) => {
						this.plugin.settings.confirmOnSave = v;
						await this.plugin.saveSettings();
					}),
			);

		// ── Spreadsheet ──────────────────────────────────────────────────
		new Setting(containerEl)
			.setName("Spreadsheets (.xlsx / .csv)")
			.setHeading();

		new Setting(containerEl)
			.setName("Toolbar position")
			.setDesc("Where to pin the spreadsheet toolbar.")
			.addDropdown((dd) =>
				dd
					.addOption("top", "Top")
					.addOption("bottom", "Bottom")
					.setValue(this.plugin.settings.spreadsheetToolbarPosition)
					.onChange(async (v) => {
						this.plugin.settings.spreadsheetToolbarPosition =
							v as ToolbarPosition;
						await this.plugin.saveSettings();
					}),
			);

		// ── PowerPoint ────────────────────────────────────────────────────
		new Setting(containerEl).setName("PowerPoint (.pptx)").setHeading();

		new Setting(containerEl)
			.setName("Toolbar position")
			.setDesc("Where to pin the slide toolbar.")
			.addDropdown((dd) =>
				dd
					.addOption("top", "Top")
					.addOption("bottom", "Bottom")
					.setValue(this.plugin.settings.pptxToolbarPosition)
					.onChange(async (v) => {
						this.plugin.settings.pptxToolbarPosition =
							v as ToolbarPosition;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Renderer mode")
			.setDesc(
				"Choose PPTX rendering engine. Library mode prioritizes visual fidelity; custom mode keeps in-plugin editable shape pipeline.",
			)
			.addDropdown((dd) =>
				dd
					.addOption("library", "Library (high fidelity)")
					.addOption("custom", "Custom (editable pipeline)")
					.setValue(this.plugin.settings.pptxRendererMode)
					.onChange(async (v) => {
						this.plugin.settings.pptxRendererMode =
							v as PptxRendererMode;
						await this.plugin.saveSettings();
					}),
			);

		// ── PDF — General ─────────────────────────────────────────────────
		new Setting(containerEl)
			.setName("Open mode")
			.setDesc("Where to open PDF files.")
			.addDropdown((dd) =>
				dd
					.addOption("tab", "New tab")
					.addOption("sidebar-right", "Right sidebar")
					.setValue(this.plugin.settings.pdfOpenMode)
					.onChange(async (v) => {
						this.plugin.settings.pdfOpenMode = v as OpenMode;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Toolbar position")
			.setDesc("Where to pin the PDF toolbar.")
			.addDropdown((dd) =>
				dd
					.addOption("top", "Top")
					.addOption("bottom", "Bottom")
					.setValue(this.plugin.settings.pdfToolbarPosition)
					.onChange(async (v) => {
						this.plugin.settings.pdfToolbarPosition =
							v as ToolbarPosition;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Default zoom")
			.setDesc("Zoom level when a PDF is first opened.")
			.addDropdown((dd) =>
				dd
					.addOption("0.5", "50%")
					.addOption("0.75", "75%")
					.addOption("1.0", "100%")
					.addOption("1.25", "125%")
					.addOption("1.5", "150%")
					.addOption("2.0", "200%")
					.setValue(String(this.plugin.settings.pdfDefaultZoom))
					.onChange(async (v) => {
						this.plugin.settings.pdfDefaultZoom = parseFloat(v);
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Default annotation tool")
			.setDesc("Tool to activate when a PDF is opened.")
			.addDropdown((dd) =>
				dd
					.addOption("none", "None (view only)")
					.addOption("pen", "Pen")
					.addOption("highlighter", "Highlighter")
					.setValue(this.plugin.settings.pdfDefaultTool)
					.onChange(async (v) => {
						this.plugin.settings.pdfDefaultTool = v as
							| "none"
							| "pen"
							| "highlighter";
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Show table of contents on open")
			.setDesc(
				"Automatically expand the table of contents sidebar when a PDF is opened (only if the PDF has an outline).",
			)
			.addToggle((t) =>
				t
					.setValue(this.plugin.settings.showTocOnOpen)
					.onChange(async (v) => {
						this.plugin.settings.showTocOnOpen = v;
						await this.plugin.saveSettings();
					}),
			);

		// ── PDF — Annotation Tools ────────────────────────────────────────
		new Setting(containerEl)
			.setName("PDF files — annotation tools")
			.setHeading();

		new Setting(containerEl).setName("Pen color").addColorPicker((cp) =>
			cp.setValue(this.plugin.settings.penColor).onChange(async (v) => {
				this.plugin.settings.penColor = v;
				await this.plugin.saveSettings();
			}),
		);

		new Setting(containerEl).setName("Pen width").addSlider((sl) =>
			sl
				.setLimits(1, 20, 1)
				.setValue(this.plugin.settings.penWidth)
				.setDynamicTooltip()
				.onChange(async (v) => {
					this.plugin.settings.penWidth = v;
					await this.plugin.saveSettings();
				}),
		);

		new Setting(containerEl)
			.setName("Highlighter color")
			.addColorPicker((cp) =>
				cp
					.setValue(this.plugin.settings.highlighterColor)
					.onChange(async (v) => {
						this.plugin.settings.highlighterColor = v;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl).setName("Highlighter width").addSlider((sl) =>
			sl
				.setLimits(10, 40, 2)
				.setValue(this.plugin.settings.highlighterWidth)
				.setDynamicTooltip()
				.onChange(async (v) => {
					this.plugin.settings.highlighterWidth = v;
					await this.plugin.saveSettings();
				}),
		);

		new Setting(containerEl)
			.setName("Highlighter opacity")
			.addSlider((sl) =>
				sl
					.setLimits(0.1, 1.0, 0.05)
					.setValue(this.plugin.settings.highlighterOpacity)
					.setDynamicTooltip()
					.onChange(async (v) => {
						this.plugin.settings.highlighterOpacity = v;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl).setName("Eraser width").addSlider((sl) =>
			sl
				.setLimits(5, 60, 5)
				.setValue(this.plugin.settings.eraserWidth)
				.setDynamicTooltip()
				.onChange(async (v) => {
					this.plugin.settings.eraserWidth = v;
					await this.plugin.saveSettings();
				}),
		);

		new Setting(containerEl)
			.setName("Default note color")
			.setDesc("Background color for newly placed text notes.")
			.addColorPicker((cp) =>
				cp
					.setValue(this.plugin.settings.noteDefaultColor)
					.onChange(async (v) => {
						this.plugin.settings.noteDefaultColor = v;
						await this.plugin.saveSettings();
					}),
			);

		// ── PDF — Snap ────────────────────────────────────────────────────
		new Setting(containerEl).setName("PDF files — snap").setHeading();

		new Setting(containerEl)
			.setName("Snap activation key")
			.setDesc(
				"Hold this key while drawing to constrain the stroke direction.",
			)
			.addDropdown((dd) =>
				dd
					.addOption("Alt", "Alt")
					.addOption("Shift", "Shift")
					.setValue(this.plugin.settings.snapActivateKey)
					.onChange(async (v) => {
						this.plugin.settings.snapActivateKey =
							v as SnapModifier;
						await this.plugin.saveSettings();
					}),
			);

		new Setting(containerEl)
			.setName("Default snap direction")
			.setDesc("Snap direction applied when opening a PDF.")
			.addDropdown((dd) =>
				dd
					.addOption("horizontal", "⟷ horizontal")
					.addOption("vertical", "↕ vertical")
					.addOption("slope", "↗ 45°")
					.setValue(this.plugin.settings.snapDefaultDirection)
					.onChange(async (v) => {
						this.plugin.settings.snapDefaultDirection =
							v as SnapDirection;
						await this.plugin.saveSettings();
					}),
			);

		// ── Keyboard Shortcuts (PDF) ──────────────────────────────────────
		new Setting(containerEl).setName("PDF keyboard shortcuts").setHeading();
		containerEl.createEl("p", {
			text: "Single character keys (no modifier). Avoid letters already used by Obsidian globally.",
			cls: "setting-item-description",
		});

		const shortcutEntry = (
			name: string,
			desc: string,
			get: () => string,
			set: (v: string) => void,
		) => {
			new Setting(containerEl)
				.setName(name)
				.setDesc(desc)
				.addText((t) =>
					t
						.setValue(get())
						.setPlaceholder("Single key")
						.onChange(async (raw) => {
							const k = normKey(raw);
							if (k) {
								set(k);
								t.setValue(k);
								await this.plugin.saveSettings();
							}
						}),
				);
		};

		shortcutEntry(
			"View tool",
			"Switch to view/pan mode.",
			() => this.plugin.settings.keyToolView,
			(v) => {
				this.plugin.settings.keyToolView = v;
			},
		);
		shortcutEntry(
			"Pen tool",
			"Switch to freehand pen.",
			() => this.plugin.settings.keyToolPen,
			(v) => {
				this.plugin.settings.keyToolPen = v;
			},
		);
		shortcutEntry(
			"Highlighter tool",
			"Switch to highlighter.",
			() => this.plugin.settings.keyToolHighlight,
			(v) => {
				this.plugin.settings.keyToolHighlight = v;
			},
		);
		shortcutEntry(
			"Eraser tool",
			"Switch to eraser.",
			() => this.plugin.settings.keyToolErase,
			(v) => {
				this.plugin.settings.keyToolErase = v;
			},
		);
		shortcutEntry(
			"Note tool",
			"Switch to sticky-note placement.",
			() => this.plugin.settings.keyToolNote,
			(v) => {
				this.plugin.settings.keyToolNote = v;
			},
		);
		shortcutEntry(
			"Cycle snap direction",
			`Press together with the snap activation key (e.g. ${this.plugin.settings.snapActivateKey}+key) to cycle snap mode.`,
			() => this.plugin.settings.keySnapCycle,
			(v) => {
				this.plugin.settings.keySnapCycle = v;
			},
		);
		shortcutEntry(
			"Open search bar",
			"Press with Ctrl/Cmd to open the text search bar.",
			() => this.plugin.settings.keySearch,
			(v) => {
				this.plugin.settings.keySearch = v;
			},
		);
	}
}
