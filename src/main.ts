import { App, Plugin } from "obsidian";
import {
	DEFAULT_SETTINGS,
	PluginSettings,
	ViewItAllSettingTab,
} from "./settings";
import {
	VIEW_TYPE_DOCX,
	VIEW_TYPE_PDF,
	VIEW_TYPE_SPREADSHEET,
	VIEW_TYPE_PPTX,
} from "./types";
import { DocxView } from "./views/DocxView";
import { PdfView } from "./views/PdfView";
import { SpreadsheetView } from "./views/SpreadsheetView";
import { PptxView } from "./views/PptxView";

interface AppWithRegistry extends App {
	viewRegistry: {
		unregisterExtensions(exts: string[]): void;
		registerExtensions(exts: string[], viewType: string): void;
	};
}

export default class ViewItAllPlugin extends Plugin {
	settings: PluginSettings;

	async onload() {
		await this.loadSettings();

		// ── Register view types ───────────────────────────────────────────
		this.registerView(VIEW_TYPE_DOCX, (leaf) => new DocxView(leaf, this));
		this.registerView(VIEW_TYPE_PDF, (leaf) => new PdfView(leaf, this));
		this.registerView(
			VIEW_TYPE_SPREADSHEET,
			(leaf) => new SpreadsheetView(leaf, this),
		);
		this.registerView(VIEW_TYPE_PPTX, (leaf) => new PptxView(leaf, this));

		// ── Register file extensions (respecting enable toggles) ──────────
		if (this.settings.enableDocx) {
			this.registerExtensions(["docx"], VIEW_TYPE_DOCX);
		}

		if (this.settings.enablePdf) {
			this.overridePdfExtension();
		}

		const sheetExts: string[] = [];
		if (this.settings.enableXlsx) sheetExts.push("xlsx");
		if (this.settings.enableCsv) sheetExts.push("csv");
		if (sheetExts.length > 0) {
			this.registerExtensions(sheetExts, VIEW_TYPE_SPREADSHEET);
		}

		if (this.settings.enablePptx) {
			this.registerExtensions(["pptx"], VIEW_TYPE_PPTX);
		}

		this.addSettingTab(new ViewItAllSettingTab(this.app, this));
	}

	onunload() {
		// Restore Obsidian's built-in PDF viewer when this plugin is disabled.
		if (this.settings.enablePdf) {
			try {
				const reg = (this.app as AppWithRegistry).viewRegistry;
				reg.unregisterExtensions(["pdf"]);
				reg.registerExtensions(["pdf"], "pdf");
			} catch {
				// ignore — built-in will be re-registered on app restart
			}
		}
	}

	private overridePdfExtension() {
		try {
			(this.app as AppWithRegistry).viewRegistry.unregisterExtensions([
				"pdf",
			]);
		} catch {
			// already unregistered or API unavailable — proceed anyway
		}
		this.registerExtensions(["pdf"], VIEW_TYPE_PDF);
	}

	async loadSettings() {
		this.settings = Object.assign(
			{},
			DEFAULT_SETTINGS,
			(await this.loadData()) as Partial<PluginSettings>,
		);
	}

	async saveSettings() {
		await this.saveData(this.settings);
	}
}
