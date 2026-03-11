import { App, PluginSettingTab, Setting } from 'obsidian';
import type ViewItAllPlugin from './main';

export type OpenMode = 'tab' | 'sidebar-right';
export type ToolbarPosition = 'top' | 'bottom';

export interface PluginSettings {
	docxOpenMode: OpenMode;
	docxToolbarPosition: ToolbarPosition;
	docxDefaultEditMode: boolean;
	confirmOnSave: boolean;
	pdfOpenMode: OpenMode;
	pdfToolbarPosition: ToolbarPosition;
	pdfDefaultTool: 'none' | 'pen' | 'highlighter';
	penColor: string;
	penWidth: number;
	highlighterColor: string;
	highlighterWidth: number;
	highlighterOpacity: number;
	eraserWidth: number;
}

export const DEFAULT_SETTINGS: PluginSettings = {
	docxOpenMode: 'tab',
	docxToolbarPosition: 'top',
	docxDefaultEditMode: false,
	confirmOnSave: true,
	pdfOpenMode: 'tab',
	pdfToolbarPosition: 'top',
	pdfDefaultTool: 'none',
	penColor: '#e03131',
	penWidth: 2,
	highlighterColor: '#ffd43b',
	highlighterWidth: 16,
	highlighterOpacity: 0.4,
	eraserWidth: 20,
};

export class ViewItAllSettingTab extends PluginSettingTab {
	plugin: ViewItAllPlugin;

	constructor(app: App, plugin: ViewItAllPlugin) {
		super(app, plugin);
		this.plugin = plugin;
	}

	display(): void {
		const { containerEl } = this;
		containerEl.empty();

		containerEl.createEl('h2', { text: 'ViewItAll Settings' });

		// ── DOCX ────────────────────────────────────────────────────────────
		containerEl.createEl('h3', { text: 'Word Documents (.docx)' });

		new Setting(containerEl)
			.setName('Open mode')
			.setDesc('Where to open .docx files.')
			.addDropdown(dd =>
				dd
					.addOption('tab', 'New tab')
					.addOption('sidebar-right', 'Right sidebar')
					.setValue(this.plugin.settings.docxOpenMode)
					.onChange(async v => {
						this.plugin.settings.docxOpenMode = v as OpenMode;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Open in edit mode by default')
			.setDesc('When enabled, .docx files open ready to edit.')
			.addToggle(t =>
				t
					.setValue(this.plugin.settings.docxDefaultEditMode)
					.onChange(async v => {
						this.plugin.settings.docxDefaultEditMode = v;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Toolbar position')
			.setDesc('Where to pin the DOCX toolbar.')
			.addDropdown(dd =>
				dd
					.addOption('top', 'Top')
					.addOption('bottom', 'Bottom')
					.setValue(this.plugin.settings.docxToolbarPosition)
					.onChange(async v => {
						this.plugin.settings.docxToolbarPosition = v as ToolbarPosition;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Confirm before saving')
			.setDesc('Show a confirmation dialog before overwriting the original .docx file.')
			.addToggle(t =>
				t
					.setValue(this.plugin.settings.confirmOnSave)
					.onChange(async v => {
						this.plugin.settings.confirmOnSave = v;
						await this.plugin.saveSettings();
					})
			);

		// ── PDF ─────────────────────────────────────────────────────────────
		containerEl.createEl('h3', { text: 'PDF Files (.pdf)' });

		new Setting(containerEl)
			.setName('Open mode')
			.setDesc('Where to open .pdf files.')
			.addDropdown(dd =>
				dd
					.addOption('tab', 'New tab')
					.addOption('sidebar-right', 'Right sidebar')
					.setValue(this.plugin.settings.pdfOpenMode)
					.onChange(async v => {
						this.plugin.settings.pdfOpenMode = v as OpenMode;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Default annotation tool')
			.setDesc('Tool to activate when a PDF is opened.')
			.addDropdown(dd =>
				dd
					.addOption('none', 'None (view only)')
					.addOption('pen', 'Pen')
					.addOption('highlighter', 'Highlighter')
					.setValue(this.plugin.settings.pdfDefaultTool)
					.onChange(async v => {
						this.plugin.settings.pdfDefaultTool = v as 'none' | 'pen' | 'highlighter';
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Toolbar position')
			.setDesc('Where to pin the PDF toolbar.')
			.addDropdown(dd =>
				dd
					.addOption('top', 'Top')
					.addOption('bottom', 'Bottom')
					.setValue(this.plugin.settings.pdfToolbarPosition)
					.onChange(async v => {
						this.plugin.settings.pdfToolbarPosition = v as ToolbarPosition;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Pen color')
			.addColorPicker(cp =>
				cp
					.setValue(this.plugin.settings.penColor)
					.onChange(async v => {
						this.plugin.settings.penColor = v;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Pen width')
			.addSlider(sl =>
				sl
					.setLimits(1, 20, 1)
					.setValue(this.plugin.settings.penWidth)
					.setDynamicTooltip()
					.onChange(async v => {
						this.plugin.settings.penWidth = v;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Highlighter color')
			.addColorPicker(cp =>
				cp
					.setValue(this.plugin.settings.highlighterColor)
					.onChange(async v => {
						this.plugin.settings.highlighterColor = v;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Highlighter width')
			.addSlider(sl =>
				sl
					.setLimits(10, 40, 2)
					.setValue(this.plugin.settings.highlighterWidth)
					.setDynamicTooltip()
					.onChange(async v => {
						this.plugin.settings.highlighterWidth = v;
						await this.plugin.saveSettings();
					})
			);

		new Setting(containerEl)
			.setName('Highlighter opacity')
			.addSlider(sl =>
				sl
					.setLimits(0.1, 1.0, 0.05)
					.setValue(this.plugin.settings.highlighterOpacity)
					.setDynamicTooltip()
					.onChange(async v => {
						this.plugin.settings.highlighterOpacity = v;
						await this.plugin.saveSettings();
					})
			);
	}
}
