import { Plugin } from "obsidian";
import {
	DEFAULT_SETTINGS,
	PluginSettings,
	ViewItAllSettingTab,
} from "./settings";
import { VIEW_TYPE_DOCX } from "./types";
import { DocxView } from "./views/DocxView";

export default class ViewItAllPlugin extends Plugin {
	settings: PluginSettings;

	async onload() {
		await this.loadSettings();

		this.registerView(VIEW_TYPE_DOCX, (leaf) => new DocxView(leaf, this));

		if (this.settings.enableDocx) {
			this.registerExtensions(["docx"], VIEW_TYPE_DOCX);
		}

		this.addSettingTab(new ViewItAllSettingTab(this.app, this));
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
