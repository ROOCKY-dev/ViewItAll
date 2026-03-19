import { App, PluginSettingTab, Setting } from "obsidian";
import type ViewItAllPlugin from "./main";

export type ToolbarPosition = "top" | "bottom";

export interface PluginSettings {
	enableDocx: boolean;
	docxToolbarPosition: ToolbarPosition;
	docxDefaultZoom: number;
}

export const DEFAULT_SETTINGS: PluginSettings = {
	enableDocx: true,
	docxToolbarPosition: "top",
	docxDefaultZoom: 1.0,
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

		new Setting(containerEl).setName("File types").setHeading();

		new Setting(containerEl)
			.setName("Word (.docx)")
			.setDesc("Enable viewing .docx files. Restart required after change.")
			.addToggle((t) =>
				t.setValue(this.plugin.settings.enableDocx).onChange(async (v) => {
					this.plugin.settings.enableDocx = v;
					await this.plugin.saveSettings();
				}),
			);

		new Setting(containerEl).setName("Word documents (.docx)").setHeading();

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
			.setName("Default zoom")
			.setDesc("Zoom level when a .docx file is first opened.")
			.addDropdown((dd) =>
				dd
					.addOption("0.5", "50%")
					.addOption("0.75", "75%")
					.addOption("1.0", "100%")
					.addOption("1.25", "125%")
					.addOption("1.5", "150%")
					.addOption("2.0", "200%")
					.setValue(String(this.plugin.settings.docxDefaultZoom))
					.onChange(async (v) => {
						this.plugin.settings.docxDefaultZoom = parseFloat(v);
						await this.plugin.saveSettings();
					}),
			);
	}
}
