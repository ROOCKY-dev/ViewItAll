import { App, TFile } from "obsidian";

/** Returns the app:// resource URL Obsidian/Electron uses to serve vault files. */
export function getResourcePath(app: App, file: TFile): string {
	return app.vault.adapter.getResourcePath(file.path);
}

/** Returns the absolute OS path for a vault file. */
export function getAbsolutePath(app: App, file: TFile): string {
	const adapter = app.vault.adapter as { basePath?: string };
	if (adapter.basePath) {
		return `${adapter.basePath}/${file.path}`;
	}
	return file.path;
}

/** Returns the vault-relative path for a companion file (e.g. .annotations.json). */
export function getCompanionPath(file: TFile, suffix: string): string {
	return file.path + suffix;
}
