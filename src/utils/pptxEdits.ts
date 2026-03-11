import { App, TFile } from 'obsidian';
import { getCompanionPath } from './fileUtils';

export interface PptxShapeEditEntry {
	translateX: number;
	translateY: number;
	widthPx: number | null;
	heightPx: number | null;
	fillToken: string | null;
	lineToken: string | null;
	textContent: string | null;
}

export interface PptxEditsFile {
	version: 1;
	shapes: Record<string, PptxShapeEditEntry>;
}

const SUFFIX = '.pptx.edits.json';

export async function loadPptxEdits(app: App, file: TFile): Promise<PptxEditsFile> {
	const path = getCompanionPath(file, SUFFIX);
	const existing = app.vault.getAbstractFileByPath(path);
	if (!(existing instanceof TFile)) {
		return { version: 1, shapes: {} };
	}

	try {
		const raw = await app.vault.read(existing);
		const parsed = JSON.parse(raw) as Partial<PptxEditsFile>;
		if (parsed.version !== 1 || !parsed.shapes || typeof parsed.shapes !== 'object') {
			return { version: 1, shapes: {} };
		}
		return {
			version: 1,
			shapes: parsed.shapes,
		};
	} catch {
		return { version: 1, shapes: {} };
	}
}

export async function savePptxEdits(app: App, file: TFile, data: PptxEditsFile): Promise<void> {
	const path = getCompanionPath(file, SUFFIX);
	const raw = JSON.stringify(data, null, 2);
	const existing = app.vault.getAbstractFileByPath(path);
	if (existing instanceof TFile) {
		await app.vault.modify(existing, raw);
		return;
	}
	await app.vault.create(path, raw);
}
