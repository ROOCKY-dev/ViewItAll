import { App, TFile } from "obsidian";
import { AnnotationFile, PageAnnotations } from "../types";
import { getCompanionPath } from "./fileUtils";

const SUFFIX = ".annotations.json";

export async function loadAnnotations(
	app: App,
	file: TFile,
): Promise<AnnotationFile> {
	const path = getCompanionPath(file, SUFFIX);
	try {
		const exists = await app.vault.adapter.exists(path);
		if (!exists) return { version: 1, pages: [] };
		const raw = await app.vault.adapter.read(path);
		return JSON.parse(raw) as AnnotationFile;
	} catch {
		return { version: 1, pages: [] };
	}
}

export async function saveAnnotations(
	app: App,
	file: TFile,
	data: AnnotationFile,
): Promise<void> {
	const path = getCompanionPath(file, SUFFIX);
	const raw = JSON.stringify(data, null, 2);
	const exists = await app.vault.adapter.exists(path);
	if (exists) {
		await app.vault.adapter.write(path, raw);
	} else {
		await app.vault.adapter.write(path, raw);
	}
}

export function getPageAnnotations(
	data: AnnotationFile,
	pageNum: number,
): PageAnnotations {
	return (
		data.pages.find((p) => p.page === pageNum) ?? {
			page: pageNum,
			paths: [],
		}
	);
}

export function setPageAnnotations(
	data: AnnotationFile,
	pageAnnotations: PageAnnotations,
): AnnotationFile {
	const pages = data.pages.filter((p) => p.page !== pageAnnotations.page);
	if (pageAnnotations.paths.length > 0) pages.push(pageAnnotations);
	pages.sort((a, b) => a.page - b.page);
	return { ...data, pages };
}
