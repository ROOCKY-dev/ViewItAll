import type { PptxEditsFile, PptxShapeEditEntry } from '../../utils/pptxEdits';
import type { JSZipInstance } from './pptxTypes';

const EMU_PER_PIXEL = 9525;

export async function applyEditsToPptxZip(
	zip: JSZipInstance,
	edits: PptxEditsFile
): Promise<number> {
	const slideFiles = zip
		.file(/^ppt\/slides\/slide\d+\.xml$/)
		.map((f) => f.name)
		.sort((a, b) => {
			const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? '0', 10);
			const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? '0', 10);
			return numA - numB;
		});

	let updatedCount = 0;

	for (let i = 0; i < slideFiles.length; i++) {
		const slidePath = slideFiles[i];
		if (!slidePath) continue;

		const entry = zip.file(slidePath);
		if (!entry) continue;
		const xml = await entry.async('string');
		const updatedXml = applyEditsToSlideXml(xml, i + 1, edits);
		if (updatedXml === null) continue;

		zip.file(slidePath, updatedXml);
		updatedCount += 1;
	}

	return updatedCount;
}

function applyEditsToSlideXml(xml: string, slideIndex: number, edits: PptxEditsFile): string | null {
	const parser = new DOMParser();
	const doc = parser.parseFromString(xml, 'application/xml');
	const spElements = doc.getElementsByTagName('p:sp');
	let changed = false;

	for (let i = 0; i < spElements.length; i++) {
		const sp = spElements[i];
		if (!sp) continue;
		const shapeId = getShapeId(sp, i);
		const shapeKey = `${slideIndex}:${shapeId}`;
		const edit = edits.shapes[shapeKey];
		if (!edit) continue;

		const spPr = getOrCreateChildElement(doc, sp, 'p:spPr');
		const wasTransformChanged = applyTransformEdit(doc, spPr, edit);
		const wasStyleChanged = applyStyleEdit(doc, spPr, edit);
		changed = changed || wasTransformChanged || wasStyleChanged;
	}

	if (!changed) return null;
	return new XMLSerializer().serializeToString(doc);
}

function getShapeId(sp: Element, fallbackIndex: number): string {
	const nvSpPr = sp.getElementsByTagName('p:nvSpPr')[0];
	const cNvPr = nvSpPr?.getElementsByTagName('p:cNvPr')[0];
	const idAttr = cNvPr?.getAttribute('id');
	if (idAttr && idAttr.trim().length > 0) return idAttr;
	return `shape-${fallbackIndex + 1}`;
}

function applyTransformEdit(doc: XMLDocument, spPr: Element, edit: PptxShapeEditEntry): boolean {
	const xfrm = getOrCreateChildElement(doc, spPr, 'a:xfrm');
	const off = getOrCreateChildElement(doc, xfrm, 'a:off');
	const ext = getOrCreateChildElement(doc, xfrm, 'a:ext');

	const currentX = parseInt(off.getAttribute('x') ?? '0', 10);
	const currentY = parseInt(off.getAttribute('y') ?? '0', 10);
	const safeCurrentX = Number.isFinite(currentX) ? currentX : 0;
	const safeCurrentY = Number.isFinite(currentY) ? currentY : 0;

	const deltaX = Math.round(edit.translateX * EMU_PER_PIXEL);
	const deltaY = Math.round(edit.translateY * EMU_PER_PIXEL);
	const nextX = safeCurrentX + deltaX;
	const nextY = safeCurrentY + deltaY;

	off.setAttribute('x', String(nextX));
	off.setAttribute('y', String(nextY));

	if (edit.widthPx !== null) {
		ext.setAttribute('cx', String(Math.round(Math.max(1, edit.widthPx) * EMU_PER_PIXEL)));
	}
	if (edit.heightPx !== null) {
		ext.setAttribute('cy', String(Math.round(Math.max(1, edit.heightPx) * EMU_PER_PIXEL)));
	}

	return true;
}

function applyStyleEdit(doc: XMLDocument, spPr: Element, edit: PptxShapeEditEntry): boolean {
	let changed = false;
	if (edit.fillToken !== null) {
		applyFillToken(doc, spPr, edit.fillToken);
		changed = true;
	}
	if (edit.lineToken !== null) {
		const lineEl = getOrCreateChildElement(doc, spPr, 'a:ln');
		applyFillToken(doc, lineEl, edit.lineToken);
		changed = true;
	}
	return changed;
}

function applyFillToken(doc: XMLDocument, parent: Element, token: string): void {
	removeChildIfExists(parent, 'a:solidFill');
	removeChildIfExists(parent, 'a:noFill');

	if (token === 'none') {
		parent.appendChild(doc.createElement('a:noFill'));
		return;
	}

	const schemeValue = tokenToScheme(token);
	const solidFill = doc.createElement('a:solidFill');
	const schemeClr = doc.createElement('a:schemeClr');
	schemeClr.setAttribute('val', schemeValue);
	solidFill.appendChild(schemeClr);
	parent.appendChild(solidFill);
}

function tokenToScheme(token: string): string {
	if (token === 'accent') return 'accent1';
	if (token === 'muted') return 'accent3';
	if (token === 'normal') return 'tx1';
	return 'accent1';
}

function removeChildIfExists(parent: Element, tagName: string): void {
	const nodes = parent.getElementsByTagName(tagName);
	if (nodes.length === 0) return;
	const node = nodes[0];
	if (node && node.parentElement === parent) parent.removeChild(node);
}

function getOrCreateChildElement(doc: XMLDocument, parent: Element, tagName: string): Element {
	const existing = parent.getElementsByTagName(tagName)[0];
	if (existing && existing.parentElement === parent) return existing;
	const created = doc.createElement(tagName);
	parent.appendChild(created);
	return created;
}
