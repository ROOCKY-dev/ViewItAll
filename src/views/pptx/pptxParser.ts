import type {
	JSZipInstance,
	ParagraphData,
	RunData,
	ShapeBounds,
	ShapeData,
	ShapePlaceholderType,
	ShapeStyle,
	SlideData,
	TableData,
	TableRowData,
} from './pptxTypes';

export async function parseSlidesFromZip(zip: JSZipInstance): Promise<SlideData[]> {
	const slideFiles = zip
		.file(/^ppt\/slides\/slide\d+\.xml$/)
		.map((f) => f.name)
		.sort((a, b) => {
			const numA = parseInt(a.match(/slide(\d+)/)?.[1] ?? '0', 10);
			const numB = parseInt(b.match(/slide(\d+)/)?.[1] ?? '0', 10);
			return numA - numB;
		});

	const slides: SlideData[] = [];

	for (let i = 0; i < slideFiles.length; i++) {
		const slideFile = slideFiles[i];
		if (!slideFile) continue;
		const slideXml = await readZipFile(zip, slideFile);
		if (!slideXml) {
			slides.push({ index: i + 1, backgroundToken: null, shapes: [], tables: [], imageDataUrls: [] });
			continue;
		}

		const shapes = extractShapes(slideXml);
		const tables = extractTables(slideXml);
		const backgroundToken = await extractSlideBackgroundToken(zip, slideFile, slideXml);
		const imageDataUrls = await extractImages(zip, slideFile);
		slides.push({ index: i + 1, backgroundToken, shapes, tables, imageDataUrls });
	}

	return slides;
}

async function readZipFile(zip: JSZipInstance, path: string): Promise<string | null> {
	const entry = zip.file(path);
	if (!entry) return null;
	return entry.async('string');
}

function extractShapes(xml: string): ShapeData[] {
	const parser = new DOMParser();
	const doc = parser.parseFromString(xml, 'application/xml');
	const spElements = doc.getElementsByTagName('p:sp');
	const shapes: ShapeData[] = [];

	for (let s = 0; s < spElements.length; s++) {
		const sp = spElements[s];
		if (!sp) continue;

		const type = detectShapeType(sp);
		const id = getShapeId(sp, s);
		const bounds = extractBounds(sp);
		const style = extractStyle(sp);

		const txBodyNodes = sp.getElementsByTagName('p:txBody');
		if (txBodyNodes.length === 0) {
			shapes.push({
				id,
				type,
				zIndex: s,
				bounds,
				style,
				paragraphs: [],
			});
			continue;
		}

		const body = txBodyNodes[0];
		if (!body) continue;

		const paragraphs: ParagraphData[] = [];
		const pElements = body.getElementsByTagName('a:p');

		for (let p = 0; p < pElements.length; p++) {
			const pEl = pElements[p];
			if (!pEl) continue;
			if (pEl.parentNode !== body) continue;

			const level = getParagraphLevel(pEl);
			const isBullet = detectBullet(pEl, level);
			const runs = extractRuns(pEl);
			paragraphs.push({ runs, isBullet, level });
		}

		shapes.push({
			id,
			type,
			zIndex: s,
			bounds,
			style,
			paragraphs,
		});
	}

	return shapes;
}

function detectShapeType(sp: Element): ShapePlaceholderType {
	const phElements = sp.getElementsByTagName('p:ph');
	if (phElements.length === 0) return 'other';
	const ph = phElements[0];
	if (!ph) return 'other';
	const type = ph.getAttribute('type') ?? '';
	if (type === 'title') return 'title';
	if (type === 'ctrTitle') return 'ctrTitle';
	if (type === 'subTitle') return 'subTitle';
	if (type === 'body') return 'body';
	return 'other';
}

function getShapeId(sp: Element, fallbackIndex: number): string {
	const nvSpPr = sp.getElementsByTagName('p:nvSpPr')[0];
	const cNvPr = nvSpPr?.getElementsByTagName('p:cNvPr')[0];
	const idAttr = cNvPr?.getAttribute('id');
	if (idAttr && idAttr.trim().length > 0) return idAttr;
	return `shape-${fallbackIndex + 1}`;
}

function extractBounds(sp: Element): ShapeBounds | null {
	const spPr = sp.getElementsByTagName('p:spPr')[0];
	if (!spPr) return null;
	const xfrm = spPr.getElementsByTagName('a:xfrm')[0];
	if (!xfrm) return null;

	const off = xfrm.getElementsByTagName('a:off')[0];
	const ext = xfrm.getElementsByTagName('a:ext')[0];
	if (!off || !ext) return null;

	const xEmu = parseInt(off.getAttribute('x') ?? '0', 10);
	const yEmu = parseInt(off.getAttribute('y') ?? '0', 10);
	const widthEmu = parseInt(ext.getAttribute('cx') ?? '0', 10);
	const heightEmu = parseInt(ext.getAttribute('cy') ?? '0', 10);
	const rotRaw = parseInt(xfrm.getAttribute('rot') ?? '0', 10);

	return {
		xEmu: Number.isFinite(xEmu) ? xEmu : 0,
		yEmu: Number.isFinite(yEmu) ? yEmu : 0,
		widthEmu: Number.isFinite(widthEmu) ? widthEmu : 0,
		heightEmu: Number.isFinite(heightEmu) ? heightEmu : 0,
		rotationDeg: Number.isFinite(rotRaw) ? rotRaw / 60000 : 0,
	};
}

function extractStyle(sp: Element): ShapeStyle {
	const spPr = sp.getElementsByTagName('p:spPr')[0];
	if (!spPr) {
		return {
			fillToken: null,
			lineToken: null,
			lineWidth: null,
			lineDash: null,
		};
	}

	const fillToken = getColorTokenFromNode(spPr.getElementsByTagName('a:solidFill')[0] ?? null);
	const lineEl = spPr.getElementsByTagName('a:ln')[0] ?? null;
	const lineToken = getColorTokenFromNode(lineEl?.getElementsByTagName('a:solidFill')[0] ?? null);
	const lineWidth = lineEl ? parseInt(lineEl.getAttribute('w') ?? '', 10) : NaN;
	const lineDashEl = lineEl?.getElementsByTagName('a:prstDash')[0] ?? null;
	const lineDash = lineDashEl?.getAttribute('val') ?? null;

	return {
		fillToken,
		lineToken,
		lineWidth: Number.isFinite(lineWidth) ? lineWidth : null,
		lineDash,
	};
}

function getParagraphLevel(pEl: Element): number {
	const pPr = pEl.getElementsByTagName('a:pPr')[0];
	if (!pPr) return 0;
	const lvl = parseInt(pPr.getAttribute('lvl') ?? '0', 10);
	return Number.isFinite(lvl) ? Math.max(0, lvl) : 0;
}

function detectBullet(pEl: Element, level: number): boolean {
	const pPr = pEl.getElementsByTagName('a:pPr')[0];
	if (!pPr) return false;
	if (pPr.getElementsByTagName('a:buNone').length > 0) return false;
	if (pPr.getElementsByTagName('a:buChar').length > 0) return true;
	if (pPr.getElementsByTagName('a:buAutoNum').length > 0) return true;
	return level > 0;
}

function extractRuns(pEl: Element): RunData[] {
	const runs: RunData[] = [];

	for (let i = 0; i < pEl.childNodes.length; i++) {
		const child = pEl.childNodes[i];
		if (!child) continue;

		if (child.nodeName === 'a:r') {
			const rEl = child as Element;
			runs.push(...extractTextRunsFromRunElement(rEl));
		}

		if (child.nodeName === 'a:fld') {
			const fEl = child as Element;
			runs.push(...extractFieldRuns(fEl));
		}
	}

	return runs;
}

function extractTextRunsFromRunElement(rEl: Element): RunData[] {
	const rPr = rEl.getElementsByTagName('a:rPr')[0] ?? null;
	const bold = rPr?.getAttribute('b') === '1';
	const italic = rPr?.getAttribute('i') === '1';
	const underline = (rPr?.getAttribute('u') ?? 'none') !== 'none';
	const fontSizeRaw = parseInt(rPr?.getAttribute('sz') ?? '', 10);
	const fontSizePt = Number.isFinite(fontSizeRaw) ? fontSizeRaw / 100 : null;
	const colorToken = getColorTokenFromNode(rPr?.getElementsByTagName('a:solidFill')[0] ?? null);

	const result: RunData[] = [];
	const tEls = rEl.getElementsByTagName('a:t');
	for (let t = 0; t < tEls.length; t++) {
		const tEl = tEls[t];
		if (!tEl) continue;
		const text = tEl.textContent ?? '';
		if (!text) continue;
		result.push({
			text,
			bold,
			italic,
			underline,
			fontSizePt,
			colorToken,
		});
	}
	return result;
}

function extractFieldRuns(fEl: Element): RunData[] {
	const runs: RunData[] = [];
	const tEls = fEl.getElementsByTagName('a:t');
	for (let t = 0; t < tEls.length; t++) {
		const tEl = tEls[t];
		if (!tEl) continue;
		const text = tEl.textContent ?? '';
		if (!text) continue;
		runs.push({
			text,
			bold: false,
			italic: false,
			underline: false,
			fontSizePt: null,
			colorToken: null,
		});
	}
	return runs;
}

function getColorTokenFromNode(fillEl: Element | null): string | null {
	if (!fillEl) return null;
	const srgb = fillEl.getElementsByTagName('a:srgbClr')[0];
	if (srgb) {
		const val = srgb.getAttribute('val');
		if (val && val.trim().length > 0) return `#${val}`;
	}
	const scheme = fillEl.getElementsByTagName('a:schemeClr')[0];
	if (scheme) {
		const val = scheme.getAttribute('val');
		if (val && val.trim().length > 0) return `scheme:${val}`;
	}
	return null;
}

function extractTables(xml: string): TableData[] {
	const parser = new DOMParser();
	const doc = parser.parseFromString(xml, 'application/xml');
	const frames = doc.getElementsByTagName('p:graphicFrame');
	const tables: TableData[] = [];

	for (let i = 0; i < frames.length; i++) {
		const frame = frames[i];
		if (!frame) continue;
		const tbl = frame.getElementsByTagName('a:tbl')[0];
		if (!tbl) continue;

		const id = getGraphicFrameId(frame, i);
		const bounds = extractGraphicFrameBounds(frame);
		const rows = extractTableRows(tbl);
		tables.push({ id, zIndex: i, bounds, rows });
	}

	return tables;
}

function extractTableRows(tbl: Element): TableRowData[] {
	const rowElements = tbl.getElementsByTagName('a:tr');
	const rows: TableRowData[] = [];

	for (let r = 0; r < rowElements.length; r++) {
		const rowEl = rowElements[r];
		if (!rowEl || rowEl.parentNode !== tbl) continue;
		const cellElements = rowEl.getElementsByTagName('a:tc');
		const cells: { text: string }[] = [];

		for (let c = 0; c < cellElements.length; c++) {
			const cellEl = cellElements[c];
			if (!cellEl || cellEl.parentNode !== rowEl) continue;
			const txBody = cellEl.getElementsByTagName('a:txBody')[0];
			if (!txBody) {
				cells.push({ text: '' });
				continue;
			}
			const paragraphs = txBody.getElementsByTagName('a:p');
			const lines: string[] = [];
			for (let p = 0; p < paragraphs.length; p++) {
				const pEl = paragraphs[p];
				if (!pEl || pEl.parentNode !== txBody) continue;
				const text = extractRuns(pEl).map((run) => run.text).join('');
				if (text.trim().length > 0) lines.push(text);
			}
			cells.push({ text: lines.join('\n') });
		}

		rows.push({ cells });
	}

	return rows;
}

function getGraphicFrameId(frame: Element, fallbackIndex: number): string {
	const nvGraphicFramePr = frame.getElementsByTagName('p:nvGraphicFramePr')[0];
	const cNvPr = nvGraphicFramePr?.getElementsByTagName('p:cNvPr')[0];
	const idAttr = cNvPr?.getAttribute('id');
	if (idAttr && idAttr.trim().length > 0) return idAttr;
	return `table-${fallbackIndex + 1}`;
}

function extractGraphicFrameBounds(frame: Element): ShapeBounds | null {
	const xfrm = frame.getElementsByTagName('p:xfrm')[0];
	if (!xfrm) return null;
	const off = xfrm.getElementsByTagName('a:off')[0];
	const ext = xfrm.getElementsByTagName('a:ext')[0];
	if (!off || !ext) return null;

	const xEmu = parseInt(off.getAttribute('x') ?? '0', 10);
	const yEmu = parseInt(off.getAttribute('y') ?? '0', 10);
	const widthEmu = parseInt(ext.getAttribute('cx') ?? '0', 10);
	const heightEmu = parseInt(ext.getAttribute('cy') ?? '0', 10);
	const rotRaw = parseInt(xfrm.getAttribute('rot') ?? '0', 10);

	return {
		xEmu: Number.isFinite(xEmu) ? xEmu : 0,
		yEmu: Number.isFinite(yEmu) ? yEmu : 0,
		widthEmu: Number.isFinite(widthEmu) ? widthEmu : 0,
		heightEmu: Number.isFinite(heightEmu) ? heightEmu : 0,
		rotationDeg: Number.isFinite(rotRaw) ? rotRaw / 60000 : 0,
	};
}

async function extractSlideBackgroundToken(
	zip: JSZipInstance,
	slidePath: string,
	slideXml: string
): Promise<string | null> {
	const direct = extractBackgroundTokenFromSlideXml(slideXml);
	if (direct) return direct;

	const slideFilename = slidePath.split('/').pop() ?? '';
	const slideRelsPath = `ppt/slides/_rels/${slideFilename}.rels`;
	const slideRelsXml = await readZipFile(zip, slideRelsPath);
	if (!slideRelsXml) return null;

	const parser = new DOMParser();
	const relDoc = parser.parseFromString(slideRelsXml, 'application/xml');
	const relNodes = relDoc.getElementsByTagName('Relationship');

	let layoutPath: string | null = null;
	for (let i = 0; i < relNodes.length; i++) {
		const rel = relNodes[i];
		if (!rel) continue;
		const type = rel.getAttribute('Type') ?? '';
		if (!type.includes('/slideLayout')) continue;
		const target = rel.getAttribute('Target') ?? '';
		layoutPath = resolvePath('ppt/slides/', target);
		break;
	}
	if (!layoutPath) return null;

	const layoutXml = await readZipFile(zip, layoutPath);
	if (!layoutXml) return null;
	const fromLayout = extractBackgroundTokenFromSlideXml(layoutXml);
	if (fromLayout) return fromLayout;

	const layoutFilename = layoutPath.split('/').pop() ?? '';
	const layoutRelsPath = `ppt/slideLayouts/_rels/${layoutFilename}.rels`;
	const layoutRelsXml = await readZipFile(zip, layoutRelsPath);
	if (!layoutRelsXml) return null;

	const layoutRelDoc = parser.parseFromString(layoutRelsXml, 'application/xml');
	const layoutRelNodes = layoutRelDoc.getElementsByTagName('Relationship');
	let masterPath: string | null = null;
	for (let i = 0; i < layoutRelNodes.length; i++) {
		const rel = layoutRelNodes[i];
		if (!rel) continue;
		const type = rel.getAttribute('Type') ?? '';
		if (!type.includes('/slideMaster')) continue;
		const target = rel.getAttribute('Target') ?? '';
		masterPath = resolvePath('ppt/slideLayouts/', target);
		break;
	}
	if (!masterPath) return null;

	const masterXml = await readZipFile(zip, masterPath);
	if (!masterXml) return null;
	return extractBackgroundTokenFromSlideXml(masterXml);
}

function extractBackgroundTokenFromSlideXml(xml: string): string | null {
	const parser = new DOMParser();
	const doc = parser.parseFromString(xml, 'application/xml');
	const bgPr = doc.getElementsByTagName('p:bgPr')[0];
	if (!bgPr) return null;
	return getColorTokenFromNode(bgPr.getElementsByTagName('a:solidFill')[0] ?? null);
}

async function extractImages(zip: JSZipInstance, slidePath: string): Promise<string[]> {
	const slideFilename = slidePath.split('/').pop() ?? '';
	const relsPath = `ppt/slides/_rels/${slideFilename}.rels`;

	const relsXml = await readZipFile(zip, relsPath);
	if (!relsXml) return [];

	const parser = new DOMParser();
	const doc = parser.parseFromString(relsXml, 'application/xml');
	const rels = doc.getElementsByTagName('Relationship');
	const dataUrls: string[] = [];

	for (let i = 0; i < rels.length; i++) {
		const rel = rels[i];
		if (!rel) continue;
		const type = rel.getAttribute('Type') ?? '';
		const target = rel.getAttribute('Target') ?? '';
		if (!type.includes('/image')) continue;

		const imagePath = resolvePath('ppt/slides/', target);
		const imageEntry = zip.file(imagePath);
		if (!imageEntry) continue;

		try {
			const imgData = await imageEntry.async('base64');
			const ext = imagePath.split('.').pop()?.toLowerCase() ?? 'png';
			const mime =
				ext === 'jpg' || ext === 'jpeg'
					? 'image/jpeg'
					: ext === 'png'
						? 'image/png'
						: ext === 'gif'
							? 'image/gif'
							: ext === 'svg'
								? 'image/svg+xml'
								: ext === 'webp'
									? 'image/webp'
									: 'image/png';
			dataUrls.push(`data:${mime};base64,${imgData}`);
		} catch {
			// Skip unreadable images.
		}
	}

	return dataUrls;
}

function resolvePath(base: string, relative: string): string {
	const baseParts = base.replace(/\/$/, '').split('/');
	const relParts = relative.split('/');

	for (const part of relParts) {
		if (part === '..') {
			baseParts.pop();
		} else if (part !== '.') {
			baseParts.push(part);
		}
	}

	return baseParts.join('/');
}
