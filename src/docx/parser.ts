/**
 * ViewItAll — OOXML Document Parser
 *
 * Main orchestrator: extracts ZIP → parses XML → builds DocxDocument model.
 * Uses JSZip (dynamic import) for ZIP extraction and browser DOMParser for XML.
 */

import type {
	DocxDocument,
	DocxBlockElement,
	DocxParagraph,
	DocxInlineElement,
	DocxRun,
	DocxHyperlink,
	DocxImage,
	DocxTable,
	DocxTableRow,
	DocxTableCell,
	DocxRunProperties,
} from "./model";
import { defaultParagraphProperties } from "./model";
import {
	NS_W,
	NS_R,
	NS_WP,
	NS_A,
	getElement,
	getDirectChild,
	getDirectChildren,
	getVal,
	getWAttr,
	getAttr,
	parseXml,
} from "../utils/xml";
import { emuToPx, safeParseInt } from "../utils/units";
import { parseRelationships, REL_HYPERLINK, REL_IMAGE } from "./relationships";
import { parseStyles, parseParagraphProperties, parseRunProperties } from "./styles";
import { parseNumbering } from "./numbering";

/**
 * Result of parsing a .docx file — includes both the document model
 * and the raw JSZip instance for round-trip serialization.
 */
export interface ParseResult {
	doc: DocxDocument;
	zip: unknown; // JSZip instance, typed as unknown to avoid top-level import
}

/**
 * Parse a .docx ArrayBuffer into a DocxDocument model.
 */
export async function parseDocx(data: ArrayBuffer): Promise<ParseResult> {
	// Dynamic import — JSZip is a heavy lib
	const JSZip = (await import("jszip")).default;
	const zip = await JSZip.loadAsync(data);

	// ── Extract core XML files ──────────────────────────────────────────
	const documentXml = await readZipText(zip, "word/document.xml");
	if (!documentXml) {
		throw new Error("Invalid .docx: missing word/document.xml");
	}

	const relsXml = await readZipText(zip, "word/_rels/document.xml.rels");
	const stylesXml = await readZipText(zip, "word/styles.xml");
	const numberingXml = await readZipText(zip, "word/numbering.xml");

	// ── Parse supporting structures ─────────────────────────────────────
	const relationships = relsXml
		? parseRelationships(relsXml)
		: new Map<string, { type: string; target: string; targetMode: string | undefined }>();

	const styles = stylesXml ? parseStyles(stylesXml) : new Map();
	const numbering = numberingXml ? parseNumbering(numberingXml) : new Map();

	// ── Extract images ──────────────────────────────────────────────────
	const images = new Map<string, Blob>();
	for (const [id, rel] of relationships) {
		if (rel.type === REL_IMAGE) {
			const imagePath = `word/${rel.target}`;
			const imageFile = zip.file(imagePath);
			if (imageFile) {
				const imageData = await imageFile.async("arraybuffer");
				const ext = rel.target.split(".").pop()?.toLowerCase() ?? "";
				const mimeType = getMimeType(ext);
				images.set(id, new Blob([imageData], { type: mimeType }));
			}
		}
	}

	// ── Parse document body ─────────────────────────────────────────────
	const docXml = parseXml(documentXml);
	const bodyEl = getElement(docXml, NS_W, "body");
	if (!bodyEl) {
		throw new Error("Invalid .docx: missing w:body element");
	}

	const relMap = new Map<string, { type: string; target: string }>();
	for (const [id, rel] of relationships) {
		relMap.set(id, { type: rel.type, target: rel.target });
	}

	const body = parseBody(bodyEl, relationships);

	return {
		doc: {
			body,
			styles,
			images,
			numbering,
			relationships: relMap,
		},
		zip,
	};
}

// ── Body Parser ─────────────────────────────────────────────────────────────

function parseBody(
	bodyEl: Element,
	relationships: Map<string, { type: string; target: string; targetMode: string | undefined }>,
): DocxBlockElement[] {
	const blocks: DocxBlockElement[] = [];

	for (let i = 0; i < bodyEl.childNodes.length; i++) {
		const node = bodyEl.childNodes.item(i);
		if (!node || node.nodeType !== Node.ELEMENT_NODE) continue;
		const el = node as Element;

		if (el.localName === "p" && el.namespaceURI === NS_W) {
			blocks.push(parseParagraph(el, relationships));
		} else if (el.localName === "tbl" && el.namespaceURI === NS_W) {
			blocks.push(parseTable(el, relationships));
		} else if (el.localName === "sdt" && el.namespaceURI === NS_W) {
			// Structured document tags — unwrap and parse content
			const sdtContent = getDirectChild(el, NS_W, "sdtContent");
			if (sdtContent) {
				const inner = parseBody(sdtContent, relationships);
				blocks.push(...inner);
			}
		}
	}

	return blocks;
}

// ── Paragraph Parser ────────────────────────────────────────────────────────

function parseParagraph(
	pEl: Element,
	relationships: Map<string, { type: string; target: string; targetMode: string | undefined }>,
): DocxParagraph {
	const pPrEl = getDirectChild(pEl, NS_W, "pPr");

	// Style ID
	let styleId: string | undefined;
	if (pPrEl) {
		const pStyleEl = getDirectChild(pPrEl, NS_W, "pStyle");
		if (pStyleEl) {
			styleId = getVal(pStyleEl) ?? undefined;
		}
	}

	// Paragraph properties (includes numbering from direct pPr)
	const properties = pPrEl
		? { ...defaultParagraphProperties(), ...parseParagraphProperties(pPrEl) }
		: defaultParagraphProperties();

	// Numbering: use direct pPr numbering if present
	const numberingId = properties.numberingId;
	const numberingLevel = properties.numberingLevel ?? 0;

	// Check for page break in paragraph properties
	const children: DocxInlineElement[] = [];

	// Parse inline children
	for (let i = 0; i < pEl.childNodes.length; i++) {
		const node = pEl.childNodes.item(i);
		if (!node || node.nodeType !== Node.ELEMENT_NODE) continue;
		const el = node as Element;

		if (el.localName === "r" && el.namespaceURI === NS_W) {
			const inlines = parseRun(el);
			children.push(...inlines);
		} else if (el.localName === "hyperlink" && el.namespaceURI === NS_W) {
			const hyperlink = parseHyperlink(el, relationships);
			if (hyperlink) children.push(hyperlink);
		} else if (el.localName === "sdt" && el.namespaceURI === NS_W) {
			// Structured document tags — unwrap and parse inline content
			const sdtContent = getDirectChild(el, NS_W, "sdtContent");
			if (sdtContent) {
				for (let j = 0; j < sdtContent.childNodes.length; j++) {
					const sdtNode = sdtContent.childNodes.item(j);
					if (!sdtNode || sdtNode.nodeType !== Node.ELEMENT_NODE) continue;
					const sdtEl = sdtNode as Element;
					if (sdtEl.localName === "r" && sdtEl.namespaceURI === NS_W) {
						children.push(...parseRun(sdtEl));
					} else if (sdtEl.localName === "hyperlink" && sdtEl.namespaceURI === NS_W) {
						const hl = parseHyperlink(sdtEl, relationships);
						if (hl) children.push(hl);
					}
				}
			}
		}
	}

	return {
		type: "paragraph",
		styleId,
		properties,
		children,
		numberingId,
		numberingLevel,
	};
}

// ── Run Parser ──────────────────────────────────────────────────────────────

function parseRun(rEl: Element): DocxInlineElement[] {
	const results: DocxInlineElement[] = [];

	// Run properties — keep as Partial so mergeRunProps knows what was explicit
	const rPrEl = getDirectChild(rEl, NS_W, "rPr");
	const properties: Partial<DocxRunProperties> = rPrEl
		? parseRunProperties(rPrEl)
		: {};

	for (let i = 0; i < rEl.childNodes.length; i++) {
		const node = rEl.childNodes.item(i);
		if (!node || node.nodeType !== Node.ELEMENT_NODE) continue;
		const child = node as Element;

		if (child.localName === "t" && child.namespaceURI === NS_W) {
			const text = child.textContent ?? "";
			if (text) {
				results.push({ type: "run", text, properties });
			}
		} else if (child.localName === "br" && child.namespaceURI === NS_W) {
			const breakType = getWAttr(child, "type");
			if (breakType === "page") {
				results.push({ type: "break", breakType: "page" });
			} else if (breakType === "column") {
				results.push({ type: "break", breakType: "column" });
			} else {
				results.push({ type: "break", breakType: "line" });
			}
		} else if (child.localName === "tab" && child.namespaceURI === NS_W) {
			results.push({ type: "tab" });
		} else if (child.localName === "drawing" && child.namespaceURI === NS_W) {
			const image = parseDrawing(child);
			if (image) results.push(image);
		} else if (child.localName === "pict" && child.namespaceURI === NS_W) {
			// Legacy VML images — skip for MVP
		}
	}

	return results;
}

// ── Hyperlink Parser ────────────────────────────────────────────────────────

function parseHyperlink(
	hlEl: Element,
	relationships: Map<string, { type: string; target: string; targetMode: string | undefined }>,
): DocxHyperlink | null {
	// Resolve URL from relationship ID
	const rId =
		getAttr(hlEl, NS_R, "id") ??
		hlEl.getAttribute("r:id");

	let url = "";
	if (rId) {
		const rel = relationships.get(rId);
		if (rel && rel.type === REL_HYPERLINK) {
			url = rel.target;
		}
	}

	// Also check for anchor (internal bookmark)
	if (!url) {
		const anchor = getWAttr(hlEl, "anchor");
		if (anchor) {
			url = `#${anchor}`;
		}
	}

	// Parse child runs
	const children: DocxRun[] = [];
	const runElements = getDirectChildren(hlEl, NS_W, "r");
	for (const rEl of runElements) {
		const inlines = parseRun(rEl);
		for (const inline of inlines) {
			if (inline.type === "run") {
				children.push(inline);
			}
		}
	}

	if (children.length === 0) return null;

	return { type: "hyperlink", url, children };
}

// ── Drawing (Image) Parser ──────────────────────────────────────────────────

function parseDrawing(drawingEl: Element): DocxImage | null {
	// Inline images: wp:inline or wp:anchor
	const inlineEl =
		getElement(drawingEl, NS_WP, "inline") ??
		getElement(drawingEl, NS_WP, "anchor");
	if (!inlineEl) return null;

	// Dimensions from wp:extent
	const extentEl = getElement(inlineEl, NS_WP, "extent");
	let width = 0;
	let height = 0;
	if (extentEl) {
		const cx = safeParseInt(extentEl.getAttribute("cx"));
		const cy = safeParseInt(extentEl.getAttribute("cy"));
		if (cx) width = emuToPx(cx);
		if (cy) height = emuToPx(cy);
	}

	// Alt text from wp:docPr
	const docPrEl = getElement(inlineEl, NS_WP, "docPr");
	const altText = docPrEl?.getAttribute("descr") ?? undefined;

	// Find the blip (actual image reference)
	const blipEl = getElement(drawingEl, NS_A, "blip");
	if (!blipEl) return null;

	const rId =
		getAttr(blipEl, NS_R, "embed") ??
		blipEl.getAttribute("r:embed") ??
		getAttr(blipEl, NS_R, "link") ??
		blipEl.getAttribute("r:link");

	if (!rId) return null;

	return {
		type: "image",
		relationshipId: rId,
		width,
		height,
		altText,
	};
}

// ── Table Parser ────────────────────────────────────────────────────────────

function parseTable(
	tblEl: Element,
	relationships: Map<string, { type: string; target: string; targetMode: string | undefined }>,
): DocxTable {
	// Table properties
	const tblPrEl = getDirectChild(tblEl, NS_W, "tblPr");
	let tableWidth: number | undefined;
	let tableAlignment: "left" | "center" | "right" | undefined;

	if (tblPrEl) {
		const tblWEl = getDirectChild(tblPrEl, NS_W, "tblW");
		if (tblWEl) {
			tableWidth = safeParseInt(getWAttr(tblWEl, "w")) ?? undefined;
		}
		const jcEl = getDirectChild(tblPrEl, NS_W, "jc");
		if (jcEl) {
			const val = getVal(jcEl);
			if (val === "left" || val === "center" || val === "right") {
				tableAlignment = val;
			}
		}
	}

	// Parse rows
	const rows: DocxTableRow[] = [];
	const trElements = getDirectChildren(tblEl, NS_W, "tr");

	for (const trEl of trElements) {
		const trPrEl = getDirectChild(trEl, NS_W, "trPr");
		const isHeader = trPrEl
			? getDirectChild(trPrEl, NS_W, "tblHeader") !== null
			: false;

		const cells: DocxTableCell[] = [];
		const tcElements = getDirectChildren(trEl, NS_W, "tc");

		for (const tcEl of tcElements) {
			cells.push(parseTableCell(tcEl, relationships));
		}

		rows.push({ cells, isHeader });
	}

	return {
		type: "table",
		rows,
		properties: {
			width: tableWidth,
			alignment: tableAlignment,
		},
	};
}

function parseTableCell(
	tcEl: Element,
	relationships: Map<string, { type: string; target: string; targetMode: string | undefined }>,
): DocxTableCell {
	const tcPrEl = getDirectChild(tcEl, NS_W, "tcPr");

	let width: number | undefined;
	let verticalMerge: "restart" | "continue" | undefined;
	let gridSpan = 1;
	let shading: string | undefined;
	let verticalAlign: "top" | "center" | "bottom" | undefined;

	if (tcPrEl) {
		const tcWEl = getDirectChild(tcPrEl, NS_W, "tcW");
		if (tcWEl) {
			width = safeParseInt(getWAttr(tcWEl, "w")) ?? undefined;
		}

		const vMergeEl = getDirectChild(tcPrEl, NS_W, "vMerge");
		if (vMergeEl) {
			const val = getVal(vMergeEl);
			verticalMerge = val === "restart" ? "restart" : "continue";
		}

		const gridSpanEl = getDirectChild(tcPrEl, NS_W, "gridSpan");
		if (gridSpanEl) {
			gridSpan = safeParseInt(getVal(gridSpanEl)) ?? 1;
		}

		const shdEl = getDirectChild(tcPrEl, NS_W, "shd");
		if (shdEl) {
			const fill = getWAttr(shdEl, "fill");
			if (fill && fill !== "auto") {
				shading = fill;
			}
		}

		const vAlignEl = getDirectChild(tcPrEl, NS_W, "vAlign");
		if (vAlignEl) {
			const val = getVal(vAlignEl);
			if (val === "top" || val === "center" || val === "bottom") {
				verticalAlign = val;
			}
		}
	}

	// Parse paragraphs inside the cell
	const paragraphs: DocxParagraph[] = [];
	const pElements = getDirectChildren(tcEl, NS_W, "p");
	for (const pEl of pElements) {
		paragraphs.push(parseParagraph(pEl, relationships));
	}

	return {
		paragraphs,
		properties: {
			width,
			verticalMerge,
			gridSpan,
			shading,
			verticalAlign,
		},
	};
}

// ── Helpers ─────────────────────────────────────────────────────────────────

async function readZipText(
	zip: { file(path: string): { async(type: "string"): Promise<string> } | null },
	path: string,
): Promise<string | null> {
	const file = zip.file(path);
	if (!file) return null;
	return file.async("string");
}

function getMimeType(ext: string): string {
	switch (ext) {
		case "png":
			return "image/png";
		case "jpg":
		case "jpeg":
			return "image/jpeg";
		case "gif":
			return "image/gif";
		case "svg":
			return "image/svg+xml";
		case "bmp":
			return "image/bmp";
		case "tiff":
		case "tif":
			return "image/tiff";
		case "webp":
			return "image/webp";
		default:
			return "application/octet-stream";
	}
}
