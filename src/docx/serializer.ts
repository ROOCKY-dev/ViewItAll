/**
 * ViewItAll — OOXML Document Serializer
 *
 * Converts a DocxDocument model back into OOXML XML and repacks the .docx ZIP.
 * Only word/document.xml is regenerated; all other files (styles, numbering,
 * images, theme, fonts, settings) are carried over from the original ZIP.
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
	DocxParagraphProperties,
} from "./model";

// ── OOXML Namespace Declarations ────────────────────────────────────────────

const DOCUMENT_NAMESPACES = [
	'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"',
	'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
	'xmlns:o="urn:schemas-microsoft-com:office:office"',
	'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
	'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"',
	'xmlns:v="urn:schemas-microsoft-com:vml"',
	'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"',
	'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
	'xmlns:w10="urn:schemas-microsoft-com:office:word"',
	'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
	'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
	'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
	'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"',
	'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"',
	'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
	'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
	'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"',
].join(" ");

// ── Public API ──────────────────────────────────────────────────────────────

/**
 * Serialize a DocxDocument back into a .docx ArrayBuffer.
 * Replaces only word/document.xml in the original ZIP; everything else is
 * carried over unchanged for maximum round-trip fidelity.
 *
 * @param doc     The document model to serialize.
 * @param origZip The original JSZip instance from parsing (typed as unknown).
 * @returns       An ArrayBuffer containing the new .docx.
 */
export async function serializeDocx(
	doc: DocxDocument,
	origZip: unknown,
): Promise<ArrayBuffer> {
	// Dynamic import — JSZip is a heavy lib
	const JSZip = (await import("jszip")).default;

	// Clone the original ZIP so we don't mutate it
	const zipData = await (origZip as { generateAsync(opts: { type: "arraybuffer" }): Promise<ArrayBuffer> })
		.generateAsync({ type: "arraybuffer" });
	const zip = await JSZip.loadAsync(zipData);

	// Generate new document.xml from the model
	const documentXml = serializeDocument(doc);
	zip.file("word/document.xml", documentXml);

	// Write any new images that were inserted during editing
	for (const [rId, rel] of doc.relationships) {
		if (rel.type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image") {
			const imagePath = `word/${rel.target}`;
			// Only write if file doesn't already exist in the ZIP
			if (!zip.file(imagePath)) {
				const blob = doc.images.get(rId);
				if (blob) {
					const buf = await blob.arrayBuffer();
					zip.file(imagePath, buf);

					// Also ensure the relationship is in the rels file
					await ensureRelationship(zip, rId, rel.type, rel.target);
				}
			}
		}
	}

	// Generate the output
	return zip.generateAsync({
		type: "arraybuffer",
		compression: "DEFLATE",
		compressionOptions: { level: 6 },
	});
}

// ── Document Serialization ──────────────────────────────────────────────────

function serializeDocument(doc: DocxDocument): string {
	const bodyXml = serializeBody(doc.body);
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
		`<w:document ${DOCUMENT_NAMESPACES}>\n` +
		`<w:body>\n${bodyXml}</w:body>\n` +
		`</w:document>`;
}

function serializeBody(blocks: DocxBlockElement[]): string {
	let xml = "";
	for (const block of blocks) {
		switch (block.type) {
			case "paragraph":
				xml += serializeParagraph(block);
				break;
			case "table":
				xml += serializeTable(block);
				break;
			case "pageBreak":
				xml += '<w:p><w:r><w:br w:type="page"/></w:r></w:p>\n';
				break;
			case "sectionBreak":
				// Section breaks are complex; skip for now
				break;
		}
	}
	return xml;
}

// ── Paragraph Serialization ─────────────────────────────────────────────────

function serializeParagraph(para: DocxParagraph): string {
	let xml = "<w:p>";

	// Paragraph properties
	const pPr = serializeParagraphProperties(
		para.properties,
		para.styleId,
		para.numberingId,
		para.numberingLevel,
	);
	if (pPr) xml += pPr;

	// Inline children
	for (const child of para.children) {
		xml += serializeInline(child);
	}

	xml += "</w:p>\n";
	return xml;
}

function serializeParagraphProperties(
	props: DocxParagraphProperties,
	styleId: string | undefined,
	numId: string | undefined,
	numLevel: number,
): string {
	const parts: string[] = [];

	if (styleId) {
		parts.push(`<w:pStyle w:val="${escXml(styleId)}"/>`);
	}

	if (numId && numId !== "0") {
		parts.push(
			"<w:numPr>" +
			`<w:ilvl w:val="${numLevel}"/>` +
			`<w:numId w:val="${escXml(numId)}"/>` +
			"</w:numPr>",
		);
	}

	if (props.alignment) {
		parts.push(`<w:jc w:val="${props.alignment}"/>`);
	}

	if (props.headingLevel !== undefined && props.headingLevel >= 1) {
		parts.push(`<w:outlineLvl w:val="${props.headingLevel - 1}"/>`);
	}

	if (props.indentation) {
		const attrs: string[] = [];
		if (props.indentation.left) attrs.push(`w:left="${props.indentation.left}"`);
		if (props.indentation.right) attrs.push(`w:right="${props.indentation.right}"`);
		if (props.indentation.firstLine) attrs.push(`w:firstLine="${props.indentation.firstLine}"`);
		if (props.indentation.hanging) attrs.push(`w:hanging="${props.indentation.hanging}"`);
		if (attrs.length > 0) {
			parts.push(`<w:ind ${attrs.join(" ")}/>`);
		}
	}

	if (props.spacing) {
		const attrs: string[] = [];
		if (props.spacing.before) attrs.push(`w:before="${props.spacing.before}"`);
		if (props.spacing.after) attrs.push(`w:after="${props.spacing.after}"`);
		if (props.spacing.line) attrs.push(`w:line="${props.spacing.line}"`);
		if (props.spacing.lineRule && props.spacing.lineRule !== "auto") {
			attrs.push(`w:lineRule="${props.spacing.lineRule}"`);
		}
		if (attrs.length > 0) {
			parts.push(`<w:spacing ${attrs.join(" ")}/>`);
		}
	}

	if (parts.length === 0) return "";
	return `<w:pPr>${parts.join("")}</w:pPr>`;
}

// ── Inline Serialization ────────────────────────────────────────────────────

function serializeInline(inline: DocxInlineElement): string {
	switch (inline.type) {
		case "run":
			return serializeRun(inline);
		case "hyperlink":
			return serializeHyperlink(inline);
		case "image":
			return serializeImage(inline);
		case "break":
			return serializeBreak(inline);
		case "tab":
			return "<w:r><w:tab/></w:r>";
	}
}

function serializeRun(run: DocxRun): string {
	const rPr = serializeRunProperties(run.properties);
	const textNeedSpace = run.text.startsWith(" ") || run.text.endsWith(" ");
	const spaceAttr = textNeedSpace ? ' xml:space="preserve"' : "";
	return `<w:r>${rPr}<w:t${spaceAttr}>${escXml(run.text)}</w:t></w:r>`;
}

function serializeRunProperties(props: Partial<DocxRunProperties>): string {
	const parts: string[] = [];

	if (props.bold) parts.push("<w:b/>");
	if (props.italic) parts.push("<w:i/>");
	if (props.underline) parts.push('<w:u w:val="single"/>');
	if (props.strikethrough) parts.push("<w:strike/>");

	if (props.fontFamily) {
		const f = escXml(props.fontFamily);
		parts.push(`<w:rFonts w:ascii="${f}" w:hAnsi="${f}"/>`);
	}

	if (props.fontSize !== undefined) {
		parts.push(`<w:sz w:val="${props.fontSize}"/>`);
	}

	if (props.color !== undefined) {
		if (props.color === "auto") {
			parts.push('<w:color w:val="auto"/>');
		} else {
			parts.push(`<w:color w:val="${escXml(props.color)}"/>`);
		}
	}

	if (props.highlight) {
		parts.push(`<w:highlight w:val="${escXml(props.highlight)}"/>`);
	}

	if (props.vertAlign) {
		parts.push(`<w:vertAlign w:val="${props.vertAlign}"/>`);
	}

	if (parts.length === 0) return "";
	return `<w:rPr>${parts.join("")}</w:rPr>`;
}

function serializeHyperlink(link: DocxHyperlink): string {
	// Find the relationship ID for this URL in the document
	// If no rId found, we still emit the hyperlink with an anchor
	let xml = "<w:hyperlink";

	if (link.url.startsWith("#")) {
		xml += ` w:anchor="${escXml(link.url.slice(1))}"`;
	}
	// External hyperlinks need r:id — they should already have one from parsing.
	// For new hyperlinks, the editing layer must create the relationship first.

	xml += ">";

	for (const run of link.children) {
		xml += serializeRun(run);
	}

	xml += "</w:hyperlink>";
	return xml;
}

function serializeImage(image: DocxImage): string {
	// Emit a drawing with inline picture referencing the relationship ID
	const cx = pxToEmu(image.width);
	const cy = pxToEmu(image.height);
	const alt = escXml(image.altText ?? "");
	const rId = escXml(image.relationshipId);

	return `<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>` +
		`<wp:inline distT="0" distB="0" distL="0" distR="0">` +
		`<wp:extent cx="${cx}" cy="${cy}"/>` +
		`<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
		`<wp:docPr id="1" name="Picture" descr="${alt}"/>` +
		`<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
		`<pic:nvPicPr><pic:cNvPr id="0" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>` +
		`<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
		`<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
		`</pic:pic></a:graphicData></a:graphic>` +
		`</wp:inline></w:drawing></w:r>`;
}

function serializeBreak(brk: { breakType: string }): string {
	if (brk.breakType === "page") {
		return '<w:r><w:br w:type="page"/></w:r>';
	} else if (brk.breakType === "column") {
		return '<w:r><w:br w:type="column"/></w:r>';
	}
	return "<w:r><w:br/></w:r>";
}

// ── Table Serialization ─────────────────────────────────────────────────────

function serializeTable(table: DocxTable): string {
	let xml = "<w:tbl>";

	// Table properties
	const tblPr: string[] = [];
	if (table.properties.width) {
		tblPr.push(`<w:tblW w:w="${table.properties.width}" w:type="dxa"/>`);
	}
	if (table.properties.alignment) {
		tblPr.push(`<w:jc w:val="${table.properties.alignment}"/>`);
	}
	if (tblPr.length > 0) {
		xml += `<w:tblPr>${tblPr.join("")}</w:tblPr>`;
	}

	// Rows
	for (const row of table.rows) {
		xml += serializeTableRow(row);
	}

	xml += "</w:tbl>\n";
	return xml;
}

function serializeTableRow(row: DocxTableRow): string {
	let xml = "<w:tr>";

	if (row.isHeader) {
		xml += "<w:trPr><w:tblHeader/></w:trPr>";
	}

	for (const cell of row.cells) {
		xml += serializeTableCell(cell);
	}

	xml += "</w:tr>";
	return xml;
}

function serializeTableCell(cell: DocxTableCell): string {
	let xml = "<w:tc>";

	// Cell properties
	const tcPr: string[] = [];
	if (cell.properties.width) {
		tcPr.push(`<w:tcW w:w="${cell.properties.width}" w:type="dxa"/>`);
	}
	if (cell.properties.gridSpan > 1) {
		tcPr.push(`<w:gridSpan w:val="${cell.properties.gridSpan}"/>`);
	}
	if (cell.properties.verticalMerge) {
		if (cell.properties.verticalMerge === "restart") {
			tcPr.push('<w:vMerge w:val="restart"/>');
		} else {
			tcPr.push("<w:vMerge/>");
		}
	}
	if (cell.properties.shading) {
		tcPr.push(`<w:shd w:val="clear" w:fill="${escXml(cell.properties.shading)}"/>`);
	}
	if (cell.properties.verticalAlign) {
		tcPr.push(`<w:vAlign w:val="${cell.properties.verticalAlign}"/>`);
	}
	if (tcPr.length > 0) {
		xml += `<w:tcPr>${tcPr.join("")}</w:tcPr>`;
	}

	// Cell paragraphs
	for (const para of cell.paragraphs) {
		xml += serializeParagraph(para);
	}

	// Ensure at least one paragraph (OOXML requires it)
	if (cell.paragraphs.length === 0) {
		xml += "<w:p/>";
	}

	xml += "</w:tc>";
	return xml;
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/** Escape XML special characters in text content and attribute values. */
function escXml(s: string): string {
	return s
		.replace(/&/g, "&amp;")
		.replace(/</g, "&lt;")
		.replace(/>/g, "&gt;")
		.replace(/"/g, "&quot;");
}

/** Convert CSS pixels to EMU (English Metric Units). */
function pxToEmu(px: number): number {
	// 1 inch = 96 px = 914400 EMU
	return Math.round((px / 96) * 914400);
}

/**
 * Ensure a relationship entry exists in word/_rels/document.xml.rels.
 * Used when new images are inserted during editing.
 */
async function ensureRelationship(
	zip: { file(path: string): { async(type: "string"): Promise<string> } | null; file(path: string, data: string): void },
	rId: string,
	type: string,
	target: string,
): Promise<void> {
	const relsPath = "word/_rels/document.xml.rels";
	const relsFile = zip.file(relsPath);
	if (!relsFile) return;

	let relsXml = await relsFile.async("string");

	// Check if this rId already exists
	if (relsXml.includes(`Id="${rId}"`)) return;

	// Insert new Relationship before closing </Relationships>
	const newRel = `<Relationship Id="${escXml(rId)}" Type="${escXml(type)}" Target="${escXml(target)}"/>`;
	relsXml = relsXml.replace("</Relationships>", `${newRel}\n</Relationships>`);

	zip.file(relsPath, relsXml);
}
