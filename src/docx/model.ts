/**
 * ViewItAll — DOCX Document Model
 *
 * Pure TypeScript interfaces representing the structure of a .docx (OOXML)
 * document. No DOM, no Obsidian — this is a data-only layer.
 */

// ── Block-level elements ────────────────────────────────────────────────────

export type DocxBlockElement =
	| DocxParagraph
	| DocxTable
	| DocxPageBreak
	| DocxSectionBreak;

export interface DocxParagraph {
	type: "paragraph";
	styleId: string | undefined;
	properties: DocxParagraphProperties;
	children: DocxInlineElement[];
	numberingId: string | undefined;
	numberingLevel: number;
}

export interface DocxParagraphProperties {
	alignment: "left" | "center" | "right" | "both" | undefined;
	headingLevel: number | undefined;
	indentation:
		| { left: number; right: number; firstLine: number; hanging: number }
		| undefined;
	spacing:
		| { before: number; after: number; line: number; lineRule: string }
		| undefined;
	/** Numbering ID from style or direct pPr */
	numberingId: string | undefined;
	/** Numbering indent level */
	numberingLevel: number | undefined;
}

export interface DocxPageBreak {
	type: "pageBreak";
}

export interface DocxSectionBreak {
	type: "sectionBreak";
}

// ── Inline elements ─────────────────────────────────────────────────────────

export type DocxInlineElement =
	| DocxRun
	| DocxHyperlink
	| DocxImage
	| DocxBreak
	| DocxTab;

export interface DocxRun {
	type: "run";
	text: string;
	/** Run-level overrides — only properties explicitly in the XML are set. */
	properties: Partial<DocxRunProperties>;
}

export interface DocxRunProperties {
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strikethrough: boolean;
	fontFamily: string | undefined;
	fontSize: number | undefined;
	color: string | undefined;
	highlight: string | undefined;
	vertAlign: "superscript" | "subscript" | undefined;
}

export interface DocxHyperlink {
	type: "hyperlink";
	url: string;
	children: DocxRun[];
}

export interface DocxImage {
	type: "image";
	relationshipId: string;
	width: number;
	height: number;
	altText: string | undefined;
}

export interface DocxBreak {
	type: "break";
	breakType: "line" | "page" | "column";
}

export interface DocxTab {
	type: "tab";
}

// ── Table ───────────────────────────────────────────────────────────────────

export interface DocxTable {
	type: "table";
	rows: DocxTableRow[];
	properties: DocxTableProperties;
}

export interface DocxTableRow {
	cells: DocxTableCell[];
	isHeader: boolean;
}

export interface DocxTableCell {
	paragraphs: DocxParagraph[];
	properties: DocxTableCellProperties;
}

export interface DocxTableProperties {
	width: number | undefined;
	alignment: "left" | "center" | "right" | undefined;
}

export interface DocxTableCellProperties {
	width: number | undefined;
	verticalMerge: "restart" | "continue" | undefined;
	gridSpan: number;
	shading: string | undefined;
	verticalAlign: "top" | "center" | "bottom" | undefined;
}

// ── Styles ──────────────────────────────────────────────────────────────────

export interface DocxStyle {
	id: string;
	name: string;
	type: "paragraph" | "character" | "table" | "numbering";
	basedOn: string | undefined;
	paragraphProperties: Partial<DocxParagraphProperties>;
	runProperties: Partial<DocxRunProperties>;
}

// ── Numbering ───────────────────────────────────────────────────────────────

export interface DocxNumberingLevel {
	format:
		| "decimal"
		| "bullet"
		| "lowerLetter"
		| "upperLetter"
		| "lowerRoman"
		| "upperRoman"
		| "none";
	text: string;
	indentLeft: number;
}

export interface DocxNumberingDef {
	abstractNumId: string;
	levels: Map<number, DocxNumberingLevel>;
}

// ── Root document ───────────────────────────────────────────────────────────

export interface DocxDocument {
	body: DocxBlockElement[];
	styles: Map<string, DocxStyle>;
	images: Map<string, Blob>;
	numbering: Map<string, DocxNumberingDef>;
	relationships: Map<string, { type: string; target: string }>;
}

// ── Default factories ───────────────────────────────────────────────────────

export function defaultRunProperties(): DocxRunProperties {
	return {
		bold: false,
		italic: false,
		underline: false,
		strikethrough: false,
		fontFamily: undefined,
		fontSize: undefined,
		color: undefined,
		highlight: undefined,
		vertAlign: undefined,
	};
}

export function defaultParagraphProperties(): DocxParagraphProperties {
	return {
		alignment: undefined,
		headingLevel: undefined,
		indentation: undefined,
		spacing: undefined,
		numberingId: undefined,
		numberingLevel: undefined,
	};
}
