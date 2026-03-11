/** Minimal JSZip type surface used by PPTX parsing. */
export interface JSZipObject {
	async(type: 'string'): Promise<string>;
	async(type: 'uint8array'): Promise<Uint8Array>;
	async(type: 'arraybuffer'): Promise<ArrayBuffer>;
	async(type: 'base64'): Promise<string>;
}

export interface JSZipInstance {
	file(path: string): JSZipObject | null;
	file(path: string, data: string | ArrayBuffer | Uint8Array): JSZipInstance;
	file(regex: RegExp): { name: string }[];
	generateAsync(options: { type: 'arraybuffer' }): Promise<ArrayBuffer>;
	loadAsync(data: ArrayBuffer | Uint8Array): Promise<JSZipInstance>;
}

export interface JSZipConstructor {
	new (): JSZipInstance;
	loadAsync(data: ArrayBuffer | Uint8Array): Promise<JSZipInstance>;
}

/** Run-level text data with lightweight formatting tokens. */
export interface RunData {
	text: string;
	bold: boolean;
	italic: boolean;
	underline: boolean;
	fontSizePt: number | null;
	colorToken: string | null;
}

/** A paragraph within a shape. */
export interface ParagraphData {
	runs: RunData[];
	isBullet: boolean;
	level: number;
}

export type ShapePlaceholderType = 'title' | 'ctrTitle' | 'subTitle' | 'body' | 'other';

export interface ShapeBounds {
	xEmu: number;
	yEmu: number;
	widthEmu: number;
	heightEmu: number;
	rotationDeg: number;
}

export interface ShapeStyle {
	fillToken: string | null;
	lineToken: string | null;
	lineWidth: number | null;
	lineDash: string | null;
}

export interface TableCellData {
	text: string;
}

export interface TableRowData {
	cells: TableCellData[];
}

export interface TableData {
	id: string;
	zIndex: number;
	bounds: ShapeBounds | null;
	rows: TableRowData[];
}

/** A shape extracted from a slide. */
export interface ShapeData {
	id: string;
	type: ShapePlaceholderType;
	zIndex: number;
	bounds: ShapeBounds | null;
	style: ShapeStyle;
	paragraphs: ParagraphData[];
}

/** Parsed slide data. */
export interface SlideData {
	index: number;
	backgroundToken: string | null;
	shapes: ShapeData[];
	tables: TableData[];
	imageDataUrls: string[];
}
