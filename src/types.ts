export const VIEW_TYPE_DOCX = "viewitall-docx";
export const VIEW_TYPE_PDF = "viewitall-pdf";
export const VIEW_TYPE_SPREADSHEET = "viewitall-spreadsheet";
export const VIEW_TYPE_PPTX = "viewitall-pptx";

export interface AnnotationPath {
	tool: "pen" | "highlighter" | "eraser";
	color: string;
	width: number;
	opacity?: number;
	points: { x: number; y: number }[];
}

export interface PageAnnotations {
	page: number;
	paths: AnnotationPath[];
}

/** A text note pinned to a normalised (0-1) position on a PDF page. */
export interface TextNote {
	id: string;
	page: number;
	x: number; // normalised 0-1
	y: number; // normalised 0-1
	text: string;
	color?: string; // background swatch color, defaults to '#ffd43b'
}

export interface AnnotationFile {
	version: 1;
	pages: PageAnnotations[];
	notes?: TextNote[];
}
