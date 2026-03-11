export const VIEW_TYPE_DOCX = 'viewitall-docx';
export const VIEW_TYPE_PDF  = 'viewitall-pdf';

export interface AnnotationPath {
	tool: 'pen' | 'highlighter' | 'eraser';
	color: string;
	width: number;
	opacity?: number;
	points: { x: number; y: number }[];
}

export interface PageAnnotations {
	page: number;
	paths: AnnotationPath[];
}

export interface AnnotationFile {
	version: 1;
	pages: PageAnnotations[];
}
