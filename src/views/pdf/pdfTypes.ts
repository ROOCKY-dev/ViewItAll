/** Shared types for PDF page rendering, used by PdfView and PdfSearchController. */

export type PageRenderState = "placeholder" | "rendering" | "rendered";

export interface PageCtx {
	pageNum: number;
	state: PageRenderState;
	container: HTMLElement;
	pdfCanvas: HTMLCanvasElement | null;
	annotCanvas: HTMLCanvasElement | null;
	searchCanvas: HTMLCanvasElement | null;
	w: number;
	h: number;
}

export interface SearchMatch {
	pageNum: number;
	/** Normalised 0-1 fractions of page viewport */
	x: number;
	y: number;
	w: number;
	h: number;
}
