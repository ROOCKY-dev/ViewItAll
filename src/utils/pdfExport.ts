import { App, TFile, Notice } from 'obsidian';
import * as pdfjsLib from 'pdfjs-dist';
import { PDFDocument, rgb, LineCapStyle } from 'pdf-lib';
import type { AnnotationFile } from '../types';

/**
 * Embed pen/highlighter annotations as vector SVG paths into a copy of the
 * source PDF and save it as `<basename>.annotated.pdf` next to the original.
 */
export async function exportAnnotatedPdf(
	app:         App,
	currentFile: TFile,
	pdfDoc:      pdfjsLib.PDFDocumentProxy,
	annotData:   AnnotationFile,
): Promise<void> {
	new Notice('Exporting PDF with annotations…');

	try {
		const srcBuffer = await app.vault.adapter.readBinary(currentFile.path);
		const pdfLibDoc = await PDFDocument.load(srcBuffer);
		const pages     = pdfLibDoc.getPages();

		const hexToRgb = (hex: string) => {
			const n = parseInt(hex.replace('#', ''), 16);
			return rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
		};

		for (const pa of annotData.pages) {
			const libPage = pages[pa.page - 1];
			if (!libPage || pa.paths.length === 0) continue;

			const pjsPage = await pdfDoc.getPage(pa.page);
			const vp   = pjsPage.getViewport({ scale: 1.0 });
			const pdfW = libPage.getWidth();
			const pdfH = libPage.getHeight();
			const scaleX = pdfW / vp.width;
			const scaleY = pdfH / vp.height;

			for (const path of pa.paths) {
				if (path.tool === 'eraser' || path.points.length < 2) continue;

				const pts = path.points.map((p: { x: number; y: number }) => ({
					px: p.x * vp.width  * scaleX,
					// PDF coordinate system has bottom-left origin (y-flipped vs canvas)
					py: pdfH - p.y * vp.height * scaleY,
				}));

				const lineWidth = (path.width * scaleX + path.width * scaleY) / 2;
				const opacity   = path.tool === 'highlighter' ? (path.opacity ?? 0.35) : 1;

				let d = `M ${pts[0]!.px.toFixed(2)} ${pts[0]!.py.toFixed(2)}`;
				for (let i = 1; i < pts.length; i++) {
					d += ` L ${pts[i]!.px.toFixed(2)} ${pts[i]!.py.toFixed(2)}`;
				}

				libPage.drawSvgPath(d, {
					borderColor:   hexToRgb(path.color),
					borderWidth:   lineWidth,
					borderOpacity: opacity,
					borderLineCap: LineCapStyle.Round,
					opacity: 0, // fill: none — stroke-only path
				});
			}
		}

		const exportBytes  = await pdfLibDoc.save();
		const exportBuffer = exportBytes.buffer as ArrayBuffer;

		const dir        = currentFile.parent?.path ?? '';
		const base       = currentFile.basename;
		const exportPath = dir ? `${dir}/${base}.annotated.pdf` : `${base}.annotated.pdf`;

		const existing = app.vault.getAbstractFileByPath(exportPath);
		if (existing && existing instanceof TFile) {
			await app.vault.modifyBinary(existing, exportBuffer);
		} else {
			await app.vault.createBinary(exportPath, exportBuffer);
		}

		new Notice(`✅ Exported to "${base}.annotated.pdf"`);
	} catch (err) {
		new Notice(`❌ Export failed: ${String(err)}`);
	}
}
