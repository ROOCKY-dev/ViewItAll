import mammoth from 'mammoth';
// @ts-expect-error — html-to-docx has no bundled type declarations
import HtmlToDocx from 'html-to-docx';

/**
 * Converts a .docx ArrayBuffer to an HTML string for display/editing.
 * Returns both the html content and any conversion messages/warnings.
 */
export async function readDocxAsHtml(
	buffer: ArrayBuffer
): Promise<{ html: string; messages: string[] }> {
	const result = await mammoth.convertToHtml({ arrayBuffer: buffer });
	// Filter out pure style-mapping noise — "Unrecognised paragraph/run style" messages
	// only mean mammoth doesn't know the custom Word style name; content is unaffected.
	const messages = result.messages
		.filter(m => m.type === 'warning')
		.filter(m => !m.message.startsWith('Unrecognised paragraph style') &&
		             !m.message.startsWith('Unrecognised run style'))
		.map(m => m.message);
	return { html: result.value, messages };
}

/**
 * Converts an HTML string back to a .docx ArrayBuffer for saving.
 * Note: Complex formatting (merged table cells, custom styles) may be simplified.
 */
export async function saveHtmlAsDocx(html: string): Promise<ArrayBuffer> {
	// HtmlToDocx returns Blob in pure-browser contexts, Buffer in Electron/Node
	const result = await HtmlToDocx(html, undefined, {
		table: { row: { cantSplit: true } },
		footer: false,
		pageNumber: false,
	});
	if (result instanceof Blob) {
		return result.arrayBuffer();
	}
	// Node.js Buffer → ArrayBuffer (zero-copy slice)
	const buf = result as Buffer;
	return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}
