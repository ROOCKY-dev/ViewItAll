import mammoth from "mammoth";
// @ts-expect-error — html-to-docx has no bundled type declarations
import htmlToDocxModule from "html-to-docx";

/** Shape of a Node.js Buffer as returned by html-to-docx in Electron/Node contexts. */
interface NodeJsBuffer {
	buffer: ArrayBuffer;
	byteOffset: number;
	byteLength: number;
}

type HtmlToDocxFn = (
	html: string,
	headerHtml: string | undefined,
	options: Record<string, unknown>,
) => Promise<Blob | NodeJsBuffer>;

// Cast the untyped import to its known function signature.
const htmlToDocx = htmlToDocxModule as unknown as HtmlToDocxFn;

/**
 * Converts a .docx ArrayBuffer to an HTML string for display/editing.
 * Returns both the html content and any conversion messages/warnings.
 */
export async function readDocxAsHtml(
	buffer: ArrayBuffer,
): Promise<{ html: string; messages: string[] }> {
	const result = await mammoth.convertToHtml({ arrayBuffer: buffer });
	// Filter out pure style-mapping noise — "Unrecognised paragraph/run style" messages
	// only mean mammoth doesn't know the custom Word style name; content is unaffected.
	const messages = result.messages
		.filter((m) => m.type === "warning")
		.filter(
			(m) =>
				!m.message.startsWith("Unrecognised paragraph style") &&
				!m.message.startsWith("Unrecognised run style"),
		)
		.map((m) => m.message);
	return { html: result.value, messages };
}

/**
 * Converts an HTML string back to a .docx ArrayBuffer for saving.
 * Note: Complex formatting (merged table cells, custom styles) may be simplified.
 */
export async function saveHtmlAsDocx(html: string): Promise<ArrayBuffer> {
	// htmlToDocx returns Blob in browser contexts, NodeJsBuffer in Electron/Node
	const result = await htmlToDocx(html, undefined, {
		table: { row: { cantSplit: true } },
		footer: false,
		pageNumber: false,
	});
	if (result instanceof Blob) {
		return result.arrayBuffer();
	}
	// Zero-copy slice of the Node.js Buffer backing array
	return result.buffer.slice(
		result.byteOffset,
		result.byteOffset + result.byteLength,
	);
}
