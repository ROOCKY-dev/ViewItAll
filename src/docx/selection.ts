/**
 * ViewItAll — DOCX Selection Utilities
 *
 * Maps between browser Selection/Range and model coordinates.
 * Used by the formatting toolbar to determine what is selected
 * and by the editing controller to restore caret after re-renders.
 */

import type {
	DocxDocument,
	DocxRunProperties,
	DocxParagraphProperties,
} from "./model";
import { defaultRunProperties, defaultParagraphProperties } from "./model";
import { resolveStyle } from "./styles";

// ── Model Selection ──────────────────────────────────────────────────────────

export interface ModelSelection {
	startBlock: number;
	startRun: number;
	startOffset: number;
	endBlock: number;
	endRun: number;
	endOffset: number;
	collapsed: boolean;
}

/**
 * Convert browser Selection to model coordinates.
 */
export function domSelectionToModel(
	contentEl: HTMLElement,
	doc: DocxDocument,
): ModelSelection | null {
	const sel = window.getSelection();
	if (!sel || !sel.rangeCount) return null;

	const range = sel.getRangeAt(0);

	const start = resolveNodeToModel(range.startContainer, range.startOffset, contentEl, doc);
	if (!start) return null;

	if (sel.isCollapsed) {
		return {
			...start,
			endBlock: start.startBlock,
			endRun: start.startRun,
			endOffset: start.startOffset,
			collapsed: true,
		};
	}

	const end = resolveNodeToModel(range.endContainer, range.endOffset, contentEl, doc);
	if (!end) return null;

	return {
		startBlock: start.startBlock,
		startRun: start.startRun,
		startOffset: start.startOffset,
		endBlock: end.startBlock,
		endRun: end.startRun,
		endOffset: end.startOffset,
		collapsed: false,
	};
}

/**
 * Get the resolved run properties at the current caret/selection start.
 * Merges style-level + run-level properties.
 */
export function getRunPropertiesAtSelection(
	msel: ModelSelection,
	doc: DocxDocument,
	styleCache: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>,
): DocxRunProperties {
	const block = doc.body[msel.startBlock];
	if (!block || block.type !== "paragraph") return defaultRunProperties();

	const para = block;

	// Get style-level run props
	let styleRProps = defaultRunProperties();
	if (para.styleId) {
		const resolved = resolveStyle(para.styleId, doc.styles, styleCache);
		styleRProps = resolved.rProps;
	}

	// Get the run at selection start
	const child = para.children[msel.startRun];
	if (!child || child.type !== "run") return styleRProps;

	// Merge run-level overrides onto style-level base
	const merged = { ...styleRProps };
	const override = child.properties;
	if (override.bold !== undefined) merged.bold = override.bold;
	if (override.italic !== undefined) merged.italic = override.italic;
	if (override.underline !== undefined) merged.underline = override.underline;
	if (override.strikethrough !== undefined) merged.strikethrough = override.strikethrough;
	if (override.fontFamily !== undefined) merged.fontFamily = override.fontFamily;
	if (override.fontSize !== undefined) merged.fontSize = override.fontSize;
	if (override.color !== undefined) {
		merged.color = override.color === "auto" ? undefined : override.color;
	}
	if (override.highlight !== undefined) merged.highlight = override.highlight;
	if (override.vertAlign !== undefined) merged.vertAlign = override.vertAlign;

	return merged;
}

/**
 * Get paragraph properties at the current selection start.
 */
export function getParagraphPropertiesAtSelection(
	msel: ModelSelection,
	doc: DocxDocument,
): DocxParagraphProperties {
	const block = doc.body[msel.startBlock];
	if (!block || block.type !== "paragraph") return defaultParagraphProperties();
	return block.properties;
}

// ── Helpers ──────────────────────────────────────────────────────────────────

function resolveNodeToModel(
	node: Node,
	offset: number,
	contentEl: HTMLElement,
	doc: DocxDocument,
): { startBlock: number; startRun: number; startOffset: number } | null {
	// Walk up to find paragraph with data-block-idx
	let paraEl: HTMLElement | null = null;
	let current: HTMLElement | null = node instanceof HTMLElement ? node : node.parentElement;

	while (current && current !== contentEl) {
		if (current.dataset.blockIdx !== undefined) {
			const tag = current.tagName.toLowerCase();
			if (tag === "p" || /^h[1-6]$/.test(tag)) {
				paraEl = current;
				break;
			}
		}
		current = current.parentElement;
	}

	if (!paraEl) return null;

	const blockIdx = parseInt(paraEl.dataset.blockIdx ?? "", 10);
	if (isNaN(blockIdx)) return null;

	const block = doc.body[blockIdx];
	if (!block || block.type !== "paragraph") return null;

	// Find which run this node belongs to
	let runSpan: HTMLElement | null = null;
	current = node instanceof HTMLElement ? node : node.parentElement;
	while (current && current !== paraEl) {
		if (current.dataset.runIdx !== undefined) {
			runSpan = current;
			break;
		}
		current = current.parentElement;
	}

	const runIdx = runSpan ? parseInt(runSpan.dataset.runIdx ?? "", 10) : 0;

	// Calculate offset within the run
	let runOffset = offset;
	if (node.nodeType === Node.TEXT_NODE && runSpan) {
		// Count text before this text node within the run span
		const textWalker = document.createTreeWalker(runSpan, NodeFilter.SHOW_TEXT);
		let textBefore = 0;
		let n = textWalker.nextNode();
		while (n && n !== node) {
			textBefore += (n.textContent ?? "").length;
			n = textWalker.nextNode();
		}
		runOffset = textBefore + offset;
	}

	return {
		startBlock: blockIdx,
		startRun: isNaN(runIdx) ? 0 : runIdx,
		startOffset: runOffset,
	};
}

/**
 * Restore a model selection back to DOM range.
 */
export function modelSelectionToDom(
	msel: ModelSelection,
	contentEl: HTMLElement,
): void {
	const sel = window.getSelection();
	if (!sel) return;

	const startNode = findTextNodeAtPosition(
		msel.startBlock,
		msel.startRun,
		msel.startOffset,
		contentEl,
	);

	if (!startNode) return;

	const range = document.createRange();
	range.setStart(startNode.node, startNode.offset);

	if (msel.collapsed) {
		range.collapse(true);
	} else {
		const endNode = findTextNodeAtPosition(
			msel.endBlock,
			msel.endRun,
			msel.endOffset,
			contentEl,
		);
		if (endNode) {
			range.setEnd(endNode.node, endNode.offset);
		}
	}

	sel.removeAllRanges();
	sel.addRange(range);
}

function findTextNodeAtPosition(
	blockIdx: number,
	runIdx: number,
	offset: number,
	contentEl: HTMLElement,
): { node: Node; offset: number } | null {
	const paraEl = contentEl.querySelector<HTMLElement>(`[data-block-idx="${blockIdx}"]`);
	if (!paraEl) return null;

	const runSpan = paraEl.querySelector<HTMLElement>(`[data-run-idx="${runIdx}"]`);
	const target = runSpan ?? paraEl;

	const walker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT, {
		acceptNode: (node: Node) => {
			const parent = node.parentElement;
			if (parent?.classList.contains("via-docx-num-prefix")) {
				return NodeFilter.FILTER_REJECT;
			}
			return NodeFilter.FILTER_ACCEPT;
		},
	});

	let remaining = offset;
	let textNode = walker.nextNode();
	while (textNode) {
		const len = (textNode.textContent ?? "").length;
		if (remaining <= len) {
			return { node: textNode, offset: remaining };
		}
		remaining -= len;
		textNode = walker.nextNode();
	}

	// Fallback: last text node at end
	if (target.lastChild) {
		const lastWalker = document.createTreeWalker(target, NodeFilter.SHOW_TEXT);
		let last: Node | null = null;
		let n = lastWalker.nextNode();
		while (n) { last = n; n = lastWalker.nextNode(); }
		if (last) {
			return { node: last, offset: (last.textContent ?? "").length };
		}
	}

	return null;
}
