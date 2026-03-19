/**
 * ViewItAll — DOCX Editing Controller
 *
 * Bridges contentEditable DOM events to model mutations.
 * Uses per-paragraph contentEditable for text input while
 * keeping the model as the source of truth for all structure.
 */

import type {
	DocxDocument,
	DocxParagraph,
	DocxInlineElement,
	DocxRun,
	DocxRunProperties,
	DocxParagraphProperties,
} from "./model";
import { defaultParagraphProperties } from "./model";
import { renderDocument, rerenderParagraph } from "./renderer";
import type { FormatCommand } from "./toolbar";
import {
	domSelectionToModel,
	modelSelectionToDom,
	type ModelSelection,
} from "./selection";
import { createEditHistory, type EditHistory } from "./history";

// ── Public Interface ─────────────────────────────────────────────────────────

export interface EditingController {
	enable(
		contentEl: HTMLElement,
		doc: DocxDocument,
		styleCache: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>,
		onChange: () => void,
	): void;
	disable(): void;
	isEnabled(): boolean;
	applyFormat(cmd: FormatCommand): void;
	undo(): boolean;
	redo(): boolean;
}

export function createEditingController(): EditingController {
	let enabled = false;
	let contentEl: HTMLElement | null = null;
	let doc: DocxDocument | null = null;
	let styleCache: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }> | null = null;
	let onChange: (() => void) | null = null;
	let blobUrls: string[] = [];
	const history: EditHistory = createEditHistory();

	// Bound handlers for cleanup
	let inputHandler: ((e: Event) => void) | null = null;
	let keydownHandler: ((e: KeyboardEvent) => void) | null = null;
	let mousedownHandler: ((e: MouseEvent) => void) | null = null;

	// Debounce timer for input snapshots
	let inputSnapshotTimer: ReturnType<typeof setTimeout> | null = null;

	// Image resize state
	let resizeState: {
		imageRid: string;
		wrap: HTMLElement;
		img: HTMLImageElement;
		startX: number;
		startWidth: number;
		aspectRatio: number;
		moveHandler: (e: MouseEvent) => void;
		upHandler: (e: MouseEvent) => void;
	} | null = null;

	function enable(
		el: HTMLElement,
		d: DocxDocument,
		sc: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>,
		cb: () => void,
	): void {
		contentEl = el;
		doc = d;
		styleCache = sc;
		onChange = cb;
		enabled = true;

		// Add editing class
		contentEl.classList.add("via-docx-content--editing");

		// Make all paragraphs and headings editable
		setEditableState(contentEl, true);

		// Register event handlers via delegation
		inputHandler = (e: Event) => handleInput(e);
		keydownHandler = (e: KeyboardEvent) => handleKeydown(e);
		mousedownHandler = (e: MouseEvent) => handleResizeMousedown(e);
		contentEl.addEventListener("input", inputHandler);
		contentEl.addEventListener("keydown", keydownHandler);
		contentEl.addEventListener("mousedown", mousedownHandler);
	}

	function disable(): void {
		if (contentEl) {
			contentEl.classList.remove("via-docx-content--editing");
			setEditableState(contentEl, false);

			if (inputHandler) {
				contentEl.removeEventListener("input", inputHandler);
				inputHandler = null;
			}
			if (keydownHandler) {
				contentEl.removeEventListener("keydown", keydownHandler);
				keydownHandler = null;
			}
			if (mousedownHandler) {
				contentEl.removeEventListener("mousedown", mousedownHandler);
				mousedownHandler = null;
			}
		}

		// Clean up any active resize
		if (resizeState) {
			document.removeEventListener("mousemove", resizeState.moveHandler);
			document.removeEventListener("mouseup", resizeState.upHandler);
			resizeState = null;
		}

		// Clear history and timers
		history.clear();
		if (inputSnapshotTimer) {
			clearTimeout(inputSnapshotTimer);
			inputSnapshotTimer = null;
		}

		// Revoke any blob URLs created during editing
		for (const url of blobUrls) {
			URL.revokeObjectURL(url);
		}
		blobUrls = [];

		enabled = false;
		contentEl = null;
		doc = null;
		styleCache = null;
		onChange = null;
	}

	function isEnabled(): boolean {
		return enabled;
	}

	// ── ContentEditable State ────────────────────────────────────────────

	function setEditableState(container: HTMLElement, editable: boolean): void {
		const editableValue = editable ? "true" : "false";
		const paragraphs = Array.from(container.querySelectorAll<HTMLElement>(
			"p[data-block-idx], h1[data-block-idx], h2[data-block-idx], " +
			"h3[data-block-idx], h4[data-block-idx], h5[data-block-idx], h6[data-block-idx]",
		));
		for (const p of paragraphs) {
			if (editable) {
				p.setAttribute("contenteditable", editableValue);
			} else {
				p.removeAttribute("contenteditable");
			}
		}

		// Also make table cells editable
		const cells = Array.from(container.querySelectorAll<HTMLElement>(".via-docx-cell"));
		for (const cell of cells) {
			if (editable) {
				cell.setAttribute("contenteditable", editableValue);
			} else {
				cell.removeAttribute("contenteditable");
			}
		}
	}

	// ── Input Handler ────────────────────────────────────────────────────

	function handleInput(e: Event): void {
		if (!doc || !contentEl || !styleCache) return;

		const target = e.target as HTMLElement;
		const paraEl = findParagraphEl(target);
		if (!paraEl) return;

		const blockIdx = parseInt(paraEl.dataset.blockIdx ?? "", 10);
		if (isNaN(blockIdx)) return;

		const block = doc.body[blockIdx];
		if (!block || block.type !== "paragraph") return;

		// Debounced snapshot — only snapshot every 500ms of typing
		if (inputSnapshotTimer) clearTimeout(inputSnapshotTimer);
		inputSnapshotTimer = setTimeout(() => {
			if (doc) history.snapshot(doc.body);
		}, 500);

		// Sync DOM text content back to model runs
		syncDomToModel(paraEl, block);
		onChange?.();
	}

	// ── Keydown Handler ──────────────────────────────────────────────────

	function handleKeydown(e: KeyboardEvent): void {
		if (!doc || !contentEl || !styleCache) return;

		// Formatting and history shortcuts
		if (e.ctrlKey || e.metaKey) {
			if (e.key === "b" || e.key === "B") {
				e.preventDefault();
				applyFormat({ type: "bold" });
				return;
			}
			if (e.key === "i" || e.key === "I") {
				e.preventDefault();
				applyFormat({ type: "italic" });
				return;
			}
			if (e.key === "u" || e.key === "U") {
				e.preventDefault();
				applyFormat({ type: "underline" });
				return;
			}
			if (e.key === "z" || e.key === "Z") {
				e.preventDefault();
				if (e.shiftKey) {
					redo();
				} else {
					undo();
				}
				return;
			}
			if (e.key === "y" || e.key === "Y") {
				e.preventDefault();
				redo();
				return;
			}
		}

		if (e.key === "Enter" && !e.shiftKey) {
			e.preventDefault();
			handleEnter();
		} else if (e.key === "Backspace") {
			handleBackspace(e);
		} else if (e.key === "Delete") {
			handleDelete(e);
		} else if (e.key === "Tab") {
			// In table cells, move to next/prev cell
			handleTab(e);
		}
	}

	function handleEnter(): void {
		if (!doc || !contentEl || !styleCache || !onChange) return;
		history.snapshot(doc.body);

		const sel = window.getSelection();
		if (!sel || !sel.rangeCount) return;

		const range = sel.getRangeAt(0);
		const paraEl = findParagraphEl(range.startContainer as HTMLElement);
		if (!paraEl) return;

		const blockIdx = parseInt(paraEl.dataset.blockIdx ?? "", 10);
		if (isNaN(blockIdx)) return;

		const block = doc.body[blockIdx];
		if (!block || block.type !== "paragraph") return;

		// First, sync any pending changes
		syncDomToModel(paraEl, block);

		// Find which run and offset the caret is at
		const caretPos = getCaretPosition(paraEl);

		// Split the paragraph at the caret position
		const { before, after } = splitParagraphChildren(block, caretPos);

		// Update current paragraph's children
		block.children = before;

		// Create new paragraph
		const newPara: DocxParagraph = {
			type: "paragraph",
			styleId: block.styleId,
			properties: { ...defaultParagraphProperties(), alignment: block.properties.alignment },
			children: after,
			numberingId: undefined,
			numberingLevel: 0,
		};

		// Insert after current block
		doc.body.splice(blockIdx + 1, 0, newPara);

		// Re-render everything and restore caret
		reRenderAll();
		setEditableState(contentEl, true);

		// Place caret at start of new paragraph
		const newParaEl = contentEl.querySelector<HTMLElement>(`[data-block-idx="${blockIdx + 1}"]`);
		if (newParaEl) {
			placeCaretAtStart(newParaEl);
		}

		onChange();
	}

	function handleBackspace(e: KeyboardEvent): void {
		if (!doc || !contentEl || !styleCache || !onChange) return;

		const sel = window.getSelection();
		if (!sel || !sel.rangeCount) return;

		// Handle cross-paragraph or multi-character selection deletion
		if (!sel.isCollapsed) {
			e.preventDefault();
			deleteSelection();
			return;
		}

		const range = sel.getRangeAt(0);
		const paraEl = findParagraphEl(range.startContainer as HTMLElement);
		if (!paraEl) return;

		const blockIdx = parseInt(paraEl.dataset.blockIdx ?? "", 10);
		if (isNaN(blockIdx) || blockIdx === 0) return;

		// Check if caret is at the very start
		const caretPos = getCaretPosition(paraEl);
		if (caretPos > 0) return; // Not at start, let browser handle it

		e.preventDefault();
		history.snapshot(doc.body);

		const prevBlock = doc.body[blockIdx - 1];
		if (!prevBlock || prevBlock.type !== "paragraph") return;

		// Sync current paragraph first
		const block = doc.body[blockIdx];
		if (!block || block.type !== "paragraph") return;
		syncDomToModel(paraEl, block);

		// Remember where to place caret (end of previous paragraph)
		const prevTextLen = getTextLength(prevBlock);

		// Merge: append current children to previous paragraph
		// Filter out empty runs from current before appending
		const nonEmptyChildren = block.children.filter(
			(c) => c.type !== "run" || c.text.length > 0,
		);
		prevBlock.children.push(...nonEmptyChildren);

		// Remove current block from model
		doc.body.splice(blockIdx, 1);

		// Re-render
		reRenderAll();
		setEditableState(contentEl, true);

		// Place caret at the merge point
		const mergedEl = contentEl.querySelector<HTMLElement>(`[data-block-idx="${blockIdx - 1}"]`);
		if (mergedEl) {
			placeCaretAtOffset(mergedEl, prevTextLen);
		}

		onChange();
	}

	function handleDelete(e: KeyboardEvent): void {
		if (!doc || !contentEl || !styleCache || !onChange) return;

		const sel = window.getSelection();
		if (!sel || !sel.rangeCount) return;

		// Handle cross-paragraph or multi-character selection deletion
		if (!sel.isCollapsed) {
			e.preventDefault();
			deleteSelection();
			return;
		}

		const range = sel.getRangeAt(0);
		const paraEl = findParagraphEl(range.startContainer as HTMLElement);
		if (!paraEl) return;

		const blockIdx = parseInt(paraEl.dataset.blockIdx ?? "", 10);
		if (isNaN(blockIdx) || blockIdx >= doc.body.length - 1) return;

		const block = doc.body[blockIdx];
		if (!block || block.type !== "paragraph") return;

		// Check if caret is at the very end
		syncDomToModel(paraEl, block);
		const textLen = getTextLength(block);
		const caretPos = getCaretPosition(paraEl);
		if (caretPos < textLen) return; // Not at end, let browser handle

		e.preventDefault();
		history.snapshot(doc.body);

		const nextBlock = doc.body[blockIdx + 1];
		if (!nextBlock || nextBlock.type !== "paragraph") return;

		// Merge: append next children to current paragraph
		block.children.push(...nextBlock.children);

		// Remove next block
		doc.body.splice(blockIdx + 1, 1);

		// Re-render
		reRenderAll();
		setEditableState(contentEl, true);

		// Place caret at the position before the merge
		const mergedEl = contentEl.querySelector<HTMLElement>(`[data-block-idx="${blockIdx}"]`);
		if (mergedEl) {
			placeCaretAtOffset(mergedEl, textLen);
		}

		onChange();
	}

	function handleTab(e: KeyboardEvent): void {
		// Check if we're inside a table cell
		const sel = window.getSelection();
		if (!sel || !sel.rangeCount) return;

		let current: HTMLElement | null = sel.anchorNode instanceof HTMLElement
			? sel.anchorNode
			: sel.anchorNode?.parentElement ?? null;

		let cell: HTMLElement | null = null;
		while (current && current !== contentEl) {
			if (current.classList.contains("via-docx-cell")) {
				cell = current;
				break;
			}
			current = current.parentElement;
		}

		if (!cell) return;

		e.preventDefault();

		// Find all cells in the table
		const table = cell.closest("table");
		if (!table) return;
		const cells = Array.from(table.querySelectorAll<HTMLElement>(".via-docx-cell"));
		const idx = cells.indexOf(cell);
		if (idx === -1) return;

		const nextIdx = e.shiftKey ? idx - 1 : idx + 1;
		const nextCell = cells[nextIdx];
		if (nextCell) {
			nextCell.focus();
			// Place caret at start of cell
			const range = document.createRange();
			range.selectNodeContents(nextCell);
			range.collapse(!e.shiftKey); // start for Tab, end for Shift+Tab
			sel.removeAllRanges();
			sel.addRange(range);
		}
	}

	// ── Selection Deletion ───────────────────────────────────────────────

	/**
	 * Delete the current browser selection from the model, handling
	 * cross-paragraph selections properly (merge first + last paragraphs,
	 * remove everything in between).
	 */
	function deleteSelection(): void {
		if (!doc || !contentEl || !styleCache || !onChange) return;

		const msel = domSelectionToModel(contentEl, doc);
		if (!msel || msel.collapsed) return;

		history.snapshot(doc.body);

		if (msel.startBlock === msel.endBlock) {
			// Same paragraph — delete selected text within runs
			const block = doc.body[msel.startBlock];
			if (block && block.type === "paragraph") {
				deleteTextInParagraph(block, msel);
			}
		} else {
			// Cross-paragraph: trim first para, trim last para, remove middle, merge
			const firstBlock = doc.body[msel.startBlock];
			const lastBlock = doc.body[msel.endBlock];
			if (!firstBlock || firstBlock.type !== "paragraph") return;
			if (!lastBlock || lastBlock.type !== "paragraph") return;

			// Trim first paragraph: keep everything before selection start
			trimParagraphAfter(firstBlock, msel.startRun, msel.startOffset);

			// Trim last paragraph: keep everything after selection end
			trimParagraphBefore(lastBlock, msel.endRun, msel.endOffset);

			// Remember caret position (end of first paragraph's remaining text)
			const caretOffset = getTextLength(firstBlock);

			// Merge last paragraph's remaining content into first
			firstBlock.children.push(...lastBlock.children);

			// Remove all blocks from startBlock+1 through endBlock
			const removeCount = msel.endBlock - msel.startBlock;
			doc.body.splice(msel.startBlock + 1, removeCount);

			// Re-render and restore caret
			reRenderAll();
			setEditableState(contentEl, true);

			const mergedEl = contentEl.querySelector<HTMLElement>(
				`[data-block-idx="${msel.startBlock}"]`,
			);
			if (mergedEl) {
				placeCaretAtOffset(mergedEl, caretOffset);
			}

			onChange();
			return;
		}

		// Single-paragraph case: re-render just that paragraph
		reRenderAll();
		setEditableState(contentEl, true);

		// Place caret at deletion point
		const el = contentEl.querySelector<HTMLElement>(
			`[data-block-idx="${msel.startBlock}"]`,
		);
		if (el) {
			const offset = getOffsetUpToRun(
				doc.body[msel.startBlock] as DocxParagraph,
				msel.startRun,
			) + msel.startOffset;
			placeCaretAtOffset(el, offset);
		}

		onChange();
	}

	/** Delete selected text within a single paragraph's runs. */
	function deleteTextInParagraph(
		para: DocxParagraph,
		msel: ModelSelection,
	): void {
		const selStart = getOffsetUpToRun(para, msel.startRun) + msel.startOffset;
		const selEnd = getOffsetUpToRun(para, msel.endRun) + msel.endOffset;

		const newChildren: DocxInlineElement[] = [];
		let pos = 0;

		for (const child of para.children) {
			if (child.type !== "run") {
				newChildren.push(child);
				continue;
			}

			const runStart = pos;
			const runEnd = pos + child.text.length;
			pos = runEnd;

			if (runEnd <= selStart || runStart >= selEnd) {
				// Entirely outside selection — keep
				newChildren.push(child);
			} else {
				// Partially or fully inside selection
				const keepBefore = child.text.slice(0, Math.max(0, selStart - runStart));
				const keepAfter = child.text.slice(Math.min(child.text.length, selEnd - runStart));
				const combined = keepBefore + keepAfter;
				if (combined.length > 0) {
					newChildren.push({
						type: "run",
						text: combined,
						properties: { ...child.properties },
					});
				}
			}
		}

		para.children = newChildren.length > 0
			? newChildren
			: [{ type: "run", text: "", properties: {} }];
	}

	/** Keep only content before a given run+offset in a paragraph. */
	function trimParagraphAfter(
		para: DocxParagraph,
		runIdx: number,
		offset: number,
	): void {
		const cutPos = getOffsetUpToRun(para, runIdx) + offset;
		const newChildren: DocxInlineElement[] = [];
		let pos = 0;

		for (const child of para.children) {
			if (child.type !== "run") {
				if (pos < cutPos) newChildren.push(child);
				continue;
			}
			const runEnd = pos + child.text.length;
			if (runEnd <= cutPos) {
				newChildren.push(child);
			} else if (pos < cutPos) {
				newChildren.push({
					type: "run",
					text: child.text.slice(0, cutPos - pos),
					properties: { ...child.properties },
				});
			}
			pos = runEnd;
		}

		para.children = newChildren.length > 0
			? newChildren
			: [{ type: "run", text: "", properties: {} }];
	}

	/** Keep only content after a given run+offset in a paragraph. */
	function trimParagraphBefore(
		para: DocxParagraph,
		runIdx: number,
		offset: number,
	): void {
		const cutPos = getOffsetUpToRun(para, runIdx) + offset;
		const newChildren: DocxInlineElement[] = [];
		let pos = 0;

		for (const child of para.children) {
			if (child.type !== "run") {
				if (pos >= cutPos) newChildren.push(child);
				continue;
			}
			const runEnd = pos + child.text.length;
			if (pos >= cutPos) {
				newChildren.push(child);
			} else if (runEnd > cutPos) {
				newChildren.push({
					type: "run",
					text: child.text.slice(cutPos - pos),
					properties: { ...child.properties },
				});
			}
			pos = runEnd;
		}

		para.children = newChildren.length > 0
			? newChildren
			: [{ type: "run", text: "", properties: {} }];
	}

	// ── Image Resize ────────────────────────────────────────────────────

	function handleResizeMousedown(e: MouseEvent): void {
		const target = e.target as HTMLElement;
		if (!target.classList.contains("via-docx-image-resize-handle")) return;
		if (!doc || !onChange) return;

		e.preventDefault();
		e.stopPropagation();

		const imageRid = target.dataset.imageRid;
		if (!imageRid) return;

		const wrap = target.parentElement;
		if (!wrap || !wrap.classList.contains("via-docx-image-wrap")) return;

		const img = wrap.querySelector<HTMLImageElement>(".via-docx-image");
		if (!img) return;

		history.snapshot(doc.body);

		const startX = e.clientX;
		const startWidth = wrap.offsetWidth;
		const naturalWidth = img.naturalWidth || startWidth;
		const naturalHeight = img.naturalHeight || img.offsetHeight;
		const aspectRatio = naturalWidth > 0 ? naturalHeight / naturalWidth : 1;

		target.classList.add("via-resizing");
		wrap.classList.add("via-resizing");

		const moveHandler = (ev: MouseEvent): void => {
			const dx = ev.clientX - startX;
			const newWidth = Math.max(50, startWidth + dx);
			wrap.style.width = `${newWidth}px`;
		};

		const upHandler = (_ev: MouseEvent): void => {
			document.removeEventListener("mousemove", moveHandler);
			document.removeEventListener("mouseup", upHandler);
			target.classList.remove("via-resizing");
			wrap.classList.remove("via-resizing");

			// Commit new dimensions to the model
			const finalWidth = wrap.offsetWidth;
			const finalHeight = Math.round(finalWidth * aspectRatio);

			updateImageModelSize(imageRid, finalWidth, finalHeight);
			resizeState = null;
			onChange?.();
		};

		resizeState = {
			imageRid,
			wrap,
			img,
			startX,
			startWidth,
			aspectRatio,
			moveHandler,
			upHandler,
		};

		document.addEventListener("mousemove", moveHandler);
		document.addEventListener("mouseup", upHandler);
	}

	function updateImageModelSize(
		rid: string,
		width: number,
		height: number,
	): void {
		if (!doc) return;

		for (const block of doc.body) {
			if (block.type === "paragraph") {
				for (const child of block.children) {
					if (child.type === "image" && child.relationshipId === rid) {
						child.width = Math.round(width);
						child.height = Math.round(height);
						return;
					}
				}
			} else if (block.type === "table") {
				for (const row of block.rows) {
					for (const cell of row.cells) {
						for (const para of cell.paragraphs) {
							for (const child of para.children) {
								if (child.type === "image" && child.relationshipId === rid) {
									child.width = Math.round(width);
									child.height = Math.round(height);
									return;
								}
							}
						}
					}
				}
			}
		}
	}

	// ── Formatting Application ───────────────────────────────────────────

	function applyFormat(cmd: FormatCommand): void {
		if (!doc || !contentEl || !styleCache || !onChange) return;
		history.snapshot(doc.body);

		const msel = domSelectionToModel(contentEl, doc);
		if (!msel) return;

		if (cmd.type === "alignment") {
			applyAlignmentFormat(msel, cmd.value);
		} else if (cmd.type === "clearFormatting") {
			applyClearFormatting(msel);
		} else {
			applyRunFormat(msel, cmd);
		}

		onChange();
	}

	function applyAlignmentFormat(
		msel: ModelSelection,
		value: "left" | "center" | "right" | "both",
	): void {
		if (!doc || !contentEl || !styleCache) return;

		// Apply to all paragraphs in selection
		for (let bi = msel.startBlock; bi <= msel.endBlock; bi++) {
			const block = doc.body[bi];
			if (block && block.type === "paragraph") {
				block.properties.alignment = value;
			}
		}

		// Re-render affected paragraphs
		for (let bi = msel.startBlock; bi <= msel.endBlock; bi++) {
			const el = contentEl.querySelector<HTMLElement>(`[data-block-idx="${bi}"]`);
			if (el && styleCache) {
				const result = rerenderParagraph(doc, bi, el, styleCache);
				blobUrls.push(...result.blobUrls);
			}
		}
		setEditableState(contentEl, true);
		modelSelectionToDom(msel, contentEl);
	}

	function applyClearFormatting(msel: ModelSelection): void {
		if (!doc || !contentEl || !styleCache) return;

		for (let bi = msel.startBlock; bi <= msel.endBlock; bi++) {
			const block = doc.body[bi];
			if (!block || block.type !== "paragraph") continue;

			for (const child of block.children) {
				if (child.type === "run") {
					child.properties = {};
				}
			}
		}

		reRenderAll();
		setEditableState(contentEl, true);
		modelSelectionToDom(msel, contentEl);
	}

	function applyRunFormat(
		msel: ModelSelection,
		cmd: FormatCommand,
	): void {
		if (!doc || !contentEl || !styleCache) return;

		if (msel.collapsed) {
			// Nothing selected — no text to format
			return;
		}

		// Iterate through affected paragraphs
		for (let bi = msel.startBlock; bi <= msel.endBlock; bi++) {
			const block = doc.body[bi];
			if (!block || block.type === "paragraph") {
				const para = block as DocxParagraph;
				applyRunFormatToParagraph(para, bi, msel, cmd);
			}
		}

		reRenderAll();
		setEditableState(contentEl, true);
		modelSelectionToDom(msel, contentEl);
	}

	function applyRunFormatToParagraph(
		para: DocxParagraph,
		blockIdx: number,
		msel: ModelSelection,
		cmd: FormatCommand,
	): void {
		const newChildren: DocxInlineElement[] = [];
		let runTextOffset = 0;

		for (let ri = 0; ri < para.children.length; ri++) {
			const child = para.children[ri];
			if (!child) continue;
			if (child.type !== "run") {
				newChildren.push(child);
				continue;
			}

			const runStart = runTextOffset;
			const runEnd = runStart + child.text.length;

			// Calculate intersection with selection
			let selStart: number;
			let selEnd: number;

			if (blockIdx === msel.startBlock && blockIdx === msel.endBlock) {
				// Single paragraph selection
				selStart = getOffsetUpToRun(para, msel.startRun) + msel.startOffset;
				selEnd = getOffsetUpToRun(para, msel.endRun) + msel.endOffset;
			} else if (blockIdx === msel.startBlock) {
				selStart = getOffsetUpToRun(para, msel.startRun) + msel.startOffset;
				selEnd = getTotalTextLength(para);
			} else if (blockIdx === msel.endBlock) {
				selStart = 0;
				selEnd = getOffsetUpToRun(para, msel.endRun) + msel.endOffset;
			} else {
				// Fully selected paragraph
				selStart = 0;
				selEnd = getTotalTextLength(para);
			}

			// No intersection with this run
			if (selEnd <= runStart || selStart >= runEnd) {
				newChildren.push(child);
				runTextOffset = runEnd;
				continue;
			}

			// Split into before / selected / after
			const splitStart = Math.max(selStart - runStart, 0);
			const splitEnd = Math.min(selEnd - runStart, child.text.length);

			// Before portion
			if (splitStart > 0) {
				newChildren.push({
					type: "run",
					text: child.text.slice(0, splitStart),
					properties: { ...child.properties },
				});
			}

			// Selected portion — apply format
			const selectedRun: DocxRun = {
				type: "run",
				text: child.text.slice(splitStart, splitEnd),
				properties: { ...child.properties },
			};
			applyFormatToRunProps(selectedRun.properties, cmd);
			newChildren.push(selectedRun);

			// After portion
			if (splitEnd < child.text.length) {
				newChildren.push({
					type: "run",
					text: child.text.slice(splitEnd),
					properties: { ...child.properties },
				});
			}

			runTextOffset = runEnd;
		}

		para.children = newChildren;
	}

	function applyFormatToRunProps(
		props: Partial<DocxRunProperties>,
		cmd: FormatCommand,
	): void {
		switch (cmd.type) {
			case "bold":
				props.bold = !props.bold;
				break;
			case "italic":
				props.italic = !props.italic;
				break;
			case "underline":
				props.underline = !props.underline;
				break;
			case "strikethrough":
				props.strikethrough = !props.strikethrough;
				break;
			case "fontSize":
				props.fontSize = cmd.value;
				break;
			case "fontFamily":
				props.fontFamily = cmd.value;
				break;
			case "color":
				props.color = cmd.value;
				break;
			case "highlight":
				props.highlight = cmd.value || undefined;
				break;
		}
	}

	function getOffsetUpToRun(para: DocxParagraph, runIdx: number): number {
		let offset = 0;
		for (let i = 0; i < runIdx && i < para.children.length; i++) {
			const child = para.children[i];
			if (child && child.type === "run") {
				offset += child.text.length;
			} else if (child && child.type === "tab") {
				offset += 1;
			}
		}
		return offset;
	}

	function getTotalTextLength(para: DocxParagraph): number {
		let len = 0;
		for (const child of para.children) {
			if (child.type === "run") len += child.text.length;
			else if (child.type === "tab") len += 1;
		}
		return len;
	}

	// ── DOM → Model Sync ─────────────────────────────────────────────────

	function syncDomToModel(paraEl: HTMLElement, para: DocxParagraph): void {
		// Walk text nodes in the paragraph and rebuild runs
		const newChildren: DocxInlineElement[] = [];

		// Collect text from run spans (data-run-idx) and other nodes
		const walker = document.createTreeWalker(
			paraEl,
			NodeFilter.SHOW_TEXT,
			{
				acceptNode: (node: Node) => {
					// Skip the numbering prefix span
					const parent = node.parentElement;
					if (parent?.classList.contains("via-docx-num-prefix")) {
						return NodeFilter.FILTER_REJECT;
					}
					return NodeFilter.FILTER_ACCEPT;
				},
			},
		);

		let currentRunIdx = -1;
		let currentText = "";
		let currentProps: Partial<DocxRunProperties> = {};

		let textNode = walker.nextNode();
		while (textNode) {
			const text = textNode.textContent ?? "";
			if (!text) {
				textNode = walker.nextNode();
				continue;
			}

			// Find the nearest run span with data-run-idx
			const runSpan = findRunSpan(textNode as Text);
			const runIdx = runSpan ? parseInt(runSpan.dataset.runIdx ?? "", 10) : -1;

			if (runIdx !== currentRunIdx && currentText) {
				// Flush previous run
				newChildren.push({ type: "run", text: currentText, properties: currentProps });
				currentText = "";
			}

			if (!isNaN(runIdx) && runIdx >= 0 && runIdx < para.children.length) {
				const origChild = para.children[runIdx];
				if (origChild && origChild.type === "run") {
					currentProps = origChild.properties;
				}
			}

			currentRunIdx = runIdx;
			currentText += text;
			textNode = walker.nextNode();
		}

		// Flush last run
		if (currentText) {
			newChildren.push({ type: "run", text: currentText, properties: currentProps });
		}

		// Preserve non-run inline children (images, breaks, tabs, hyperlinks)
		for (const child of para.children) {
			if (child.type === "image" || child.type === "break" || child.type === "tab" || child.type === "hyperlink") {
				newChildren.push(child);
			}
		}

		// Always update — if empty, use a single empty run so model stays in sync
		if (newChildren.length === 0) {
			newChildren.push({ type: "run", text: "", properties: {} });
		}
		para.children = newChildren;
	}

	// ── Re-render All ────────────────────────────────────────────────────

	function reRenderAll(): void {
		if (!doc || !contentEl || !styleCache) return;

		contentEl.empty();
		const newBlobUrls = renderDocument(doc, contentEl);
		blobUrls.push(...newBlobUrls);
	}

	// ── Helpers ──────────────────────────────────────────────────────────

	function findParagraphEl(node: HTMLElement | Node): HTMLElement | null {
		let current: HTMLElement | null = node instanceof HTMLElement ? node : node.parentElement;
		while (current && current !== contentEl) {
			if (current.dataset.blockIdx !== undefined) {
				const tag = current.tagName.toLowerCase();
				if (tag === "p" || /^h[1-6]$/.test(tag)) {
					return current;
				}
			}
			current = current.parentElement;
		}
		return null;
	}

	function findRunSpan(textNode: Text): HTMLElement | null {
		let current: HTMLElement | null = textNode.parentElement;
		while (current && current !== contentEl) {
			if (current.dataset.runIdx !== undefined) {
				return current;
			}
			current = current.parentElement;
		}
		return null;
	}

	function getCaretPosition(paraEl: HTMLElement): number {
		const sel = window.getSelection();
		if (!sel || !sel.rangeCount) return 0;

		const range = sel.getRangeAt(0);

		// Create a range from start of paragraph to caret
		const preRange = document.createRange();
		preRange.selectNodeContents(paraEl);
		preRange.setEnd(range.startContainer, range.startOffset);

		return preRange.toString().length;
	}

	function getTextLength(para: DocxParagraph): number {
		let len = 0;
		for (const child of para.children) {
			if (child.type === "run") len += child.text.length;
			else if (child.type === "tab") len += 1;
		}
		return len;
	}

	function splitParagraphChildren(
		para: DocxParagraph,
		offset: number,
	): { before: DocxInlineElement[]; after: DocxInlineElement[] } {
		const before: DocxInlineElement[] = [];
		const after: DocxInlineElement[] = [];
		let pos = 0;
		let split = false;

		for (const child of para.children) {
			if (split) {
				after.push(child);
				continue;
			}

			if (child.type === "run") {
				const end = pos + child.text.length;
				if (offset <= pos) {
					split = true;
					after.push(child);
				} else if (offset >= end) {
					before.push(child);
				} else {
					// Split this run
					const splitAt = offset - pos;
					before.push({
						type: "run",
						text: child.text.slice(0, splitAt),
						properties: { ...child.properties },
					});
					const afterText = child.text.slice(splitAt);
					if (afterText) {
						after.push({
							type: "run",
							text: afterText,
							properties: { ...child.properties },
						});
					}
					split = true;
				}
				pos = end;
			} else if (child.type === "tab") {
				if (offset <= pos) {
					split = true;
					after.push(child);
				} else {
					before.push(child);
				}
				pos += 1;
			} else {
				if (offset <= pos) {
					split = true;
					after.push(child);
				} else {
					before.push(child);
				}
			}
		}

		// Ensure at least one empty run in each side for editability
		if (before.length === 0) {
			before.push({ type: "run", text: "", properties: {} });
		}
		if (after.length === 0) {
			after.push({ type: "run", text: "", properties: {} });
		}

		return { before, after };
	}

	function placeCaretAtStart(el: HTMLElement): void {
		const sel = window.getSelection();
		if (!sel) return;
		const range = document.createRange();

		// Find first text node
		const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
		const firstText = walker.nextNode();
		if (firstText) {
			range.setStart(firstText, 0);
			range.collapse(true);
		} else {
			range.selectNodeContents(el);
			range.collapse(true);
		}
		sel.removeAllRanges();
		sel.addRange(range);
	}

	function placeCaretAtOffset(el: HTMLElement, offset: number): void {
		const sel = window.getSelection();
		if (!sel) return;

		const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
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
				const range = document.createRange();
				range.setStart(textNode, remaining);
				range.collapse(true);
				sel.removeAllRanges();
				sel.addRange(range);
				return;
			}
			remaining -= len;
			textNode = walker.nextNode();
		}

		// Fallback: place at end
		const range = document.createRange();
		range.selectNodeContents(el);
		range.collapse(false);
		sel.removeAllRanges();
		sel.addRange(range);
	}

	// ── Undo / Redo ─────────────────────────────────────────────────────

	function undo(): boolean {
		if (!doc || !contentEl || !styleCache || !onChange) return false;

		const prevBody = history.undo(doc.body);
		if (!prevBody) return false;

		doc.body = prevBody;
		reRenderAll();
		setEditableState(contentEl, true);
		onChange();
		return true;
	}

	function redo(): boolean {
		if (!doc || !contentEl || !styleCache || !onChange) return false;

		// Save current state before redo so undo can get back here
		history.snapshot(doc.body);

		const nextBody = history.redo();
		if (!nextBody) return false;

		doc.body = nextBody;
		reRenderAll();
		setEditableState(contentEl, true);
		onChange();
		return true;
	}

	return { enable, disable, isEnabled, applyFormat, undo, redo };
}
