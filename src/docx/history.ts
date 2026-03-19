/**
 * ViewItAll — DOCX Edit History (Undo/Redo)
 *
 * Snapshot-based undo/redo using deep-cloned model states.
 * Each snapshot stores the full body array, enabling reliable
 * undo without complex inverse operations.
 */

import type { DocxBlockElement } from "./model";

// ── Public Interface ─────────────────────────────────────────────────────────

export interface EditHistory {
	/** Take a snapshot of the current state (call before a mutation). */
	snapshot(body: DocxBlockElement[]): void;
	/** Undo: restore the previous snapshot. Returns the body to apply, or null. */
	undo(currentBody: DocxBlockElement[]): DocxBlockElement[] | null;
	/** Redo: re-apply the next snapshot. Returns the body to apply, or null. */
	redo(): DocxBlockElement[] | null;
	canUndo(): boolean;
	canRedo(): boolean;
	clear(): void;
}

const MAX_HISTORY = 100;

export function createEditHistory(): EditHistory {
	/** Past states (most recent at end). */
	const undoStack: string[] = [];
	/** Future states (most recent at end). */
	const redoStack: string[] = [];

	function snapshot(body: DocxBlockElement[]): void {
		const serialized = JSON.stringify(body);
		// Don't push duplicate states
		if (undoStack.length > 0 && undoStack[undoStack.length - 1] === serialized) {
			return;
		}
		undoStack.push(serialized);
		if (undoStack.length > MAX_HISTORY) {
			undoStack.shift();
		}
		// New mutation invalidates redo stack
		redoStack.length = 0;
	}

	function undo(currentBody: DocxBlockElement[]): DocxBlockElement[] | null {
		if (undoStack.length === 0) return null;

		// Save current state for redo
		redoStack.push(JSON.stringify(currentBody));

		const prev = undoStack.pop();
		if (!prev) return null;

		return JSON.parse(prev) as DocxBlockElement[];
	}

	function redo(): DocxBlockElement[] | null {
		if (redoStack.length === 0) return null;

		const next = redoStack.pop();
		if (!next) return null;

		// Save current state to undo stack (will be pushed by caller via snapshot)
		// Actually, we need to save the current state before redo
		const parsed = JSON.parse(next) as DocxBlockElement[];

		return parsed;
	}

	function canUndo(): boolean {
		return undoStack.length > 0;
	}

	function canRedo(): boolean {
		return redoStack.length > 0;
	}

	function clear(): void {
		undoStack.length = 0;
		redoStack.length = 0;
	}

	return { snapshot, undo, redo, canUndo, canRedo, clear };
}
