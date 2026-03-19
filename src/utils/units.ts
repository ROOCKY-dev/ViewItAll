/**
 * ViewItAll — Unit Conversion Utilities
 *
 * OOXML uses several unit systems. These helpers convert to
 * CSS-friendly values (px, pt, em).
 * Pure functions — no Obsidian state.
 */

/** 1 inch = 914400 EMU (English Metric Units) */
const EMU_PER_INCH = 914400;

/** Standard CSS pixels per inch */
const PX_PER_INCH = 96;

/** Convert EMU to CSS pixels. */
export function emuToPx(emu: number): number {
	return Math.round((emu / EMU_PER_INCH) * PX_PER_INCH);
}

/**
 * Convert half-points to CSS points.
 * OOXML `w:sz` stores font size in half-points (e.g. 24 = 12pt).
 */
export function halfPointsToPt(halfPts: number): number {
	return halfPts / 2;
}

/**
 * Convert twips to CSS pixels.
 * 1 inch = 1440 twips. Used for indentation, spacing, table widths.
 */
export function twipsToPx(twips: number): number {
	return Math.round((twips / 1440) * PX_PER_INCH);
}

/**
 * Convert twentieths of a point to CSS points.
 * Same as twips but expressed differently in OOXML docs.
 */
export function twipsToPt(twips: number): number {
	return twips / 20;
}

/**
 * Convert DXA (twentieths of a point) to CSS pixels.
 * DXA is used for table cell widths in OOXML.
 */
export function dxaToPx(dxa: number): number {
	return twipsToPx(dxa);
}

/**
 * Parse a numeric string safely, returning undefined if not a valid number.
 */
export function safeParseInt(value: string | null | undefined): number | undefined {
	if (value == null) return undefined;
	const n = parseInt(value, 10);
	return isNaN(n) ? undefined : n;
}

/**
 * Parse a numeric float string safely.
 */
export function safeParseFloat(value: string | null | undefined): number | undefined {
	if (value == null) return undefined;
	const n = parseFloat(value);
	return isNaN(n) ? undefined : n;
}
