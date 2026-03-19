/**
 * ViewItAll — OOXML Styles Parser
 *
 * Parses `word/styles.xml` and resolves style inheritance chains
 * to produce fully-resolved DocxStyle objects.
 */

import type {
	DocxStyle,
	DocxParagraphProperties,
	DocxRunProperties,
} from "./model";
import { defaultParagraphProperties, defaultRunProperties } from "./model";
import {
	NS_W,
	getElements,
	getDirectChild,
	getVal,
	getWAttr,
	parseXml,
} from "../utils/xml";
import { safeParseInt } from "../utils/units";

/**
 * Parse styles.xml content into a map of style ID → DocxStyle.
 */
export function parseStyles(stylesXml: string): Map<string, DocxStyle> {
	const doc = parseXml(stylesXml);
	const styleMap = new Map<string, DocxStyle>();

	const styleElements = getElements(doc, NS_W, "style");
	for (const el of styleElements) {
		const id = getWAttr(el, "styleId");
		if (!id) continue;

		const typeAttr = getWAttr(el, "type") ?? "paragraph";
		const nameEl = getDirectChild(el, NS_W, "name");
		const name = nameEl ? (getVal(nameEl) ?? id) : id;
		const basedOnEl = getDirectChild(el, NS_W, "basedOn");
		const basedOn = basedOnEl ? (getVal(basedOnEl) ?? undefined) : undefined;

		const pPr = getDirectChild(el, NS_W, "pPr");
		const rPr = getDirectChild(el, NS_W, "rPr");

		styleMap.set(id, {
			id,
			name,
			type: typeAttr as DocxStyle["type"],
			basedOn,
			paragraphProperties: pPr
				? parseParagraphProperties(pPr)
				: {},
			runProperties: rPr ? parseRunProperties(rPr) : {},
		});
	}

	return styleMap;
}

/**
 * Resolve a style's full properties by walking the basedOn chain.
 * Caches resolved results to avoid repeated walks.
 */
export function resolveStyle(
	styleId: string,
	styles: Map<string, DocxStyle>,
	cache: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>,
): { pProps: DocxParagraphProperties; rProps: DocxRunProperties } {
	const cached = cache.get(styleId);
	if (cached) return cached;

	const style = styles.get(styleId);
	if (!style) {
		const fallback = {
			pProps: defaultParagraphProperties(),
			rProps: defaultRunProperties(),
		};
		cache.set(styleId, fallback);
		return fallback;
	}

	let pProps: DocxParagraphProperties;
	let rProps: DocxRunProperties;

	if (style.basedOn && style.basedOn !== styleId) {
		const parent = resolveStyle(style.basedOn, styles, cache);
		pProps = { ...parent.pProps };
		rProps = { ...parent.rProps };
	} else {
		pProps = defaultParagraphProperties();
		rProps = defaultRunProperties();
	}

	mergeParagraphProperties(pProps, style.paragraphProperties);
	mergeRunProperties(rProps, style.runProperties);

	const result = { pProps, rProps };
	cache.set(styleId, result);
	return result;
}

/**
 * Determine heading level from a style name or outlineLvl.
 */
export function resolveHeadingLevel(
	style: DocxStyle,
): number | undefined {
	const nameMatch = /^heading\s+(\d)$/i.exec(style.name);
	if (nameMatch && nameMatch[1]) {
		const level = parseInt(nameMatch[1], 10);
		if (level >= 1 && level <= 6) return level;
	}

	const headingLevel = style.paragraphProperties.headingLevel;
	if (headingLevel !== undefined && headingLevel >= 0 && headingLevel <= 5) {
		return headingLevel + 1;
	}

	return undefined;
}

// ── Property Parsers ────────────────────────────────────────────────────────

/**
 * Parse paragraph properties from a w:pPr element.
 */
export function parseParagraphProperties(
	pPr: Element,
): Partial<DocxParagraphProperties> {
	const props: Partial<DocxParagraphProperties> = {};

	const jcEl = getDirectChild(pPr, NS_W, "jc");
	if (jcEl) {
		const val = getVal(jcEl);
		if (val === "left" || val === "center" || val === "right" || val === "both") {
			props.alignment = val;
		}
	}

	const outlineLvlEl = getDirectChild(pPr, NS_W, "outlineLvl");
	if (outlineLvlEl) {
		const val = safeParseInt(getVal(outlineLvlEl));
		if (val !== undefined && val >= 0 && val <= 5) {
			props.headingLevel = val + 1;
		}
	}

	// Numbering from pPr (used by styles and direct paragraph formatting)
	const numPrEl = getDirectChild(pPr, NS_W, "numPr");
	if (numPrEl) {
		const numIdEl = getDirectChild(numPrEl, NS_W, "numId");
		const ilvlEl = getDirectChild(numPrEl, NS_W, "ilvl");
		const numId = numIdEl ? (getVal(numIdEl) ?? undefined) : undefined;
		const ilvl = ilvlEl ? safeParseInt(getVal(ilvlEl)) : undefined;
		if (numId) {
			props.numberingId = numId;
			props.numberingLevel = ilvl ?? 0;
		}
	}

	const indEl = getDirectChild(pPr, NS_W, "ind");
	if (indEl) {
		props.indentation = {
			left: safeParseInt(getWAttr(indEl, "left")) ?? 0,
			right: safeParseInt(getWAttr(indEl, "right")) ?? 0,
			firstLine: safeParseInt(getWAttr(indEl, "firstLine")) ?? 0,
			hanging: safeParseInt(getWAttr(indEl, "hanging")) ?? 0,
		};
	}

	const spacingEl = getDirectChild(pPr, NS_W, "spacing");
	if (spacingEl) {
		props.spacing = {
			before: safeParseInt(getWAttr(spacingEl, "before")) ?? 0,
			after: safeParseInt(getWAttr(spacingEl, "after")) ?? 0,
			line: safeParseInt(getWAttr(spacingEl, "line")) ?? 0,
			lineRule: getWAttr(spacingEl, "lineRule") ?? "auto",
		};
	}

	return props;
}

/**
 * Parse run properties from a w:rPr element.
 */
export function parseRunProperties(
	rPr: Element,
): Partial<DocxRunProperties> {
	const props: Partial<DocxRunProperties> = {};

	const bEl = getDirectChild(rPr, NS_W, "b");
	if (bEl) {
		const val = getVal(bEl);
		props.bold = val !== "0" && val !== "false";
	}

	const iEl = getDirectChild(rPr, NS_W, "i");
	if (iEl) {
		const val = getVal(iEl);
		props.italic = val !== "0" && val !== "false";
	}

	const uEl = getDirectChild(rPr, NS_W, "u");
	if (uEl) {
		const val = getVal(uEl);
		props.underline = val !== "none" && val !== undefined;
	}

	const strikeEl = getDirectChild(rPr, NS_W, "strike");
	if (strikeEl) {
		const val = getVal(strikeEl);
		props.strikethrough = val !== "0" && val !== "false";
	}

	const rFontsEl = getDirectChild(rPr, NS_W, "rFonts");
	if (rFontsEl) {
		props.fontFamily =
			getWAttr(rFontsEl, "ascii") ??
			getWAttr(rFontsEl, "hAnsi") ??
			getWAttr(rFontsEl, "cs") ??
			undefined;
	}

	const szEl = getDirectChild(rPr, NS_W, "sz");
	if (szEl) {
		props.fontSize = safeParseInt(getVal(szEl));
	}

	const colorEl = getDirectChild(rPr, NS_W, "color");
	if (colorEl) {
		const val = getVal(colorEl);
		if (val === "auto") {
			// "auto" means reset to default (black) — override inherited colors
			props.color = "auto";
		} else if (val) {
			props.color = val;
		}
	}

	const highlightEl = getDirectChild(rPr, NS_W, "highlight");
	if (highlightEl) {
		props.highlight = getVal(highlightEl) ?? undefined;
	}

	const vertAlignEl = getDirectChild(rPr, NS_W, "vertAlign");
	if (vertAlignEl) {
		const val = getVal(vertAlignEl);
		if (val === "superscript" || val === "subscript") {
			props.vertAlign = val;
		}
	}

	return props;
}

// ── Merge Helpers ───────────────────────────────────────────────────────────

function mergeParagraphProperties(
	target: DocxParagraphProperties,
	source: Partial<DocxParagraphProperties>,
): void {
	if (source.alignment !== undefined) target.alignment = source.alignment;
	if (source.headingLevel !== undefined) target.headingLevel = source.headingLevel;
	if (source.indentation !== undefined) target.indentation = source.indentation;
	if (source.spacing !== undefined) target.spacing = source.spacing;
	if (source.numberingId !== undefined) target.numberingId = source.numberingId;
	if (source.numberingLevel !== undefined) target.numberingLevel = source.numberingLevel;
}

function mergeRunProperties(
	target: DocxRunProperties,
	source: Partial<DocxRunProperties>,
): void {
	if (source.bold !== undefined) target.bold = source.bold;
	if (source.italic !== undefined) target.italic = source.italic;
	if (source.underline !== undefined) target.underline = source.underline;
	if (source.strikethrough !== undefined) target.strikethrough = source.strikethrough;
	if (source.fontFamily !== undefined) target.fontFamily = source.fontFamily;
	if (source.fontSize !== undefined) target.fontSize = source.fontSize;
	if (source.color !== undefined) {
		// "auto" resets color to default, overriding any inherited color
		target.color = source.color === "auto" ? undefined : source.color;
	}
	if (source.highlight !== undefined) target.highlight = source.highlight;
	if (source.vertAlign !== undefined) target.vertAlign = source.vertAlign;
}
