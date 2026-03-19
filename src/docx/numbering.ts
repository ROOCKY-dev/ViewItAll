/**
 * ViewItAll — OOXML Numbering Parser
 *
 * Parses `word/numbering.xml` to resolve list definitions.
 */

import type { DocxNumberingDef, DocxNumberingLevel } from "./model";
import {
	NS_W,
	getElements,
	getDirectChild,
	getDirectChildren,
	getVal,
	getWAttr,
	parseXml,
} from "../utils/xml";
import { safeParseInt } from "../utils/units";

/**
 * Parse numbering.xml content into a map of numId → DocxNumberingDef.
 */
export function parseNumbering(
	numberingXml: string,
): Map<string, DocxNumberingDef> {
	const doc = parseXml(numberingXml);
	const result = new Map<string, DocxNumberingDef>();

	const abstractMap = new Map<string, Map<number, DocxNumberingLevel>>();
	const abstractElements = getElements(doc, NS_W, "abstractNum");

	for (const absEl of abstractElements) {
		const absId = getWAttr(absEl, "abstractNumId");
		if (!absId) continue;

		const levels = new Map<number, DocxNumberingLevel>();
		const lvlElements = getDirectChildren(absEl, NS_W, "lvl");

		for (const lvlEl of lvlElements) {
			const ilvl = safeParseInt(getWAttr(lvlEl, "ilvl"));
			if (ilvl === undefined) continue;

			const numFmtEl = getDirectChild(lvlEl, NS_W, "numFmt");
			const lvlTextEl = getDirectChild(lvlEl, NS_W, "lvlText");
			const pPrEl = getDirectChild(lvlEl, NS_W, "pPr");

			const format = parseNumFormat(numFmtEl ? getVal(numFmtEl) : null);
			const text = lvlTextEl ? (getVal(lvlTextEl) ?? "") : "";

			let indentLeft = 0;
			if (pPrEl) {
				const indEl = getDirectChild(pPrEl, NS_W, "ind");
				if (indEl) {
					indentLeft = safeParseInt(getWAttr(indEl, "left")) ?? 0;
				}
			}

			levels.set(ilvl, { format, text, indentLeft });
		}

		abstractMap.set(absId, levels);
	}

	const numElements = getElements(doc, NS_W, "num");
	for (const numEl of numElements) {
		const numId = getWAttr(numEl, "numId");
		if (!numId) continue;

		const abstractNumIdEl = getDirectChild(numEl, NS_W, "abstractNumId");
		const abstractNumId = abstractNumIdEl
			? (getVal(abstractNumIdEl) ?? "")
			: "";

		const levels = abstractMap.get(abstractNumId) ?? new Map();

		result.set(numId, {
			abstractNumId: abstractNumId,
			levels,
		});
	}

	return result;
}

function parseNumFormat(
	val: string | null,
): DocxNumberingLevel["format"] {
	switch (val) {
		case "decimal":
			return "decimal";
		case "bullet":
			return "bullet";
		case "lowerLetter":
			return "lowerLetter";
		case "upperLetter":
			return "upperLetter";
		case "lowerRoman":
			return "lowerRoman";
		case "upperRoman":
			return "upperRoman";
		case "none":
			return "none";
		default:
			return "bullet";
	}
}
