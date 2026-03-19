/**
 * ViewItAll — OOXML Relationships Parser
 *
 * Parses `word/_rels/document.xml.rels` to resolve hyperlink URLs
 * and image file paths referenced by relationship IDs.
 */

import { NS_RELS, parseXml } from "../utils/xml";

export interface Relationship {
	type: string;
	target: string;
	targetMode: string | undefined;
}

/** Relationship type URIs */
export const REL_HYPERLINK =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
export const REL_IMAGE =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
export const REL_NUMBERING =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
export const REL_STYLES =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

/**
 * Parse a .rels XML string into a map of relationship ID → Relationship.
 */
export function parseRelationships(
	relsXml: string,
): Map<string, Relationship> {
	const doc = parseXml(relsXml);
	const map = new Map<string, Relationship>();

	const elements = doc.getElementsByTagNameNS(NS_RELS, "Relationship");
	const fallback =
		elements.length === 0
			? doc.getElementsByTagName("Relationship")
			: elements;

	for (let i = 0; i < fallback.length; i++) {
		const el = fallback[i];
		if (!el) continue;

		const id = el.getAttribute("Id");
		const type = el.getAttribute("Type");
		const target = el.getAttribute("Target");

		if (id && type && target) {
			map.set(id, {
				type,
				target,
				targetMode: el.getAttribute("TargetMode") ?? undefined,
			});
		}
	}

	return map;
}
