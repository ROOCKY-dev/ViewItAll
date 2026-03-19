/**
 * ViewItAll — XML Utility Helpers
 *
 * OOXML namespace constants and typed DOM helpers.
 * Pure functions — no Obsidian state.
 */

/** Word processing main namespace (w:) */
export const NS_W =
	"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** Relationships namespace (r:) */
export const NS_R =
	"http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/** DrawingML main namespace (a:) */
export const NS_A =
	"http://schemas.openxmlformats.org/drawingml/2006/main";

/** Word Drawing namespace (wp:) */
export const NS_WP =
	"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

/** Picture namespace (pic:) */
export const NS_PIC =
	"http://schemas.openxmlformats.org/drawingml/2006/picture";

/** Package relationships namespace */
export const NS_RELS =
	"http://schemas.openxmlformats.org/package/2006/relationships";

/**
 * Get all child elements matching a namespace + local name.
 * Returns a plain array (not a live NodeList).
 */
export function getElements(
	parent: Element | Document,
	ns: string,
	localName: string,
): Element[] {
	return Array.from(parent.getElementsByTagNameNS(ns, localName));
}

/**
 * Get the first child element matching a namespace + local name, or null.
 */
export function getElement(
	parent: Element | Document,
	ns: string,
	localName: string,
): Element | null {
	return parent.getElementsByTagNameNS(ns, localName).item(0);
}

/**
 * Get an attribute value from an element in a specific namespace.
 * Falls back to checking without namespace (some parsers strip ns prefixes).
 */
export function getAttr(
	el: Element,
	ns: string,
	localName: string,
): string | null {
	return el.getAttributeNS(ns, localName) ?? el.getAttribute(localName);
}

/**
 * Get an attribute value from the w: namespace.
 */
export function getWAttr(el: Element, localName: string): string | null {
	const val = el.getAttributeNS(NS_W, localName);
	if (val) return val;
	return el.getAttribute(`w:${localName}`) ?? el.getAttribute(localName);
}

/**
 * Get the `w:val` attribute from an element (very common in OOXML).
 */
export function getVal(el: Element): string | null {
	return getWAttr(el, "val");
}

/**
 * Get direct child elements (not descendants) matching namespace + local name.
 */
export function getDirectChildren(
	parent: Element,
	ns: string,
	localName: string,
): Element[] {
	const results: Element[] = [];
	for (let i = 0; i < parent.childNodes.length; i++) {
		const node = parent.childNodes.item(i);
		if (
			node &&
			node.nodeType === Node.ELEMENT_NODE &&
			(node as Element).localName === localName &&
			(node as Element).namespaceURI === ns
		) {
			results.push(node as Element);
		}
	}
	return results;
}

/**
 * Get the first direct child element matching namespace + local name.
 */
export function getDirectChild(
	parent: Element,
	ns: string,
	localName: string,
): Element | null {
	for (let i = 0; i < parent.childNodes.length; i++) {
		const node = parent.childNodes.item(i);
		if (
			node &&
			node.nodeType === Node.ELEMENT_NODE &&
			(node as Element).localName === localName &&
			(node as Element).namespaceURI === ns
		) {
			return node as Element;
		}
	}
	return null;
}

/**
 * Parse an XML string into a Document using the browser's DOMParser.
 */
export function parseXml(xmlString: string): Document {
	const parser = new DOMParser();
	return parser.parseFromString(xmlString, "application/xml");
}
