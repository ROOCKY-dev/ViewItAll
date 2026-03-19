/**
 * ViewItAll — DOCX Document Renderer
 *
 * Converts a DocxDocument model into native Obsidian DOM using createEl().
 * Never uses innerHTML — all elements are built programmatically.
 */

import type {
	DocxDocument,
	DocxBlockElement,
	DocxParagraph,
	DocxInlineElement,
	DocxRun,
	DocxHyperlink,
	DocxImage,
	DocxTable,
	DocxTableRow,
	DocxTableCell,
	DocxRunProperties,
	DocxParagraphProperties,
} from "./model";
import { defaultRunProperties, defaultParagraphProperties } from "./model";
import { resolveStyle, resolveHeadingLevel } from "./styles";
import { halfPointsToPt, twipsToPx, dxaToPx } from "../utils/units";

// ── Word highlight color → CSS color map ────────────────────────────────────

const HIGHLIGHT_COLORS: Record<string, string> = {
	yellow: "rgba(255, 255, 0, 0.4)",
	green: "rgba(0, 255, 0, 0.4)",
	cyan: "rgba(0, 255, 255, 0.4)",
	magenta: "rgba(255, 0, 255, 0.4)",
	blue: "rgba(0, 0, 255, 0.4)",
	red: "rgba(255, 0, 0, 0.4)",
	darkBlue: "rgba(0, 0, 139, 0.4)",
	darkCyan: "rgba(0, 139, 139, 0.4)",
	darkGreen: "rgba(0, 100, 0, 0.4)",
	darkMagenta: "rgba(139, 0, 139, 0.4)",
	darkRed: "rgba(139, 0, 0, 0.4)",
	darkYellow: "rgba(139, 139, 0, 0.4)",
	darkGray: "rgba(169, 169, 169, 0.4)",
	lightGray: "rgba(211, 211, 211, 0.4)",
	black: "rgba(0, 0, 0, 0.4)",
};

/** Render context passed through the rendering pipeline. */
interface RenderContext {
	doc: DocxDocument;
	blobUrls: string[];
	styleCache: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>;
	/**
	 * Numbering counters: key = "numId:level", value = current count.
	 * Reset when a different numId is encountered at the same level.
	 */
	numberingCounters: Map<string, number>;
	/** Track last numId seen at each level to detect resets */
	lastNumIdAtLevel: Map<number, string>;
}

/**
 * Render a DocxDocument into a DOM container element.
 * Returns an array of blob URLs that must be revoked on cleanup.
 */
export function renderDocument(
	doc: DocxDocument,
	container: HTMLElement,
): string[] {
	const ctx: RenderContext = {
		doc,
		blobUrls: [],
		styleCache: new Map(),
		numberingCounters: new Map(),
		lastNumIdAtLevel: new Map(),
	};

	renderBlocks(doc.body, container, ctx);

	return ctx.blobUrls;
}

// ── Block Rendering ─────────────────────────────────────────────────────────

function renderBlocks(
	blocks: DocxBlockElement[],
	container: HTMLElement,
	ctx: RenderContext,
): void {
	for (const block of blocks) {
		if (!block) continue;

		if (block.type === "paragraph") {
			renderParagraph(block, container, ctx);
		} else if (block.type === "table") {
			renderTable(block, container, ctx);
		} else if (block.type === "pageBreak") {
			container.createEl("hr", { cls: "via-docx-page-break" });
		}
	}
}

// ── Paragraph Rendering ─────────────────────────────────────────────────────

function renderParagraph(
	para: DocxParagraph,
	container: HTMLElement,
	ctx: RenderContext,
): void {
	// Resolve heading level from style
	let headingLevel = para.properties.headingLevel;
	if (!headingLevel && para.styleId) {
		const style = ctx.doc.styles.get(para.styleId);
		if (style) {
			headingLevel = resolveHeadingLevel(style);
		}
	}

	// Also resolve full style properties
	let resolvedRProps = defaultRunProperties();
	let resolvedPProps = defaultParagraphProperties();
	if (para.styleId) {
		const resolved = resolveStyle(para.styleId, ctx.doc.styles, ctx.styleCache);
		resolvedRProps = resolved.rProps;
		resolvedPProps = resolved.pProps;
		// Merge paragraph-level heading from style
		if (!headingLevel && resolved.pProps.headingLevel) {
			headingLevel = resolved.pProps.headingLevel;
		}
		// Merge alignment from style if not set on paragraph
		if (!para.properties.alignment && resolved.pProps.alignment) {
			para.properties.alignment = resolved.pProps.alignment;
		}
	}

	// Resolve numbering: direct paragraph > style-inherited
	let numId = para.numberingId;
	let numLevel = para.numberingLevel;
	if (!numId && resolvedPProps.numberingId) {
		numId = resolvedPProps.numberingId;
		numLevel = resolvedPProps.numberingLevel ?? 0;
	}

	// Generate numbering text prefix if applicable
	let numberingPrefix = "";
	if (numId && numId !== "0") {
		numberingPrefix = getNumberingText(numId, numLevel, ctx);
	}

	// Create the element
	let el: HTMLElement;
	if (headingLevel && headingLevel >= 1 && headingLevel <= 6) {
		const tag = `h${headingLevel}` as keyof HTMLElementTagNameMap;
		el = container.createEl(tag, { cls: "via-docx-heading" });
		// Inline style beats any Obsidian theme CSS that colors headings
		el.style.color = "inherit";
	} else {
		el = container.createEl("p");
	}

	// Apply paragraph styles
	applyParagraphStyles(el, para.properties);

	// Check if paragraph is empty (just whitespace/empty runs)
	const hasContent = para.children.some((child) => {
		if (child.type === "run") return child.text.trim().length > 0;
		if (child.type === "hyperlink") return child.children.length > 0;
		if (child.type === "image") return true;
		if (child.type === "break") return true;
		return false;
	});

	if (!hasContent && para.children.length === 0) {
		// Empty paragraph — add a zero-width space to preserve spacing
		el.createEl("br");
		return;
	}

	// Prepend numbering text if present
	if (numberingPrefix) {
		const numSpan = el.createEl("span", {
			cls: "via-docx-num-prefix",
		});
		numSpan.textContent = numberingPrefix;
		numSpan.style.color = "inherit";
	}

	// Render inline children
	for (const child of para.children) {
		renderInline(child, el, ctx, resolvedRProps);
	}
}

function applyParagraphStyles(
	el: HTMLElement,
	props: DocxParagraphProperties,
): void {
	if (props.alignment) {
		const align = props.alignment === "both" ? "justify" : props.alignment;
		el.style.textAlign = align;
	}

	if (props.indentation) {
		const left = props.indentation.left
			? twipsToPx(props.indentation.left)
			: 0;
		const right = props.indentation.right
			? twipsToPx(props.indentation.right)
			: 0;
		const firstLine = props.indentation.firstLine
			? twipsToPx(props.indentation.firstLine)
			: 0;
		const hanging = props.indentation.hanging
			? twipsToPx(props.indentation.hanging)
			: 0;

		if (left > 0) el.style.marginLeft = `${left}px`;
		if (right > 0) el.style.marginRight = `${right}px`;
		if (firstLine > 0) el.style.textIndent = `${firstLine}px`;
		if (hanging > 0) el.style.textIndent = `-${hanging}px`;
	}

	if (props.spacing) {
		if (props.spacing.before > 0) {
			el.style.marginTop = `${twipsToPx(props.spacing.before)}px`;
		}
		if (props.spacing.after > 0) {
			el.style.marginBottom = `${twipsToPx(props.spacing.after)}px`;
		}
	}
}

// ── Inline Rendering ────────────────────────────────────────────────────────

function renderInline(
	inline: DocxInlineElement,
	container: HTMLElement,
	ctx: RenderContext,
	styleRProps: DocxRunProperties,
): void {
	switch (inline.type) {
		case "run":
			renderRun(inline, container, styleRProps);
			break;
		case "hyperlink":
			renderHyperlink(inline, container, ctx, styleRProps);
			break;
		case "image":
			renderImage(inline, container, ctx);
			break;
		case "break":
			renderBreak(inline, container);
			break;
		case "tab":
			container.createEl("span", {
				cls: "via-docx-tab",
				text: "\u00A0\u00A0\u00A0\u00A0",
			});
			break;
	}
}

function renderRun(
	run: DocxRun,
	container: HTMLElement,
	styleRProps: DocxRunProperties,
): void {
	if (!run.text) return;
	// Merge style-level run properties with run-level overrides
	const props = mergeRunProps(styleRProps, run.properties);

	let el: HTMLElement = container.createEl("span");
	el.textContent = run.text;

	// Apply formatting by wrapping in semantic elements
	if (props.bold) {
		const strong = container.createEl("strong");
		strong.appendChild(el);
		el = strong;
	}
	if (props.italic) {
		const em = container.createEl("em");
		em.appendChild(el);
		el = em;
	}
	if (props.underline) {
		const u = container.createEl("u");
		u.appendChild(el);
		el = u;
	}
	if (props.strikethrough) {
		const s = container.createEl("s");
		s.appendChild(el);
		el = s;
	}
	if (props.vertAlign === "superscript") {
		const sup = container.createEl("sup");
		sup.appendChild(el);
		el = sup;
	} else if (props.vertAlign === "subscript") {
		const sub = container.createEl("sub");
		sub.appendChild(el);
		el = sub;
	}

	// Force color inheritance so Obsidian themes can't override via CSS on
	// <strong>, <em>, etc. applyRunStyles will overwrite if doc sets color.
	el.style.color = "inherit";

	// Apply inline styles for color, size, highlight, font
	applyRunStyles(el, props);

	// Highlight wraps the run in a <mark>
	if (props.highlight) {
		const cssColor = HIGHLIGHT_COLORS[props.highlight];
		if (cssColor) {
			const mark = container.createEl("mark", { cls: "via-docx-highlight" });
			mark.style.backgroundColor = cssColor;
			mark.appendChild(el);
			el = mark;
		}
	}

	// The element is already appended by createEl, but if we wrapped it,
	// we need to ensure the outermost wrapper is in the container
	if (el.parentElement !== container) {
		container.appendChild(el);
	}
}

function applyRunStyles(el: HTMLElement, props: DocxRunProperties): void {
	if (props.color && props.color !== "auto") {
		el.style.color = `#${props.color}`;
	}

	if (props.fontSize) {
		const pt = halfPointsToPt(props.fontSize);
		// Use em units relative to parent for native feel
		const em = pt / 12; // Assume 12pt base
		if (Math.abs(em - 1) > 0.05) {
			el.style.fontSize = `${em.toFixed(2)}em`;
		}
	}

	if (props.fontFamily) {
		el.style.fontFamily = `"${props.fontFamily}", var(--font-text)`;
	}
}

function renderHyperlink(
	link: DocxHyperlink,
	container: HTMLElement,
	ctx: RenderContext,
	styleRProps: DocxRunProperties,
): void {
	const a = container.createEl("a", {
		cls: "via-docx-link",
		attr: { href: link.url },
	});

	// External links open in browser
	if (link.url.startsWith("http://") || link.url.startsWith("https://")) {
		a.setAttribute("target", "_blank");
		a.setAttribute("rel", "noopener noreferrer");
	}

	for (const run of link.children) {
		renderRun(run, a, styleRProps);
	}
}

function renderImage(
	image: DocxImage,
	container: HTMLElement,
	ctx: RenderContext,
): void {
	const blob = ctx.doc.images.get(image.relationshipId);
	if (!blob) {
		// Unsupported or missing image — show placeholder
		container.createEl("span", {
			cls: "via-docx-image-placeholder",
			text: `[Image: ${image.altText ?? "missing"}]`,
		});
		return;
	}

	const url = URL.createObjectURL(blob);
	ctx.blobUrls.push(url);

	const img = container.createEl("img", {
		cls: "via-docx-image",
		attr: {
			src: url,
			alt: image.altText ?? "",
		},
	});

	if (image.width > 0) img.style.maxWidth = `${image.width}px`;
	if (image.height > 0) img.style.maxHeight = `${image.height}px`;
}

function renderBreak(
	brk: { breakType: string },
	container: HTMLElement,
): void {
	if (brk.breakType === "page") {
		container.createEl("hr", { cls: "via-docx-page-break" });
	} else {
		container.createEl("br");
	}
}

// ── Numbering Text Generation ───────────────────────────────────────────────

/**
 * Generate the numbering text prefix for a paragraph (e.g., "Part 1:", "Step 2:", "a.").
 * Tracks and increments counters per numId+level combination.
 */
function getNumberingText(
	numId: string,
	level: number,
	ctx: RenderContext,
): string {
	const numDef = ctx.doc.numbering.get(numId);
	if (!numDef) return "";

	const lvlDef = numDef.levels.get(level);
	if (!lvlDef) return "";

	// format "none" means no visible numbering
	if (lvlDef.format === "none") return "";

	// Bullet format — return a bullet character
	if (lvlDef.format === "bullet") {
		return "\u2022 ";
	}

	// Increment the counter for this numId+level
	const counterKey = `${numId}:${level}`;
	const prev = ctx.numberingCounters.get(counterKey) ?? 0;
	const current = prev + 1;
	ctx.numberingCounters.set(counterKey, current);

	// Reset sub-level counters when a higher level increments
	for (const [key] of ctx.numberingCounters) {
		const [kNumId, kLevel] = key.split(":");
		if (kNumId === numId && Number(kLevel) > level) {
			ctx.numberingCounters.delete(key);
		}
	}

	// Build the numbering text from lvlText template
	// lvlText uses %1, %2, %3 etc. where %N refers to level N's counter
	let text = lvlDef.text;
	if (!text) {
		// Fallback: just use the counter
		return `${formatNumber(current, lvlDef.format)} `;
	}

	// Replace %N placeholders with the counter value for that level
	text = text.replace(/%(\d+)/g, (_match: string, levelStr: string) => {
		const refLevel = parseInt(levelStr, 10) - 1; // %1 = level 0, %2 = level 1
		const refKey = `${numId}:${refLevel}`;
		const refCount = ctx.numberingCounters.get(refKey) ?? 1;
		// Use the format of the referenced level
		const refLvlDef = numDef.levels.get(refLevel);
		const refFormat = refLvlDef?.format ?? "decimal";
		return formatNumber(refCount, refFormat);
	});

	return text + " ";
}

/** Format a number according to the numbering format type. */
function formatNumber(
	num: number,
	format: string,
): string {
	switch (format) {
		case "decimal":
			return String(num);
		case "lowerLetter":
			return String.fromCharCode(96 + ((num - 1) % 26) + 1);
		case "upperLetter":
			return String.fromCharCode(64 + ((num - 1) % 26) + 1);
		case "lowerRoman":
			return toRoman(num).toLowerCase();
		case "upperRoman":
			return toRoman(num);
		default:
			return String(num);
	}
}

/** Convert integer to Roman numeral. */
function toRoman(num: number): string {
	const vals = [1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1];
	const syms = ["M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I"];
	let result = "";
	let remaining = num;
	for (let i = 0; i < vals.length; i++) {
		const v = vals[i];
		const s = syms[i];
		if (v === undefined || s === undefined) continue;
		while (remaining >= v) {
			result += s;
			remaining -= v;
		}
	}
	return result;
}

// ── Table Rendering ─────────────────────────────────────────────────────────

function renderTable(
	table: DocxTable,
	container: HTMLElement,
	ctx: RenderContext,
): void {
	const tableEl = container.createEl("table", { cls: "via-docx-table" });

	// Track vertical merge state: column index → "consumed" flag
	const vMergeState: Map<number, boolean> = new Map();

	for (const row of table.rows) {
		renderTableRow(row, tableEl, ctx, vMergeState);
	}
}

function renderTableRow(
	row: DocxTableRow,
	tableEl: HTMLElement,
	ctx: RenderContext,
	vMergeState: Map<number, boolean>,
): void {
	const trEl = tableEl.createEl("tr");
	let colIdx = 0;

	for (const cell of row.cells) {
		// Skip cells that are vertically merged continuations
		if (cell.properties.verticalMerge === "continue") {
			// Increment the rowspan of the cell that started this merge
			colIdx += cell.properties.gridSpan;
			continue;
		}

		const tag = row.isHeader ? "th" : "td";
		const tdEl = trEl.createEl(tag, { cls: "via-docx-cell" });

		// Column span
		if (cell.properties.gridSpan > 1) {
			tdEl.setAttribute("colspan", String(cell.properties.gridSpan));
		}

		// Cell shading (background color)
		if (cell.properties.shading) {
			tdEl.style.backgroundColor = `#${cell.properties.shading}`;
		}

		// Vertical alignment
		if (cell.properties.verticalAlign) {
			tdEl.style.verticalAlign = cell.properties.verticalAlign;
		}

		// Cell width
		if (cell.properties.width) {
			tdEl.style.width = `${dxaToPx(cell.properties.width)}px`;
		}

		// Render cell content
		for (const para of cell.paragraphs) {
			renderParagraph(para, tdEl, ctx);
		}

		colIdx += cell.properties.gridSpan;
	}
}

// ── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Merge style-level run properties with run-level overrides.
 * Only properties that were explicitly set in the run XML take precedence.
 * Uses Partial so that `false` (explicitly turned off) beats the base.
 */
function mergeRunProps(
	base: DocxRunProperties,
	override: Partial<DocxRunProperties>,
): DocxRunProperties {
	return {
		bold: override.bold !== undefined ? override.bold : base.bold,
		italic: override.italic !== undefined ? override.italic : base.italic,
		underline: override.underline !== undefined ? override.underline : base.underline,
		strikethrough: override.strikethrough !== undefined ? override.strikethrough : base.strikethrough,
		fontFamily: override.fontFamily !== undefined ? override.fontFamily : base.fontFamily,
		fontSize: override.fontSize !== undefined ? override.fontSize : base.fontSize,
		color: override.color !== undefined ? override.color : base.color,
		highlight: override.highlight !== undefined ? override.highlight : base.highlight,
		vertAlign: override.vertAlign !== undefined ? override.vertAlign : base.vertAlign,
	};
}
