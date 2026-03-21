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
	DocxRunProperties,
	DocxParagraphProperties,
	DocxTableBorder,
} from "./model";
import { defaultRunProperties, defaultParagraphProperties } from "./model";
import { resolveStyle, resolveHeadingLevel } from "./styles";
import { halfPointsToPt, twipsToPx, twipsToPt, dxaToPx } from "../utils/units";
import { readabilityGuard, adaptShadingForTheme } from "./themeAdapter";

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
	/** Whether the current Obsidian theme is dark mode */
	isDark: boolean;
}

/**
 * Render a DocxDocument into a DOM container element.
 * Returns an array of blob URLs that must be revoked on cleanup.
 */
export function renderDocument(
	doc: DocxDocument,
	container: HTMLElement,
): string[] {
	const isDark = document.body.classList.contains("theme-dark");
	const ctx: RenderContext = {
		doc,
		blobUrls: [],
		styleCache: new Map(),
		numberingCounters: new Map(),
		lastNumIdAtLevel: new Map(),
		isDark,
	};

	renderBlocks(doc.body, container, ctx);

	return ctx.blobUrls;
}

/**
 * Re-render a single paragraph element in place.
 * Used by the editing controller after model mutations.
 * Returns the new DOM element that replaced the old one.
 */
export function rerenderParagraph(
	doc: DocxDocument,
	blockIdx: number,
	oldEl: HTMLElement,
	styleCache: Map<string, { pProps: DocxParagraphProperties; rProps: DocxRunProperties }>,
): { el: HTMLElement; blobUrls: string[] } {
	const block = doc.body[blockIdx];
	if (!block || block.type !== "paragraph") {
		return { el: oldEl, blobUrls: [] };
	}

	const isDark = document.body.classList.contains("theme-dark");
	const ctx: RenderContext = {
		doc,
		blobUrls: [],
		styleCache,
		numberingCounters: new Map(),
		lastNumIdAtLevel: new Map(),
		isDark,
	};

	// Create a temporary container to render into
	const parent = oldEl.parentElement;
	if (!parent) return { el: oldEl, blobUrls: [] };

	const tempDiv = document.createElement("div");
	renderParagraph(block, tempDiv, ctx, blockIdx);

	const newEl = tempDiv.firstElementChild as HTMLElement;
	if (!newEl) return { el: oldEl, blobUrls: [] };

	parent.replaceChild(newEl, oldEl);
	return { el: newEl, blobUrls: ctx.blobUrls };
}

// ── Block Rendering ─────────────────────────────────────────────────────────

function renderBlocks(
	blocks: DocxBlockElement[],
	container: HTMLElement,
	ctx: RenderContext,
): void {
	for (let blockIdx = 0; blockIdx < blocks.length; blockIdx++) {
		const block = blocks[blockIdx];
		if (!block) continue;

		if (block.type === "paragraph") {
			renderParagraph(block, container, ctx, blockIdx);
		} else if (block.type === "table") {
			renderTable(block, container, ctx, blockIdx);
		} else if (block.type === "pageBreak") {
			const hr = container.createEl("hr", { cls: "via-docx-page-break" });
			hr.dataset.blockIdx = String(blockIdx);
		}
	}
}

// ── Paragraph Rendering ─────────────────────────────────────────────────────

function renderParagraph(
	para: DocxParagraph,
	container: HTMLElement,
	ctx: RenderContext,
	blockIdx?: number,
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
	}

	// Compute effective paragraph properties without mutating the model.
	// Direct properties take precedence over style-resolved properties.
	const effectiveProps: DocxParagraphProperties = {
		alignment: para.properties.alignment ?? resolvedPProps.alignment,
		headingLevel: para.properties.headingLevel ?? resolvedPProps.headingLevel,
		indentation: para.properties.indentation ?? resolvedPProps.indentation,
		spacing: para.properties.spacing ?? resolvedPProps.spacing,
		numberingId: para.properties.numberingId ?? resolvedPProps.numberingId,
		numberingLevel: para.properties.numberingLevel ?? resolvedPProps.numberingLevel,
		shading: para.properties.shading ?? resolvedPProps.shading,
	};

	// Resolve numbering: direct paragraph > style-inherited
	let numId = para.numberingId;
	let numLevel = para.numberingLevel;
	if (!numId && effectiveProps.numberingId) {
		numId = effectiveProps.numberingId;
		numLevel = effectiveProps.numberingLevel ?? 0;
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
	} else {
		el = container.createEl("p");
	}

	// Data attribute for DOM ↔ model binding (editing support)
	if (blockIdx !== undefined) {
		el.dataset.blockIdx = String(blockIdx);
	}

	// Apply paragraph styles using effective (merged) properties
	applyParagraphStyles(el, effectiveProps, ctx.isDark);

	// Check if paragraph is empty (just whitespace/empty runs)
	const hasContent = para.children.some((child) => {
		if (child.type === "run") return child.text.trim().length > 0;
		if (child.type === "hyperlink") return child.children.length > 0;
		if (child.type === "image") return true;
		if (child.type === "break") return true;
		return false;
	});

	if (!hasContent) {
		// Empty paragraph — add a <br> to preserve spacing and caret position
		el.createEl("br");
		return;
	}

	// Prepend numbering text if present
	if (numberingPrefix) {
		const numSpan = el.createEl("span", {
			cls: "via-docx-num-prefix",
		});
		numSpan.textContent = numberingPrefix;
	}

	// Render inline children
	for (let runIdx = 0; runIdx < para.children.length; runIdx++) {
		const child = para.children[runIdx];
		if (!child) continue;
		renderInline(child, el, ctx, resolvedRProps, runIdx, effectiveProps.shading);
	}
}

function applyParagraphStyles(
	el: HTMLElement,
	props: DocxParagraphProperties,
	isDark: boolean,
): void {
	if (props.alignment) {
		const align = props.alignment === "both" ? "justify" : props.alignment;
		el.setCssStyles({ textAlign: align });
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

		if (left > 0) el.setCssStyles({ marginLeft: `${left}px` });
		if (right > 0) el.setCssStyles({ marginRight: `${right}px` });
		if (firstLine > 0) el.setCssStyles({ textIndent: `${firstLine}px` });
		if (hanging > 0) el.setCssStyles({ textIndent: `-${hanging}px` });
	}

	if (props.spacing) {
		if (props.spacing.before > 0) {
			el.setCssStyles({ marginTop: `${twipsToPx(props.spacing.before)}px` });
		}
		if (props.spacing.after > 0) {
			el.setCssStyles({ marginBottom: `${twipsToPx(props.spacing.after)}px` });
		}
		// B3: Apply line spacing
		if (props.spacing.line > 0) {
			if (props.spacing.lineRule === "exact") {
				el.setCssStyles({ lineHeight: `${twipsToPt(props.spacing.line)}pt` });
			} else if (props.spacing.lineRule === "atLeast") {
				el.setCssStyles({ minHeight: `${twipsToPt(props.spacing.line)}pt`, lineHeight: `${twipsToPt(props.spacing.line)}pt` });
			} else {
				const factor = props.spacing.line / 240;
				if (Math.abs(factor - 1) > 0.05) {
					el.setCssStyles({ lineHeight: `${factor.toFixed(2)}` });
				}
			}
		}
	}

	// B1: Paragraph background shading — C3: adapted for theme
	if (props.shading) {
		const adapted = adaptShadingForTheme(props.shading, isDark);
		el.setCssStyles({ backgroundColor: `#${adapted}`, padding: "4px 8px", borderRadius: "var(--radius-s)" });
	}
}

// ── Inline Rendering ────────────────────────────────────────────────────────

function renderInline(
	inline: DocxInlineElement,
	container: HTMLElement,
	ctx: RenderContext,
	styleRProps: DocxRunProperties,
	runIdx?: number,
	parentBgHex?: string,
): void {
	switch (inline.type) {
		case "run":
			renderRun(inline, container, styleRProps, ctx, runIdx, parentBgHex);
			break;
		case "hyperlink":
			renderHyperlink(inline, container, ctx, styleRProps, parentBgHex);
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
	ctx: RenderContext,
	runIdx?: number,
	parentBgHex?: string,
): void {
	if (!run.text) return;
	// Merge style-level run properties with run-level overrides
	const props = mergeRunProps(styleRProps, run.properties);

	let el: HTMLElement = container.createEl("span");
	el.textContent = run.text;

	// Data attribute for DOM ↔ model binding (editing support)
	if (runIdx !== undefined) {
		el.dataset.runIdx = String(runIdx);
	}

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

	// CSS class "via-docx-run" handles color inheritance;
	// applyRunStyles will set --via-color if the document specifies a color.
	el.classList.add("via-docx-run");

	// Apply inline styles for color, size, highlight, font
	applyRunStyles(el, props, ctx.isDark, parentBgHex);

	// Highlight wraps the run in a <mark>
	if (props.highlight) {
		const cssColor = HIGHLIGHT_COLORS[props.highlight];
		if (cssColor) {
			const mark = container.createEl("mark", { cls: "via-docx-highlight" });
			mark.setCssStyles({ backgroundColor: cssColor });
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

function applyRunStyles(
	el: HTMLElement,
	props: DocxRunProperties,
	isDark = false,
	parentBgHex?: string,
): void {
	if (props.color && props.color !== "auto") {
		// C4: readability guard — prevent invisible text
		const safeColor = readabilityGuard(props.color, parentBgHex, isDark);
		if (safeColor) {
			el.setCssStyles({ color: `#${safeColor}` });
		}
		// If safeColor is undefined, text uses theme default (inherited)
	}

	if (props.fontSize) {
		const pt = halfPointsToPt(props.fontSize);
		// B2: Use absolute pt units — avoids incorrect scaling from assumed base
		el.setCssStyles({ fontSize: `${pt}pt` });
	}

	if (props.fontFamily) {
		el.setCssStyles({ fontFamily: `"${props.fontFamily}", var(--font-text)` });
	}
}

function renderHyperlink(
	link: DocxHyperlink,
	container: HTMLElement,
	ctx: RenderContext,
	styleRProps: DocxRunProperties,
	parentBgHex?: string,
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
		renderRun(run, a, styleRProps, ctx, undefined, parentBgHex);
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

	// Wrap image in a resizable container
	const wrap = container.createEl("span", {
		cls: "via-docx-image-wrap",
	});
	wrap.dataset.imageRid = image.relationshipId;
	wrap.setAttribute("contenteditable", "false");

	const img = wrap.createEl("img", {
		cls: "via-docx-image",
		attr: {
			src: url,
			alt: image.altText ?? "",
		},
	});

	// Set explicit width to preserve Word's intended size.
	// Height auto-scales to maintain aspect ratio.
	// The CSS max-width: 100% ensures it never overflows the container.
	if (image.width > 0 && image.height > 0) {
		wrap.setCssStyles({ width: `${image.width}px`, display: "inline-block" });
		img.setCssStyles({ width: "100%", height: "auto", aspectRatio: `${image.width} / ${image.height}` });
	} else if (image.width > 0) {
		wrap.setCssStyles({ width: `${image.width}px`, display: "inline-block" });
		img.setCssStyles({ width: "100%", height: "auto" });
	} else if (image.height > 0) {
		img.setCssStyles({ height: `${image.height}px`, width: "auto" });
	}

	// Resize handle (bottom-right corner)
	const handle = wrap.createEl("div", { cls: "via-docx-image-resize-handle" });
	handle.dataset.imageRid = image.relationshipId;
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

function applyBorder(
	el: HTMLElement,
	side: "top" | "bottom" | "left" | "right",
	border: DocxTableBorder,
): void {
	const cssProp = `border${side.charAt(0).toUpperCase()}${side.slice(1)}` as keyof CSSStyleDeclaration;
	if (border.val === "nil" || border.val === "none") {
		el.setCssStyles({ [cssProp]: "none" } as Partial<CSSStyleDeclaration>);
		return;
	}
	const color =
		border.color && border.color !== "auto"
			? `#${border.color}`
			: "var(--table-border-color, var(--background-modifier-border))";
	const sz = border.sz ? Math.max(1, Math.round(border.sz / 8)) : 1;
	let style = "solid";
	if (border.val === "dashed" || border.val === "dotted" || border.val === "double") {
		style = border.val;
	}
	el.setCssStyles({ [cssProp]: `${sz}px ${style} ${color}` } as Partial<CSSStyleDeclaration>);
}

function renderTable(
	table: DocxTable,
	container: HTMLElement,
	ctx: RenderContext,
	blockIdx?: number,
): void {
	const tableEl = container.createEl("table", { cls: "via-docx-table" });
	if (blockIdx !== undefined) {
		tableEl.dataset.blockIdx = String(blockIdx);
	}

	if (table.properties.styleId) {
		tableEl.classList.add(`via-docx-table--${table.properties.styleId}`);
	}

	if (table.properties.borders) {
		const tb = table.properties.borders;
		if (tb.top) applyBorder(tableEl, "top", tb.top);
		if (tb.bottom) applyBorder(tableEl, "bottom", tb.bottom);
		if (tb.left) applyBorder(tableEl, "left", tb.left);
		if (tb.right) applyBorder(tableEl, "right", tb.right);
	}

	// Track vertical merge state: visual column index → td element to update rowspan
	const vMergeState: Map<number, HTMLTableCellElement> = new Map();

	for (const row of table.rows) {
		renderTableRow(row, tableEl, ctx, vMergeState);
	}
}

function renderTableRow(
	row: DocxTableRow,
	tableEl: HTMLElement,
	ctx: RenderContext,
	vMergeState: Map<number, HTMLTableCellElement>,
): void {
	const trEl = tableEl.createEl("tr");
	let colIdx = 0;

	for (const cell of row.cells) {
		const currentColIdx = colIdx;
		colIdx += cell.properties.gridSpan;

		// Skip cells that are vertically merged continuations but increment their master's rowspan
		if (cell.properties.verticalMerge === "continue") {
			const rootTd = vMergeState.get(currentColIdx);
			if (rootTd) {
				const currentSpan = parseInt(rootTd.getAttribute("rowspan") || "1", 10);
				rootTd.setAttribute("rowspan", String(currentSpan + 1));
			}
			continue;
		}

		const tag = row.isHeader ? "th" : "td";
		const tdEl = trEl.createEl(tag, { cls: "via-docx-cell" }) as HTMLTableCellElement;

		if (cell.properties.verticalMerge === "restart") {
			vMergeState.set(currentColIdx, tdEl);
		} else {
			vMergeState.delete(currentColIdx);
		}

		// Column span
		if (cell.properties.gridSpan > 1) {
			tdEl.setAttribute("colspan", String(cell.properties.gridSpan));
		}

		// Cell shading (background color)
		if (cell.properties.shading) {
			const adapted = adaptShadingForTheme(cell.properties.shading, ctx.isDark);
			tdEl.setCssStyles({ backgroundColor: `#${adapted}` });
		}

		// Vertical alignment
		if (cell.properties.verticalAlign) {
			tdEl.setCssStyles({ verticalAlign: cell.properties.verticalAlign });
		}

		// Cell width
		if (cell.properties.width) {
			tdEl.setCssStyles({ width: `${dxaToPx(cell.properties.width)}px` });
		}

		// Cell borders
		if (cell.properties.borders) {
			const b = cell.properties.borders;
			if (b.top) applyBorder(tdEl, "top", b.top);
			if (b.bottom) applyBorder(tdEl, "bottom", b.bottom);
			if (b.left) applyBorder(tdEl, "left", b.left);
			if (b.right) applyBorder(tdEl, "right", b.right);
		}

		// Cell margins
		if (cell.properties.margins) {
			const m = cell.properties.margins;
			if (m.top !== undefined && m.top > 0) tdEl.setCssStyles({ paddingTop: `${dxaToPx(m.top)}px` });
			if (m.bottom !== undefined && m.bottom > 0) tdEl.setCssStyles({ paddingBottom: `${dxaToPx(m.bottom)}px` });
			if (m.left !== undefined && m.left > 0) tdEl.setCssStyles({ paddingLeft: `${dxaToPx(m.left)}px` });
			if (m.right !== undefined && m.right > 0) tdEl.setCssStyles({ paddingRight: `${dxaToPx(m.right)}px` });
		}

		// Render cell content
		for (const para of cell.paragraphs) {
			// Inherit cell shading into child paragraph if not overriden
			if (!para.properties.shading && cell.properties.shading) {
				para.properties.shading = cell.properties.shading;
			}
			renderParagraph(para, tdEl, ctx);
		}
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
