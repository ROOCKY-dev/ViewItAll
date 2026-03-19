/**
 * ViewItAll — DOCX Formatting Toolbar
 *
 * Selection-aware formatting controls: bold, italic, underline,
 * strikethrough, font, size, color, highlight, alignment, clear.
 * All buttons use Lucide icons via setIcon().
 */

import { setIcon, setTooltip } from "obsidian";
import type { DocxRunProperties, DocxParagraphProperties } from "./model";

// ── Format Commands ──────────────────────────────────────────────────────────

export type FormatCommand =
	| { type: "bold" }
	| { type: "italic" }
	| { type: "underline" }
	| { type: "strikethrough" }
	| { type: "fontSize"; value: number }
	| { type: "fontFamily"; value: string }
	| { type: "color"; value: string }
	| { type: "highlight"; value: string }
	| { type: "alignment"; value: "left" | "center" | "right" | "both" }
	| { type: "clearFormatting" };

// ── Public Interface ─────────────────────────────────────────────────────────

export interface FormattingToolbar {
	build(container: HTMLElement, onFormat: (cmd: FormatCommand) => void): void;
	updateState(rProps: DocxRunProperties, pProps: DocxParagraphProperties): void;
	show(): void;
	hide(): void;
	destroy(): void;
}

// ── Font sizes (display as pt) ───────────────────────────────────────────────

const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72];

const FONT_FAMILIES = [
	"Arial",
	"Calibri",
	"Cambria",
	"Comic Sans MS",
	"Consolas",
	"Courier New",
	"Georgia",
	"Helvetica",
	"Impact",
	"Lucida Console",
	"Palatino Linotype",
	"Segoe UI",
	"Tahoma",
	"Times New Roman",
	"Trebuchet MS",
	"Verdana",
];

// ── Full OOXML Highlight Color Palette ───────────────────────────────────────

const HIGHLIGHT_COLORS: { name: string; value: string; css: string }[] = [
	{ name: "Yellow", value: "yellow", css: "#FFFF00" },
	{ name: "Bright Green", value: "green", css: "#00FF00" },
	{ name: "Cyan", value: "cyan", css: "#00FFFF" },
	{ name: "Magenta", value: "magenta", css: "#FF00FF" },
	{ name: "Blue", value: "blue", css: "#0000FF" },
	{ name: "Red", value: "red", css: "#FF0000" },
	{ name: "Dark Blue", value: "darkBlue", css: "#00008B" },
	{ name: "Dark Cyan", value: "darkCyan", css: "#008B8B" },
	{ name: "Dark Green", value: "darkGreen", css: "#006400" },
	{ name: "Dark Magenta", value: "darkMagenta", css: "#8B008B" },
	{ name: "Dark Red", value: "darkRed", css: "#8B0000" },
	{ name: "Dark Yellow", value: "darkYellow", css: "#808000" },
	{ name: "Dark Gray", value: "darkGray", css: "#A9A9A9" },
	{ name: "Light Gray", value: "lightGray", css: "#D3D3D3" },
	{ name: "Black", value: "black", css: "#000000" },
	{ name: "White", value: "white", css: "#FFFFFF" },
];

// ── Implementation ───────────────────────────────────────────────────────────

export function createFormattingToolbar(): FormattingToolbar {
	let toolbarEl: HTMLElement | null = null;
	let onFormat: ((cmd: FormatCommand) => void) | null = null;

	// Button refs for active state
	let boldBtn: HTMLElement | null = null;
	let italicBtn: HTMLElement | null = null;
	let underlineBtn: HTMLElement | null = null;
	let strikeBtn: HTMLElement | null = null;
	let fontSelect: HTMLSelectElement | null = null;
	let sizeSelect: HTMLSelectElement | null = null;
	let colorInput: HTMLInputElement | null = null;
	let colorSwatch: HTMLElement | null = null;
	let highlightMenu: HTMLElement | null = null;
	let alignBtns: Map<string, HTMLElement> = new Map();
	let outsideClickHandler: ((e: MouseEvent) => void) | null = null;

	function build(container: HTMLElement, cb: (cmd: FormatCommand) => void): void {
		onFormat = cb;

		toolbarEl = container.createEl("div", {
			cls: "via-docx-format-toolbar",
		});

		// ── Group 1: Text style toggles ──
		const styleGroup = toolbarEl.createEl("div", { cls: "via-docx-format-group" });

		boldBtn = makeIconBtn(styleGroup, "bold", "Bold (Ctrl+B)", () => {
			onFormat?.({ type: "bold" });
		});

		italicBtn = makeIconBtn(styleGroup, "italic", "Italic (Ctrl+I)", () => {
			onFormat?.({ type: "italic" });
		});

		underlineBtn = makeIconBtn(styleGroup, "underline", "Underline (Ctrl+U)", () => {
			onFormat?.({ type: "underline" });
		});

		strikeBtn = makeIconBtn(styleGroup, "strikethrough", "Strikethrough", () => {
			onFormat?.({ type: "strikethrough" });
		});

		// Separator
		toolbarEl.createEl("div", { cls: "via-toolbar-sep" });

		// ── Group 2: Font family + size ──
		const fontGroup = toolbarEl.createEl("div", { cls: "via-docx-format-group" });

		fontSelect = fontGroup.createEl("select", { cls: "via-docx-format-select" });
		const defaultFontOpt = fontSelect.createEl("option", { text: "Font", attr: { value: "" } });
		defaultFontOpt.disabled = true;
		defaultFontOpt.selected = true;
		for (const font of FONT_FAMILIES) {
			const opt = fontSelect.createEl("option", { text: font, attr: { value: font } });
			opt.style.fontFamily = font;
		}
		fontSelect.addEventListener("change", () => {
			if (fontSelect?.value) {
				onFormat?.({ type: "fontFamily", value: fontSelect.value });
			}
		});

		sizeSelect = fontGroup.createEl("select", { cls: "via-docx-format-select via-docx-format-select--size" });
		const defaultSizeOpt = sizeSelect.createEl("option", { text: "Size", attr: { value: "" } });
		defaultSizeOpt.disabled = true;
		defaultSizeOpt.selected = true;
		for (const size of FONT_SIZES) {
			sizeSelect.createEl("option", {
				text: String(size),
				attr: { value: String(Math.round(size * 2)) }, // half-points
			});
		}
		sizeSelect.addEventListener("change", () => {
			const val = parseInt(sizeSelect?.value ?? "", 10);
			if (!isNaN(val)) {
				onFormat?.({ type: "fontSize", value: val });
			}
		});

		// Separator
		toolbarEl.createEl("div", { cls: "via-toolbar-sep" });

		// ── Group 3: Color + Highlight ──
		const colorGroup = toolbarEl.createEl("div", { cls: "via-docx-format-group" });

		// Text color — visible swatch + hidden native picker
		const colorWrap = colorGroup.createEl("div", { cls: "via-docx-format-color-wrap" });
		const colorBtn = colorWrap.createEl("div", { cls: "clickable-icon via-docx-format-color-btn" });
		setTooltip(colorBtn, "Text color");

		// Icon + swatch bar
		const colorIconWrap = colorBtn.createEl("div", { cls: "via-docx-format-color-icon-wrap" });
		setIcon(colorIconWrap, "type");
		colorSwatch = colorBtn.createEl("div", { cls: "via-docx-format-color-swatch" });
		colorSwatch.setCssStyles({ "background": "#000000" });

		colorInput = colorWrap.createEl("input", {
			cls: "via-docx-format-color-input",
			attr: { type: "color", value: "#000000" },
		});
		colorBtn.addEventListener("click", () => colorInput?.click());
		// Use both input (live) and change (confirmed) for responsiveness
		colorInput.addEventListener("input", handleColorChange);
		colorInput.addEventListener("change", handleColorChange);

		// Highlight — grid palette popup (appended to document.body to escape overflow clips)
		const highlightWrap = colorGroup.createEl("div", { cls: "via-docx-format-highlight-wrap" });
		const highlightBtn = highlightWrap.createEl("div", { cls: "clickable-icon" });
		setIcon(highlightBtn, "highlighter");
		setTooltip(highlightBtn, "Highlight color");

		// Append to body so it's never clipped by parent overflow
		highlightMenu = document.body.createEl("div", {
			cls: "via-docx-format-highlight-menu via-hidden",
		});

		// Grid layout for colors
		const grid = highlightMenu.createEl("div", { cls: "via-docx-format-highlight-grid" });
		for (const opt of HIGHLIGHT_COLORS) {
			const swatch = grid.createEl("div", {
				cls: "via-docx-format-highlight-swatch",
				attr: { title: opt.name },
			});
			swatch.style.background = opt.css;
			if (["darkBlue", "darkGreen", "darkMagenta", "darkRed", "black"].includes(opt.value)) {
				swatch.classList.add("via-docx-format-highlight-swatch--dark");
			}
			swatch.addEventListener("click", () => {
				closeHighlightMenu();
				onFormat?.({ type: "highlight", value: opt.value });
			});
		}

		// "No highlight" button below grid
		const noHighlight = highlightMenu.createEl("div", {
			cls: "via-docx-format-highlight-none",
		});
		noHighlight.createEl("span", { text: "No highlight" });
		noHighlight.addEventListener("click", () => {
			closeHighlightMenu();
			onFormat?.({ type: "highlight", value: "" });
		});

		highlightBtn.addEventListener("click", (e) => {
			e.stopPropagation();
			if (highlightMenu?.classList.contains("via-hidden")) {
				// Position the menu below the button using fixed positioning
				const rect = highlightBtn.getBoundingClientRect();
				highlightMenu.style.top = `${rect.bottom + 4}px`;
				highlightMenu.style.left = `${rect.left}px`;
				highlightMenu.classList.remove("via-hidden");
			} else {
				closeHighlightMenu();
			}
		});

		// Close highlight menu on outside click
		outsideClickHandler = (e: MouseEvent) => {
			if (highlightMenu && !highlightMenu.contains(e.target as Node) && !highlightWrap.contains(e.target as Node)) {
				closeHighlightMenu();
			}
		};
		document.addEventListener("click", outsideClickHandler);

		// Separator
		toolbarEl.createEl("div", { cls: "via-toolbar-sep" });

		// ── Group 4: Alignment ──
		const alignGroup = toolbarEl.createEl("div", { cls: "via-docx-format-group" });

		const alignments: Array<{ icon: string; value: "left" | "center" | "right" | "both"; tip: string }> = [
			{ icon: "align-left", value: "left", tip: "Align left" },
			{ icon: "align-center", value: "center", tip: "Center" },
			{ icon: "align-right", value: "right", tip: "Align right" },
			{ icon: "align-justify", value: "both", tip: "Justify" },
		];

		for (const a of alignments) {
			const btn = makeIconBtn(alignGroup, a.icon, a.tip, () => {
				onFormat?.({ type: "alignment", value: a.value });
			});
			alignBtns.set(a.value, btn);
		}

		// Separator
		toolbarEl.createEl("div", { cls: "via-toolbar-sep" });

		// ── Clear formatting ──
		makeIconBtn(toolbarEl, "eraser", "Clear formatting", () => {
			onFormat?.({ type: "clearFormatting" });
		});
	}

	function closeHighlightMenu(): void {
		highlightMenu?.classList.add("via-hidden");
	}

	function handleColorChange(): void {
		if (!colorInput?.value) return;
		const hex = colorInput.value.replace("#", "");

		// Update the swatch immediately
		if (colorSwatch) {
			colorSwatch.style.background = colorInput.value;
		}

		onFormat?.({ type: "color", value: hex });
	}

	function updateState(rProps: DocxRunProperties, pProps: DocxParagraphProperties): void {
		// Toggle buttons
		boldBtn?.classList.toggle("is-active", rProps.bold);
		italicBtn?.classList.toggle("is-active", rProps.italic);
		underlineBtn?.classList.toggle("is-active", rProps.underline);
		strikeBtn?.classList.toggle("is-active", rProps.strikethrough);

		// Font family
		if (fontSelect) {
			if (rProps.fontFamily) {
				fontSelect.value = rProps.fontFamily;
			} else {
				fontSelect.selectedIndex = 0;
			}
		}

		// Font size (half-points)
		if (sizeSelect) {
			if (rProps.fontSize) {
				sizeSelect.value = String(rProps.fontSize);
			} else {
				sizeSelect.selectedIndex = 0;
			}
		}

		// Text color — update swatch and input
		if (rProps.color && rProps.color !== "auto") {
			const hex = `#${rProps.color}`;
			if (colorInput) colorInput.value = hex;
			if (colorSwatch) colorSwatch.setCssStyles({ "background": hex });
		} else {
			if (colorInput) colorInput.value = "#000000";
			if (colorSwatch) colorSwatch.setCssStyles({ "background": "var(--text-normal)" });
		}

		// Alignment
		for (const [val, btn] of alignBtns) {
			btn.classList.toggle("is-active", pProps.alignment === val);
		}
	}

	function show(): void {
		toolbarEl?.classList.remove("via-hidden");
	}

	function hide(): void {
		toolbarEl?.classList.add("via-hidden");
		closeHighlightMenu();
	}

	function destroy(): void {
		if (outsideClickHandler) {
			document.removeEventListener("click", outsideClickHandler);
			outsideClickHandler = null;
		}
		if (highlightMenu) {
			highlightMenu.remove();
			highlightMenu = null;
		}
		if (toolbarEl) {
			toolbarEl.remove();
			toolbarEl = null;
		}
		onFormat = null;
		boldBtn = null;
		italicBtn = null;
		underlineBtn = null;
		strikeBtn = null;
		fontSelect = null;
		sizeSelect = null;
		colorInput = null;
		colorSwatch = null;
		alignBtns.clear();
	}

	return { build, updateState, show, hide, destroy };
}

// ── Helper ───────────────────────────────────────────────────────────────────

function makeIconBtn(
	parent: HTMLElement,
	icon: string,
	tooltip: string,
	onClick: () => void,
): HTMLElement {
	const btn = parent.createEl("div", { cls: "clickable-icon via-docx-format-btn" });
	setIcon(btn, icon);
	setTooltip(btn, tooltip);
	btn.addEventListener("click", onClick);
	return btn;
}
