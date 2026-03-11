import type * as pdfjsLib from "pdfjs-dist";
import type { TextItem as PdfTextItem } from "pdfjs-dist/types/src/display/api";
import { setIcon, setTooltip } from "obsidian";
import type { PageCtx, SearchMatch } from "./pdfTypes";

/**
 * Manages the text-search bar, match highlighting, and page-text caching for
 * a PdfView instance. Call `setContext()` each time a new PDF is loaded.
 */
export class PdfSearchController {
	private pdfDoc: pdfjsLib.PDFDocumentProxy | null = null;
	private pages: PageCtx[] = [];
	private wrapperEl: HTMLElement | null = null;
	private bodyEl: HTMLElement | null = null;

	// DOM elements (null when bar is closed)
	private searchBarEl: HTMLElement | null = null;
	private searchInputEl: HTMLInputElement | null = null;
	private searchMatchCountEl: HTMLElement | null = null;

	// State
	private matches: SearchMatch[] = [];
	private currentIdx = -1;
	private textCache = new Map<number, PdfTextItem[]>();
	private debounceTimer: ReturnType<typeof setTimeout> | null = null;

	// ── Lifecycle ─────────────────────────────────────────────────────────

	/**
	 * Call this each time a new PDF is loaded into the view.
	 * Closes any open search bar and resets all state.
	 */
	setContext(
		pdfDoc: pdfjsLib.PDFDocumentProxy,
		pages: PageCtx[],
		wrapperEl: HTMLElement,
		bodyEl: HTMLElement,
	): void {
		this.close();
		this.pdfDoc = pdfDoc;
		this.pages = pages;
		this.wrapperEl = wrapperEl;
		this.bodyEl = bodyEl;
		this.textCache.clear();
		this.matches = [];
		this.currentIdx = -1;
	}

	/** Call when the view is unloaded to clean up DOM and timers. */
	destroy(): void {
		this.close();
		this.pdfDoc = null;
		this.pages = [];
		this.wrapperEl = null;
		this.bodyEl = null;
	}

	// ── Public API ────────────────────────────────────────────────────────

	open(): void {
		if (this.searchBarEl) {
			this.searchInputEl?.focus();
			return;
		}
		if (!this.wrapperEl) return;

		const bar = this.wrapperEl.createEl("div", {
			cls: "via-pdf-search-bar",
		});
		this.searchBarEl = bar;

		const iconEl = bar.createEl("span", { cls: "via-pdf-search-icon" });
		setIcon(iconEl, "search");

		this.searchInputEl = bar.createEl("input");
		this.searchInputEl.type = "text";
		this.searchInputEl.placeholder = "Search…";
		this.searchInputEl.className = "via-pdf-search-input";

		this.searchMatchCountEl = bar.createEl("span", {
			cls: "via-pdf-search-count",
			text: "",
		});

		const prevBtn = bar.createEl("div", { cls: "clickable-icon" });
		setIcon(prevBtn, "chevron-up");
		setTooltip(prevBtn, "Previous match (Shift+Enter)");
		prevBtn.addEventListener("click", () =>
			this.goToMatch(this.currentIdx - 1),
		);

		const nextBtn = bar.createEl("div", { cls: "clickable-icon" });
		setIcon(nextBtn, "chevron-down");
		setTooltip(nextBtn, "Next match (Enter)");
		nextBtn.addEventListener("click", () =>
			this.goToMatch(this.currentIdx + 1),
		);

		const closeBtn = bar.createEl("div", { cls: "clickable-icon" });
		setIcon(closeBtn, "x");
		setTooltip(closeBtn, "Close search (Escape)");
		closeBtn.addEventListener("click", () => this.close());

		this.searchInputEl.addEventListener("keydown", (e) => {
			e.stopPropagation();
			if (e.key === "Enter") {
				e.preventDefault();
				e.shiftKey
					? this.goToMatch(this.currentIdx - 1)
					: this.goToMatch(this.currentIdx + 1);
			} else if (e.key === "Escape") {
				e.preventDefault();
				this.close();
			}
		});

		this.searchInputEl.addEventListener("input", () => {
			if (this.debounceTimer) clearTimeout(this.debounceTimer);
			this.debounceTimer = setTimeout(
				() => this.performSearch(this.searchInputEl!.value),
				300,
			);
		});

		// Insert search bar between toolbar and body area
		if (this.bodyEl) this.wrapperEl.insertBefore(bar, this.bodyEl);

		this.searchInputEl.focus();
	}

	close(): void {
		this.matches = [];
		this.currentIdx = -1;
		this.clearAll();
		this.updateCount();
		if (this.debounceTimer) {
			clearTimeout(this.debounceTimer);
			this.debounceTimer = null;
		}
		this.searchBarEl?.remove();
		this.searchBarEl = null;
		this.searchInputEl = null;
		this.searchMatchCountEl = null;
	}

	get hasMatches(): boolean {
		return this.matches.length > 0;
	}

	/** Call after a page canvas has been rendered (or re-rendered) to draw highlights. */
	drawHighlightsForPage(ctx: PageCtx): void {
		if (!ctx.searchCanvas) return;
		const canvas = ctx.searchCanvas;
		const c = canvas.getContext("2d")!;
		c.clearRect(0, 0, canvas.width, canvas.height);

		for (let i = 0; i < this.matches.length; i++) {
			const m = this.matches[i]!;
			if (m.pageNum !== ctx.pageNum) continue;
			c.fillStyle =
				i === this.currentIdx
					? "rgba(255, 140, 0, 0.55)"
					: "rgba(255, 220, 0, 0.38)";
			c.fillRect(
				m.x * canvas.width,
				m.y * canvas.height,
				m.w * canvas.width,
				m.h * canvas.height,
			);
		}
	}

	/** Redraw highlights on all currently-rendered pages. */
	redrawAll(): void {
		for (const ctx of this.pages) {
			if (ctx.state === "rendered") this.drawHighlightsForPage(ctx);
		}
	}

	/** Clear all search highlight canvases. */
	clearAll(): void {
		for (const ctx of this.pages) {
			if (!ctx.searchCanvas) continue;
			const c = ctx.searchCanvas.getContext("2d")!;
			c.clearRect(0, 0, ctx.searchCanvas.width, ctx.searchCanvas.height);
		}
	}

	// ── Private ───────────────────────────────────────────────────────────

	private async performSearch(query: string): Promise<void> {
		this.matches = [];
		this.currentIdx = -1;
		this.clearAll();

		if (!query.trim() || !this.pdfDoc) {
			this.updateCount();
			return;
		}

		const q = query.toLowerCase();
		const total = this.pdfDoc.numPages;

		for (let pn = 1; pn <= total; pn++) {
			const items = await this.getPageTextItems(pn);
			const page = await this.pdfDoc.getPage(pn);
			const vp = page.getViewport({ scale: 1.0 });

			for (const item of items) {
				const str = item.str.toLowerCase();
				let idx = 0;
				while ((idx = str.indexOf(q, idx)) !== -1) {
					const tx = item.transform[4] ?? 0;
					const ty = item.transform[5] ?? 0;
					const fontSize = Math.abs(item.transform[3] ?? 12);
					const charWidth =
						(item.width || 0) / (item.str.length || 1);
					const matchX = tx + charWidth * idx;
					const matchW = charWidth * query.length;

					const [cx, cy] = vp.convertToViewportPoint(matchX, ty);
					const [cx2, cy2] = vp.convertToViewportPoint(
						matchX + matchW,
						ty - fontSize,
					);

					this.matches.push({
						pageNum: pn,
						x: Math.min(cx, cx2) / vp.width,
						y: Math.min(cy, cy2) / vp.height,
						w: Math.abs(cx2 - cx) / vp.width,
						h: Math.abs(cy2 - cy) / vp.height,
					});
					idx += q.length;
				}
			}
		}

		if (this.matches.length > 0) {
			this.currentIdx = 0;
			this.scrollToMatch(0);
		}
		this.updateCount();
		this.redrawAll();
	}

	private goToMatch(idx: number): void {
		if (this.matches.length === 0) return;
		this.currentIdx =
			((idx % this.matches.length) + this.matches.length) %
			this.matches.length;
		this.scrollToMatch(this.currentIdx);
		this.updateCount();
		this.redrawAll();
	}

	private scrollToMatch(idx: number): void {
		const m = this.matches[idx];
		if (!m) return;
		const ctx = this.pages.find((p) => p.pageNum === m.pageNum);
		if (ctx)
			ctx.container.scrollIntoView({
				behavior: "smooth",
				block: "center",
			});
	}

	private updateCount(): void {
		if (!this.searchMatchCountEl) return;
		if (this.matches.length === 0) {
			this.searchMatchCountEl.textContent =
				this.searchInputEl?.value.trim() ? "No matches" : "";
		} else {
			this.searchMatchCountEl.textContent = `${this.currentIdx + 1} / ${this.matches.length}`;
		}
	}

	private async getPageTextItems(pageNum: number): Promise<PdfTextItem[]> {
		if (this.textCache.has(pageNum)) return this.textCache.get(pageNum)!;
		const page = await this.pdfDoc!.getPage(pageNum);
		const tc = await page.getTextContent();
		const items = tc.items.filter((it): it is PdfTextItem => "str" in it);
		this.textCache.set(pageNum, items);
		return items;
	}
}
