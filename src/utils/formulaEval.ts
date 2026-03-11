/**
 * Lightweight spreadsheet formula evaluator.
 * Supports basic arithmetic, cell/range references, and common functions.
 */

type CellValue = string | number | boolean | null | undefined;

interface Cell {
	v?: CellValue;
	f?: string;
	t?: string;
	w?: string;
}

type Sheet = Record<string, unknown>;

interface CellAddr {
	r: number;
	c: number;
}
interface Range {
	s: CellAddr;
	e: CellAddr;
}

interface Utils {
	encode_cell(addr: CellAddr): string;
	decode_cell(addr: string): CellAddr;
	decode_range(range: string): Range;
}

// ── Public API ────────────────────────────────────────────────────────────────

/**
 * Evaluate every formula cell in the sheet (multi-pass to resolve dependencies).
 * Modifies cell `.v`, `.w`, `.t` in place.
 */
export function evaluateSheet(
	sheet: Sheet,
	utils: Utils,
	rangeStr: string | undefined,
): void {
	if (!rangeStr) return;
	const range = utils.decode_range(rangeStr);

	// Up to 3 passes so that formulas referencing other formulas stabilise
	for (let pass = 0; pass < 3; pass++) {
		for (let r = range.s.r; r <= range.e.r; r++) {
			for (let c = range.s.c; c <= range.e.c; c++) {
				const ref = utils.encode_cell({ r, c });
				const cell = sheet[ref] as Cell | undefined;
				if (!cell?.f) continue;
				try {
					const val = evalFormula(cell.f, sheet, utils, ref);
					cell.v = val;
					cell.w = fmtVal(val);
					cell.t =
						typeof val === "number"
							? "n"
							: typeof val === "boolean"
								? "b"
								: "s";
				} catch {
					// Leave pre-computed value if evaluation fails
				}
			}
		}
	}
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function fmtVal(v: number | string | boolean): string {
	if (typeof v === "number") {
		return Number.isInteger(v)
			? String(v)
			: parseFloat(v.toFixed(10)).toString();
	}
	return String(v);
}

function cellValue(sheet: Sheet, ref: string): number | string | boolean {
	const c = sheet[ref] as Cell | undefined;
	if (!c) return 0;

	// Prefer the raw value if it's a clean number or boolean
	if (typeof c.v === "number") return c.v;
	if (typeof c.v === "boolean") return c.v;

	// v might be a formatted string ("20.00%", "€100.00") — try to extract a number
	const raw = c.v != null ? String(c.v) : (c.w ?? "");
	return parseFormattedNumber(raw);
}

/** Parse a potentially formatted string into a number, or return it as-is. */
function parseFormattedNumber(s: string): number | string {
	const trimmed = s.trim();
	if (trimmed === "") return 0;

	// Percentage: "20.00%" → 0.2
	if (trimmed.endsWith("%")) {
		const n = Number(trimmed.slice(0, -1));
		if (!isNaN(n)) return n / 100;
	}

	// Strip common currency symbols / thousand separators then try parse
	const stripped = trimmed
		.replace(/^[€$£¥₹₽¥₩₪₫₴₸₺₼₻₥₦₧₨₩₿\s]+/, "")
		.replace(/,/g, "");
	if (stripped !== "" && !isNaN(Number(stripped))) return Number(stripped);

	// Plain number
	const n = Number(trimmed);
	if (!isNaN(n)) return n;

	// Not numeric — return as string
	return trimmed;
}

function resolveRange(
	sheet: Sheet,
	utils: Utils,
	rangeStr: string,
): (number | string | boolean)[] {
	const rng = utils.decode_range(rangeStr);
	const out: (number | string | boolean)[] = [];
	for (let r = rng.s.r; r <= rng.e.r; r++) {
		for (let c = rng.s.c; c <= rng.e.c; c++) {
			out.push(cellValue(sheet, utils.encode_cell({ r, c })));
		}
	}
	return out;
}

// ── Tokeniser ─────────────────────────────────────────────────────────────────

type Token =
	| { type: "num"; v: number }
	| { type: "str"; v: string }
	| { type: "bool"; v: boolean }
	| { type: "cell"; ref: string }
	| { type: "range"; ref: string }
	| { type: "func"; name: string }
	| { type: "op"; v: string }
	| { type: "paren"; v: "(" | ")" }
	| { type: "comma" };

const RE_RANGE = /^\$?[A-Z]{1,3}\$?\d{1,7}:\$?[A-Z]{1,3}\$?\d{1,7}/i;
const RE_CELL = /^\$?[A-Z]{1,3}\$?\d{1,7}/i;
const RE_FUNC = /^[A-Z_][A-Z0-9_.]*(?=\()/i;
const RE_NUM = /^\d+(\.\d+)?([eE][+-]?\d+)?/;

function tokenise(f: string): Token[] {
	const tokens: Token[] = [];
	let i = 0;
	while (i < f.length) {
		const s = f.slice(i);

		if (/^\s/.test(s)) {
			i++;
			continue;
		}

		// String literal
		if (s[0] === '"') {
			let j = 1;
			while (j < s.length && s[j] !== '"') j++;
			tokens.push({ type: "str", v: s.slice(1, j) });
			i += j + 1;
			continue;
		}

		// Boolean
		const bm = s.match(/^(TRUE|FALSE)\b/i);
		if (bm) {
			tokens.push({ type: "bool", v: bm[1]!.toUpperCase() === "TRUE" });
			i += bm[0].length;
			continue;
		}

		// Function (before cell/range so SUM( is caught)
		const fm = s.match(RE_FUNC);
		if (fm) {
			tokens.push({ type: "func", name: fm[0].toUpperCase() });
			i += fm[0].length;
			continue;
		}

		// Range
		const rm = s.match(RE_RANGE);
		if (rm) {
			tokens.push({
				type: "range",
				ref: rm[0].replace(/\$/g, "").toUpperCase(),
			});
			i += rm[0].length;
			continue;
		}

		// Cell ref
		const cm = s.match(RE_CELL);
		if (cm) {
			tokens.push({
				type: "cell",
				ref: cm[0].replace(/\$/g, "").toUpperCase(),
			});
			i += cm[0].length;
			continue;
		}

		// Number
		const nm = s.match(RE_NUM);
		if (nm) {
			tokens.push({ type: "num", v: parseFloat(nm[0]) });
			i += nm[0].length;
			continue;
		}

		// Two-char operators
		const two = s.slice(0, 2);
		if ([">=", "<=", "<>", "!="].includes(two)) {
			tokens.push({ type: "op", v: two });
			i += 2;
			continue;
		}

		// Single-char operators
		if ("+-*/^%&=<>".includes(s[0]!)) {
			tokens.push({ type: "op", v: s[0]! });
			i++;
			continue;
		}

		if (s[0] === "(" || s[0] === ")") {
			tokens.push({ type: "paren", v: s[0] });
			i++;
			continue;
		}

		if (s[0] === ",") {
			tokens.push({ type: "comma" });
			i++;
			continue;
		}

		// Skip unknown
		i++;
	}
	return tokens;
}

// ── Recursive-descent parser / evaluator ──────────────────────────────────────

type Arg = number | string | boolean | (number | string | boolean)[];

function evalFormula(
	formula: string,
	sheet: Sheet,
	utils: Utils,
	selfRef: string,
): number | string | boolean {
	const tokens = tokenise(formula);
	if (tokens.length === 0) return 0;
	let pos = 0;

	const peek = (): Token | undefined => tokens[pos];
	const next = (): Token => tokens[pos++]!;

	function expect(type: string, val?: string): Token {
		const t = next();
		if (!t || t.type !== type) throw new Error("Unexpected token");
		if (val !== undefined && "v" in t && (t as { v: unknown }).v !== val)
			throw new Error("Expected " + val);
		return t;
	}

	// ── Precedence layers ─────────────────────────────────────────────────

	function parseExpr(): number | string | boolean {
		return parseComparison();
	}

	function parseComparison(): number | string | boolean {
		let left = parseConcat();
		while (
			peek()?.type === "op" &&
			["=", "<", ">", ">=", "<=", "<>", "!="].includes(
				(peek() as { v: string }).v,
			)
		) {
			const op = (next() as { v: string }).v;
			const right = parseConcat();
			const l = typeof left === "number" ? left : Number(left);
			const r = typeof right === "number" ? right : Number(right);
			switch (op) {
				case "=":
					left = left === right;
					break;
				case "<":
					left = l < r;
					break;
				case ">":
					left = l > r;
					break;
				case ">=":
					left = l >= r;
					break;
				case "<=":
					left = l <= r;
					break;
				case "<>":
				case "!=":
					left = left !== right;
					break;
			}
		}
		return left;
	}

	function parseConcat(): number | string | boolean {
		let left = parseAddSub();
		while (peek()?.type === "op" && (peek() as { v: string }).v === "&") {
			next();
			const right = parseAddSub();
			left = String(left) + String(right);
		}
		return left;
	}

	function parseAddSub(): number | string | boolean {
		let left = parseMulDiv();
		while (
			peek()?.type === "op" &&
			["+", "-"].includes((peek() as { v: string }).v)
		) {
			const op = (next() as { v: string }).v;
			const right = parseMulDiv();
			const l = typeof left === "number" ? left : Number(left);
			const r = typeof right === "number" ? right : Number(right);
			left = op === "+" ? l + r : l - r;
		}
		return left;
	}

	function parseMulDiv(): number | string | boolean {
		let left = parseUnary();
		while (
			peek()?.type === "op" &&
			["*", "/"].includes((peek() as { v: string }).v)
		) {
			const op = (next() as { v: string }).v;
			const right = parseUnary();
			const l = typeof left === "number" ? left : Number(left);
			const r = typeof right === "number" ? right : Number(right);
			left = op === "*" ? l * r : r === 0 ? NaN : l / r;
		}
		return left;
	}

	function parseUnary(): number | string | boolean {
		if (peek()?.type === "op" && (peek() as { v: string }).v === "-") {
			next();
			const v = parsePower();
			return typeof v === "number" ? -v : -Number(v);
		}
		if (peek()?.type === "op" && (peek() as { v: string }).v === "+") {
			next();
			return parsePower();
		}
		return parsePower();
	}

	function parsePower(): number | string | boolean {
		let left = parsePostfix();
		if (peek()?.type === "op" && (peek() as { v: string }).v === "^") {
			next();
			const right = parseUnary();
			left = Math.pow(
				typeof left === "number" ? left : Number(left),
				typeof right === "number" ? right : Number(right),
			);
		}
		return left;
	}

	/** Handle postfix `%` operator: 50% → 0.5, A1% → A1/100 */
	function parsePostfix(): number | string | boolean {
		let left = parsePrimary();
		while (peek()?.type === "op" && (peek() as { v: string }).v === "%") {
			next();
			left = (typeof left === "number" ? left : Number(left)) / 100;
		}
		return left;
	}

	function parsePrimary(): number | string | boolean {
		const t = peek();
		if (!t) throw new Error("Unexpected end");

		switch (t.type) {
			case "num":
				next();
				return t.v;
			case "str":
				next();
				return t.v;
			case "bool":
				next();
				return t.v;

			case "cell": {
				next();
				if (t.ref === selfRef) return 0; // self-ref guard
				return cellValue(sheet, t.ref);
			}

			case "range": {
				next();
				const rng = utils.decode_range(t.ref);
				return cellValue(sheet, utils.encode_cell(rng.s));
			}

			case "func":
				return parseFunc();

			case "paren":
				if (t.v === "(") {
					next();
					const v = parseExpr();
					expect("paren", ")");
					return v;
				}
				throw new Error("Unexpected )");

			default:
				throw new Error("Unexpected token");
		}
	}

	// ── Function calls ────────────────────────────────────────────────────

	function parseFunc(): number | string | boolean {
		const name = (next() as { name: string }).name;
		expect("paren", "(");
		const args = parseFuncArgs();
		expect("paren", ")");
		return evalFunc(name, args);
	}

	function parseFuncArgs(): Arg[] {
		const args: Arg[] = [];
		if (peek()?.type === "paren" && (peek() as { v: string }).v === ")")
			return args;
		args.push(parseFuncArg());
		while (peek()?.type === "comma") {
			next();
			args.push(parseFuncArg());
		}
		return args;
	}

	function parseFuncArg(): Arg {
		if (peek()?.type === "range") {
			const t = next() as { ref: string };
			return resolveRange(sheet, utils, t.ref);
		}
		return parseExpr();
	}

	function flatNums(args: Arg[]): number[] {
		const out: number[] = [];
		for (const a of args) {
			if (Array.isArray(a)) {
				for (const v of a) {
					if (typeof v === "number") out.push(v);
					else if (typeof v === "boolean") out.push(v ? 1 : 0);
				}
			} else if (typeof a === "number") {
				out.push(a);
			} else if (typeof a === "boolean") {
				out.push(a ? 1 : 0);
			} else {
				const n = Number(a);
				if (!isNaN(n)) out.push(n);
			}
		}
		return out;
	}

	function evalFunc(name: string, args: Arg[]): number | string | boolean {
		switch (name) {
			case "SUM":
				return flatNums(args).reduce((a, b) => a + b, 0);

			case "AVERAGE": {
				const n = flatNums(args);
				return n.length > 0
					? n.reduce((a, b) => a + b, 0) / n.length
					: 0;
			}

			case "MIN": {
				const n = flatNums(args);
				return n.length > 0 ? Math.min(...n) : 0;
			}
			case "MAX": {
				const n = flatNums(args);
				return n.length > 0 ? Math.max(...n) : 0;
			}

			case "COUNT":
				return flatNums(args).length;
			case "COUNTA": {
				let c = 0;
				for (const a of args) c += Array.isArray(a) ? a.length : 1;
				return c;
			}

			case "ABS":
				return Math.abs(Number(args[0] ?? 0));
			case "INT":
				return Math.floor(Number(args[0] ?? 0));
			case "SQRT":
				return Math.sqrt(Number(args[0] ?? 0));
			case "POWER":
				return Math.pow(Number(args[0] ?? 0), Number(args[1] ?? 0));
			case "PI":
				return Math.PI;
			case "LOG": {
				const val = Number(args[0] ?? 1);
				const base = args[1] !== undefined ? Number(args[1]) : 10;
				return Math.log(val) / Math.log(base);
			}
			case "LN":
				return Math.log(Number(args[0] ?? 1));

			case "ROUND": {
				const v = Number(args[0] ?? 0),
					p = Number(args[1] ?? 0),
					f = Math.pow(10, p);
				return Math.round(v * f) / f;
			}
			case "ROUNDUP": {
				const v = Number(args[0] ?? 0),
					p = Number(args[1] ?? 0),
					f = Math.pow(10, p);
				return Math.ceil(v * f) / f;
			}
			case "ROUNDDOWN": {
				const v = Number(args[0] ?? 0),
					p = Number(args[1] ?? 0),
					f = Math.pow(10, p);
				return Math.floor(v * f) / f;
			}

			case "MOD": {
				const n = Number(args[0] ?? 0),
					d = Number(args[1] ?? 1);
				return d === 0 ? NaN : n % d;
			}

			case "IF": {
				const cond = args[0];
				const truthy =
					typeof cond === "number"
						? cond !== 0
						: typeof cond === "boolean"
							? cond
							: Boolean(cond);
				return truthy
					? ((args[1] ?? true) as number | string | boolean)
					: ((args[2] ?? false) as number | string | boolean);
			}

			case "AND": {
				for (const a of args) {
					if (Array.isArray(a)) {
						if (a.some((v) => !v)) return false;
					} else if (!a && a !== 0) return false;
				}
				return true;
			}

			case "OR": {
				for (const a of args) {
					if (Array.isArray(a)) {
						if (a.some((v) => !!v || v === 0)) return true;
					} else if (a || a === 0) return true;
				}
				return false;
			}

			case "NOT":
				return !args[0];

			case "LEN":
				return String(args[0] ?? "").length;
			case "UPPER":
				return String(args[0] ?? "").toUpperCase();
			case "LOWER":
				return String(args[0] ?? "").toLowerCase();
			case "TRIM":
				return String(args[0] ?? "").trim();
			case "CONCATENATE":
			case "CONCAT":
				return args
					.map((a) =>
						Array.isArray(a) ? a.map(String).join("") : String(a),
					)
					.join("");
			case "LEFT":
				return String(args[0] ?? "").slice(0, Number(args[1] ?? 1));
			case "RIGHT": {
				const s = String(args[0] ?? "");
				return s.slice(Math.max(0, s.length - Number(args[1] ?? 1)));
			}
			case "MID": {
				const s = String(args[0] ?? "");
				const start = Number(args[1] ?? 1) - 1;
				return s.slice(start, start + Number(args[2] ?? 1));
			}

			case "TODAY":
				return new Date().toISOString().split("T")[0]!;
			case "NOW":
				return new Date().toISOString();

			default:
				throw new Error(`Unknown function: ${name}`);
		}
	}

	// ── Run ───────────────────────────────────────────────────────────────
	return parseExpr();
}
