// How values look: number formats, conditional rules, and the effective style of a cell.

import { contains } from "./a1.ts";
import { compare, Err, formatGeneral, isErr, literal, toNumber, type Scalar } from "./formula/values.ts";
import { key, type CondOp, type CondStyle, type Format, type Sheet } from "./model.ts";

interface NumFmt {
  prefix: string;
  suffix: string;
  decimals: number;
  group: boolean;
  percent: boolean;
  sci: boolean;
}

const fmtCache = new Map<string, NumFmt | null>();

/** Parse a small, Sheets-compatible subset: [currency]#,##0.00[%] and 0.00E+00. "@" or "text" shows raw text. */
function parseNf(nf: string): NumFmt | null {
  let f = fmtCache.get(nf);
  if (f !== undefined) return f;
  const m = /^([^0#.,%]*)([#0,]*)(?:\.(0+))?(E\+0+)?(%?)([^0#.,%]*)$/.exec(nf);
  f = m && m[2]
    ? {
        prefix: m[1]!.replace(/"/g, ""),
        suffix: m[6]!.replace(/"/g, ""),
        decimals: m[3]?.length ?? 0,
        group: m[2].includes(","),
        percent: m[5] === "%",
        sci: !!m[4],
      }
    : null;
  fmtCache.set(nf, f);
  return f;
}

export function formatNumber(n: number, nf?: string): string {
  if (!nf || nf === "general") return formatGeneral(n);
  const f = parseNf(nf);
  if (!f) return formatGeneral(n);
  if (f.sci) return f.prefix + n.toExponential(f.decimals).replace(/e([+-])(\d)$/, "E$10$2").replace("e", "E") + f.suffix;
  const x = f.percent ? n * 100 : n;
  let s = Math.abs(x).toFixed(f.decimals);
  if (f.group) {
    const [int, frac] = s.split(".");
    s = int!.replace(/\B(?=(\d{3})+(?!\d))/g, ",") + (frac ? "." + frac : "");
  }
  const neg = x < 0 && Number(s.replace(/,/g, "")) !== 0;
  return (neg ? "-" : "") + f.prefix + s + (f.percent ? "%" : "") + f.suffix;
}

/** Display text for a computed value. */
export function displayValue(v: Scalar, nf?: string): string {
  if (v === null) return "";
  if (v instanceof Err) return v.code;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatNumber(v, nf);
  return v;
}

/** Does a value satisfy a condition (used by conditional fills and filters)? */
export function matches(op: CondOp, v: Scalar, a?: string, b?: string): boolean {
  if (op === "empty") return v === null || v === "";
  if (op === "notEmpty") return !(v === null || v === "");
  if (v === null || isErr(v)) return false;
  const x = literal(a ?? "");
  switch (op) {
    case "eq": return compare(v, x) === 0;
    case "ne": return compare(v, x) !== 0;
    case "contains": return displayValue(v).toLowerCase().includes(String(a ?? "").toLowerCase());
  }
  const n = toNumber(v);
  const lo = toNumber(x);
  if (typeof v !== "number" || isErr(n) || isErr(lo)) return false;
  switch (op) {
    case "gt": return n > lo;
    case "gte": return n >= lo;
    case "lt": return n < lo;
    case "lte": return n <= lo;
    case "between": {
      const hi = toNumber(literal(b ?? ""));
      return !isErr(hi) && n >= Math.min(lo, hi) && n <= Math.max(lo, hi);
    }
  }
  return false;
}

/** First matching conditional rule's style for a cell, if any. */
export function condStyle(sheet: Sheet, r: number, c: number, v: Scalar): CondStyle | null {
  for (const rule of sheet.cond) {
    if (contains(rule.range, r, c) && matches(rule.op, v, rule.value, rule.value2)) return rule.style;
  }
  return null;
}

/** Stored format with conditional styling layered on top. */
export function effectiveFormat(sheet: Sheet, r: number, c: number, v: Scalar): Format | undefined {
  const base = sheet.cells.get(key(r, c))?.f;
  if (!sheet.cond.length) return base;
  const cs = condStyle(sheet, r, c, v);
  return cs ? { ...base, ...cs } : base;
}
