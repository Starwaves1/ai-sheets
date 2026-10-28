// Rewrite references inside formula text. Used when rows/cols are inserted or
// deleted, when formulas are copied to a new place, and when sheets are renamed or deleted.

import { colName, MAX_COLS, MAX_ROWS, quoteSheet } from "../a1.ts";
import { tokenize, type RefEnd, type Tok } from "./lexer.ts";

type RefTok = Extract<Tok, { t: "ref" }>;

/** Returns new text for a reference, "#REF!" to invalidate it, or null to keep it. */
type RefMapper = (ref: RefTok) => string | null;

export function isFormula(input: string | undefined): input is string {
  return input !== undefined && input.length > 1 && input[0] === "=";
}

function rewrite(formula: string, map: RefMapper): string {
  let toks: Tok[];
  try {
    toks = tokenize(formula, 1);
  } catch {
    return formula; // unparseable formulas are left alone; they evaluate to an error anyway
  }
  let out = "";
  let last = 0;
  for (const t of toks) {
    if (t.t !== "ref") continue;
    const rep = map(t);
    if (rep === null) continue;
    out += formula.slice(last, t.s) + rep;
    last = t.e;
  }
  return last === 0 ? formula : out + formula.slice(last);
}

function endText(e: RefEnd): string {
  let s = "";
  if (e.c !== undefined) s += (e.ca ? "$" : "") + colName(e.c);
  if (e.r !== undefined) s += (e.ra ? "$" : "") + (e.r + 1);
  return s;
}

function refText(sheet: string | undefined, a: RefEnd, b?: RefEnd): string {
  const prefix = sheet === undefined ? "" : quoteSheet(sheet) + "!";
  return prefix + endText(a) + (b ? ":" + endText(b) : "");
}

const targets = (ref: RefTok, homeSheet: string, sheet: string) =>
  (ref.sheet ?? homeSheet).toLowerCase() === sheet.toLowerCase();

/**
 * Adjust references to `sheet` after inserting (n > 0) or deleting (n < 0) |n|
 * rows or columns at index `at`. `homeSheet` is the sheet the formula lives on.
 */
export function shiftRefs(
  formula: string, homeSheet: string, sheet: string, axis: "row" | "col", at: number, n: number,
): string {
  const k = axis === "row" ? "r" : "c";
  return rewrite(formula, (ref) => {
    if (!targets(ref, homeSheet, sheet)) return null;
    const a = { ...ref.a };
    const b = ref.b ? { ...ref.b } : undefined;
    if (a[k] === undefined) return null; // whole-row ref unaffected by column shifts, and vice versa
    if (n > 0) {
      if (a[k]! >= at) a[k]! += n;
      if (b && b[k]! >= at) b[k]! += n;
    } else {
      const del = -n;
      const end = at + del; // first index after the deleted block
      if (!b) {
        if (a[k]! >= at && a[k]! < end) return "#REF!";
        if (a[k]! >= end) a[k]! -= del;
      } else {
        let lo = a[k]!, hi = b[k]!;
        const swap = lo > hi;
        if (swap) [lo, hi] = [hi, lo];
        if (lo >= at && hi < end) return "#REF!";
        const nlo = lo >= end ? lo - del : lo >= at ? at : lo;
        const nhi = hi >= end ? hi - del : hi >= at ? at - 1 : hi;
        if (swap) {
          a[k] = nhi;
          b[k] = nlo;
        } else {
          a[k] = nlo;
          b[k] = nhi;
        }
      }
    }
    return refText(ref.sheet, a, b);
  });
}

/** Offset relative references by (dr, dc), as when copying a formula. Out-of-sheet refs become #REF!. */
export function moveRefs(formula: string, dr: number, dc: number): string {
  if (dr === 0 && dc === 0) return formula;
  return rewrite(formula, (ref) => {
    const shift = (e: RefEnd): RefEnd | null => {
      const o = { ...e };
      if (o.r !== undefined && !o.ra) o.r += dr;
      if (o.c !== undefined && !o.ca) o.c += dc;
      if (o.r !== undefined && (o.r < 0 || o.r >= MAX_ROWS)) return null;
      if (o.c !== undefined && (o.c < 0 || o.c >= MAX_COLS)) return null;
      return o;
    };
    const a = shift(ref.a);
    const b = ref.b ? shift(ref.b) : undefined;
    if (!a || b === null) return "#REF!";
    return refText(ref.sheet, a, b);
  });
}

/** Point explicit references from sheet `from` to `to`; to = null invalidates them (sheet deleted). */
export function renameSheetRefs(formula: string, from: string, to: string | null): string {
  return rewrite(formula, (ref) => {
    if (ref.sheet === undefined || ref.sheet.toLowerCase() !== from.toLowerCase()) return null;
    return to === null ? "#REF!" : refText(to, ref.a, ref.b);
  });
}
