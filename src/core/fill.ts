// Autofill: continue the pattern in a block of cells into the cells next to it,
// as dragging a spreadsheet's fill handle does. Pure; used by the browser and the CLI.
//
//   extendSeries(values, n)   the next n values of one line (numbers, "Item 7", Monday, cycles)
//   autofillOps(...)          ops that fill a target block from a source block, with formats

import type { Range } from "./a1.ts";
import { isFormula, moveRefs } from "./formula/rewrite.ts";
import { key, type CondRule, type Flag, type Format, type Sheet } from "./model.ts";
import type { Op } from "./ops.ts";

const NUMERIC = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
const LISTS = [
  ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
  ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
  ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
  ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
];

/** Constant difference between consecutive numbers, or null if there is none. */
function constantStep(xs: number[]): number | null {
  if (xs.length < 2) return null;
  const step = xs[1]! - xs[0]!;
  for (let i = 2; i < xs.length; i++) if (Math.abs(xs[i]! - xs[i - 1]! - step) > 1e-9 * Math.max(1, Math.abs(step))) return null;
  return step;
}

const decimals = (s: string) => (s.includes(".") ? s.split(".")[1]!.length : 0);

function sameCase(word: string, like: string): string {
  if (like.length > 1 && like === like.toUpperCase()) return word.toUpperCase();
  if (like === like.toLowerCase()) return word.toLowerCase();
  return word;
}

/**
 * The next `n` values after `values` (raw inputs, read in the fill direction).
 * `alt` is the Ctrl/Alt toggle: a lone number counts up instead of copying, and a series copies instead.
 * Returns null where the pattern is "repeat the source" (formulas, mixed content): the caller copies
 * the source cell at the same position in the cycle.
 */
export function extendSeries(values: Array<string | null>, n: number, alt = false): Array<string | null> {
  const cycle = () => Array.from({ length: n }, (_, k) => values[k % values.length] ?? null);
  if (!values.length || values.some((v) => v === null || v === "" || isFormula(v))) return cycle();
  const vs = values as string[];

  // Numbers: constant step continues; a single number copies (alt: counts by 1).
  if (vs.every((v) => NUMERIC.test(v.trim()))) {
    const xs = vs.map(Number);
    const step = xs.length === 1 ? (alt ? 1 : null) : alt ? null : constantStep(xs);
    if (step === null) return cycle();
    const d = Math.max(...vs.map(decimals), decimals(String(step)));
    const last = xs[xs.length - 1]!;
    return Array.from({ length: n }, (_, k) => {
      const x = last + step * (k + 1);
      return d ? x.toFixed(d) : String(Math.round(x));
    });
  }
  if (alt) return cycle();

  // Weekday and month names, full or short, keeping the case of the source.
  for (const list of LISTS) {
    const idx = vs.map((v) => list.findIndex((w) => w.toLowerCase() === v.trim().toLowerCase()));
    if (idx.some((i) => i < 0)) continue;
    const step = idx.length === 1 ? 1 : constantStep(idx);
    if (step === null) break;
    const len = list.length;
    const last = idx[idx.length - 1]!;
    const like = vs[vs.length - 1]!.trim();
    return Array.from({ length: n }, (_, k) => sameCase(list[(((last + step * (k + 1)) % len) + len) % len]!, like));
  }

  // Text with a number at the end ("Item 7", "run-01") or the start ("1st place").
  for (const re of [/^(.*?)(\d+)(\D*)$/, /^()(\d+)(\D.*)$/]) {
    const ms = vs.map((v) => re.exec(v));
    if (ms.some((m) => !m)) continue;
    const pre = ms[0]![1]!, post = ms[0]![3]!;
    if (ms.some((m) => m![1] !== pre || m![3] !== post)) continue;
    const xs = ms.map((m) => Number(m![2]));
    const step = xs.length === 1 ? 1 : constantStep(xs);
    if (step === null) break;
    const lastDigits = ms[ms.length - 1]![2]!;
    const pad = lastDigits.length > 1 && lastDigits.startsWith("0") ? lastDigits.length : 0;
    const last = xs[xs.length - 1]!;
    return Array.from({ length: n }, (_, k) => pre + String(Math.abs(last + step * (k + 1))).padStart(pad, "0") + post);
  }
  return cycle();
}

export type FillDir = "down" | "up" | "right" | "left";

/**
 * Ops that fill `target` from `source` on one sheet. `target` must sit directly beside `source`
 * (below, above, right or left) and span the same columns or rows. `rowAt`/`colAt` map display
 * positions to stored rows/cols (identity unless the sheet is sorted or has hidden rows);
 * both ranges are in display positions. One set op, one fmt op, and a cond op when a
 * conditional rule covering the source can grow to cover the target.
 */
export function autofillOps(
  sheet: Sheet, source: Range, target: Range, alt = false,
  rowAt: (i: number) => number = (i) => i, colAt: (j: number) => number = (j) => j,
): Op[] {
  const dir = fillDirection(source, target);
  if (!dir) throw new Error("the fill target must be next to the source, with the same rows or columns");
  const vertical = dir === "down" || dir === "up";
  const lines = vertical ? range(source.c1, source.c2) : range(source.r1, source.r2);
  // Source positions read away from the target's far side, i.e. in the direction of the fill.
  const along = vertical ? range(source.r1, source.r2) : range(source.c1, source.c2);
  const into = vertical ? range(target.r1, target.r2) : range(target.c1, target.c2);
  if (dir === "up" || dir === "left") {
    along.reverse();
    into.reverse();
  }
  const cells: Array<[number, number, string | null]> = [];
  const fmts: Array<[number, number, Format | null]> = [];
  const flags: Array<[number, number, Flag | null]> = [];
  let flagged = false;
  const at = (line: number, pos: number) => (vertical ? { r: rowAt(pos), c: colAt(line) } : { r: rowAt(line), c: colAt(pos) });
  for (const line of lines) {
    const src = along.map((pos) => at(line, pos));
    const raw = src.map((p) => sheet.cells.get(key(p.r, p.c))?.v ?? null);
    const next = extendSeries(raw, into.length, alt);
    into.forEach((pos, k) => {
      const t = at(line, pos);
      const from = src[k % src.length]!;
      let v = next[k] ?? null;
      if (v !== null && isFormula(v)) v = moveRefs(v, t.r - from.r, t.c - from.c);
      cells.push([t.r, t.c, v]);
      const cell = sheet.cells.get(key(from.r, from.c));
      fmts.push([t.r, t.c, cell?.f ?? null]);
      flags.push([t.r, t.c, cell?.flag ?? null]);
      if (cell?.flag) flagged = true;
    });
  }
  const ops: Op[] = [{ t: "set", s: sheet.id, cells }, { t: "fmt", s: sheet.id, cells: fmts }];
  if (flagged) ops.push({ t: "flag", s: sheet.id, cells: flags });
  const cond = growRules(sheet.cond, bbox(cells), bbox(along.flatMap((pos) => lines.map((line) => at(line, pos)))), vertical);
  if (cond) ops.push({ t: "cond", s: sheet.id, rules: cond });
  return ops;
}

export function fillDirection(source: Range, target: Range): FillDir | null {
  const sameCols = target.c1 === source.c1 && target.c2 === source.c2;
  const sameRows = target.r1 === source.r1 && target.r2 === source.r2;
  if (sameCols && target.r1 === source.r2 + 1) return "down";
  if (sameCols && target.r2 === source.r1 - 1) return "up";
  if (sameRows && target.c1 === source.c2 + 1) return "right";
  if (sameRows && target.c2 === source.c1 - 1) return "left";
  return null;
}

/**
 * A conditional rule that spans the source along the fill axis (and overlaps it across) grows
 * along that axis to cover the target, keeping its own width.
 */
function growRules(rules: CondRule[], target: Range, source: Range, vertical: boolean): CondRule[] | null {
  let changed = false;
  const out = rules.map((rule) => {
    const g = rule.range;
    const spans = vertical ? g.r1 <= source.r1 && g.r2 >= source.r2 && g.c1 <= source.c2 && g.c2 >= source.c1
      : g.c1 <= source.c1 && g.c2 >= source.c2 && g.r1 <= source.r2 && g.r2 >= source.r1;
    if (!spans) return rule;
    const u = vertical ? { ...g, r1: Math.min(g.r1, target.r1), r2: Math.max(g.r2, target.r2) }
      : { ...g, c1: Math.min(g.c1, target.c1), c2: Math.max(g.c2, target.c2) };
    if (u.r1 === g.r1 && u.r2 === g.r2 && u.c1 === g.c1 && u.c2 === g.c2) return rule;
    changed = true;
    return { ...rule, range: u };
  });
  return changed ? out : null;
}

function bbox(cells: Array<{ r: number; c: number } | [number, number, ...unknown[]]>): Range {
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  for (const x of cells) {
    const r = Array.isArray(x) ? x[0] : x.r, c = Array.isArray(x) ? x[1] : x.c;
    r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c);
  }
  return { r1, c1, r2, c2 };
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}
