// Ops: the only way a Doc changes. Each op is small (a few cells, one insert, one
// format patch) so concurrent writers touching different cells never clobber each other.
//
//   apply(doc, op)        mutates doc, returns the inverse ops (for undo) and what to recalc
//   transform(op, prior)  rewrites op made concurrently with `prior` so it applies after it

import type { Range } from "./a1.ts";
import { isFormula, renameSheetRefs, shiftRefs } from "./formula/rewrite.ts";
import {
  COND_OPS, FORMAT_KEYS, getSheet, key, keyCol, keyRow, OpError, sheetFromJSON, sheetToJSON,
  type CondRule, type Doc, type Format, type Sheet, type SheetJSON, type ViewState,
} from "./model.ts";

export type Axis = "row" | "col";
export type FormatPatch = { [K in keyof Format]?: Format[K] | null };

export type Op =
  | { t: "set"; s: string; cells: Array<[number, number, string | null]> }
  | { t: "fmt"; s: string; cells: Array<[number, number, Format | null]> }
  | { t: "style"; s: string; range: Range; patch: FormatPatch }
  | { t: "insert"; s: string; axis: Axis; at: number; n: number }
  | { t: "delete"; s: string; axis: Axis; at: number; n: number }
  | { t: "size"; s: string; axis: Axis; sizes: Array<[number, number | null]> }
  | { t: "hide"; s: string; axis: Axis; idx: number[]; hidden: boolean }
  | { t: "freeze"; s: string; rows: number; cols: number }
  | { t: "view"; s: string; view: ViewState | null }
  | { t: "cond"; s: string; rules: CondRule[] }
  | { t: "grid"; s: string; rows: number; cols: number }
  | { t: "addSheet"; sheet: SheetJSON; index?: number }
  | { t: "deleteSheet"; s: string }
  | { t: "renameSheet"; s: string; name: string }
  | { t: "moveSheet"; s: string; index: number }
  | { t: "rename"; name: string };

/** Cells whose computed values may have changed; "all" after structural edits. */
export type Dirty = "all" | Array<[string, number, number]>;

export interface Applied {
  inverse: Op[];
  dirty: Dirty;
}

const MAX_GRID_ROWS = 1_000_000;
const MAX_GRID_COLS = 16_384;

// ---- apply ----------------------------------------------------------------

export function apply(doc: Doc, op: Op): Applied {
  switch (op.t) {
    case "set": {
      const sh = getSheet(doc, op.s);
      const old: Array<[number, number, string | null]> = [];
      const dirty: Array<[string, number, number]> = [];
      for (const [r, c, v] of op.cells) {
        checkCell(r, c);
        const k = key(r, c);
        const cell = sh.cells.get(k);
        old.push([r, c, cell?.v ?? null]);
        dirty.push([sh.id, r, c]);
        if (v === null || v === "") {
          if (cell) {
            delete cell.v;
            if (!cell.f) sh.cells.delete(k);
          }
        } else {
          if (cell) cell.v = v;
          else sh.cells.set(k, { v });
          if (r >= sh.rows) sh.rows = r + 1;
          if (c >= sh.cols) sh.cols = c + 1;
        }
      }
      return { inverse: [{ t: "set", s: op.s, cells: old.reverse() }], dirty };
    }
    case "fmt": {
      const sh = getSheet(doc, op.s);
      const old: Array<[number, number, Format | null]> = [];
      for (const [r, c, f] of op.cells) {
        checkCell(r, c);
        old.push([r, c, sh.cells.get(key(r, c))?.f ?? null]);
        setFormat(sh, r, c, f);
      }
      return { inverse: [{ t: "fmt", s: op.s, cells: old.reverse() }], dirty: [] };
    }
    case "style": {
      const sh = getSheet(doc, op.s);
      const g = op.range;
      const old: Array<[number, number, Format | null]> = [];
      const r2 = Math.min(g.r2, sh.rows - 1), c2 = Math.min(g.c2, sh.cols - 1);
      for (let r = g.r1; r <= r2; r++) {
        for (let c = g.c1; c <= c2; c++) {
          const prev = sh.cells.get(key(r, c))?.f ?? null;
          old.push([r, c, prev]);
          const next: Format = { ...prev };
          for (const [k, v] of Object.entries(op.patch)) {
            if (v === null || v === undefined || v === false) delete next[k as keyof Format];
            else (next as Record<string, unknown>)[k] = v;
          }
          setFormat(sh, r, c, next);
        }
      }
      return { inverse: [{ t: "fmt", s: op.s, cells: old }], dirty: [] };
    }
    case "insert": {
      const sh = getSheet(doc, op.s);
      checkSpan(sh, op.axis, op.at, op.n, true);
      shift(doc, sh, op.axis, op.at, op.n);
      return { inverse: [{ t: "delete", s: op.s, axis: op.axis, at: op.at, n: op.n }], dirty: "all" };
    }
    case "delete": {
      const sh = getSheet(doc, op.s);
      checkSpan(sh, op.axis, op.at, op.n, false);
      const row = op.axis === "row";
      const end = op.at + op.n;
      const inSpan = (k: number) => {
        const i = row ? keyRow(k) : keyCol(k);
        return i >= op.at && i < end;
      };
      const vals: Array<[number, number, string | null]> = [];
      const fmts: Array<[number, number, Format | null]> = [];
      for (const [k, cell] of sh.cells) {
        if (!inSpan(k)) continue;
        if (cell.v !== undefined) vals.push([keyRow(k), keyCol(k), cell.v]);
        if (cell.f) fmts.push([keyRow(k), keyCol(k), cell.f]);
      }
      const sizesMap = row ? sh.rowH : sh.colW;
      const hiddenSet = row ? sh.hiddenRows : sh.hiddenCols;
      const sizes: Array<[number, number | null]> = [];
      const hidden: number[] = [];
      for (let i = op.at; i < end; i++) {
        if (sizesMap.has(i)) sizes.push([i, sizesMap.get(i)!]);
        if (hiddenSet.has(i)) hidden.push(i);
      }
      const restore: Op[] = [
        { t: "cond", s: sh.id, rules: structuredClone(sh.cond) },
        { t: "view", s: sh.id, view: sh.view ? structuredClone(sh.view) : null },
        { t: "freeze", s: sh.id, rows: sh.frozenRows, cols: sh.frozenCols },
      ];
      const changed = shift(doc, sh, op.axis, op.at, -op.n);
      const inverse: Op[] = [{ t: "insert", s: op.s, axis: op.axis, at: op.at, n: op.n }];
      if (vals.length) inverse.push({ t: "set", s: sh.id, cells: vals });
      if (fmts.length) inverse.push({ t: "fmt", s: sh.id, cells: fmts });
      for (const [sid, cells] of changed) inverse.push({ t: "set", s: sid, cells });
      if (sizes.length) inverse.push({ t: "size", s: sh.id, axis: op.axis, sizes });
      if (hidden.length) inverse.push({ t: "hide", s: sh.id, axis: op.axis, idx: hidden, hidden: true });
      inverse.push(...restore);
      return { inverse, dirty: "all" };
    }
    case "size": {
      const sh = getSheet(doc, op.s);
      const m = op.axis === "row" ? sh.rowH : sh.colW;
      const old: Array<[number, number | null]> = [];
      for (const [i, size] of op.sizes) {
        checkIndex(op.axis, i);
        old.push([i, m.get(i) ?? null]);
        if (size === null) m.delete(i);
        else {
          if (!(size >= 2 && size <= 2000)) throw new OpError(`size ${size} out of range 2..2000`);
          m.set(i, Math.round(size));
        }
      }
      return { inverse: [{ t: "size", s: op.s, axis: op.axis, sizes: old.reverse() }], dirty: [] };
    }
    case "hide": {
      const sh = getSheet(doc, op.s);
      const set = op.axis === "row" ? sh.hiddenRows : sh.hiddenCols;
      const flipped = op.idx.filter((i) => set.has(i) !== op.hidden);
      for (const i of flipped) {
        checkIndex(op.axis, i);
        if (op.hidden) set.add(i);
        else set.delete(i);
      }
      return { inverse: [{ t: "hide", s: op.s, axis: op.axis, idx: flipped, hidden: !op.hidden }], dirty: [] };
    }
    case "freeze": {
      const sh = getSheet(doc, op.s);
      if (!(op.rows >= 0 && op.rows <= 100 && op.cols >= 0 && op.cols <= 50)) throw new OpError("freeze at most 100 rows and 50 columns");
      const inverse: Op = { t: "freeze", s: op.s, rows: sh.frozenRows, cols: sh.frozenCols };
      sh.frozenRows = op.rows;
      sh.frozenCols = op.cols;
      return { inverse: [inverse], dirty: [] };
    }
    case "view": {
      const sh = getSheet(doc, op.s);
      const inverse: Op = { t: "view", s: op.s, view: sh.view };
      sh.view = op.view ? structuredClone(op.view) : null;
      return { inverse: [inverse], dirty: [] };
    }
    case "cond": {
      const sh = getSheet(doc, op.s);
      for (const rule of op.rules) {
        if (!COND_OPS.includes(rule.op)) throw new OpError(`unknown condition ${rule.op}`);
      }
      const inverse: Op = { t: "cond", s: op.s, rules: sh.cond };
      sh.cond = structuredClone(op.rules);
      return { inverse: [inverse], dirty: [] };
    }
    case "grid": {
      const sh = getSheet(doc, op.s);
      if (!(op.rows >= 1 && op.rows <= MAX_GRID_ROWS && op.cols >= 1 && op.cols <= MAX_GRID_COLS)) {
        throw new OpError("grid size out of range");
      }
      const inverse: Op = { t: "grid", s: op.s, rows: sh.rows, cols: sh.cols };
      sh.rows = op.rows;
      sh.cols = op.cols;
      return { inverse: [inverse], dirty: "all" };
    }
    case "addSheet": {
      const name = op.sheet.name.trim();
      if (!name) throw new OpError("sheet name is empty");
      if (doc.sheets.some((s) => s.id === op.sheet.id)) throw new OpError(`sheet id ${op.sheet.id} exists`);
      if (doc.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) {
        throw new OpError(`a sheet named "${name}" already exists`);
      }
      const sh = sheetFromJSON({ ...op.sheet, name });
      const index = Math.max(0, Math.min(op.index ?? doc.sheets.length, doc.sheets.length));
      doc.sheets.splice(index, 0, sh);
      return { inverse: [{ t: "deleteSheet", s: sh.id }], dirty: "all" };
    }
    case "deleteSheet": {
      const sh = getSheet(doc, op.s);
      if (doc.sheets.length === 1) throw new OpError("cannot delete the only sheet");
      const index = doc.sheets.indexOf(sh);
      doc.sheets.splice(index, 1);
      // Formulas keep their text and show #REF! until a sheet with that name exists again.
      return { inverse: [{ t: "addSheet", sheet: sheetToJSON(sh), index }], dirty: "all" };
    }
    case "renameSheet": {
      const sh = getSheet(doc, op.s);
      const name = op.name.trim();
      if (!name || name.includes("!")) throw new OpError("sheet names must be non-empty and contain no '!'");
      if (doc.sheets.some((s) => s !== sh && s.name.toLowerCase() === name.toLowerCase())) {
        throw new OpError(`a sheet named "${name}" already exists`);
      }
      const old = sh.name;
      for (const s of doc.sheets) {
        for (const cell of s.cells.values()) {
          if (isFormula(cell.v)) cell.v = renameSheetRefs(cell.v, old, name);
        }
      }
      sh.name = name;
      return { inverse: [{ t: "renameSheet", s: op.s, name: old }], dirty: "all" };
    }
    case "moveSheet": {
      const sh = getSheet(doc, op.s);
      const from = doc.sheets.indexOf(sh);
      const to = Math.max(0, Math.min(op.index, doc.sheets.length - 1));
      doc.sheets.splice(from, 1);
      doc.sheets.splice(to, 0, sh);
      return { inverse: [{ t: "moveSheet", s: op.s, index: from }], dirty: [] };
    }
    case "rename": {
      const name = op.name.trim();
      if (!name) throw new OpError("workbook name is empty");
      const old = doc.name;
      doc.name = name;
      return { inverse: [{ t: "rename", name: old }], dirty: [] };
    }
  }
}

/** Apply a list of ops atomically: on error, everything applied so far is undone. */
export function applyAll(doc: Doc, ops: Op[]): { inverse: Op[]; dirty: Dirty } {
  const inverses: Op[][] = [];
  let dirty: Dirty = [];
  try {
    for (const op of ops) {
      const res = apply(doc, op);
      inverses.push(res.inverse);
      dirty = mergeDirty(dirty, res.dirty);
    }
  } catch (e) {
    for (const inv of inverses.reverse()) for (const op of inv) apply(doc, op);
    throw e;
  }
  return { inverse: inverses.reverse().flat(), dirty };
}

export function mergeDirty(a: Dirty, b: Dirty): Dirty {
  if (a === "all" || b === "all") return "all";
  if (!a.length) return b;
  if (!b.length) return a;
  return a.concat(b);
}

function checkCell(r: number, c: number): void {
  if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0 || r >= MAX_GRID_ROWS || c >= MAX_GRID_COLS) {
    throw new OpError(`cell (${r},${c}) out of range`);
  }
}

function checkIndex(axis: Axis, i: number): void {
  const max = axis === "row" ? MAX_GRID_ROWS : MAX_GRID_COLS;
  if (!Number.isInteger(i) || i < 0 || i >= max) throw new OpError(`${axis} ${i} out of range`);
}

function checkSpan(sh: Sheet, axis: Axis, at: number, n: number, insert: boolean): void {
  const count = axis === "row" ? sh.rows : sh.cols;
  const max = axis === "row" ? MAX_GRID_ROWS : MAX_GRID_COLS;
  if (!Number.isInteger(at) || !Number.isInteger(n) || n < 1 || at < 0) throw new OpError("bad insert/delete span");
  if (insert ? at > count || count + n > max : at + n > count || n >= count) {
    throw new OpError(`cannot ${insert ? "insert" : "delete"} ${n} ${axis}s at ${at}`);
  }
}

function setFormat(sh: Sheet, r: number, c: number, f: Format | null): void {
  const k = key(r, c);
  const cell = sh.cells.get(k);
  const clean: Format = {};
  if (f) {
    for (const name of FORMAT_KEYS) {
      const v = f[name];
      if (v !== undefined && v !== null && v !== false && v !== "") (clean as Record<string, unknown>)[name] = v;
    }
  }
  const has = Object.keys(clean).length > 0;
  if (cell) {
    if (has) cell.f = clean;
    else delete cell.f;
    if (cell.v === undefined && !cell.f) sh.cells.delete(k);
  } else if (has) {
    sh.cells.set(k, { f: clean });
  }
}

// ---- index shifting (shared by apply and transform) ----------------------

/** New index after inserting (n>0) or deleting (n<0) at `at`; null if deleted. */
export function shiftIndex(i: number, at: number, n: number): number | null {
  if (n > 0) return i >= at ? i + n : i;
  const end = at - n;
  if (i >= at && i < end) return null;
  return i >= end ? i + n : i;
}

/** Shift one dimension [lo, hi] of a range; null if it is entirely deleted. */
function shiftSpan(lo: number, hi: number, at: number, n: number): [number, number] | null {
  if (n > 0) return [lo >= at ? lo + n : lo, hi >= at ? hi + n : hi];
  const end = at - n;
  if (lo >= at && hi < end) return null;
  const nlo = lo >= end ? lo + n : lo >= at ? at : lo;
  const nhi = hi >= end ? hi + n : hi >= at ? at - 1 : hi;
  return [nlo, nhi];
}

export function shiftRange(g: Range, axis: Axis, at: number, n: number): Range | null {
  if (axis === "row") {
    const s = shiftSpan(g.r1, g.r2, at, n);
    return s && { r1: s[0], r2: s[1], c1: g.c1, c2: g.c2 };
  }
  const s = shiftSpan(g.c1, g.c2, at, n);
  return s && { r1: g.r1, r2: g.r2, c1: s[0], c2: s[1] };
}

function shiftView(view: ViewState, axis: Axis, at: number, n: number): ViewState | null {
  if (axis === "row") {
    const h = shiftIndex(view.headerRow, at, n);
    return h === null ? null : { ...view, headerRow: h };
  }
  const filters: ViewState["filters"] = {};
  for (const [c, f] of Object.entries(view.filters)) {
    const nc = shiftIndex(Number(c), at, n);
    if (nc !== null) filters[nc] = f;
  }
  const out: ViewState = { headerRow: view.headerRow, filters };
  if (view.sort) {
    const sc = shiftIndex(view.sort.col, at, n);
    if (sc !== null) out.sort = { ...view.sort, col: sc };
  }
  return out;
}

/**
 * Insert (n>0) or delete (n<0) rows/cols in `sh`, rewriting formulas everywhere.
 * Returns the old text of every formula whose text changed, by sheet id, at pre-shift coordinates.
 */
function shift(doc: Doc, sh: Sheet, axis: Axis, at: number, n: number): Map<string, Array<[number, number, string]>> {
  const row = axis === "row";
  const changed = new Map<string, Array<[number, number, string]>>();
  for (const s of doc.sheets) {
    for (const [k, cell] of s.cells) {
      if (!isFormula(cell.v)) continue;
      const next = shiftRefs(cell.v, s.name, sh.name, axis, at, n);
      if (next === cell.v) continue;
      let list = changed.get(s.id);
      if (!list) changed.set(s.id, (list = []));
      list.push([keyRow(k), keyCol(k), cell.v]);
      cell.v = next;
    }
  }
  const cells: Sheet["cells"] = new Map();
  for (const [k, cell] of sh.cells) {
    const r = keyRow(k), c = keyCol(k);
    const i = shiftIndex(row ? r : c, at, n);
    if (i === null) continue;
    cells.set(row ? key(i, c) : key(r, i), cell);
  }
  sh.cells = cells;
  const remap = <T>(m: Map<number, T>) => {
    const out = new Map<number, T>();
    for (const [i, v] of m) {
      const j = shiftIndex(i, at, n);
      if (j !== null) out.set(j, v);
    }
    return out;
  };
  const reset = (s: Set<number>) => new Set([...remap(new Map([...s].map((i) => [i, true]))).keys()]);
  if (row) {
    sh.rowH = remap(sh.rowH);
    sh.hiddenRows = reset(sh.hiddenRows);
    sh.rows += n;
    if (at < sh.frozenRows) sh.frozenRows = Math.max(at, sh.frozenRows + n);
  } else {
    sh.colW = remap(sh.colW);
    sh.hiddenCols = reset(sh.hiddenCols);
    sh.cols += n;
    if (at < sh.frozenCols) sh.frozenCols = Math.max(at, sh.frozenCols + n);
  }
  sh.cond = sh.cond.flatMap((rule) => {
    const g = shiftRange(rule.range, axis, at, n);
    return g ? [{ ...rule, range: g }] : [];
  });
  if (sh.view) sh.view = shiftView(sh.view, axis, at, n);
  return changed;
}

// ---- transform --------------------------------------------------------------

/**
 * Rewrite `op` (made without seeing `prior`) so it means the same thing after `prior`
 * was applied. Only structural ops move coordinates; everything else passes through.
 * Returns null when the op no longer has a target (its cells or sheet were deleted).
 * Two inserts at the same index: the one the server ordered first stays put, so pass
 * priorFirst=false when moving an earlier-ordered op past a later one.
 */
export function transform(op: Op, prior: Op, doc?: Doc, priorFirst = true): Op | null {
  if (prior.t === "deleteSheet") return "s" in op && op.s === prior.s ? null : op;
  if (prior.t !== "insert" && prior.t !== "delete") return op;
  if (!("s" in op) || op.t === "renameSheet" || op.t === "moveSheet" || op.t === "deleteSheet") return op;
  const n = prior.t === "insert" ? prior.n : -prior.n;
  const { axis, at } = prior;
  const row = axis === "row";
  const sameSheet = op.s === prior.s;
  const sheetName = doc?.sheets.find((s) => s.id === prior.s)?.name;
  const homeName = doc?.sheets.find((s) => s.id === op.s)?.name;
  switch (op.t) {
    case "set": {
      const cells: Array<[number, number, string | null]> = [];
      for (const [r, c, v] of op.cells) {
        let nr = r, nc = c;
        if (sameSheet) {
          const i = shiftIndex(row ? r : c, at, n);
          if (i === null) continue;
          if (row) nr = i;
          else nc = i;
        }
        const text = isFormula(v ?? undefined) && sheetName && homeName ? shiftRefs(v!, homeName, sheetName, axis, at, n) : v;
        cells.push([nr, nc, text]);
      }
      return cells.length ? { ...op, cells } : null;
    }
    case "fmt": {
      if (!sameSheet) return op;
      const cells: Array<[number, number, Format | null]> = [];
      for (const [r, c, f] of op.cells) {
        const i = shiftIndex(row ? r : c, at, n);
        if (i !== null) cells.push(row ? [i, c, f] : [r, i, f]);
      }
      return cells.length ? { ...op, cells } : null;
    }
    case "style": {
      if (!sameSheet) return op;
      const g = shiftRange(op.range, axis, at, n);
      return g ? { ...op, range: g } : null;
    }
    case "insert":
    case "delete": {
      if (!sameSheet || op.axis !== axis) return op;
      if (op.t === "insert") {
        const i = op.at > at || (op.at === at && n > 0 && priorFirst) ? shiftIndex(op.at, at, n) : op.at;
        return { ...op, at: i ?? at };
      }
      const g = shiftSpan(op.at, op.at + op.n - 1, at, n);
      return g ? { ...op, at: g[0], n: g[1] - g[0] + 1 } : null;
    }
    case "size": {
      if (!sameSheet || op.axis !== axis) return op;
      const sizes = op.sizes.flatMap(([i, s]) => {
        const j = shiftIndex(i, at, n);
        return j === null ? [] : [[j, s] as [number, number | null]];
      });
      return sizes.length ? { ...op, sizes } : null;
    }
    case "hide": {
      if (!sameSheet || op.axis !== axis) return op;
      const idx = op.idx.flatMap((i) => {
        const j = shiftIndex(i, at, n);
        return j === null ? [] : [j];
      });
      return idx.length ? { ...op, idx } : null;
    }
    case "cond": {
      if (!sameSheet) return op;
      return {
        ...op,
        rules: op.rules.flatMap((rule) => {
          const g = shiftRange(rule.range, axis, at, n);
          return g ? [{ ...rule, range: g }] : [];
        }),
      };
    }
    case "view":
      return sameSheet && op.view ? { ...op, view: shiftView(op.view, axis, at, n) } : op;
    default:
      return op;
  }
}

export function transformAll(ops: Op[], prior: Op[], doc?: Doc, priorFirst = true): Op[] {
  let out = ops;
  for (const p of prior) {
    out = out.flatMap((op) => {
      const t = transform(op, p, doc, priorFirst);
      return t ? [t] : [];
    });
  }
  return out;
}

/** Does this op change anything a user would want highlighted? Returns the touched cells. */
export function touchedCells(op: Op): { s: string; cells: Array<[number, number]> } | null {
  if (op.t === "set" || op.t === "fmt") return { s: op.s, cells: op.cells.map(([r, c]) => [r, c]) };
  if (op.t === "style") {
    const out: Array<[number, number]> = [];
    const g = op.range;
    for (let r = g.r1; r <= Math.min(g.r2, g.r1 + 999); r++) {
      for (let c = g.c1; c <= Math.min(g.c2, g.c1 + 99); c++) out.push([r, c]);
    }
    return { s: op.s, cells: out };
  }
  return null;
}
