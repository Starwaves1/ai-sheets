// Version history, pure and shared by server, CLI and browser.
//
//   summarize / describe   what a change did, in plain words ("Claude set 12 cells in S2:BP19")
//   groupSessions          changes by the same person close in time, even when others edit in between
//   rangeBefore            where a range was before a change (to filter history by cells)
//   revertOps              undo chosen changes as new ops, keeping later edits by others
//   restoreOps             ops that turn the current workbook back into an earlier one
//   changedCells           cells that differ between two versions, matched across row/col moves

import { cellName, colName, rangeName, type Range } from "./a1.ts";
import { docFromJSON, docToJSON, key, keyCol, keyRow, sheetToJSON, type Doc, type Flag, type Format, type Sheet } from "./model.ts";
import { applyAll, moveHull, movedStart, shiftRange, touchedCells, transform, transformAll, type Op } from "./ops.ts";

/** Why a change was made, when it was not a plain edit. Shown instead of the ops. */
export type Meta =
  | { kind: "revert"; of: number[]; actors: string[]; since: number } // since: time of the oldest reverted change
  | { kind: "restore"; to: number; at: number; sheet?: string }
  | { kind: "cell"; ref: string; at: number }; // one cell set back to an earlier value

export interface Change {
  v: number;
  ops: Op[];
  actor: string;
  client: string;
  ts: number;
  seq?: number;
  inv?: Op[] | null; // inverse ops (undo), recorded when the change was applied
  meta?: Meta | null;
}

// ---- summaries ---------------------------------------------------------------------

/** What one change touched, small enough to send for thousands of changes. */
export interface Summary {
  v: number;
  actor: string;
  ts: number;
  meta?: Meta;
  /** Cell edits per sheet: bounds, count, a few keys for de-duplication, the value when it was one cell. */
  cells: Array<{ s: string; kind: "set" | "fmt" | "flag"; range: Range; n: number; keys?: number[]; value?: string | null; note?: string; done?: boolean }>;
  /** Everything else, already in words ("inserted 2 rows at 5"). */
  other: string[];
  undoable: boolean; // false for changes recorded before inverses were kept
}

const plural = (n: number, w: string) => `${n} ${w}${n === 1 ? "" : "s"}`;
const span = (axis: "row" | "col", at: number, n: number) =>
  axis === "row" ? (n === 1 ? `row ${at + 1}` : `rows ${at + 1}–${at + n}`) : n === 1 ? `column ${colName(at)}` : `columns ${colName(at)}–${colName(at + n - 1)}`;

export function summarize(c: Change, doc: Doc): Summary {
  const name = (id: string) => doc.sheets.find((s) => s.id === id)?.name ?? "a deleted sheet";
  const header = (sid: string, col: number, row: number) => {
    const s = doc.sheets.find((x) => x.id === sid);
    const v = s?.cells.get(key(row, col))?.v;
    return v && !v.startsWith("=") ? v : colName(col);
  };
  const out: Summary = { v: c.v, actor: c.actor, ts: c.ts, cells: [], other: [], undoable: !!c.inv };
  if (c.meta) out.meta = c.meta;
  for (const op of c.ops) {
    const t = touchedCells(op);
    if (t && (op.t === "set" || op.t === "fmt" || op.t === "flag" || op.t === "style")) {
      if (!t.cells.length) continue;
      const kind = op.t === "style" ? "fmt" : op.t;
      let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
      for (const [r, cc] of t.cells) {
        r1 = Math.min(r1, r); c1 = Math.min(c1, cc); r2 = Math.max(r2, r); c2 = Math.max(c2, cc);
      }
      const range = op.t === "style" ? op.range : { r1, c1, r2, c2 };
      const n = op.t === "style" ? (range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1) : t.cells.length;
      const entry: Summary["cells"][number] = { s: op.s, kind, range, n };
      if (n <= 200 && op.t !== "style") entry.keys = t.cells.map(([r, cc]) => key(r, cc));
      if (op.t === "set" && op.cells.length === 1) entry.value = op.cells[0]![2];
      if (op.t === "flag") {
        const f = op.cells[0]![2];
        if (f) {
          entry.note = f.note;
          if (f.done) entry.done = true;
        }
      }
      out.cells.push(entry);
      continue;
    }
    switch (op.t) {
      case "insert": out.other.push(`inserted ${plural(op.n, op.axis === "row" ? "row" : "column")} at ${op.axis === "row" ? op.at + 1 : colName(op.at)}`); break;
      case "delete": out.other.push(`deleted ${span(op.axis, op.at, op.n)}`); break;
      case "move": {
        const start = movedStart(op.from, op.n, op.to);
        out.other.push(`moved ${span(op.axis, op.from, op.n)} to ${op.axis === "row" ? `row ${start + 1}` : `column ${colName(start)}`}`);
        break;
      }
      case "size": out.other.push(`resized ${plural(op.sizes.length, op.axis === "row" ? "row" : "column")}`); break;
      case "hide": out.other.push(`${op.hidden ? "hid" : "showed"} ${plural(op.idx.length, op.axis === "row" ? "row" : "column")}`); break;
      case "freeze": out.other.push(op.rows || op.cols ? `froze ${[op.rows ? plural(op.rows, "row") : "", op.cols ? plural(op.cols, "column") : ""].filter(Boolean).join(" and ")}` : "unfroze panes"); break;
      case "view": {
        if (!op.view) out.other.push("cleared sort and filters");
        else {
          if (op.view.sort) out.other.push(`sorted by ${header(op.s, op.view.sort.col, op.view.headerRow)}${op.view.sort.desc ? " (Z→A)" : ""}`);
          const f = Object.keys(op.view.filters).map((x) => header(op.s, Number(x), op.view!.headerRow));
          if (f.length) out.other.push(`filtered by ${f.join(", ")}`);
          if (!op.view.sort && !f.length) out.other.push("turned on filters");
        }
        break;
      }
      case "cond": out.other.push("changed conditional fill"); break;
      case "grid": out.other.push(`resized ${name(op.s)} to ${op.rows}×${op.cols}`); break;
      case "addSheet": out.other.push(`added sheet ${op.sheet.name}`); break;
      case "deleteSheet": out.other.push(`deleted sheet ${name(op.s)}`); break;
      case "renameSheet": out.other.push(`renamed a sheet to ${op.name}`); break;
      case "moveSheet": out.other.push(`moved sheet ${name(op.s)}`); break;
      case "rename": out.other.push(`renamed the workbook to ${op.name}`); break;
    }
  }
  return out;
}

export type TimeFormat = (ts: number) => string;

export const clockTime: TimeFormat = (ts) => {
  const d = new Date(ts);
  const now = new Date();
  const time = d.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
  return d.toDateString() === now.toDateString() ? time : `${d.toLocaleDateString("en-US", { month: "short", day: "numeric" })}, ${time}`;
};

/**
 * One line for a group of changes by one person, without the name: "set 12 cells in S2:BP19,
 * sorted by pass_rate". `sheetName` adds "Runs!" prefixes when the workbook has several sheets.
 */
export function describe(sums: Summary[], sheetName: (id: string) => string | null, time: TimeFormat = clockTime): string {
  const metas = sums.filter((s) => s.meta);
  if (metas.length === sums.length && metas.length === 1) {
    const m = metas[0]!.meta!;
    if (m.kind === "restore") return `restored ${m.sheet ? m.sheet : "the workbook"} to ${time(m.at)}`;
    if (m.kind === "cell") return `set ${m.ref} back to its value from ${time(m.at)}`;
    return `reverted ${plural(m.of.length, "change")} by ${m.actors.join(", ")} from ${time(m.since)}`;
  }
  const phrases: string[] = [];
  const where = (s: string, g: Range) => `${sheetName(s) ? `${sheetName(s)}!` : ""}${rangeName(g)}`;
  for (const kind of ["set", "fmt", "flag"] as const) {
    const bySheet = new Map<string, { range: Range; n: number; keys: Set<number> | null; value?: string | null; note?: string; done?: boolean; count: number }>();
    for (const s of sums) {
      for (const c of s.cells) {
        if (c.kind !== kind) continue;
        let g = bySheet.get(c.s);
        if (!g) bySheet.set(c.s, (g = { range: { ...c.range }, n: 0, keys: new Set(), count: 0 }));
        else g.range = { r1: Math.min(g.range.r1, c.range.r1), c1: Math.min(g.range.c1, c.range.c1), r2: Math.max(g.range.r2, c.range.r2), c2: Math.max(g.range.c2, c.range.c2) };
        g.n += c.n;
        g.count++;
        if (c.keys && g.keys) for (const k of c.keys) g.keys.add(k);
        else g.keys = null;
        g.value = c.value;
        g.note = c.note;
        g.done = c.done;
      }
    }
    for (const [sid, g] of bySheet) {
      const n = g.keys ? g.keys.size : g.n;
      const one = n === 1 && g.range.r1 === g.range.r2 && g.range.c1 === g.range.c2;
      const at = where(sid, g.range);
      if (kind === "set") {
        if (one && g.count === 1 && g.value !== undefined) phrases.push(g.value === null ? `cleared ${at}` : `set ${at} to ${JSON.stringify(g.value.length > 40 ? g.value.slice(0, 39) + "…" : g.value)}`);
        else phrases.push(one ? `edited ${at}${g.count > 1 ? ` ${g.count} times` : ""}` : `set ${plural(n, "cell")} in ${at}`);
      } else if (kind === "fmt") phrases.push(`formatted ${at}`);
      else phrases.push(g.note === undefined ? `removed a flag on ${at}` : g.done ? `resolved the flag on ${at}` : `flagged ${at}: ${JSON.stringify(g.note.length > 50 ? g.note.slice(0, 49) + "…" : g.note)}`);
    }
  }
  // Repeated resizes and hides read as one: "resized 1 column" + "resized 6 columns" = "resized 7 columns".
  const counted = new Map<string, number>();
  for (const s of sums) {
    for (const o of s.other) {
      const m = /^(resized|hid|showed) (\d+) (row|column)s?$/.exec(o);
      if (m) counted.set(`${m[1]} ${m[3]}`, (counted.get(`${m[1]} ${m[3]}`) ?? 0) + Number(m[2]));
      else if (!phrases.includes(o)) phrases.push(o);
    }
  }
  for (const [k, n] of counted) {
    const [verb, w] = k.split(" ");
    phrases.push(`${verb} ${plural(n, w!)}`);
  }
  for (const s of metas) phrases.push(describe([s], sheetName, time));
  if (!phrases.length) return "made no visible change";
  return phrases.length > 3 ? `${phrases.slice(0, 3).join(", ")} and ${plural(phrases.length - 3, "more change")}` : phrases.join(", ");
}

// ---- sessions ----------------------------------------------------------------------

export interface Session {
  id: string; // "s<last version>"
  actor: string;
  versions: number[]; // ascending
  first: number;
  last: number;
  start: number; // ts
  end: number;
  sums: Summary[];
}

/**
 * Group changes into sessions: one person's edits with less than `gap` between them. Other people's
 * edits in between do not split a session. Reverts and restores stand alone. Newest session first.
 */
export function groupSessions(sums: Summary[], gap = 10 * 60_000): Session[] {
  const out: Session[] = [];
  const open = new Map<string, Session>();
  for (const s of [...sums].sort((a, b) => a.v - b.v)) {
    let cur = s.meta ? undefined : open.get(s.actor);
    if (!cur || s.ts - cur.end > gap) {
      cur = { id: "", actor: s.actor, versions: [], first: s.v, last: s.v, start: s.ts, end: s.ts, sums: [] };
      out.push(cur);
      if (s.meta) open.delete(s.actor);
      else open.set(s.actor, cur);
    }
    cur.versions.push(s.v);
    cur.sums.push(s);
    cur.last = s.v;
    cur.end = s.ts;
  }
  for (const s of out) s.id = `s${s.last}`;
  return out.sort((a, b) => b.last - a.last);
}

// ---- following cells through structural changes -------------------------------------

/** Where `range` on `sheet` was before `ops` ran (null: none of it existed yet). */
export function rangeBefore(range: Range, sheet: string, ops: Op[]): Range | null {
  let g: Range | null = range;
  for (let i = ops.length - 1; i >= 0 && g; i--) {
    const op = ops[i]!;
    if (op.t === "addSheet" && op.sheet.id === sheet) return null;
    if (!("s" in op) || op.s !== sheet) continue;
    if (op.t === "insert") g = shiftRange(g, op.axis, op.at, -op.n);
    else if (op.t === "delete") g = shiftRange(g, op.axis, op.at, op.n);
    else if (op.t === "move") {
      const start = movedStart(op.from, op.n, op.to);
      const back = start > op.from ? op.from : op.from + op.n;
      const [lo, hi] = moveHull(op.axis === "row" ? g.r1 : g.c1, op.axis === "row" ? g.r2 : g.c2, start, op.n, back);
      g = op.axis === "row" ? { ...g, r1: lo, r2: hi } : { ...g, c1: lo, c2: hi };
    }
  }
  return g;
}

const inRange = (g: Range, r: number, c: number) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2;

/** Does change `c` touch cells of `range` (given in the frame right after `c`)? */
export function touches(c: Change, sheet: string, range: Range): boolean {
  let g: Range | null = range;
  for (let i = c.ops.length - 1; i >= 0 && g; i--) {
    const op = c.ops[i]!;
    if ("s" in op && op.s === sheet) {
      if (op.t === "style" && !(op.range.r2 < g.r1 || op.range.r1 > g.r2 || op.range.c2 < g.c1 || op.range.c1 > g.c2)) return true;
      if ((op.t === "set" || op.t === "fmt" || op.t === "flag") && op.cells.some(([r, cc]) => inRange(g!, r, cc))) return true;
      if ((op.t === "insert" || op.t === "delete" || op.t === "move")) {
        const lo = op.axis === "row" ? g.r1 : g.c1, hi = op.axis === "row" ? g.r2 : g.c2;
        const a = op.t === "move" ? op.from : op.at;
        if (a <= hi && a + op.n - 1 >= lo) return true;
      }
    }
    g = rangeBefore(g, sheet, [op]);
  }
  return false;
}

/** Map cells (sheet, row, col) through later ops into the latest frame; dropped cells are deleted. */
export function carryCells(cells: Array<{ s: string; r: number; c: number }>, later: Op[], doc?: Doc): Array<{ s: string; r: number; c: number } | null> {
  const bySheet = new Map<string, number[]>();
  cells.forEach((x, i) => {
    const list = bySheet.get(x.s) ?? [];
    list.push(i);
    bySheet.set(x.s, list);
  });
  const out: Array<{ s: string; r: number; c: number } | null> = cells.map(() => null);
  for (const [s, idx] of bySheet) {
    // Carry a probe op whose third field remembers which input each cell was.
    let probe: Op | null = { t: "flag", s, cells: idx.map((i) => [cells[i]!.r, cells[i]!.c, { note: String(i), by: "", ts: 0 }]) };
    for (const p of later) {
      if (!probe) break;
      probe = transform(probe, p, doc);
    }
    if (probe?.t === "flag") for (const [r, c, f] of probe.cells) out[Number(f!.note)] = { s, r, c };
  }
  return out;
}

// ---- selective revert ------------------------------------------------------------------

export interface Skipped {
  what: string; // "B4, C5" or "conditional fill"
  sheet: string; // sheet id ("" for workbook-level)
  by: string; // who changed it later
}

/** Keys an op writes, for conflict checks: cell values, formats and flags, and whole settings. */
function footprint(op: Op): { cells: Map<string, Set<number>>; ranges: Array<{ k: string; g: Range }>; whole: Set<string> } {
  const cells = new Map<string, Set<number>>(), ranges: Array<{ k: string; g: Range }> = [], whole = new Set<string>();
  const add = (k: string, r: number, c: number) => {
    let set = cells.get(k);
    if (!set) cells.set(k, (set = new Set()));
    set.add(key(r, c));
  };
  switch (op.t) {
    case "set": for (const [r, c] of op.cells) add(`v:${op.s}`, r, c); break;
    case "fmt": for (const [r, c] of op.cells) add(`f:${op.s}`, r, c); break;
    case "flag": for (const [r, c] of op.cells) add(`g:${op.s}`, r, c); break;
    case "style": ranges.push({ k: `f:${op.s}`, g: op.range }); break;
    case "size": for (const [i] of op.sizes) whole.add(`size:${op.s}:${op.axis}:${i}`); break;
    case "hide": for (const i of op.idx) whole.add(`hide:${op.s}:${op.axis}:${i}`); break;
    case "cond": case "view": case "freeze": case "grid": whole.add(`${op.t}:${op.s}`); break;
    case "renameSheet": case "moveSheet": whole.add(`${op.t}:${op.s}`); break;
    case "rename": whole.add("rename"); break;
  }
  return { cells, ranges, whole };
}

/**
 * Rebase the inverse sequence `inv` past the later op `p0`; with `carveOut`, set aside what `p0`
 * overwrote. Each op of `inv` is in the frame its predecessors leave (insert the row, then refill
 * it), so `p0` is carried through the sequence and compared with each op in that op's own frame.
 */
function pastLater(inv: Op[], p0: Op, doc: Doc, carveOut: boolean): { keep: Op[]; skip: Op[]; gone: Array<{ s: string; what: string }> } {
  const keep: Op[] = [], skip: Op[] = [];
  const gone: Array<{ s: string; what: string }> = [];
  let p: Op | null = p0;
  for (const op of inv) {
    if (!p) {
      keep.push(op);
      continue;
    }
    const parts = carveOut ? carveOne(op, p) : { keep: [op], skip: [] };
    for (const x of parts.keep) {
      const t = transform(x, p, doc);
      if (t) keep.push(t);
      if (carveOut) gone.push(...lostTo(x, p));
    }
    for (const x of parts.skip) {
      const t = transform(x, p, doc);
      if (t) skip.push(t);
    }
    p = transform(p, op, doc, false);
  }
  return { keep, skip, gone };
}

/** Cells of `x` (set/fmt/flag) that the delete `p` removed, so their revert has nowhere to go. */
function lostTo(x: Op, p: Op): Array<{ s: string; what: string }> {
  if ((x.t !== "set" && x.t !== "fmt" && x.t !== "flag") || !("s" in p) || p.s !== x.s) return [];
  if (p.t === "deleteSheet") return [{ s: x.s, what: "(sheet deleted since)" }];
  if (p.t !== "delete") return [];
  const lost = x.cells.filter(([r, c]) => {
    const i = p.axis === "row" ? r : c;
    return i >= p.at && i < p.at + p.n;
  });
  return lost.map(([r, c]) => ({ s: x.s, what: `${cellName(r, c)} (deleted since)` }));
}

/**
 * Split one inverse op into what can still be applied and what `p` (in the same frame) overwrote.
 * Cell lists split cell by cell; whole settings are kept or set aside as a unit.
 */
function carveOne(op: Op, p: Op): { keep: Op[]; skip: Op[] } {
  const fp = footprint(p);
  const hit = (k: string, r: number, c: number) => !!fp.cells.get(k)?.has(key(r, c)) || fp.ranges.some((x) => x.k === k && inRange(x.g, r, c));
  const split = (yes: Op | null, no: Op | null) => ({ keep: no ? [no] : [], skip: yes ? [yes] : [] });
  if (op.t === "set" || op.t === "fmt" || op.t === "flag") {
    const k = `${op.t === "set" ? "v" : op.t === "fmt" ? "f" : "g"}:${op.s}`;
    const yes = op.cells.filter(([r, c]) => hit(k, r, c)), no = op.cells.filter(([r, c]) => !hit(k, r, c));
    return split(yes.length ? ({ ...op, cells: yes } as Op) : null, no.length ? ({ ...op, cells: no } as Op) : null);
  }
  if (op.t === "size") {
    const bad = (i: number) => fp.whole.has(`size:${op.s}:${op.axis}:${i}`);
    const yes = op.sizes.filter(([i]) => bad(i)), no = op.sizes.filter(([i]) => !bad(i));
    return split(yes.length ? { ...op, sizes: yes } : null, no.length ? { ...op, sizes: no } : null);
  }
  if (op.t === "hide") {
    const bad = (i: number) => fp.whole.has(`hide:${op.s}:${op.axis}:${i}`);
    const yes = op.idx.filter(bad), no = op.idx.filter((i) => !bad(i));
    return split(yes.length ? { ...op, idx: yes } : null, no.length ? { ...op, idx: no } : null);
  }
  if (op.t === "delete") {
    // Undoing an insert: keep the rows/cols if someone has since written into them.
    const written = [...fp.cells.entries()].some(([k, set]) => k.endsWith(`:${op.s}`) && [...set].some((x) => {
      const i = op.axis === "row" ? keyRow(x) : keyCol(x);
      return i >= op.at && i < op.at + op.n;
    }));
    return written ? split(op, null) : split(null, op);
  }
  if (op.t === "deleteSheet") return "s" in p && p.s === op.s ? split(op, null) : split(null, op);
  const mine = footprint(op).whole;
  return [...mine].some((k) => fp.whole.has(k)) ? split(op, null) : split(null, op);
}

const moveOver = (ops: Op[], prior: Op, doc?: Doc): Op[] => transformAll(ops, [prior], doc);

/**
 * Ops that undo the `targets` changes as a new change, applied after everything in `log`
 * (`log` holds every change from the oldest target on, in order, with inverses). Later edits by
 * other changes win: cells they touched are left alone and reported, unless `force`.
 */
export function revertOps(log: Change[], targets: Set<number>, doc: Doc, force = false): { ops: Op[]; skipped: Skipped[] } {
  const out: Op[] = [];
  const skipped: Array<{ op: Op; by: string }> = [];
  const gone: Array<{ s: string; what: string; by: string }> = [];
  const chosen = log.filter((c) => targets.has(c.v)).sort((a, b) => b.v - a.v);
  for (const t of chosen) {
    if (!t.inv) throw new Error(`v${t.v} was recorded before history kept undo information; it can't be reverted`);
    let inv: Op[] = t.inv;
    let skip: Array<{ op: Op; by: string }> = [];
    for (const c of log) {
      if (c.v <= t.v) continue;
      for (const p of c.ops) {
        skip = skip.flatMap((x) => moveOver([x.op], p, doc).map((op) => ({ op, by: x.by })));
        const r = pastLater(inv, p, doc, !force && !targets.has(c.v));
        inv = r.keep;
        for (const op of r.skip) skip.push({ op, by: c.actor });
        for (const g of r.gone) gone.push({ ...g, by: c.actor });
      }
    }
    for (const p of out) {
      inv = moveOver(inv, p, doc); // out's ops were applied in order, so each is a prior in turn
      skip = skip.flatMap((x) => moveOver([x.op], p, doc).map((op) => ({ op, by: x.by })));
    }
    out.push(...inv);
    skipped.push(...skip);
  }
  const lost = new Map<string, Skipped>();
  for (const g of gone) {
    const k = `${g.s}\u0001${g.by}`;
    const m = lost.get(k);
    if (m) m.what += `, ${g.what}`;
    else lost.set(k, { sheet: g.s, by: g.by, what: g.what });
  }
  return { ops: out, skipped: [...reportSkipped(skipped, doc), ...lost.values()] };
}

function reportSkipped(list: Array<{ op: Op; by: string }>, doc: Doc): Skipped[] {
  const name = (id: string) => doc.sheets.find((s) => s.id === id)?.name ?? "?";
  const merged = new Map<string, { sheet: string; by: string; cells: Set<string>; other: Set<string> }>();
  for (const { op, by } of list) {
    const sheet = "s" in op ? op.s : "";
    const k = `${sheet}\u0001${by}`;
    let m = merged.get(k);
    if (!m) merged.set(k, (m = { sheet, by, cells: new Set(), other: new Set() }));
    if (op.t === "set" || op.t === "fmt" || op.t === "flag") for (const [r, c] of op.cells) m.cells.add(cellName(r, c));
    else m.other.add(op.t === "delete" ? `${op.axis === "row" ? "rows" : "columns"} edited since` : op.t === "cond" ? "conditional fill" : op.t === "view" ? "sort and filters" : op.t === "freeze" ? "frozen panes" : op.t);
  }
  return [...merged.values()].map((m) => {
    const cells = [...m.cells];
    const parts = cells.length ? [cells.length > 12 ? `${cells.slice(0, 12).join(", ")} and ${cells.length - 12} more` : cells.join(", ")] : [];
    return { sheet: m.sheet, by: m.by, what: [...parts, ...m.other].join("; ") + (m.sheet && doc.sheets.length > 1 ? ` (${name(m.sheet)})` : "") };
  });
}

// ---- restore -----------------------------------------------------------------------------

/**
 * Ops that make `cur` equal `target` (the workbook, or one sheet of it). Cells are compared by
 * position, so the result is exact; it is one ordinary change, so it can be undone or reverted.
 */
export function restoreOps(cur: Doc, target: Doc, sheetId?: string): Op[] {
  const want = sheetId ? target.sheets.filter((s) => s.id === sheetId) : target.sheets;
  if (sheetId && !want.length) throw new Error("that sheet did not exist at that version");
  // Structural steps run on a scratch copy, so each later step (and the contents diff) sees the
  // workbook exactly as the earlier ops leave it, including formulas rewritten by renames.
  const scratch = docFromJSON(docToJSON(cur));
  const ops: Op[] = [];
  const run = (op: Op) => {
    applyAll(scratch, [op]);
    ops.push(op);
  };
  const temp = (id: string) => `restoring ${id} ${Date.now().toString(36)}`;
  const sheet = (id: string) => scratch.sheets.find((s) => s.id === id);
  // 1. Sheets that are gone come back, under temporary names (their old names may be taken).
  for (const [i, t] of target.sheets.entries()) {
    if (want.includes(t) && !sheet(t.id)) run({ t: "addSheet", sheet: { ...sheetToJSON(t), name: temp(t.id) }, index: i });
  }
  // 2. Sheets that did not exist then go (whole-workbook restores only).
  if (!sheetId) for (const s of [...scratch.sheets]) if (!target.sheets.some((t) => t.id === s.id)) run({ t: "deleteSheet", s: s.id });
  // 3. Names, through temporary ones so swapped or reused names never collide.
  const renames = want.filter((t) => sheet(t.id)!.name !== t.name);
  if (sheetId) {
    for (const t of renames) if (!scratch.sheets.some((s) => s.id !== t.id && s.name.toLowerCase() === t.name.toLowerCase())) run({ t: "renameSheet", s: t.id, name: t.name });
  } else {
    for (const t of renames) if (!sheet(t.id)!.name.startsWith("restoring ")) run({ t: "renameSheet", s: t.id, name: temp(t.id) });
    for (const t of renames) run({ t: "renameSheet", s: t.id, name: t.name });
  }
  // 4. Contents, compared with the sheets as they now are.
  for (const t of want) ops.push(...sheetDiff(sheet(t.id)!, t));
  // 5. Order and the workbook's name.
  if (!sheetId) {
    target.sheets.forEach((t, i) => {
      if (scratch.sheets.indexOf(sheet(t.id)!) !== i) run({ t: "moveSheet", s: t.id, index: i });
    });
    if (cur.name !== target.name) ops.push({ t: "rename", name: target.name });
  }
  return ops;
}

function sheetDiff(s: Sheet, t: Sheet): Op[] {
  const ops: Op[] = [];
  const vals: Array<[number, number, string | null]> = [], fmts: Array<[number, number, Format | null]> = [], flags: Array<[number, number, Flag | null]> = [];
  const same = (a: unknown, b: unknown) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  for (const k of new Set([...s.cells.keys(), ...t.cells.keys()])) {
    const a = s.cells.get(k), b = t.cells.get(k);
    const r = keyRow(k), c = keyCol(k);
    if ((a?.v ?? null) !== (b?.v ?? null)) vals.push([r, c, b?.v ?? null]);
    if (!same(a?.f, b?.f)) fmts.push([r, c, b?.f ?? null]);
    if (!same(a?.flag, b?.flag)) flags.push([r, c, b?.flag ?? null]);
  }
  if (vals.length) ops.push({ t: "set", s: s.id, cells: vals });
  if (fmts.length) ops.push({ t: "fmt", s: s.id, cells: fmts });
  if (flags.length) ops.push({ t: "flag", s: s.id, cells: flags });
  for (const axis of ["row", "col"] as const) {
    const [ms, mt] = axis === "row" ? [s.rowH, t.rowH] : [s.colW, t.colW];
    const sizes: Array<[number, number | null]> = [];
    for (const i of new Set([...ms.keys(), ...mt.keys()])) if (ms.get(i) !== mt.get(i)) sizes.push([i, mt.get(i) ?? null]);
    if (sizes.length) ops.push({ t: "size", s: s.id, axis, sizes });
    const [hs, ht] = axis === "row" ? [s.hiddenRows, t.hiddenRows] : [s.hiddenCols, t.hiddenCols];
    const show = [...hs].filter((i) => !ht.has(i)), hide = [...ht].filter((i) => !hs.has(i));
    if (show.length) ops.push({ t: "hide", s: s.id, axis, idx: show, hidden: false });
    if (hide.length) ops.push({ t: "hide", s: s.id, axis, idx: hide, hidden: true });
  }
  if (s.frozenRows !== t.frozenRows || s.frozenCols !== t.frozenCols) ops.push({ t: "freeze", s: s.id, rows: t.frozenRows, cols: t.frozenCols });
  if (!same(s.view, t.view)) ops.push({ t: "view", s: s.id, view: t.view });
  if (!same(s.cond, t.cond)) ops.push({ t: "cond", s: s.id, rules: t.cond });
  if (s.rows !== t.rows || s.cols !== t.cols) ops.push({ t: "grid", s: s.id, rows: t.rows, cols: t.cols });
  return ops;
}

// ---- compare -----------------------------------------------------------------------------

export interface Mark {
  s: string;
  r: number;
  c: number;
  before: string | null; // raw input before (null: empty)
  after: string | null;
  fmt?: boolean; // formatting changed
  flag?: boolean; // flag changed
  by?: string;
}

/**
 * Cells whose input, format or flag differ between `a` (older) and `b`, in `b`'s coordinates.
 * `between` are the ops that led from a to b, used to follow rows and columns that moved.
 */
export function changedCells(a: Doc, b: Doc, between: Op[], limit = 20_000): Mark[] {
  const marks: Mark[] = [];
  for (const sb of b.sheets) {
    const sa = a.sheets.find((x) => x.id === sb.id);
    const old = new Map<number, { v: string | null; f?: Format; flag?: Flag }>();
    if (sa) {
      const list = [...sa.cells].map(([k, cell]) => ({ s: sa.id, r: keyRow(k), c: keyCol(k), cell }));
      const carried = carryCells(list, between, b);
      carried.forEach((p, i) => {
        if (p && p.s === sb.id) old.set(key(p.r, p.c), { v: list[i]!.cell.v ?? null, f: list[i]!.cell.f, flag: list[i]!.cell.flag });
      });
    }
    for (const k of new Set([...old.keys(), ...sb.cells.keys()])) {
      const x = old.get(k), y = sb.cells.get(k);
      const before = x?.v ?? null, after = y?.v ?? null;
      const fmt = JSON.stringify(x?.f ?? null) !== JSON.stringify(y?.f ?? null);
      const flag = JSON.stringify(x?.flag ?? null) !== JSON.stringify(y?.flag ?? null);
      if (before === after && !fmt && !flag) continue;
      const m: Mark = { s: sb.id, r: keyRow(k), c: keyCol(k), before, after };
      if (fmt) m.fmt = true;
      if (flag) m.flag = true;
      marks.push(m);
      if (marks.length >= limit) return marks;
    }
  }
  return marks;
}
