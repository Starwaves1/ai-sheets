// Flags: notes on cells asking someone (usually an agent) to look at them. A flag set on a
// range lives on every cell of it; groups put those cells back together for listing.

import { rangeName, type Range } from "./a1.ts";
import { keyCol, keyRow, type Doc, type Flag } from "./model.ts";

export interface FlagGroup {
  sheet: string; // sheet id
  sheetName: string;
  range: Range; // bounding box of the cells carrying this flag
  ref: string; // "Runs!B4" or "Runs!B4:C9"
  cells: Array<[number, number]>;
  flag: Flag;
}

/** Same flag = same note, author and time (set together). */
const flagKey = (f: Flag) => `${f.ts}\u0001${f.by}\u0001${f.note}`;

/**
 * Every flag in the workbook, oldest first: one entry per block of touching cells carrying the
 * same flag (set together; a pasted copy elsewhere is its own entry).
 */
export function flagGroups(doc: Doc): FlagGroup[] {
  const out: FlagGroup[] = [];
  for (const sh of doc.sheets) {
    const byFlag = new Map<string, Map<number, Flag>>();
    for (const [k, cell] of sh.cells) {
      if (!cell.flag) continue;
      const id = flagKey(cell.flag) + (cell.flag.done ? "\u0001done" : "");
      let m = byFlag.get(id);
      if (!m) byFlag.set(id, (m = new Map()));
      m.set(k, cell.flag);
    }
    const name = /^[A-Za-z_][\w.]*$/.test(sh.name) ? sh.name : `'${sh.name.replace(/'/g, "''")}'`;
    for (const cells of byFlag.values()) {
      const left = new Set(cells.keys());
      for (const start of cells.keys()) {
        if (!left.has(start)) continue;
        left.delete(start);
        const g: FlagGroup = { sheet: sh.id, sheetName: sh.name, range: { r1: Infinity, c1: Infinity, r2: -1, c2: -1 }, ref: "", cells: [], flag: cells.get(start)! };
        for (const stack = [start]; stack.length; ) {
          const k = stack.pop()!;
          const r = keyRow(k), c = keyCol(k);
          g.cells.push([r, c]);
          g.range = { r1: Math.min(g.range.r1, r), c1: Math.min(g.range.c1, c), r2: Math.max(g.range.r2, r), c2: Math.max(g.range.c2, c) };
          for (const n of [k - 16384, k + 16384, c > 0 ? k - 1 : -1, k + 1]) {
            if (left.has(n)) {
              left.delete(n);
              stack.push(n);
            }
          }
        }
        g.ref = `${name}!${rangeName(g.range)}`;
        out.push(g);
      }
    }
  }
  return out.sort((a, b) => a.flag.ts - b.flag.ts);
}

/** The group a cell's flag belongs to (all cells flagged with it at the same time). */
export function groupAt(doc: Doc, sheet: string, r: number, c: number): FlagGroup | null {
  return flagGroups(doc).find((g) => g.sheet === sheet && g.cells.some(([gr, gc]) => gr === r && gc === c)) ?? null;
}
