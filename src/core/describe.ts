// One-line, human-readable summaries of ops ("set B2:C4", "inserted 2 rows at 5").
// Used by `sheets watch` and the browser's activity line.

import { cellName, colName, rangeName, type Range } from "./a1.ts";
import type { Op } from "./ops.ts";

/** Bounding box of a list of cells. */
export function bounds(cells: Array<readonly [number, number, ...unknown[]]>): Range | null {
  if (!cells.length) return null;
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  for (const [r, c] of cells) {
    r1 = Math.min(r1, r); c1 = Math.min(c1, c);
    r2 = Math.max(r2, r); c2 = Math.max(c2, c);
  }
  return { r1, c1, r2, c2 };
}

const span = (axis: "row" | "col", at: number, n: number) =>
  axis === "row" ? (n === 1 ? `row ${at + 1}` : `rows ${at + 1}-${at + n}`) : n === 1 ? `column ${colName(at)}` : `columns ${colName(at)}-${colName(at + n - 1)}`;

export function describeOp(op: Op, sheetName: (id: string) => string, verbose = false): string {
  const where = (id: string) => sheetName(id) + "!";
  switch (op.t) {
    case "set": {
      const b = bounds(op.cells);
      if (!b) return "set nothing";
      if (op.cells.length === 1) {
        const [r, c, v] = op.cells[0]!;
        return v === null ? `cleared ${where(op.s)}${cellName(r, c)}` : `set ${where(op.s)}${cellName(r, c)} = ${JSON.stringify(v)}`;
      }
      const head = `set ${where(op.s)}${rangeName(b)} (${op.cells.length} cells)`;
      if (!verbose) return head;
      return head + op.cells.slice(0, 20).map(([r, c, v]) => `\n  ${cellName(r, c)} = ${v === null ? "(empty)" : JSON.stringify(v)}`).join("") + (op.cells.length > 20 ? "\n  …" : "");
    }
    case "fmt": {
      const b = bounds(op.cells);
      return b ? `formatted ${where(op.s)}${rangeName(b)}` : "formatted nothing";
    }
    case "flag": {
      const b = bounds(op.cells);
      if (!b) return "flagged nothing";
      const at = `${where(op.s)}${rangeName(b)}`;
      const f = op.cells[0]![2];
      if (!f) return `removed the flag on ${at}`;
      if (f.done) return `resolved the flag on ${at}${f.done.reply ? `: ${JSON.stringify(f.done.reply)}` : ""}`;
      return `flagged ${at}: ${JSON.stringify(f.note)}`;
    }
    case "style": return `formatted ${where(op.s)}${rangeName(op.range)} ${JSON.stringify(op.patch)}`;
    case "insert": return `inserted ${span(op.axis, op.at, op.n)} in ${sheetName(op.s)}`;
    case "delete": return `deleted ${span(op.axis, op.at, op.n)} in ${sheetName(op.s)}`;
    case "move": {
      const start = op.to <= op.from ? op.to : op.to - op.n;
      return `moved ${span(op.axis, op.from, op.n)} to ${op.axis === "row" ? `row ${start + 1}` : `column ${colName(start)}`} in ${sheetName(op.s)}`;
    }
    case "size": return `resized ${op.sizes.length} ${op.axis}${op.sizes.length === 1 ? "" : "s"} in ${sheetName(op.s)}`;
    case "hide": return `${op.hidden ? "hid" : "unhid"} ${op.idx.length} ${op.axis}${op.idx.length === 1 ? "" : "s"} in ${sheetName(op.s)}`;
    case "freeze": return `froze ${op.rows} row${op.rows === 1 ? "" : "s"} and ${op.cols} column${op.cols === 1 ? "" : "s"} in ${sheetName(op.s)}`;
    case "view": {
      if (!op.view) return `cleared sort and filters in ${sheetName(op.s)}`;
      const parts: string[] = [];
      if (op.view.sort) parts.push(`sorted by ${colName(op.view.sort.col)} ${op.view.sort.desc ? "Z→A" : "A→Z"}`);
      const nf = Object.keys(op.view.filters).length;
      if (nf) parts.push(`${nf} column filter${nf === 1 ? "" : "s"}`);
      return `${parts.join(", ") || "set view"} in ${sheetName(op.s)}`;
    }
    case "cond": return `set ${op.rules.length} conditional rule${op.rules.length === 1 ? "" : "s"} in ${sheetName(op.s)}`;
    case "grid": return `resized ${sheetName(op.s)} to ${op.rows}×${op.cols}`;
    case "addSheet": return `added sheet ${op.sheet.name}`;
    case "deleteSheet": return `deleted sheet ${sheetName(op.s)}`;
    case "renameSheet": return `renamed sheet ${sheetName(op.s)} to ${op.name}`;
    case "moveSheet": return `moved sheet ${sheetName(op.s)}`;
    case "rename": return `renamed workbook to ${op.name}`;
  }
}
