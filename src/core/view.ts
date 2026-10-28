// Which rows and columns are shown, in what order: hidden rows/cols, filters, and view sort.
// The data never moves; the grid maps display positions to stored rows through these lists.

import { displayValue, matches } from "./display.ts";
import { compare, type Scalar } from "./formula/values.ts";
import { key, usedRange, type Sheet } from "./model.ts";

export type ValueOf = (r: number, c: number) => Scalar;

/** Rows that pass the sheet's filters (ignores manual hiding and sort). */
export function rowPassesFilter(sheet: Sheet, r: number, valueOf: ValueOf): boolean {
  const view = sheet.view;
  if (!view || r <= view.headerRow) return true;
  for (const [cs, f] of Object.entries(view.filters)) {
    const c = Number(cs);
    const v = valueOf(r, c);
    if (f.hide?.length && f.hide.includes(displayValue(v, sheet.cells.get(key(r, c))?.f?.nf))) return false;
    if (f.op && !matches(f.op, v, f.value, f.value2)) return false;
  }
  return true;
}

/** Stored row indices in display order. */
export function visibleRows(sheet: Sheet, valueOf: ValueOf): number[] {
  const view = sheet.view;
  const out: number[] = [];
  const body: number[] = [];
  for (let r = 0; r < sheet.rows; r++) {
    if (sheet.hiddenRows.has(r)) continue;
    if (!view || r <= view.headerRow) out.push(r);
    else if (rowPassesFilter(sheet, r, valueOf)) body.push(r);
  }
  const sort = view?.sort;
  if (sort) {
    const empty = (v: Scalar) => v === null || v === "";
    body.sort((a, b) => {
      const x = valueOf(a, sort.col), y = valueOf(b, sort.col);
      // Empty cells always sink to the bottom, whichever direction.
      if (empty(x) || empty(y)) return empty(x) === empty(y) ? a - b : empty(x) ? 1 : -1;
      const d = compare(x, y);
      return (sort.desc ? -d : d) || a - b;
    });
  }
  for (const r of body) out.push(r);
  return out;
}

export function visibleCols(sheet: Sheet): number[] {
  const out: number[] = [];
  for (let c = 0; c < sheet.cols; c++) if (!sheet.hiddenCols.has(c)) out.push(c);
  return out;
}

/** Distinct display texts in a column below the header (for the filter menu), with counts. */
export function columnValues(sheet: Sheet, col: number, valueOf: ValueOf): Array<[string, number]> {
  const counts = new Map<string, number>();
  const start = (sheet.view?.headerRow ?? 0) + 1;
  const end = usedRange(sheet)?.r2 ?? -1;
  for (let r = start; r <= end; r++) {
    const t = displayValue(valueOf(r, col), sheet.cells.get(key(r, col))?.f?.nf);
    counts.set(t, (counts.get(t) ?? 0) + 1);
  }
  return [...counts].sort((a, b) => (a[0] === "" ? 1 : b[0] === "" ? -1 : compare(a[0], b[0])));
}
