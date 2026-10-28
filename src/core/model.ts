// The workbook document: plain data, JSON-serializable through toJSON/fromJSON.
// Everything mutates through ops (ops.ts); nothing else writes to a Doc.

import type { Range } from "./a1.ts";

export type Align = "left" | "center" | "right";

/** Cell formatting. Colors are palette names (see PALETTE) or #rrggbb. */
export interface Format {
  b?: boolean;
  i?: boolean;
  s?: boolean; // strikethrough
  color?: string;
  fill?: string;
  align?: Align;
  nf?: string; // number format, e.g. "0.00", "0%", "#,##0", "$#,##0.00"
  wrap?: boolean;
}

export const FORMAT_KEYS = ["b", "i", "s", "color", "fill", "align", "nf", "wrap"] as const;

/** Named colors render with theme-aware shades (see styles.css). */
export const PALETTE = ["green", "red", "amber", "blue", "violet", "grey"] as const;

export interface Cell {
  v?: string; // raw input exactly as entered: "12", "hello", "=SUM(A1:A3)"
  f?: Format;
}

export type CondOp =
  | "eq" | "ne" | "gt" | "gte" | "lt" | "lte"
  | "between" | "contains" | "empty" | "notEmpty";

export const COND_OPS: readonly CondOp[] = [
  "eq", "ne", "gt", "gte", "lt", "lte", "between", "contains", "empty", "notEmpty",
];

export interface CondStyle {
  fill?: string;
  color?: string;
  b?: boolean;
}

/** Conditional format rule; first matching rule in sheet order wins. */
export interface CondRule {
  range: Range;
  op: CondOp;
  value?: string;
  value2?: string;
  style: CondStyle;
}

/** Column filter: hide rows whose displayed text is in `hide`, and/or keep rows matching a condition. */
export interface ColFilter {
  hide?: string[];
  op?: CondOp;
  value?: string;
  value2?: string;
}

/** Shared sort/filter state. Rows at or above headerRow are never sorted or filtered. View sorting reorders display only, never data. */
export interface ViewState {
  headerRow: number;
  sort?: { col: number; desc?: boolean };
  filters: Record<number, ColFilter>;
}

export interface Sheet {
  id: string;
  name: string;
  rows: number;
  cols: number;
  cells: Map<number, Cell>;
  colW: Map<number, number>;
  rowH: Map<number, number>;
  hiddenRows: Set<number>;
  hiddenCols: Set<number>;
  frozenRows: number;
  frozenCols: number;
  view: ViewState | null;
  cond: CondRule[];
}

export interface Doc {
  name: string;
  sheets: Sheet[];
}

export const DEFAULT_ROWS = 1000;
export const DEFAULT_COLS = 26;
export const DEFAULT_COL_W = 100;
export const DEFAULT_ROW_H = 24;

// Keys pack (row, col) into one number; cols < 2^14 so rows fit easily in 2^53.
export const key = (r: number, c: number): number => r * 16384 + c;
export const keyRow = (k: number): number => Math.floor(k / 16384);
export const keyCol = (k: number): number => k % 16384;

export function newSheet(id: string, name: string, rows = DEFAULT_ROWS, cols = DEFAULT_COLS): Sheet {
  return {
    id, name, rows, cols,
    cells: new Map(),
    colW: new Map(),
    rowH: new Map(),
    hiddenRows: new Set(),
    hiddenCols: new Set(),
    frozenRows: 0,
    frozenCols: 0,
    view: null,
    cond: [],
  };
}

export function newDoc(name: string, firstSheetId = "s1"): Doc {
  return { name, sheets: [newSheet(firstSheetId, "Sheet1")] };
}

export function getSheet(doc: Doc, id: string): Sheet {
  const s = doc.sheets.find((x) => x.id === id);
  if (!s) throw new OpError(`no sheet with id ${id}`);
  return s;
}

export function findSheet(doc: Doc, nameOrId: string): Sheet | undefined {
  const lower = nameOrId.toLowerCase();
  return (
    doc.sheets.find((s) => s.id === nameOrId) ??
    doc.sheets.find((s) => s.name.toLowerCase() === lower)
  );
}

export class OpError extends Error {}

/** Smallest range covering every cell with a value, or null if the sheet is empty. */
export function usedRange(sheet: Sheet): Range | null {
  let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
  for (const [k, cell] of sheet.cells) {
    if (cell.v === undefined || cell.v === "") continue;
    const r = keyRow(k), c = keyCol(k);
    if (r < r1) r1 = r;
    if (c < c1) c1 = c;
    if (r > r2) r2 = r;
    if (c > c2) c2 = c;
  }
  return r2 < 0 ? null : { r1: 0, c1: 0, r2, c2 };
}

// ---- JSON form -------------------------------------------------------------

export interface SheetJSON {
  id: string;
  name: string;
  rows: number;
  cols: number;
  cells: Array<[number, number, string | null, Format?]>;
  colW?: Array<[number, number]>;
  rowH?: Array<[number, number]>;
  hiddenRows?: number[];
  hiddenCols?: number[];
  frozenRows?: number;
  frozenCols?: number;
  view?: ViewState | null;
  cond?: CondRule[];
}

export interface DocJSON {
  name: string;
  sheets: SheetJSON[];
}

export function sheetToJSON(s: Sheet): SheetJSON {
  const cells: SheetJSON["cells"] = [];
  for (const [k, cell] of s.cells) {
    const row: [number, number, string | null, Format?] = [keyRow(k), keyCol(k), cell.v ?? null];
    if (cell.f) row.push(cell.f);
    cells.push(row);
  }
  return {
    id: s.id, name: s.name, rows: s.rows, cols: s.cols, cells,
    colW: [...s.colW], rowH: [...s.rowH],
    hiddenRows: [...s.hiddenRows], hiddenCols: [...s.hiddenCols],
    frozenRows: s.frozenRows, frozenCols: s.frozenCols,
    view: s.view ? structuredClone(s.view) : null,
    cond: structuredClone(s.cond),
  };
}

export function sheetFromJSON(j: SheetJSON): Sheet {
  const s = newSheet(j.id, j.name, j.rows, j.cols);
  for (const [r, c, v, f] of j.cells) {
    const cell: Cell = {};
    if (v !== null && v !== "") cell.v = v;
    if (f && Object.keys(f).length) cell.f = { ...f };
    if (cell.v !== undefined || cell.f) s.cells.set(key(r, c), cell);
  }
  s.colW = new Map(j.colW ?? []);
  s.rowH = new Map(j.rowH ?? []);
  s.hiddenRows = new Set(j.hiddenRows ?? []);
  s.hiddenCols = new Set(j.hiddenCols ?? []);
  s.frozenRows = j.frozenRows ?? 0;
  s.frozenCols = j.frozenCols ?? 0;
  s.view = j.view ? structuredClone(j.view) : null;
  s.cond = structuredClone(j.cond ?? []);
  return s;
}

export const docToJSON = (d: Doc): DocJSON => ({ name: d.name, sheets: d.sheets.map(sheetToJSON) });
export const docFromJSON = (j: DocJSON): Doc => ({ name: j.name, sheets: j.sheets.map(sheetFromJSON) });
