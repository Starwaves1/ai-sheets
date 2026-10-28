// A1 notation: column letters, cell and range addresses, optional sheet prefix.
// Rows and columns are 0-based everywhere in code; A1 text is 1-based.

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

export interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export interface SheetRange extends Range {
  sheet?: string;
}

export function colName(c: number): string {
  let s = "";
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function colIndex(name: string): number {
  let n = 0;
  for (const ch of name.toUpperCase()) {
    const code = ch.charCodeAt(0);
    if (code < 65 || code > 90) return -1;
    n = n * 26 + (code - 64);
  }
  return n - 1;
}

export function cellName(r: number, c: number): string {
  return colName(c) + (r + 1);
}

/** Normalize so r1<=r2 and c1<=c2. */
export function norm(r: Range): Range {
  return {
    r1: Math.min(r.r1, r.r2),
    c1: Math.min(r.c1, r.c2),
    r2: Math.max(r.r1, r.r2),
    c2: Math.max(r.c1, r.c2),
  };
}

export function rangeName(r: Range): string {
  const n = norm(r);
  const wholeCols = n.r1 === 0 && n.r2 >= MAX_ROWS - 1;
  const wholeRows = n.c1 === 0 && n.c2 >= MAX_COLS - 1;
  if (wholeCols && !wholeRows) return `${colName(n.c1)}:${colName(n.c2)}`;
  if (wholeRows && !wholeCols) return `${n.r1 + 1}:${n.r2 + 1}`;
  const a = cellName(n.r1, n.c1);
  if (n.r1 === n.r2 && n.c1 === n.c2) return a;
  return `${a}:${cellName(n.r2, n.c2)}`;
}

export function quoteSheet(name: string): string {
  return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

export function sheetRangeName(sheet: string | undefined, r: Range): string {
  return sheet ? `${quoteSheet(sheet)}!${rangeName(r)}` : rangeName(r);
}

const CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;
const COL = /^\$?([A-Za-z]{1,3})$/;
const ROW = /^\$?(\d+)$/;

function part(s: string): { r?: number; c?: number } | null {
  let m = CELL.exec(s);
  if (m) {
    const c = colIndex(m[1]!);
    const r = Number(m[2]) - 1;
    if (c < 0 || c >= MAX_COLS || r < 0 || r >= MAX_ROWS) return null;
    return { r, c };
  }
  m = COL.exec(s);
  if (m) {
    const c = colIndex(m[1]!);
    return c >= 0 && c < MAX_COLS ? { c } : null;
  }
  m = ROW.exec(s);
  if (m) {
    const r = Number(m[1]) - 1;
    return r >= 0 && r < MAX_ROWS ? { r } : null;
  }
  return null;
}

/** Parse "B3", "A1:C9", "A:C", "2:5", optionally prefixed "Sheet 1!" or "'It''s'!". */
export function parseRange(text: string): SheetRange | null {
  let s = text.trim();
  let sheet: string | undefined;
  const bang = s.lastIndexOf("!");
  if (bang >= 0) {
    let name = s.slice(0, bang);
    if (name.startsWith("'") && name.endsWith("'") && name.length >= 2) {
      name = name.slice(1, -1).replace(/''/g, "'");
    }
    if (!name) return null;
    sheet = name;
    s = s.slice(bang + 1);
  }
  const [a, b, extra] = s.split(":");
  if (extra !== undefined || !a) return null;
  const p = part(a);
  const q = b === undefined ? p : part(b);
  if (!p || !q) return null;
  // A single part must be a cell; "A" alone or "3" alone is not a range.
  if (b === undefined && (p.r === undefined || p.c === undefined)) return null;
  const isCell = (x: { r?: number; c?: number }) => x.r !== undefined && x.c !== undefined;
  if (b !== undefined && isCell(p) !== isCell(q)) return null;
  const r1 = p.r ?? 0;
  const c1 = p.c ?? 0;
  const r2 = q.r ?? (p.r === undefined ? MAX_ROWS - 1 : 0);
  const c2 = q.c ?? (p.c === undefined ? MAX_COLS - 1 : 0);
  const out: SheetRange = { ...norm({ r1, c1, r2, c2 }) };
  if (p.r === undefined) {
    out.r1 = 0;
    out.r2 = MAX_ROWS - 1;
  }
  if (p.c === undefined) {
    out.c1 = 0;
    out.c2 = MAX_COLS - 1;
  }
  if (sheet !== undefined) out.sheet = sheet;
  return out;
}

export function parseCell(text: string): { r: number; c: number; sheet?: string } | null {
  const r = parseRange(text);
  if (!r || r.r1 !== r.r2 || r.c1 !== r.c2) return null;
  return r.sheet === undefined ? { r: r.r1, c: r.c1 } : { r: r.r1, c: r.c1, sheet: r.sheet };
}

export function clamp(r: Range, rows: number, cols: number): Range {
  return {
    r1: Math.max(0, Math.min(r.r1, rows - 1)),
    c1: Math.max(0, Math.min(r.c1, cols - 1)),
    r2: Math.max(0, Math.min(r.r2, rows - 1)),
    c2: Math.max(0, Math.min(r.c2, cols - 1)),
  };
}

export function contains(r: Range, row: number, col: number): boolean {
  return row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
}

export function intersects(a: Range, b: Range): boolean {
  return a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
}
