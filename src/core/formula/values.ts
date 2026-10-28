// Runtime values and the coercion rules shared by operators and functions.

export type ErrCode =
  | "#REF!" | "#DIV/0!" | "#NAME?" | "#VALUE!" | "#N/A" | "#NUM!" | "#NULL!" | "#CIRC!" | "#ERROR!";

export class Err {
  constructor(readonly code: ErrCode, readonly msg?: string) {}
  toString(): string {
    return this.code;
  }
}

/** A computed cell value. null is an empty cell. */
export type Scalar = number | string | boolean | null | Err;

export const isErr = (v: unknown): v is Err => v instanceof Err;

const NUMERIC = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/;

/** Interpret raw (non-formula) cell input the way a spreadsheet does. */
export function literal(input: string): Scalar {
  if (input === "") return null;
  if (input[0] === "'") return input.slice(1);
  const t = input.trim();
  if (NUMERIC.test(t)) return Number(t);
  const u = t.toUpperCase();
  if (u === "TRUE") return true;
  if (u === "FALSE") return false;
  return input;
}

export function toNumber(v: Scalar): number | Err {
  if (v === null) return 0;
  if (typeof v === "number") return v;
  if (typeof v === "boolean") return v ? 1 : 0;
  if (v instanceof Err) return v;
  const t = v.trim();
  if (t === "") return 0;
  if (NUMERIC.test(t)) return Number(t);
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)%$/.test(t)) return Number(t.slice(0, -1)) / 100;
  return new Err("#VALUE!", `"${v}" is not a number`);
}

export function toText(v: Scalar): string | Err {
  if (v === null) return "";
  if (v instanceof Err) return v;
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatGeneral(v);
  return v;
}

export function toBool(v: Scalar): boolean | Err {
  if (v === null) return false;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v instanceof Err) return v;
  const u = v.trim().toUpperCase();
  if (u === "TRUE") return true;
  if (u === "FALSE") return false;
  return new Err("#VALUE!", `"${v}" is not TRUE or FALSE`);
}

/** Number to text with up to 10 significant digits and no trailing zeros. */
export function formatGeneral(n: number): string {
  if (!Number.isFinite(n)) return "#NUM!";
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const abs = Math.abs(n);
  if (abs !== 0 && (abs >= 1e15 || abs < 1e-9)) return n.toExponential(5).replace(/\.?0+e/, "e");
  return String(Number(n.toPrecision(10)));
}

/** Spreadsheet ordering: numbers < text < booleans; empty acts as 0 or "" next to its peer. Text compares case-insensitively. */
export function compare(a: Scalar, b: Scalar): number {
  if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const rank = (v: Scalar) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  const ra = rank(a), rb = rank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === "string" && typeof b === "string") {
    const x = a.toLowerCase(), y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  const x = Number(a), y = Number(b);
  return x < y ? -1 : x > y ? 1 : 0;
}
