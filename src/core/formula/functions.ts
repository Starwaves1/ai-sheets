// Built-in functions. Each receives unevaluated argument nodes so IF/IFERROR/AND/OR
// can short-circuit; helpers on Ev evaluate them as scalars or ranges.

import type { Node } from "./parser.ts";
import { compare, Err, isErr, literal, toBool, toNumber, toText, type Scalar } from "./values.ts";

export interface RangeVal {
  rows: number;
  cols: number;
  at(i: number, j: number): Scalar;
}

export interface Ev {
  scalar(n: Node): Scalar;
  value(n: Node): Scalar | RangeVal;
}

type Fn = (args: Node[], ev: Ev) => Scalar;

export const isRange = (v: unknown): v is RangeVal =>
  typeof v === "object" && v !== null && !(v instanceof Err) && "at" in v;

function* cells(v: Scalar | RangeVal): Generator<Scalar> {
  if (!isRange(v)) {
    yield v;
    return;
  }
  for (let i = 0; i < v.rows; i++) for (let j = 0; j < v.cols; j++) yield v.at(i, j);
}

/** Numbers for aggregate functions: in ranges only real numbers count; direct arguments are coerced. */
function numbers(args: Node[], ev: Ev): number[] | Err {
  const out: number[] = [];
  for (const a of args) {
    const v = ev.value(a);
    if (isRange(v)) {
      for (const x of cells(v)) {
        if (isErr(x)) return x;
        if (typeof x === "number") out.push(x);
      }
    } else {
      if (v === null) continue;
      const n = toNumber(v);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function arity(args: Node[], min: number, max = min): Err | null {
  if (args.length < min || args.length > max) {
    return new Err("#N/A", `expects ${min === max ? min : `${min} to ${max}`} argument${max === 1 ? "" : "s"}`);
  }
  return null;
}

function num(ev: Ev, n: Node): number | Err {
  return toNumber(ev.scalar(n));
}

function roundTo(x: number, digits: number, mode: "round" | "up" | "down"): number {
  const f = 10 ** Math.trunc(digits);
  const y = Math.abs(x) * f;
  // Nudge by a relative epsilon so 1.005 rounds like people expect.
  const e = y * (1 + Number.EPSILON * 4);
  const r = mode === "round" ? Math.round(e) : mode === "up" ? Math.ceil(y - 1e-9) : Math.floor(e);
  return (Math.sign(x) * r) / f;
}

// ---- criteria (COUNTIF, SUMIF, AVERAGEIF) ---------------------------------

export function criterion(c: Scalar): (v: Scalar) => boolean {
  if (typeof c !== "string") return (v) => v !== null && compare(v, c) === 0;
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(c)!;
  const op = m[1] ?? "=";
  const rhsText = m[2]!;
  const rhs = literal(rhsText);
  if (typeof rhs === "string" && (op === "=" || op === "<>") && /[*?]/.test(rhs)) {
    const re = new RegExp(
      "^" + rhs.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$",
      "is",
    );
    return (v) => {
      const t = toText(v);
      const hit = !isErr(t) && re.test(t);
      return op === "=" ? hit : !hit;
    };
  }
  if (rhsText === "") {
    return op === "<>" ? (v) => v !== null && v !== "" : (v) => v === null || v === "";
  }
  return (v) => {
    if (isErr(v) || v === null) return op === "<>";
    // Typed comparisons only between like kinds, except = / <> which compare loosely.
    if ((typeof rhs === "number") !== (typeof v === "number") && op !== "=" && op !== "<>") return false;
    const d = compare(v, rhs);
    switch (op) {
      case "=": return d === 0;
      case "<>": return d !== 0;
      case "<": return d < 0;
      case ">": return d > 0;
      case "<=": return d <= 0;
      default: return d >= 0;
    }
  };
}

function ifAgg(args: Node[], ev: Ev, kind: "count" | "sum" | "avg"): Scalar {
  const bad = arity(args, 2, kind === "count" ? 2 : 3);
  if (bad) return bad;
  const range = ev.value(args[0]!);
  const crit = ev.scalar(args[1]!);
  if (isErr(crit)) return crit;
  const test = criterion(crit);
  const sumRange = args[2] ? ev.value(args[2]) : range;
  const rr = isRange(range) ? range : { rows: 1, cols: 1, at: () => range as Scalar };
  const sr = isRange(sumRange) ? sumRange : { rows: 1, cols: 1, at: () => sumRange as Scalar };
  let count = 0, sum = 0;
  for (let i = 0; i < rr.rows; i++) {
    for (let j = 0; j < rr.cols; j++) {
      if (!test(rr.at(i, j))) continue;
      count++;
      if (kind !== "count") {
        const v = i < sr.rows && j < sr.cols ? sr.at(i, j) : null;
        if (isErr(v)) return v;
        if (typeof v === "number") sum += v;
        else count--;
      }
    }
  }
  if (kind === "count") return count;
  if (kind === "sum") return sum;
  return count === 0 ? new Err("#DIV/0!", "no matching numbers") : sum / count;
}

// ---- table ----------------------------------------------------------------------

export const FUNCTIONS: Record<string, Fn> = {
  SUM(args, ev) {
    const ns = numbers(args, ev);
    return isErr(ns) ? ns : ns.reduce((a, b) => a + b, 0);
  },
  PRODUCT(args, ev) {
    const ns = numbers(args, ev);
    return isErr(ns) ? ns : ns.reduce((a, b) => a * b, 1);
  },
  AVERAGE(args, ev) {
    const ns = numbers(args, ev);
    if (isErr(ns)) return ns;
    return ns.length ? ns.reduce((a, b) => a + b, 0) / ns.length : new Err("#DIV/0!", "no numbers to average");
  },
  MIN(args, ev) {
    const ns = numbers(args, ev);
    return isErr(ns) ? ns : ns.length ? Math.min(...ns) : 0;
  },
  MAX(args, ev) {
    const ns = numbers(args, ev);
    return isErr(ns) ? ns : ns.length ? Math.max(...ns) : 0;
  },
  MEDIAN(args, ev) {
    const ns = numbers(args, ev);
    if (isErr(ns)) return ns;
    if (!ns.length) return new Err("#NUM!", "no numbers");
    ns.sort((a, b) => a - b);
    const m = ns.length >> 1;
    return ns.length % 2 ? ns[m]! : (ns[m - 1]! + ns[m]!) / 2;
  },
  COUNT(args, ev) {
    let n = 0;
    for (const a of args) {
      const v = ev.value(a);
      for (const x of cells(v)) {
        if (typeof x === "number") n++;
        else if (!isRange(v) && typeof x === "string" && !isErr(toNumber(x))) n++;
      }
    }
    return n;
  },
  COUNTA(args, ev) {
    let n = 0;
    for (const a of args) for (const x of cells(ev.value(a))) if (x !== null && x !== "") n++;
    return n;
  },
  COUNTBLANK(args, ev) {
    let n = 0;
    for (const a of args) for (const x of cells(ev.value(a))) if (x === null || x === "") n++;
    return n;
  },
  COUNTIF: (args, ev) => ifAgg(args, ev, "count"),
  SUMIF: (args, ev) => ifAgg(args, ev, "sum"),
  AVERAGEIF: (args, ev) => ifAgg(args, ev, "avg"),
  IF(args, ev) {
    const bad = arity(args, 2, 3);
    if (bad) return bad;
    const c = toBool(ev.scalar(args[0]!));
    if (isErr(c)) return c;
    if (c) return ev.scalar(args[1]!);
    return args[2] ? ev.scalar(args[2]) : false;
  },
  IFERROR(args, ev) {
    const bad = arity(args, 2);
    if (bad) return bad;
    const v = ev.scalar(args[0]!);
    return isErr(v) ? ev.scalar(args[1]!) : v;
  },
  AND(args, ev) {
    let any = false;
    for (const a of args) {
      for (const x of cells(ev.value(a))) {
        if (x === null) continue;
        const b = toBool(x);
        if (isErr(b)) return b;
        any = true;
        if (!b) return false;
      }
    }
    return any ? true : new Err("#VALUE!", "no logical values");
  },
  OR(args, ev) {
    let any = false;
    for (const a of args) {
      for (const x of cells(ev.value(a))) {
        if (x === null) continue;
        const b = toBool(x);
        if (isErr(b)) return b;
        any = true;
        if (b) return true;
      }
    }
    return any ? false : new Err("#VALUE!", "no logical values");
  },
  NOT(args, ev) {
    const bad = arity(args, 1);
    if (bad) return bad;
    const b = toBool(ev.scalar(args[0]!));
    return isErr(b) ? b : !b;
  },
  ROUND: (args, ev) => roundFn(args, ev, "round"),
  ROUNDUP: (args, ev) => roundFn(args, ev, "up"),
  ROUNDDOWN: (args, ev) => roundFn(args, ev, "down"),
  ABS(args, ev) {
    const bad = arity(args, 1);
    if (bad) return bad;
    const x = num(ev, args[0]!);
    return isErr(x) ? x : Math.abs(x);
  },
  LEN(args, ev) {
    const bad = arity(args, 1);
    if (bad) return bad;
    const t = toText(ev.scalar(args[0]!));
    return isErr(t) ? t : t.length;
  },
  UPPER: (args, ev) => textFn(args, ev, (s) => s.toUpperCase()),
  LOWER: (args, ev) => textFn(args, ev, (s) => s.toLowerCase()),
  TRIM: (args, ev) => textFn(args, ev, (s) => s.trim().replace(/ +/g, " ")),
  CONCAT: concat,
  CONCATENATE: concat,
  ISBLANK(args, ev) {
    const bad = arity(args, 1);
    return bad ?? ev.scalar(args[0]!) === null;
  },
  ISNUMBER(args, ev) {
    const bad = arity(args, 1);
    return bad ?? typeof ev.scalar(args[0]!) === "number";
  },
  ISERROR(args, ev) {
    const bad = arity(args, 1);
    return bad ?? isErr(ev.scalar(args[0]!));
  },
};

function roundFn(args: Node[], ev: Ev, mode: "round" | "up" | "down"): Scalar {
  const bad = arity(args, 1, 2);
  if (bad) return bad;
  const x = num(ev, args[0]!);
  if (isErr(x)) return x;
  const d = args[1] ? num(ev, args[1]) : 0;
  if (isErr(d)) return d;
  return roundTo(x, d, mode);
}

function textFn(args: Node[], ev: Ev, f: (s: string) => string): Scalar {
  const bad = arity(args, 1);
  if (bad) return bad;
  const t = toText(ev.scalar(args[0]!));
  return isErr(t) ? t : f(t);
}

function concat(args: Node[], ev: Ev): Scalar {
  let s = "";
  for (const a of args) {
    for (const x of cells(ev.value(a))) {
      const t = toText(x);
      if (isErr(t)) return t;
      s += t;
    }
  }
  return s;
}
