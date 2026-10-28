// Lazy, cached formula evaluation with dependency tracking.
// Values are computed on first read and cached. Each formula records which cells
// and ranges it read; invalidate() drops the cached values of a changed cell and
// everything that transitively depends on it.

import { clamp, contains, type Range } from "../a1.ts";
import { findSheet, key, type Doc, type Sheet } from "../model.ts";
import { FUNCTIONS, isRange, type Ev, type RangeVal } from "./functions.ts";
import { parseFormula, type Node } from "./parser.ts";
import { compare, Err, isErr, literal, toNumber, toText, type Scalar } from "./values.ts";
import { isFormula } from "./rewrite.ts";

type Id = string; // `${sheetId}:${cellKey}`
const id = (sheetId: string, r: number, c: number): Id => `${sheetId}:${key(r, c)}`;

interface RangeDep extends Range {
  sheet: string;
}

const MAX_DEPTH = 300;

/** Thrown when a dependency chain gets too deep for the JS stack; value() evaluates the deepest cell first and retries. */
class TooDeep {
  constructor(readonly sheet: Sheet, readonly r: number, readonly c: number) {}
}

export class Calc {
  private vals = new Map<Id, Scalar>();
  private asts = new Map<string, Node | Err>();
  private cellDeps = new Map<Id, Id[]>(); // formula -> cells it read
  private rangeDeps = new Map<Id, RangeDep[]>(); // formula -> ranges it read
  private readers = new Map<Id, Set<Id>>(); // cell -> formulas that read it
  private active = new Set<Id>();
  private depth = 0;

  constructor(private doc: Doc) {}

  /** Computed value of a cell. */
  value(sheetId: string, r: number, c: number): Scalar {
    const sheet = this.doc.sheets.find((s) => s.id === sheetId);
    if (!sheet) return new Err("#REF!");
    // Long chains (A2=A1+1, A3=A2+1, ...) are evaluated in slices of MAX_DEPTH, deepest first.
    const pending: Array<[Sheet, number, number]> = [[sheet, r, c]];
    const seen = new Set<Id>([id(sheet.id, r, c)]);
    for (;;) {
      const [s, rr, cc] = pending[pending.length - 1]!;
      try {
        const v = this.get(s, rr, cc);
        pending.pop();
        if (!pending.length) return v;
      } catch (e) {
        if (!(e instanceof TooDeep)) throw e;
        const deep = id(e.sheet.id, e.r, e.c);
        if (seen.has(deep)) this.vals.set(deep, new Err("#CIRC!", "circular reference"));
        else {
          seen.add(deep);
          pending.push([e.sheet, e.r, e.c]);
        }
      }
    }
  }

  /** Forget everything (after structural changes such as row insertion). */
  reset(): void {
    this.vals.clear();
    this.cellDeps.clear();
    this.rangeDeps.clear();
    this.readers.clear();
  }

  /** A cell's input changed: drop cached values of it and its dependents. */
  invalidate(cells: Iterable<[string, number, number]>): void {
    const queue: Id[] = [];
    const changed: Array<[string, number, number]> = [];
    for (const [s, r, c] of cells) {
      queue.push(id(s, r, c));
      changed.push([s, r, c]);
    }
    if (changed.length > 5000) return this.reset();
    const seen = new Set<Id>();
    const rangeHits = (sheet: string, r: number, c: number) => {
      for (const [f, ranges] of this.rangeDeps) {
        if (ranges.some((g) => g.sheet === sheet && contains(g, r, c))) queue.push(f);
      }
    };
    for (const [s, r, c] of changed) rangeHits(s, r, c);
    while (queue.length) {
      const x = queue.pop()!;
      if (seen.has(x)) continue;
      seen.add(x);
      const wasCached = this.vals.delete(x);
      const rs = this.readers.get(x);
      if (rs) for (const f of rs) queue.push(f);
      this.forgetDeps(x);
      // A formula's own value changed, so formulas reading it through a range must recompute too.
      if (wasCached && this.rangeDeps.size) {
        const [s, k] = x.split(":") as [string, string];
        const kn = Number(k);
        rangeHits(s, Math.floor(kn / 16384), kn % 16384);
      }
    }
  }

  private forgetDeps(f: Id): void {
    const cs = this.cellDeps.get(f);
    if (cs) for (const c of cs) this.readers.get(c)?.delete(f);
    this.cellDeps.delete(f);
    this.rangeDeps.delete(f);
  }

  private get(sheet: Sheet, r: number, c: number): Scalar {
    const input = sheet.cells.get(key(r, c))?.v;
    if (input === undefined) return null;
    if (!isFormula(input)) return literal(input);
    const me = id(sheet.id, r, c);
    const cached = this.vals.get(me);
    if (cached !== undefined) return cached;
    if (this.active.has(me)) return new Err("#CIRC!", "circular reference");
    if (this.depth >= MAX_DEPTH) throw new TooDeep(sheet, r, c);
    this.active.add(me);
    this.depth++;
    let v: Scalar;
    try {
      v = this.evalFormula(sheet, me, input);
    } finally {
      this.active.delete(me);
      this.depth--;
    }
    if (v === null) v = 0; // a formula pointing at an empty cell shows 0
    if (typeof v === "number" && !Number.isFinite(v)) v = new Err("#NUM!", "result is not a finite number");
    this.vals.set(me, v);
    return v;
  }

  private parse(input: string): Node | Err {
    let ast = this.asts.get(input);
    if (ast === undefined) {
      try {
        ast = parseFormula(input.slice(1));
      } catch (e) {
        ast = new Err("#ERROR!", `formula parse error: ${(e as Error).message}`);
      }
      if (this.asts.size > 50_000) this.asts.clear();
      this.asts.set(input, ast);
    }
    return ast;
  }

  private evalFormula(home: Sheet, me: Id, input: string): Scalar {
    const ast = this.parse(input);
    if (isErr(ast)) return ast;
    this.forgetDeps(me);
    const cells: Id[] = [];
    const ranges: RangeDep[] = [];
    this.cellDeps.set(me, cells);
    this.rangeDeps.set(me, ranges);

    const resolve = (name: string | undefined): Sheet | null =>
      name === undefined ? home : findSheet(this.doc, name) ?? null;

    const ev: Ev = {
      value: (n) => evalNode(n),
      scalar: (n) => {
        const v = evalNode(n);
        if (!isRange(v)) return v;
        if (v.rows === 1 && v.cols === 1) return v.at(0, 0);
        return new Err("#VALUE!", "expected a single cell, got a range");
      },
    };

    const evalNode = (n: Node): Scalar | RangeVal => {
      switch (n.k) {
        case "num": return n.v;
        case "str": return n.v;
        case "bool": return n.v;
        case "err": return new Err(n.v as Err["code"]);
        case "name": return new Err("#NAME?", `unknown name ${n.v}`);
        case "ref": {
          const sheet = resolve(n.sheet);
          if (!sheet) return new Err("#REF!", `no sheet named ${n.sheet}`);
          if (!n.area) {
            const cid = id(sheet.id, n.r1, n.c1);
            cells.push(cid);
            let set = this.readers.get(cid);
            if (!set) this.readers.set(cid, (set = new Set()));
            set.add(me);
            return this.get(sheet, n.r1, n.c1);
          }
          const g = clamp(n, sheet.rows, sheet.cols);
          ranges.push({ sheet: sheet.id, ...g });
          const rows = Math.max(0, Math.min(n.r2, sheet.rows - 1) - g.r1 + 1);
          const cols = Math.max(0, Math.min(n.c2, sheet.cols - 1) - g.c1 + 1);
          return { rows, cols, at: (i, j) => this.get(sheet, g.r1 + i, g.c1 + j) };
        }
        case "un": {
          const a = ev.scalar(n.a);
          if (isErr(a)) return a;
          const x = toNumber(a);
          if (isErr(x)) return x;
          return n.op === "-" ? -x : n.op === "%" ? x / 100 : x;
        }
        case "bin": return binary(n.op, ev.scalar(n.a), ev.scalar(n.b));
        case "call": {
          const fn = FUNCTIONS[n.name];
          if (!fn) return new Err("#NAME?", `unknown function ${n.name}`);
          return fn(n.args, ev);
        }
      }
    };

    return ev.scalar(ast);
  }
}

function binary(op: string, a: Scalar, b: Scalar): Scalar {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  switch (op) {
    case "&": {
      const x = toText(a), y = toText(b);
      return isErr(x) ? x : isErr(y) ? y : x + y;
    }
    case "=": return compare(a, b) === 0;
    case "<>": return compare(a, b) !== 0;
    case "<": return compare(a, b) < 0;
    case ">": return compare(a, b) > 0;
    case "<=": return compare(a, b) <= 0;
    case ">=": return compare(a, b) >= 0;
  }
  const x = toNumber(a), y = toNumber(b);
  if (isErr(x)) return x;
  if (isErr(y)) return y;
  switch (op) {
    case "+": return x + y;
    case "-": return x - y;
    case "*": return x * y;
    case "/": return y === 0 ? new Err("#DIV/0!", "division by zero") : x / y;
    case "^": return x ** y;
  }
  return new Err("#ERROR!", `unknown operator ${op}`);
}
