// Pratt parser: formula text (without the leading "=") to an AST.
// Precedence follows spreadsheets: comparison < & < +- < */ < ^ < unary - < %.

import { MAX_COLS, MAX_ROWS } from "../a1.ts";
import { tokenize, type Tok } from "./lexer.ts";

export type Node =
  | { k: "num"; v: number }
  | { k: "str"; v: string }
  | { k: "bool"; v: boolean }
  | { k: "err"; v: string }
  | { k: "name"; v: string }
  | { k: "ref"; sheet?: string; r1: number; c1: number; r2: number; c2: number; area: boolean }
  | { k: "un"; op: string; a: Node }
  | { k: "bin"; op: string; a: Node; b: Node }
  | { k: "call"; name: string; args: Node[] };

export class ParseError extends Error {}

const BIN: Record<string, number> = {
  "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
  "&": 2,
  "+": 3, "-": 3,
  "*": 4, "/": 4,
  "^": 5,
};
const PREFIX = 6;
const POSTFIX = 7;

function refNode(t: Extract<Tok, { t: "ref" }>): Node {
  const { a, b } = t;
  const r1 = a.r ?? 0;
  const c1 = a.c ?? 0;
  let r2 = r1;
  let c2 = c1;
  if (b) {
    r2 = b.r ?? MAX_ROWS - 1;
    c2 = b.c ?? MAX_COLS - 1;
  }
  const n: Node = {
    k: "ref",
    r1: Math.min(r1, r2), c1: Math.min(c1, c2),
    r2: Math.max(r1, r2), c2: Math.max(c1, c2),
    area: b !== undefined,
  };
  if (t.sheet !== undefined) n.sheet = t.sheet;
  return n;
}

export function parseFormula(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p]!;
  const next = () => toks[p++]!;

  function expr(minPrec: number): Node {
    let left = prefix();
    for (;;) {
      const t = peek();
      if (t.t !== "op") break;
      if (t.v === "%") {
        if (POSTFIX < minPrec) break;
        next();
        left = { k: "un", op: "%", a: left };
        continue;
      }
      const prec = BIN[t.v];
      if (prec === undefined || prec < minPrec) break;
      next();
      // ^ is left-associative in spreadsheets: 2^3^2 = 64.
      const right = expr(prec + 1);
      left = { k: "bin", op: t.v, a: left, b: right };
    }
    return left;
  }

  function prefix(): Node {
    const t = next();
    switch (t.t) {
      case "num": return { k: "num", v: t.v };
      case "str": return { k: "str", v: t.v };
      case "bool": return { k: "bool", v: t.v };
      case "err": return { k: "err", v: t.v };
      case "name": return { k: "name", v: t.v };
      case "ref": return refNode(t);
      case "op":
        if (t.v === "-" || t.v === "+") return { k: "un", op: t.v, a: expr(PREFIX) };
        throw new ParseError(`unexpected '${t.v}'`);
      case "(": {
        const e = expr(0);
        if (next().t !== ")") throw new ParseError("missing )");
        return e;
      }
      case "fn": {
        if (next().t !== "(") throw new ParseError("expected (");
        const args: Node[] = [];
        if (peek().t === ")") {
          next();
          return { k: "call", name: t.v, args };
        }
        for (;;) {
          // Empty argument, as in IF(A1,,2), is an empty string / zero.
          if (peek().t === "," || peek().t === ")") args.push({ k: "str", v: "" });
          else args.push(expr(0));
          const sep = next();
          if (sep.t === ")") break;
          if (sep.t !== ",") throw new ParseError("expected , or )");
        }
        return { k: "call", name: t.v, args };
      }
      case "end": throw new ParseError("unexpected end of formula");
      default: throw new ParseError(`unexpected '${t.t}'`);
    }
  }

  const root = expr(0);
  if (peek().t !== "end") throw new ParseError(`unexpected text at ${peek().s + 1}`);
  return root;
}
