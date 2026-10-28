// Formula tokenizer. Positions are kept so references can be rewritten in place
// (row/column insert and delete, copy-paste offsets, sheet renames).

import { colIndex, MAX_COLS, MAX_ROWS } from "../a1.ts";

/** One end of a reference. Missing r means a whole-column ref (A:A); missing c a whole-row ref (3:3). */
export interface RefEnd {
  r?: number;
  c?: number;
  ra: boolean; // $ on row
  ca: boolean; // $ on column
}

export type Tok =
  | { t: "num"; v: number; s: number; e: number }
  | { t: "str"; v: string; s: number; e: number }
  | { t: "bool"; v: boolean; s: number; e: number }
  | { t: "err"; v: string; s: number; e: number }
  | { t: "ref"; sheet?: string; a: RefEnd; b?: RefEnd; s: number; e: number }
  | { t: "fn"; v: string; s: number; e: number }
  | { t: "name"; v: string; s: number; e: number }
  | { t: "op"; v: string; s: number; e: number }
  | { t: "(" | ")" | "," | "end"; s: number; e: number };

const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_.]*))!`;
const CELLP = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
const REF_RE = new RegExp(
  `(?:${SHEET})?(${CELLP}(?::${CELLP})?|\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}|\\$?\\d+:\\$?\\d+)(?![A-Za-z0-9_(!])`,
  "y",
);
const NUM_RE = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const NAME_RE = /[A-Za-z_][A-Za-z0-9_.]*/y;
const ERR_RE = /#(?:REF!|DIV\/0!|NAME\?|VALUE!|N\/A|NUM!|NULL!|CIRC!|ERROR!)/y;
const END_RE = /^(\$?)([A-Za-z]{1,3})?(\$?)(\d+)?$/;

function refEnd(text: string): RefEnd | null {
  const m = END_RE.exec(text);
  if (!m) return null;
  const end: RefEnd = { ca: m[1] === "$", ra: m[3] === "$" };
  if (m[2]) {
    const c = colIndex(m[2]);
    if (c < 0 || c >= MAX_COLS) return null;
    end.c = c;
  }
  if (m[4]) {
    const r = Number(m[4]) - 1;
    if (r < 0 || r >= MAX_ROWS) return null;
    end.r = r;
  }
  // "$3" alone parses as column "$"+row: fix flags for whole-row refs.
  if (end.c === undefined && end.ca) {
    end.ca = false;
    end.ra = true;
  }
  return end;
}

export class LexError extends Error {}

export function tokenize(src: string, start = 0): Tok[] {
  const out: Tok[] = [];
  let i = start;
  while (i < src.length) {
    const ch = src[i]!;
    if (ch === " " || ch === "\t" || ch === "\n" || ch === "\r") {
      i++;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let v = "";
      for (;;) {
        if (j >= src.length) throw new LexError("unterminated string");
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            v += '"';
            j += 2;
            continue;
          }
          break;
        }
        v += src[j++];
      }
      out.push({ t: "str", v, s: i, e: j + 1 });
      i = j + 1;
      continue;
    }
    if (ch === "#") {
      ERR_RE.lastIndex = i;
      const m = ERR_RE.exec(src);
      if (!m) throw new LexError(`unknown error literal at ${i}`);
      out.push({ t: "err", v: m[0], s: i, e: i + m[0].length });
      i += m[0].length;
      continue;
    }
    REF_RE.lastIndex = i;
    let m = REF_RE.exec(src);
    if (m) {
      const [aText, bText] = m[3]!.split(":");
      const a = refEnd(aText!);
      const b = bText === undefined ? undefined : refEnd(bText);
      if (a && (bText === undefined || b)) {
        const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
        const tok: Tok = { t: "ref", a, s: i, e: i + m[0].length };
        if (b) tok.b = b;
        if (sheet !== undefined) tok.sheet = sheet;
        out.push(tok);
        i += m[0].length;
        continue;
      }
    }
    NUM_RE.lastIndex = i;
    m = NUM_RE.exec(src);
    if (m) {
      out.push({ t: "num", v: Number(m[0]), s: i, e: i + m[0].length });
      i += m[0].length;
      continue;
    }
    NAME_RE.lastIndex = i;
    m = NAME_RE.exec(src);
    if (m) {
      const name = m[0];
      let j = i + name.length;
      while (src[j] === " ") j++;
      const upper = name.toUpperCase();
      if (src[j] === "(") out.push({ t: "fn", v: upper, s: i, e: i + name.length });
      else if (upper === "TRUE" || upper === "FALSE") out.push({ t: "bool", v: upper === "TRUE", s: i, e: i + name.length });
      else out.push({ t: "name", v: name, s: i, e: i + name.length });
      i += name.length;
      continue;
    }
    const two = src.slice(i, i + 2);
    if (two === "<=" || two === ">=" || two === "<>") {
      out.push({ t: "op", v: two, s: i, e: i + 2 });
      i += 2;
      continue;
    }
    if ("+-*/^&=<>%".includes(ch)) {
      out.push({ t: "op", v: ch, s: i, e: i + 1 });
      i++;
      continue;
    }
    if (ch === "(" || ch === ")" || ch === ",") {
      out.push({ t: ch, s: i, e: i + 1 });
      i++;
      continue;
    }
    throw new LexError(`unexpected character '${ch}' at ${i}`);
  }
  out.push({ t: "end", s: src.length, e: src.length });
  return out;
}
