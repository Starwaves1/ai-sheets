// Shape checks for ops arriving over the network. Semantic checks (sheet exists,
// span in range, names unique) happen in apply(); this only guarantees the types.

import { COND_OPS, FORMAT_KEYS, OpError } from "./model.ts";
import type { Op } from "./ops.ts";

type Obj = Record<string, unknown>;

const isObj = (x: unknown): x is Obj => typeof x === "object" && x !== null && !Array.isArray(x);
const isInt = (x: unknown): x is number => Number.isInteger(x);
const isStr = (x: unknown): x is string => typeof x === "string";

function need(cond: unknown, msg: string): asserts cond {
  if (!cond) throw new OpError(msg);
}

function range(x: unknown, what: string): void {
  need(isObj(x) && isInt(x.r1) && isInt(x.c1) && isInt(x.r2) && isInt(x.c2), `${what}: range needs integer r1,c1,r2,c2`);
}

function format(x: unknown, what: string): void {
  need(isObj(x), `${what}: format must be an object`);
  for (const k of Object.keys(x)) need((FORMAT_KEYS as readonly string[]).includes(k), `${what}: unknown format key ${k}`);
}

const AXES = ["row", "col"];

export function validateOps(raw: unknown): Op[] {
  need(Array.isArray(raw), "ops must be an array");
  need(raw.length <= 10_000, "at most 10000 ops per change");
  raw.forEach((op, i) => validateOp(op, `ops[${i}]`));
  return raw as Op[];
}

function validateOp(op: unknown, at: string): void {
  need(isObj(op) && isStr(op.t), `${at}: op needs a string t`);
  const sheetOp = op.t !== "addSheet" && op.t !== "rename";
  if (sheetOp) need(isStr(op.s), `${at}: op needs sheet id s`);
  switch (op.t) {
    case "set":
      need(Array.isArray(op.cells), `${at}: cells must be an array`);
      for (const c of op.cells as unknown[]) {
        need(Array.isArray(c) && isInt(c[0]) && isInt(c[1]) && (c[2] === null || isStr(c[2])), `${at}: cells are [row, col, text|null]`);
      }
      return;
    case "fmt":
      need(Array.isArray(op.cells), `${at}: cells must be an array`);
      for (const c of op.cells as unknown[]) {
        need(Array.isArray(c) && isInt(c[0]) && isInt(c[1]), `${at}: cells are [row, col, format|null]`);
        if (c[2] !== null) format(c[2], at);
      }
      return;
    case "style":
      range(op.range, at);
      format(op.patch, at);
      return;
    case "insert":
    case "delete":
      need(AXES.includes(op.axis as string) && isInt(op.at) && isInt(op.n), `${at}: needs axis, at, n`);
      return;
    case "size":
      need(AXES.includes(op.axis as string) && Array.isArray(op.sizes), `${at}: needs axis and sizes`);
      for (const p of op.sizes as unknown[]) {
        need(Array.isArray(p) && isInt(p[0]) && (p[1] === null || typeof p[1] === "number"), `${at}: sizes are [index, px|null]`);
      }
      return;
    case "hide":
      need(AXES.includes(op.axis as string) && Array.isArray(op.idx) && (op.idx as unknown[]).every(isInt) && typeof op.hidden === "boolean", `${at}: needs axis, idx[], hidden`);
      return;
    case "freeze":
    case "grid":
      need(isInt(op.rows) && isInt(op.cols), `${at}: needs integer rows and cols`);
      return;
    case "view":
      if (op.view !== null) {
        need(isObj(op.view) && isInt(op.view.headerRow) && isObj(op.view.filters), `${at}: view needs headerRow and filters`);
        const sort = op.view.sort;
        if (sort !== undefined) need(isObj(sort) && isInt(sort.col), `${at}: sort needs col`);
        for (const f of Object.values(op.view.filters)) {
          need(isObj(f), `${at}: filter must be an object`);
          if (f.hide !== undefined) need(Array.isArray(f.hide) && f.hide.every(isStr), `${at}: filter hide must be strings`);
          if (f.op !== undefined) need(COND_OPS.includes(f.op as never), `${at}: unknown filter op`);
        }
      }
      return;
    case "cond":
      need(Array.isArray(op.rules), `${at}: rules must be an array`);
      for (const r of op.rules as unknown[]) {
        need(isObj(r) && COND_OPS.includes(r.op as never) && isObj(r.style), `${at}: rule needs op and style`);
        range(r.range, at);
      }
      return;
    case "addSheet":
      need(isObj(op.sheet) && isStr(op.sheet.id) && isStr(op.sheet.name) && Array.isArray(op.sheet.cells)
        && isInt(op.sheet.rows) && isInt(op.sheet.cols), `${at}: sheet needs id, name, rows, cols, cells`);
      return;
    case "deleteSheet":
      return;
    case "renameSheet":
    case "rename":
      need(isStr(op.name), `${at}: needs name`);
      return;
    case "moveSheet":
      need(isInt(op.index), `${at}: needs index`);
      return;
    default:
      throw new OpError(`${at}: unknown op type ${String(op.t)}`);
  }
}
