// A Doc plus its formula engine: the object both server and browser hold.

import { clamp, type Range } from "./a1.ts";
import { displayValue } from "./display.ts";
import { Calc } from "./formula/calc.ts";
import { Err, type Scalar } from "./formula/values.ts";
import { docFromJSON, docToJSON, getSheet, key, newDoc, type Doc, type DocJSON, type Sheet } from "./model.ts";
import { applyAll, type Dirty, type Op } from "./ops.ts";

export class Workbook {
  readonly doc: Doc;
  readonly calc: Calc;

  constructor(doc: Doc = newDoc("Untitled")) {
    this.doc = doc;
    this.calc = new Calc(doc);
  }

  static fromJSON(j: DocJSON): Workbook {
    return new Workbook(docFromJSON(j));
  }

  toJSON(): DocJSON {
    return docToJSON(this.doc);
  }

  /** Apply ops atomically and keep computed values in step. Returns the inverse ops. */
  apply(ops: Op[]): { inverse: Op[]; dirty: Dirty } {
    const res = applyAll(this.doc, ops);
    if (res.dirty === "all") this.calc.reset();
    else if (res.dirty.length) this.calc.invalidate(res.dirty);
    return res;
  }

  value(sheetId: string, r: number, c: number): Scalar {
    return this.calc.value(sheetId, r, c);
  }

  display(sheet: Sheet, r: number, c: number): string {
    return displayValue(this.value(sheet.id, r, c), sheet.cells.get(key(r, c))?.f?.nf);
  }

  /** Read a block of cells, row-major (which also keeps long formula chains shallow). */
  read(sheetId: string, range: Range, what: "values" | "inputs" | "display"): Array<Array<Scalar | string>> {
    const sheet = getSheet(this.doc, sheetId);
    const g = clamp(range, Math.max(sheet.rows, range.r1 + 1), Math.max(sheet.cols, range.c1 + 1));
    const out: Array<Array<Scalar | string>> = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const row: Array<Scalar | string> = [];
      for (let c = g.c1; c <= g.c2; c++) {
        if (what === "inputs") row.push(sheet.cells.get(key(r, c))?.v ?? "");
        else if (what === "display") row.push(this.display(sheet, r, c));
        else {
          const v = this.value(sheet.id, r, c);
          row.push(v instanceof Err ? v.code : v);
        }
      }
      out.push(row);
    }
    return out;
  }
}
