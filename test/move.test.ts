import assert from "node:assert/strict";
import { test } from "node:test";
import { docToJSON, newDoc, newSheet, sheetToJSON } from "../src/core/model.ts";
import { moveIndex, transform, type Op } from "../src/core/ops.ts";
import { Workbook } from "../src/core/workbook.ts";

function sample(): Workbook {
  const wb = new Workbook(newDoc("t"));
  wb.apply([
    { t: "set", s: "s1", cells: Array.from({ length: 8 }, (_, r) => [r, 0, String(r + 1)] as [number, number, string]).concat([[0, 1, "=A2*10"], [1, 1, "=SUM(A1:A8)"], [2, 1, "=A$3+$A1"]]) },
    { t: "fmt", s: "s1", cells: [[3, 0, { b: true }]] },
    { t: "flag", s: "s1", cells: [[4, 0, { note: "n", by: "x", ts: 1 }]] },
    { t: "size", s: "s1", axis: "row", sizes: [[5, 40]] },
    { t: "size", s: "s1", axis: "col", sizes: [[1, 150]] },
    { t: "hide", s: "s1", axis: "row", idx: [6], hidden: true },
    { t: "cond", s: "s1", rules: [{ range: { r1: 1, c1: 0, r2: 4, c2: 0 }, op: "gt", value: "3", style: { fill: "red" } }] },
    { t: "addSheet", sheet: sheetToJSON(newSheet("s2", "Other")) },
    { t: "set", s: "s2", cells: [[0, 0, "=Sheet1!A2"], [1, 0, "=SUM(Sheet1!A1:A8)"], [2, 0, "=Sheet1!B1"]] },
  ]);
  return wb;
}

const col = (wb: Workbook, s: string, c: number, n: number) => Array.from({ length: n }, (_, r) => wb.value(s, r, c));

test("moving rows carries values, formats, flags, sizes, hidden rows; references follow", () => {
  const wb = sample();
  const s = wb.doc.sheets[0]!;
  wb.apply([{ t: "move", s: "s1", axis: "row", from: 1, n: 1, to: 5 }]); // row 2 goes below row 5
  assert.deepEqual(col(wb, "s1", 0, 8), [1, 3, 4, 5, 2, 6, 7, 8]);
  assert.equal(s.cells.get(2 * 16384)?.f?.b, true); // was row 4, now row 3
  assert.equal(s.cells.get(3 * 16384)?.flag?.note, "n");
  assert.equal(s.rowH.get(5), 40);
  assert.ok(s.hiddenRows.has(6));
  assert.equal(wb.value("s2", 0, 0), 2); // =Sheet1!A2 followed the 2 to A5
  assert.equal(wb.doc.sheets[1]!.cells.get(0)?.v, "=Sheet1!A5");
  assert.equal(wb.value("s2", 1, 0), 36); // the range still covers all eight rows
  assert.equal(wb.value("s1", 0, 1), 20); // B1 (=A2*10) still points at the 2
  assert.deepEqual(s.cond[0]!.range, { r1: 1, c1: 0, r2: 4, c2: 0 });
});

test("moves undo exactly, including formula text, conditional ranges and columns", () => {
  const seed = docToJSON(sample().doc);
  let x = 7;
  const rand = (m: number) => (x = (x * 1103515245 + 12345) % 2 ** 31) % m;
  for (let k = 0; k < 200; k++) {
    const wb = sample();
    const axis = k % 2 ? "row" : "col";
    const count = axis === "row" ? wb.doc.sheets[0]!.rows : wb.doc.sheets[0]!.cols;
    const lim = axis === "row" ? 10 : 4;
    const n = 1 + rand(3), from = rand(lim);
    let to = rand(lim + 4);
    if (to >= from && to <= from + n) to = from + n + 1 + rand(3);
    if (from + n > count || to > count) continue;
    const { inverse } = wb.apply([{ t: "move", s: "s1", axis, from, n, to }]);
    wb.apply(inverse);
    assert.deepEqual(docToJSON(wb.doc), seed, `move ${axis} ${from}+${n} -> ${to}`);
  }
});

test("moveIndex is a permutation; no-op moves are refused", () => {
  for (const [from, n, to] of [[2, 3, 0], [2, 3, 9], [0, 1, 10], [7, 3, 2]] as const) {
    const out = Array.from({ length: 10 }, (_, i) => moveIndex(i, from, n, to)).sort((a, b) => a - b);
    assert.deepEqual(out, Array.from({ length: 10 }, (_, i) => i));
  }
  const wb = sample();
  assert.throws(() => wb.apply([{ t: "move", s: "s1", axis: "row", from: 2, n: 2, to: 4 }]), /leaves everything/);
});

test("transform past a move: cells follow, a split delete is dropped, concurrent moves rebase", () => {
  const wb = sample();
  const m: Op = { t: "move", s: "s1", axis: "row", from: 1, n: 2, to: 6 }; // rows 2-3 go below row 6
  assert.deepEqual(transform({ t: "set", s: "s1", cells: [[1, 0, "x"], [3, 0, "y"]] }, m, wb.doc), { t: "set", s: "s1", cells: [[4, 0, "x"], [1, 0, "y"]] });
  assert.equal(transform({ t: "delete", s: "s1", axis: "row", at: 2, n: 2 }, m, wb.doc), null);
  assert.deepEqual(transform({ t: "delete", s: "s1", axis: "row", at: 3, n: 2 }, m, wb.doc), { t: "delete", s: "s1", axis: "row", at: 1, n: 2 });
  assert.deepEqual(transform({ t: "move", s: "s1", axis: "row", from: 7, n: 1, to: 0 }, m, wb.doc), { t: "move", s: "s1", axis: "row", from: 7, n: 1, to: 0 });
  // A move made before a concurrent insert above it shifts down.
  assert.deepEqual(transform({ t: "move", s: "s1", axis: "row", from: 4, n: 1, to: 1 }, { t: "insert", s: "s1", axis: "row", at: 0, n: 2 }, wb.doc),
    { t: "move", s: "s1", axis: "row", from: 6, n: 1, to: 3 });
  assert.equal(transform({ t: "move", s: "s1", axis: "row", from: 4, n: 1, to: 1 }, { t: "delete", s: "s1", axis: "row", at: 4, n: 1 }, wb.doc), null);
});
