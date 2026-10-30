import assert from "node:assert/strict";
import { test } from "node:test";
import { parseRange } from "../src/core/a1.ts";
import { autofillOps, extendSeries } from "../src/core/fill.ts";
import { newDoc } from "../src/core/model.ts";
import { Workbook } from "../src/core/workbook.ts";

test("numeric series continue with their step; a single number copies unless toggled", () => {
  assert.deepEqual(extendSeries(["1", "2", "3"], 3), ["4", "5", "6"]);
  assert.deepEqual(extendSeries(["2", "4"], 2), ["6", "8"]);
  assert.deepEqual(extendSeries(["10", "7"], 2), ["4", "1"]);
  assert.deepEqual(extendSeries(["0.1", "0.2"], 2), ["0.3", "0.4"]);
  assert.deepEqual(extendSeries(["5"], 3), ["5", "5", "5"]);
  assert.deepEqual(extendSeries(["5"], 3, true), ["6", "7", "8"]);
  assert.deepEqual(extendSeries(["1", "2", "3"], 4, true), ["1", "2", "3", "1"]);
  assert.deepEqual(extendSeries(["1", "2", "4"], 4), ["1", "2", "4", "1"]); // no constant step: repeat
});

test("text with numbers, weekdays and months, cycles", () => {
  assert.deepEqual(extendSeries(["Item 1", "Item 2"], 2), ["Item 3", "Item 4"]);
  assert.deepEqual(extendSeries(["Item 7"], 2), ["Item 8", "Item 9"]);
  assert.deepEqual(extendSeries(["run-01"], 2), ["run-02", "run-03"]);
  assert.deepEqual(extendSeries(["run-09"], 2), ["run-10", "run-11"]);
  assert.deepEqual(extendSeries(["1st"], 2), ["2st", "3st"]);
  assert.deepEqual(extendSeries(["Monday"], 2), ["Tuesday", "Wednesday"]);
  assert.deepEqual(extendSeries(["SAT"], 3), ["SUN", "MON", "TUE"]);
  assert.deepEqual(extendSeries(["jan", "mar"], 2), ["may", "jul"]);
  assert.deepEqual(extendSeries(["November"], 3), ["December", "January", "February"]);
  assert.deepEqual(extendSeries(["a", "b", "c"], 4), ["a", "b", "c", "a"]);
  assert.deepEqual(extendSeries(["x", null], 3), ["x", null, "x"]);
});

const ref = (a: string) => {
  const g = parseRange(a)!;
  return { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 };
};

test("autofill down, up and right on a sheet: series, formulas, formats, conditional rules", () => {
  const wb = new Workbook(newDoc("t"));
  const s = wb.doc.sheets[0]!;
  wb.apply([
    { t: "set", s: "s1", cells: [[4, 0, "1"], [5, 0, "2"], [6, 0, "3"], [4, 1, "=A5*$A$5"], [5, 1, "=A6*$A$5"], [6, 1, "=A7*$A$5"]] },
    { t: "fmt", s: "s1", cells: [[4, 0, { b: true }], [5, 0, { fill: "green" }]] },
    { t: "cond", s: "s1", rules: [{ range: ref("A5:A7"), op: "gt", value: "2", style: { fill: "red" } }] },
  ]);
  wb.apply(autofillOps(s, ref("A5:B7"), ref("A8:B10")));
  const inputs = (a: string) => wb.read("s1", ref(a), "inputs").map((r) => r.join("|"));
  assert.deepEqual(inputs("A8:B10"), ["4|=A8*$A$5", "5|=A9*$A$5", "6|=A10*$A$5"]);
  assert.equal(s.cells.get(7 * 16384)?.f?.b, true); // A8 takes A5's format (cycle of 3)
  assert.equal(s.cells.get(8 * 16384)?.f?.fill, "green");
  assert.deepEqual(s.cond[0]!.range, ref("A5:A10"));

  wb.apply(autofillOps(s, ref("A5:A7"), ref("A2:A4"))); // up: 0, -1, -2 reading upward
  assert.deepEqual(inputs("A2:A4"), ["-2", "-1", "0"]);

  wb.apply([{ t: "set", s: "s1", cells: [[0, 3, "Mon"], [0, 4, "Tue"]] }]);
  wb.apply(autofillOps(s, ref("D1:E1"), ref("F1:H1")));
  assert.deepEqual(inputs("D1:H1"), ["Mon|Tue|Wed|Thu|Fri"]);
  wb.apply(autofillOps(s, ref("D1:E1"), ref("B1:C1"))); // left
  assert.deepEqual(inputs("B1:C1"), ["Sat|Sun"]);
  assert.throws(() => autofillOps(s, ref("D1:E1"), ref("D3:E3")));
});
