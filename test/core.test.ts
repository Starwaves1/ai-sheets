import assert from "node:assert/strict";
import { test } from "node:test";
import { colIndex, colName, parseRange, rangeName } from "../src/core/a1.ts";
import { parseDelimited, toDelimited } from "../src/core/csv.ts";
import { displayValue, formatNumber } from "../src/core/display.ts";
import { moveRefs, shiftRefs } from "../src/core/formula/rewrite.ts";
import { newDoc } from "../src/core/model.ts";
import { transform, type Op } from "../src/core/ops.ts";
import { visibleRows } from "../src/core/view.ts";
import { Workbook } from "../src/core/workbook.ts";

/** Workbook with Sheet1 filled from an A1 → input map. */
function book(cells: Record<string, string>): Workbook {
  const wb = new Workbook(newDoc("t"));
  wb.apply([{ t: "set", s: "s1", cells: Object.entries(cells).map(([a, v]) => {
    const g = parseRange(a)!;
    return [g.r1, g.c1, v];
  }) }]);
  return wb;
}

const val = (wb: Workbook, a1: string) => {
  const g = parseRange(a1)!;
  const v = wb.value("s1", g.r1, g.c1);
  return v instanceof Object ? String(v) : v;
};

test("A1 names round-trip", () => {
  assert.equal(colName(0), "A");
  assert.equal(colName(25), "Z");
  assert.equal(colName(26), "AA");
  assert.equal(colName(16383), "XFD");
  assert.equal(colIndex("AB"), 27);
  assert.deepEqual(parseRange("B3"), { r1: 2, c1: 1, r2: 2, c2: 1 });
  assert.deepEqual(parseRange("C9:A1"), { r1: 0, c1: 0, r2: 8, c2: 2 });
  assert.equal(parseRange("'My sheet'!A1:B2")!.sheet, "My sheet");
  assert.equal(rangeName(parseRange("B:D")!), "B:D");
  assert.equal(rangeName(parseRange("2:4")!), "2:4");
  assert.equal(parseRange("A"), null);
  assert.equal(parseRange("A1:B"), null);
});

test("arithmetic, precedence, concat, comparison", () => {
  const wb = book({
    A1: "2", A2: "3", A3: "=A1+A2*2", A4: "=-A1^2", A5: '=A1&"x"&A2', A6: "=A1<A2", A7: "=50%",
    A8: "=2^3^2", A9: '="abc"="ABC"', A10: "=(A1+A2)*2",
  });
  assert.equal(val(wb, "A3"), 8);
  assert.equal(val(wb, "A4"), 4);
  assert.equal(val(wb, "A5"), "2x3");
  assert.equal(val(wb, "A6"), true);
  assert.equal(val(wb, "A7"), 0.5);
  assert.equal(val(wb, "A8"), 64);
  assert.equal(val(wb, "A9"), true);
  assert.equal(val(wb, "A10"), 10);
});

test("aggregate and logical functions", () => {
  const wb = book({
    A1: "1", A2: "2", A3: "x", A4: "4", A5: "",
    B1: "=SUM(A1:A5)", B2: "=AVERAGE(A1:A4)", B3: "=COUNT(A1:A5)", B4: "=COUNTA(A1:A5)",
    B5: "=MIN(A1:A4)", B6: "=MAX(A:A)", B7: '=IF(A1>0,"pos","neg")', B8: "=ROUND(2.345,2)",
    B9: "=AND(TRUE,A1=1)", B10: "=OR(FALSE,NOT(TRUE))", B11: "=IFERROR(1/0,\"none\")",
    B12: '=COUNTIF(A1:A4,">1")', B13: '=SUMIF(A1:A4,">=2")', B14: '=AVERAGEIF(A1:A4,"<>2")',
    B15: '=COUNTIF(C1:C3,"pa*")', C1: "pass", C2: "Pause", C3: "fail", B16: "=ROUND(-1.5)",
  });
  assert.equal(val(wb, "B1"), 7);
  assert.equal(val(wb, "B2"), 7 / 3);
  assert.equal(val(wb, "B3"), 3);
  assert.equal(val(wb, "B4"), 4);
  assert.equal(val(wb, "B5"), 1);
  assert.equal(val(wb, "B6"), 4);
  assert.equal(val(wb, "B7"), "pos");
  assert.equal(val(wb, "B8"), 2.35);
  assert.equal(val(wb, "B9"), true);
  assert.equal(val(wb, "B10"), false);
  assert.equal(val(wb, "B11"), "none");
  assert.equal(val(wb, "B12"), 2);
  assert.equal(val(wb, "B13"), 6);
  assert.equal(val(wb, "B14"), 2.5);
  assert.equal(val(wb, "B15"), 2);
  assert.equal(val(wb, "B16"), -2);
});

test("errors: #DIV/0!, #NAME?, #REF!, #VALUE!, #CIRC!", () => {
  const wb = book({
    A1: "=1/0", A2: "=NOPE(1)", A3: "=Missing!A1", A4: '="a"+1', A5: "=A6", A6: "=A5", A7: "=A1+1", A8: "=foo",
    A9: "=1+", A10: "=A10",
  });
  assert.equal(val(wb, "A1"), "#DIV/0!");
  assert.equal(val(wb, "A2"), "#NAME?");
  assert.equal(val(wb, "A3"), "#REF!");
  assert.equal(val(wb, "A4"), "#VALUE!");
  assert.equal(val(wb, "A5"), "#CIRC!");
  assert.equal(val(wb, "A7"), "#DIV/0!");
  assert.equal(val(wb, "A8"), "#NAME?");
  assert.equal(val(wb, "A9"), "#ERROR!");
  assert.equal(val(wb, "A10"), "#CIRC!");
});

test("dependency tracking recalculates dependents, including through ranges", () => {
  const wb = book({ A1: "1", A2: "=A1*10", A3: "=SUM(A1:A2)", B1: "=A3+1" });
  assert.equal(val(wb, "B1"), 12);
  wb.apply([{ t: "set", s: "s1", cells: [[0, 0, "5"]] }]);
  assert.equal(val(wb, "A2"), 50);
  assert.equal(val(wb, "A3"), 55);
  assert.equal(val(wb, "B1"), 56);
  // Breaking a cycle clears the error.
  wb.apply([{ t: "set", s: "s1", cells: [[0, 0, "=B1"]] }]);
  assert.equal(val(wb, "B1"), "#CIRC!");
  wb.apply([{ t: "set", s: "s1", cells: [[0, 0, "2"]] }]);
  assert.equal(val(wb, "B1"), 23);
});

test("long dependency chains evaluate without blowing the stack", () => {
  const cells: Record<string, string> = { A1: "1" };
  for (let r = 2; r <= 3000; r++) cells[`A${r}`] = `=A${r - 1}+1`;
  const wb = book(cells);
  assert.equal(val(wb, "A3000"), 3000);
});

test("cross-sheet references follow renames", () => {
  const wb = book({ A1: "=Data!B2*2" });
  wb.apply([{ t: "addSheet", sheet: { id: "s2", name: "Data", rows: 10, cols: 5, cells: [[1, 1, "21"]] } }]);
  assert.equal(val(wb, "A1"), 42);
  wb.apply([{ t: "renameSheet", s: "s2", name: "My data" }]);
  assert.equal(wb.doc.sheets[0]!.cells.get(0)!.v, "='My data'!B2*2");
  assert.equal(val(wb, "A1"), 42);
});

test("insert and delete rows rewrite formulas; undo restores them", () => {
  const wb = book({ A1: "1", A2: "2", A3: "3", A4: "=SUM(A1:A3)", B1: "=A3" });
  const ins = wb.apply([{ t: "insert", s: "s1", axis: "row", at: 1, n: 2 }]);
  assert.equal(val(wb, "A6"), 6);
  assert.equal(wb.doc.sheets[0]!.cells.get(16384 * 5)!.v, "=SUM(A1:A5)");
  wb.apply(ins.inverse);
  assert.equal(val(wb, "A4"), 6);
  const del = wb.apply([{ t: "delete", s: "s1", axis: "row", at: 2, n: 1 }]);
  assert.equal(val(wb, "B1"), "#REF!");
  assert.equal(val(wb, "A3"), 3);
  wb.apply(del.inverse);
  assert.equal(val(wb, "B1"), 3);
  assert.equal(val(wb, "A4"), 6);
});

test("ref rewriting: shift and move", () => {
  assert.equal(shiftRefs("=A1+B5+$C$9", "S", "S", "row", 2, 3), "=A1+B8+$C$12");
  assert.equal(shiftRefs("=SUM(A1:A10)", "S", "S", "row", 0, -2), "=SUM(A1:A8)");
  assert.equal(shiftRefs("=Other!A5+A5", "S", "Other", "row", 0, 1), "=Other!A6+A5");
  assert.equal(shiftRefs("=SUM(B:B)", "S", "S", "row", 0, 5), "=SUM(B:B)");
  assert.equal(moveRefs("=A1+$B$2+B$3", 2, 1), "=B3+$B$2+C$3");
  assert.equal(moveRefs("=A1", -1, 0), "=#REF!");
});

test("transform shifts concurrent cell edits past a row insert", () => {
  const doc = newDoc("t");
  const ins: Op = { t: "insert", s: "s1", axis: "row", at: 3, n: 2 };
  const edit: Op = { t: "set", s: "s1", cells: [[1, 0, "keep"], [5, 0, "=A5"]] };
  assert.deepEqual(transform(edit, ins, doc), { t: "set", s: "s1", cells: [[1, 0, "keep"], [7, 0, "=A7"]] });
  const del: Op = { t: "delete", s: "s1", axis: "row", at: 1, n: 1 };
  assert.deepEqual(transform(edit, del, doc), { t: "set", s: "s1", cells: [[4, 0, "=A4"]] });
});

test("format, conditional rules and number formats", () => {
  assert.equal(formatNumber(0.3333, "0.0%"), "33.3%");
  assert.equal(formatNumber(1234567.891, "#,##0.00"), "1,234,567.89");
  assert.equal(formatNumber(-5, "$#,##0"), "-$5");
  assert.equal(formatNumber(0.1 + 0.2), "0.3");
  assert.equal(displayValue(true), "TRUE");
  const wb = book({ A1: "1", A2: "0", A3: "x" });
  wb.apply([{ t: "cond", s: "s1", rules: [
    { range: { r1: 0, c1: 0, r2: 9, c2: 0 }, op: "eq", value: "1", style: { fill: "green" } },
    { range: { r1: 0, c1: 0, r2: 9, c2: 0 }, op: "eq", value: "x", style: { fill: "grey" } },
  ] }]);
  wb.apply([{ t: "style", s: "s1", range: { r1: 0, c1: 0, r2: 1, c2: 0 }, patch: { b: true } }]);
  const sheet = wb.doc.sheets[0]!;
  assert.equal(sheet.cells.get(0)!.f!.b, true);
  assert.equal(sheet.cond[0]!.style.fill, "green");
});

test("view sort and filter reorder display rows only", () => {
  const wb = book({ A1: "name", A2: "b", A3: "c", A4: "a", B1: "n", B2: "2", B3: "3", B4: "1" });
  const sheet = wb.doc.sheets[0]!;
  wb.apply([{ t: "grid", s: "s1", rows: 4, cols: 2 }]);
  wb.apply([{ t: "view", s: "s1", view: { headerRow: 0, sort: { col: 1, desc: true }, filters: { 0: { hide: ["c"] } } } }]);
  assert.deepEqual(visibleRows(sheet, (r, c) => wb.value("s1", r, c)), [0, 1, 3]);
  assert.equal(sheet.cells.get(16384)!.v, "b"); // data did not move
});

test("CSV and TSV round-trip", () => {
  const rows = [["a", "b,c", 'say "hi"'], ["multi\nline", "", "x"]];
  assert.deepEqual(parseDelimited(toDelimited(rows)), rows);
  assert.deepEqual(parseDelimited("a\tb\n1\t2\n"), [["a", "b"], ["1", "2"]]);
});

test("failed op batches roll back atomically", () => {
  const wb = book({ A1: "1" });
  assert.throws(() => wb.apply([
    { t: "set", s: "s1", cells: [[0, 0, "changed"]] },
    { t: "deleteSheet", s: "s1" },
  ]));
  assert.equal(val(wb, "A1"), 1);
});
