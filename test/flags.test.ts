import assert from "node:assert/strict";
import { test } from "node:test";
import { flagGroups } from "../src/core/flags.ts";
import { docFromJSON, docToJSON, newDoc, type Flag } from "../src/core/model.ts";
import { transform, type Op } from "../src/core/ops.ts";
import { validateOps } from "../src/core/validate.ts";
import { Workbook } from "../src/core/workbook.ts";

const f: Flag = { note: "check", by: "Garrett", ts: 1 };

test("flags: set, undo, survive clearing values, round-trip through JSON", () => {
  const wb = new Workbook(newDoc("t"));
  const s = wb.doc.sheets[0]!;
  wb.apply([{ t: "set", s: "s1", cells: [[1, 1, "3"]] }]);
  const { inverse } = wb.apply([{ t: "flag", s: "s1", cells: [[1, 1, f], [5, 5, f]] }]);
  assert.equal(s.cells.get(16384 + 1)?.flag?.note, "check");
  wb.apply([{ t: "set", s: "s1", cells: [[1, 1, null]] }]); // clearing the value keeps the flag
  assert.equal(s.cells.get(16384 + 1)?.flag?.note, "check");
  const back = docFromJSON(docToJSON(wb.doc));
  assert.deepEqual(back.sheets[0]!.cells.get(5 * 16384 + 5)?.flag, f);
  wb.apply(inverse);
  assert.equal(s.cells.get(16384 + 1), undefined);
  assert.equal(s.cells.size, 0);
});

test("flags move with inserted rows, come back when a deleted row is restored, and transform like cells", () => {
  const wb = new Workbook(newDoc("t"));
  const s = wb.doc.sheets[0]!;
  wb.apply([{ t: "flag", s: "s1", cells: [[3, 0, f]] }]);
  wb.apply([{ t: "insert", s: "s1", axis: "row", at: 0, n: 2 }]);
  assert.equal(s.cells.get(5 * 16384)?.flag?.note, "check");
  const del = wb.apply([{ t: "delete", s: "s1", axis: "row", at: 5, n: 1 }]);
  assert.equal(flagGroups(wb.doc).length, 0);
  wb.apply(del.inverse);
  assert.equal(s.cells.get(5 * 16384)?.flag?.note, "check");
  const op: Op = { t: "flag", s: "s1", cells: [[4, 2, f]] };
  assert.deepEqual(transform(op, { t: "insert", s: "s1", axis: "row", at: 1, n: 3 }), { t: "flag", s: "s1", cells: [[7, 2, f]] });
  assert.equal(transform(op, { t: "delete", s: "s1", axis: "row", at: 4, n: 1 }), null);
});

test("flag groups: one per block flagged together; a pasted copy is separate", () => {
  const wb = new Workbook(newDoc("t"));
  wb.apply([{ t: "flag", s: "s1", cells: [[0, 0, f], [0, 1, f], [1, 0, f], [1, 1, f], [9, 0, f]] }]);
  const groups = flagGroups(wb.doc).map((g) => `${g.ref}:${g.cells.length}`).sort();
  assert.deepEqual(groups, ["Sheet1!A10:1", "Sheet1!A1:B2:4"]);
});

test("flag ops are validated", () => {
  assert.throws(() => validateOps([{ t: "flag", s: "s1", cells: [[0, 0, { note: 1 }]] }]));
  assert.throws(() => validateOps([{ t: "flag", s: "s1", cells: [[0, 0, { ...f, done: { by: "x" } }]] }]));
  validateOps([{ t: "flag", s: "s1", cells: [[0, 0, { ...f, done: { by: "claude", ts: 2, reply: "ok" } }], [0, 1, null]] }]);
});
