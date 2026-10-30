import assert from "node:assert/strict";
import { test } from "node:test";
import { parseRange } from "../src/core/a1.ts";
import { changedCells, describe, groupSessions, restoreOps, revertOps, summarize, touches, type Change } from "../src/core/history.ts";
import { docFromJSON, docToJSON, newDoc, type DocJSON } from "../src/core/model.ts";
import { transformAll, type Op } from "../src/core/ops.ts";
import { Workbook } from "../src/core/workbook.ts";

/** A tiny server: applies changes in order, keeps inverses and a JSON copy of every version. */
class Log {
  wb = new Workbook(newDoc("t"));
  changes: Change[] = [];
  versions: DocJSON[] = [docToJSON(this.wb.doc)];
  t = 1_000_000;
  edit(actor: string, ops: Op[], gapMs = 1000): number {
    const { inverse } = this.wb.apply(ops);
    const c: Change = { v: this.changes.length + 1, ops, inv: inverse, actor, client: actor, ts: (this.t += gapMs) };
    this.changes.push(c);
    this.versions.push(docToJSON(this.wb.doc));
    return c.v;
  }
  set(actor: string, cells: Record<string, string | null>): number {
    return this.edit(actor, [{ t: "set", s: "s1", cells: Object.entries(cells).map(([a, v]) => [parseRange(a)!.r1, parseRange(a)!.c1, v]) }]);
  }
  get(a: string): string {
    const g = parseRange(a)!;
    return this.wb.doc.sheets[0]!.cells.get(g.r1 * 16384 + g.c1)?.v ?? "";
  }
  revert(versions: number[], force = false, actor = "garrett") {
    const from = Math.min(...versions);
    const plan = revertOps(this.changes.filter((c) => c.v >= from), new Set(versions), this.wb.doc, force);
    this.edit(actor, plan.ops);
    return plan;
  }
}

test("revert an agent's session while keeping a person's later edits, across an inserted row", () => {
  const L = new Log();
  L.set("garrett", { B2: "orig2", B3: "orig3", B4: "orig4" });
  const a1 = L.set("claude", { B2: "c2", B3: "c3" });
  L.set("garrett", { C2: "mine" });
  const a2 = L.set("claude", { B4: "c4", D9: "c9" });
  L.edit("garrett", [{ t: "insert", s: "s1", axis: "row", at: 0, n: 1 }]); // everything moves down a row
  L.set("garrett", { B4: "garrett later" }); // was B3: Claude's c3 cell
  const plan = L.revert([a1, a2]);
  assert.equal(L.get("B3"), "orig2");
  assert.equal(L.get("B4"), "garrett later"); // later edit wins
  assert.equal(L.get("B5"), "orig4");
  assert.equal(L.get("D10"), "");
  assert.equal(L.get("C3"), "mine");
  assert.deepEqual(plan.skipped, [{ sheet: "s1", by: "garrett", what: "B4" }]);
  // Overwrite: force puts the pre-session value back too.
  const L2 = new Log();
  L2.set("garrett", { A1: "orig" });
  const v = L2.set("claude", { A1: "agent" });
  L2.set("garrett", { A1: "later" });
  L2.revert([v], true);
  assert.equal(L2.get("A1"), "orig");
});

test("reverting a row delete brings the row back, with later edits and moves around it", () => {
  const L = new Log();
  L.set("garrett", { A1: "h", A2: "keep", A3: "gone", B3: "=A3&\"!\"", A4: "after" });
  L.edit("garrett", [{ t: "fmt", s: "s1", cells: [[2, 0, { b: true }]] }, { t: "flag", s: "s1", cells: [[2, 0, { note: "n", by: "g", ts: 1 }]] }]);
  const del = L.edit("garrett", [{ t: "delete", s: "s1", axis: "row", at: 2, n: 1 }]);
  L.set("claude", { A5: "later" });
  L.edit("claude", [{ t: "move", s: "s1", axis: "row", from: 2, n: 1, to: 0 }]); // "after" moves to the top
  L.revert([del]);
  const rows = [1, 2, 3, 4, 5, 6].map((r) => L.get(`A${r}`));
  assert.deepEqual(rows, ["after", "h", "keep", "gone", "", "later"]);
  assert.equal(L.get("B4"), "=A4&\"!\"");
  const s = L.wb.doc.sheets[0]!;
  assert.equal(s.cells.get(3 * 16384)?.f?.b, true);
  assert.equal(s.cells.get(3 * 16384)?.flag?.note, "n");
});

test("reverting an insert keeps rows someone wrote into, and says so", () => {
  const L = new Log();
  L.set("garrett", { A1: "x", A2: "y" });
  const ins = L.edit("claude", [{ t: "insert", s: "s1", axis: "row", at: 1, n: 2 }]);
  L.set("garrett", { A2: "typed into the new row" });
  const plan = L.revert([ins]);
  assert.equal(L.get("A2"), "typed into the new row");
  assert.equal(plan.skipped.length, 1);
  assert.match(plan.skipped[0]!.what, /rows edited since/);
});

test("restore rebuilds an earlier version exactly, as one undoable change", () => {
  const L = new Log();
  L.set("g", { A1: "1", B2: "=A1*2" });
  L.edit("g", [{ t: "fmt", s: "s1", cells: [[0, 0, { fill: "green" }]] }, { t: "size", s: "s1", axis: "col", sizes: [[1, 160]] }]);
  const v = L.changes.length;
  L.set("c", { A1: "5", C3: "new" });
  L.edit("c", [{ t: "insert", s: "s1", axis: "row", at: 0, n: 3 }, { t: "hide", s: "s1", axis: "col", idx: [2], hidden: true }]);
  L.edit("c", [{ t: "freeze", s: "s1", rows: 1, cols: 0 }, { t: "renameSheet", s: "s1", name: "Data" }]);
  L.edit("c", [{ t: "addSheet", sheet: { id: "s2", name: "Extra", rows: 10, cols: 5, cells: [[0, 0, "z"]] } }]);
  const target = docFromJSON(L.versions[v]!);
  const ops = restoreOps(L.wb.doc, target);
  const { inverse } = L.wb.apply(ops);
  // Same content; cell order inside the JSON may differ.
  const norm = (j: DocJSON) => ({ ...j, sheets: j.sheets.map((sh) => ({ ...sh, cells: [...sh.cells].sort((x, y) => x[0] - y[0] || x[1] - y[1]) })) });
  assert.deepEqual(norm(docToJSON(L.wb.doc)), norm(L.versions[v]!));
  L.wb.apply(inverse);
  assert.deepEqual(norm(docToJSON(L.wb.doc)), norm(L.versions[L.versions.length - 1]!));
});

test("sessions group one person's edits across others' interleaved ones, and read in plain words", () => {
  const L = new Log();
  L.set("claude", { S2: "1", T3: "0" });
  L.set("garrett", { A1: "x" });
  L.set("claude", { BP19: "1" });
  L.edit("garrett", [{ t: "view", s: "s1", view: { headerRow: 0, sort: { col: 0 }, filters: {} } }]);
  L.edit("claude", [{ t: "insert", s: "s1", axis: "row", at: 4, n: 2 }], 20 * 60_000); // after a long break
  const sums = L.changes.map((c) => summarize(c, L.wb.doc));
  const ses = groupSessions(sums);
  assert.deepEqual(ses.map((s) => `${s.actor}:${s.versions.join(",")}`), ["claude:5", "garrett:2,4", "claude:1,3"]);
  const name = () => null;
  assert.equal(describe(ses[2]!.sums, name), "set 3 cells in S2:BP19");
  assert.equal(describe(ses[1]!.sums, name), 'set A1 to "x", sorted by x');
  assert.equal(describe(ses[0]!.sums, name), "inserted 2 rows at 5");
});

test("filter by range follows rows through inserts; changed cells skip rows that only moved", () => {
  const L = new Log();
  const a = L.set("c", { M2: "1" });
  const b = L.set("c", { N2: "1" });
  L.edit("g", [{ t: "insert", s: "s1", axis: "row", at: 0, n: 1 }]);
  const colM = parseRange("M:M")!;
  const later = L.changes.filter((c) => c.v > a);
  // Range M:M in today's frame, walked back to each change's frame.
  assert.equal(touches(L.changes[a - 1]!, "s1", colM), true);
  assert.equal(touches(L.changes[b - 1]!, "s1", colM), false);
  const before = docFromJSON(L.versions[1]!);
  const marks = changedCells(before, L.wb.doc, later.flatMap((c) => c.ops));
  assert.deepEqual(marks.map((m) => `${m.r},${m.c}:${m.before}->${m.after}`), ["2,13:null->1"]); // N3 new; M2 only moved to M3
});

// ---- review findings: each repro first failed on main ----

const colA = (L: Log, n: number) => Array.from({ length: n }, (_, i) => L.get(`A${i + 1}`));

test("finding 1: undoing a delete after someone moved a nearby row keeps every row", () => {
  const L = new Log();
  L.set("g", { A1: "r0", A2: "r1", A3: "r2", A4: "r3", A5: "r4" });
  const del = L.edit("g", [{ t: "delete", s: "s1", axis: "row", at: 2, n: 1 }]); // [r0,r1,r3,r4]
  L.edit("c", [{ t: "move", s: "s1", axis: "row", from: 1, n: 1, to: 4 }]); // [r0,r3,r4,r1]
  L.revert([del]);
  assert.deepEqual([...colA(L, 5)].sort(), ["r0", "r1", "r2", "r3", "r4"]);
});

test("finding 1: insert and move converge in either order (property)", () => {
  const vals = Array.from({ length: 10 }, (_, i) => `r${i}`);
  let diverged = 0, total = 0;
  for (let at = 0; at <= 10; at++) for (const ni of [1, 2]) for (let from = 0; from < 10; from++) for (const n of [1, 2, 3]) for (let to = 0; to <= 10; to++) {
    if (from + n > 10 || (to >= from && to <= from + n)) continue;
    const base = () => {
      const wb = new Workbook(newDoc("t"));
      wb.apply([{ t: "grid", s: "s1", rows: 10, cols: 2 }, { t: "set", s: "s1", cells: vals.map((v, i) => [i, 0, v] as [number, number, string]) }]);
      return wb;
    };
    const ins: Op = { t: "insert", s: "s1", axis: "row", at, n: ni };
    const mv: Op = { t: "move", s: "s1", axis: "row", from, n, to };
    const a = base(), b = base();
    a.apply([ins]);
    const mv2 = transformAll([mv], [ins], a.doc);
    a.apply(mv2);
    b.apply([mv]);
    b.apply(transformAll([ins], [mv], b.doc, false));
    const col = (wb: Workbook) => Array.from({ length: 12 }, (_, i) => wb.doc.sheets[0]!.cells.get(i * 16384)?.v ?? "").join(",");
    total++;
    if (col(a) !== col(b)) diverged++;
  }
  assert.equal(diverged, 0, `${diverged}/${total} insert/move pairs diverge`);
});
