import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { mkdtempSync, rmSync } from "node:fs";
import type { AddressInfo } from "node:net";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { after, before, test } from "node:test";
import { promisify } from "node:util";
import { startServer } from "../src/server/main.ts";

const dataDir = mkdtempSync(join(tmpdir(), "sheets-test-"));
let srv: ReturnType<typeof startServer>;
let base = "";

async function boot() {
  srv = startServer({ port: 0, host: "127.0.0.1", dataDir, publicDir: dataDir });
  await new Promise((r) => srv.server.once("listening", r));
  base = `http://127.0.0.1:${(srv.server.address() as AddressInfo).port}`;
}

before(boot);
after(async () => {
  await srv.close();
  rmSync(dataDir, { recursive: true, force: true });
});

async function req(method: string, path: string, body?: unknown, type = "application/json") {
  const res = await fetch(base + path, {
    method,
    headers: { "x-actor": "test", ...(body !== undefined ? { "content-type": type } : {}) },
    ...(body !== undefined ? { body: typeof body === "string" ? body : JSON.stringify(body) } : {}),
  });
  const text = await res.text();
  let data: any = text;
  try {
    data = JSON.parse(text);
  } catch {}
  return { status: res.status, data };
}

test("create, write with A1, read values and formulas", async () => {
  const created = await req("POST", "/api/books", { name: "Budget" });
  assert.equal(created.status, 201);
  const w = await req("PUT", "/api/books/budget/range?ref=A1", { values: [["item", "cost"], ["tea", 3], ["cake", 4.5], ["total", "=SUM(B2:B3)"]] });
  assert.equal(w.status, 200);
  assert.equal(w.data.range, "A1:B4");
  const r = await req("GET", "/api/books/Budget/range?ref=A1:B4");
  assert.deepEqual(r.data.values, [["item", "cost"], ["tea", 3], ["cake", 4.5], ["total", 7.5]]);
  const f = await req("GET", "/api/books/Budget/range?ref=B4&show=inputs");
  assert.deepEqual(f.data.values, [["=SUM(B2:B3)"]]);
  const csv = await req("GET", "/api/books/Budget/range?ref=A1:B2&as=csv");
  assert.equal(csv.data, "item,cost\ntea,3\n");
});

test("bad input gets a clear 4xx", async () => {
  assert.equal((await req("GET", "/api/books/nope/range")).status, 404);
  const bad = await req("GET", "/api/books/Budget/range?ref=Nope!A1");
  assert.equal(bad.status, 404);
  assert.match(bad.data.error, /no sheet named "Nope"/);
  const op = await req("POST", "/api/books/Budget/ops", { ops: [{ t: "explode" }] });
  assert.equal(op.status, 400);
});

test("SSE delivers changes from other writers, in order", async () => {
  const ac = new AbortController();
  const res = await fetch(`${base}/api/books/Budget/events`, { signal: ac.signal });
  const reader = res.body!.getReader();
  const seen: any[] = [];
  const done = (async () => {
    let buf = "";
    for (;;) {
      const { value, done } = await reader.read();
      if (done) return;
      buf += new TextDecoder().decode(value);
      for (const m of buf.matchAll(/event: change\ndata: (.*)\n\n/g)) seen.push(JSON.parse(m[1]!));
      if (seen.length >= 2) return;
      buf = buf.slice(buf.lastIndexOf("\n\n") + 2);
    }
  })();
  await new Promise((r) => setTimeout(r, 50));
  await req("PUT", "/api/books/Budget/range?ref=C1", { values: [["note"]] });
  await req("PUT", "/api/books/Budget/range?ref=C2", { values: [["fresh"]] });
  await done;
  ac.abort();
  assert.equal(seen.length, 2);
  assert.equal(seen[0].actor, "test");
  assert.ok(seen[1].v === seen[0].v + 1);
  assert.deepEqual(seen[1].ops[0].cells, [[1, 2, "fresh"]]);
});

test("a stale edit is shifted past a concurrent row insert", async () => {
  const b = (await req("GET", "/api/books/Budget")).data;
  const s = b.sheets[0].id;
  const v0 = b.version;
  // Writer 1 inserts a row above "cake"; writer 2, still at v0, edits the cake row.
  await req("POST", "/api/books/Budget/ops", { ops: [{ t: "insert", s, axis: "row", at: 2, n: 1 }], base: v0, client: "w1" });
  await req("POST", "/api/books/Budget/ops", { ops: [{ t: "set", s, cells: [[2, 1, "5"]] }], base: v0, client: "w2" });
  const r = await req("GET", "/api/books/Budget/range?ref=A3:B5");
  assert.deepEqual(r.data.values, [[null, null], ["cake", 5], ["total", 8]]);
});

test("a resent change (same client and seq) is applied once", async () => {
  const b = (await req("GET", "/api/books/Budget")).data;
  const s = b.sheets[0].id;
  const body = { ops: [{ t: "insert", s, axis: "row", at: 0, n: 1 }], client: "c9", seq: 1 };
  const first = await req("POST", "/api/books/Budget/ops", body);
  const again = await req("POST", "/api/books/Budget/ops", body);
  assert.equal(again.data.v, first.data.v);
  await req("POST", "/api/books/Budget/ops", { ops: [{ t: "delete", s, axis: "row", at: 0, n: 1 }] });
});

test("reads past the sheet edge return the full requested block", async () => {
  const r = await req("GET", "/api/books/Budget/range?ref=Y1:AB1");
  assert.equal(r.data.range, "Y1:AB1");
  assert.equal(r.data.values[0].length, 4);
});

test("data survives a restart", async () => {
  await srv.close();
  await boot();
  const r = await req("GET", "/api/books/Budget/range?ref=A1:B5");
  assert.equal(r.data.values[3][0], "cake");
  assert.equal(r.data.values[4][1], 8);
});

test("CSV import and the CLI agent path", async () => {
  const run = promisify(execFile);
  const cli = (...args: string[]) =>
    run("node", [join(import.meta.dirname, "..", "cli.js"), ...args], { env: { ...process.env, SHEETS_URL: base, SHEETS_ACTOR: "claude" } });
  const imp = await req("POST", "/api/books?name=Runs", "name,pass\nalpha,1\nbeta,0\n", "text/csv");
  assert.equal(imp.data.sheets[0].used, "A1:B3");
  await cli("set", "runs", "C1", "score");
  await cli("set", "runs", "C2:C3", "=B2*10");
  const out = (await cli("get", "runs", "A1:C3")).stdout;
  assert.match(out, /^# Runs \/ Sheet1!A1:C3/);
  assert.match(out, /2\talpha\t1\t10/);
  await cli("format", "runs", "A1:C1", "--bold", "--fill", "grey");
  await cli("cond", "runs", "add", "B2:B3", "eq", "1", "--fill", "green");
  const doc = (await req("GET", "/api/books/runs/doc")).data.doc;
  assert.equal(doc.sheets[0].cond[0].style.fill, "green");
  assert.equal(doc.sheets[0].cells.find((c: any) => c[0] === 0 && c[1] === 0)[3].b, true);
  const csv = (await cli("export", "runs")).stdout;
  assert.equal(csv, "name,pass,score\nalpha,1,10\nbeta,0,0\n");
});

test("flags: listed across workbooks with the current value, resolved with a reply", async () => {
  const b = await req("POST", "/api/books", { name: "Flagged" });
  const sid = b.data.sheets[0].id;
  await req("PUT", "/api/books/Flagged/range?ref=A1", { values: [["x", "=1+1"]] });
  const flag = { note: "why 2?", by: "Garrett", ts: 1000 };
  await req("POST", "/api/books/Flagged/ops", { ops: [{ t: "flag", s: sid, cells: [[0, 1, flag]] }] });
  let all = await req("GET", "/api/flags");
  const mine = all.data.filter((f: any) => f.book === "Flagged");
  assert.equal(mine.length, 1);
  assert.equal(mine[0].ref, "Sheet1!B1");
  assert.equal(mine[0].value, "2");
  assert.equal(mine[0].input, "=1+1");
  await req("POST", "/api/books/Flagged/ops", { ops: [{ t: "flag", s: sid, cells: [[0, 1, { ...flag, done: { by: "claude", ts: 2000, reply: "1+1" } }]] }] });
  all = await req("GET", "/api/flags");
  assert.equal(all.data.filter((f: any) => f.book === "Flagged").length, 0);
  const done = await req("GET", "/api/books/Flagged/flags?all=1");
  assert.equal(done.data[0].done.reply, "1+1");
});

test("history: timeline, session revert keeps later edits, restore, compare, cell history, named versions", async () => {
  const b = await req("POST", "/api/books", { name: "Hist" });
  const id = b.data.id;
  const as = (actor: string, method: string, path: string, body?: unknown) =>
    fetch(base + path, { method, headers: { "x-actor": actor, "content-type": "application/json" }, body: body === undefined ? undefined : JSON.stringify(body) }).then((r) => r.json() as Promise<any>);
  await as("garrett", "PUT", `/api/books/${id}/range?ref=A1`, { values: [["model", "pass"], ["a", "0.4"], ["b", "0.5"]] });
  const v1 = (await as("garrett", "GET", `/api/books/${id}`)).version;
  await as("claude", "PUT", `/api/books/${id}/range?ref=B2`, { values: [["0.41"], ["0.52"]] });
  await as("claude", "PUT", `/api/books/${id}/range?ref=C1`, { values: [["note"]] });
  await as("garrett", "PUT", `/api/books/${id}/range?ref=B3`, { values: [["0.55"]] }); // later edit to a cell Claude changed
  const h = await as("x", "GET", `/api/books/${id}/history`);
  assert.equal(h.summaries.length, 4);
  assert.deepEqual(h.summaries.map((s: any) => s.actor), ["garrett", "claude", "claude", "garrett"]);
  // Revert Claude's session (two changes): B2 back, C1 gone, Garrett's B3 kept and reported.
  const dry = await as("garrett", "POST", `/api/books/${id}/history/revert`, { session: h.summaries[2].v, dryRun: true });
  assert.equal(dry.applied, false);
  assert.deepEqual(dry.skipped, [{ sheet: b.data.sheets[0].id, by: "garrett", what: "B3" }]);
  const rv = await as("garrett", "POST", `/api/books/${id}/history/revert`, { session: h.summaries[2].v });
  assert.equal(rv.reverted, 2);
  const now = await as("x", "GET", `/api/books/${id}/range?ref=B2:C3&show=inputs`);
  assert.deepEqual(now.values, [["0.4", ""], ["0.55", ""]]);
  const after = await as("x", "GET", `/api/books/${id}/history`);
  assert.equal(after.summaries.at(-1).meta.kind, "revert");
  // Compare: since v1, what changed (B3 only; B2 and C1 went back).
  const at = await as("x", "GET", `/api/books/${id}/history/at?from=${v1}`);
  assert.deepEqual(at.marks.map((m: any) => `${m.r},${m.c}:${m.before}->${m.after}`), ["2,1:0.5->0.55"]);
  // A past version, rebuilt, with the cells one change touched.
  const past = await as("x", "GET", `/api/books/${id}/history/at?v=${v1 + 1}&from=${v1}&only=${v1 + 1}`);
  assert.equal(past.doc.sheets[0].cells.find((c: any) => c[0] === 1 && c[1] === 1)[2], "0.41");
  assert.equal(past.marks.length, 2);
  // Cell history, newest first.
  const ch = await as("x", "GET", `/api/books/${id}/history/cell?ref=B3`);
  assert.deepEqual(ch.entries.map((e: any) => `${e.actor}:${e.value}`), ["garrett:0.55", "claude:0.52", "garrett:0.5"]);
  // Restore v1 as a new change; the restore itself can be reverted.
  const rs = await as("garrett", "POST", `/api/books/${id}/history/restore`, { v: v1 });
  assert.equal(rs.applied, true);
  assert.deepEqual((await as("x", "GET", `/api/books/${id}/range?ref=B3&show=inputs`)).values, [["0.5"]]);
  await as("garrett", "POST", `/api/books/${id}/history/revert`, { versions: [rs.v] });
  assert.deepEqual((await as("x", "GET", `/api/books/${id}/range?ref=B3&show=inputs`)).values, [["0.55"]]);
  // Named versions.
  const cp = await as("garrett", "POST", `/api/books/${id}/history/checkpoints`, { name: "before int4 run" });
  assert.equal(cp.name, "before int4 run");
  assert.equal((await as("x", "GET", `/api/books/${id}/history`)).checkpoints[0].name, "before int4 run");
  await as("garrett", "DELETE", `/api/books/${id}/history/checkpoints/${cp.id}`);
  assert.equal((await as("x", "GET", `/api/books/${id}/history/checkpoints`)).length, 0);
  // Filter by range and by person.
  const onlyB3 = await as("x", "GET", `/api/books/${id}/history?range=B3`);
  assert.ok(onlyB3.summaries.every((s: any) => s.v !== v1 + 2));
  const byClaude = await as("x", "GET", `/api/books/${id}/history?by=claude`);
  assert.equal(byClaude.summaries.length, 2);
  // Survives a restart: inverses and snapshots are on disk.
  await srv.close();
  await boot();
  const again = await as("garrett", "POST", `/api/books/${id}/history/revert`, { versions: [v1 + 1], dryRun: true });
  assert.equal(again.applied, false);
  assert.equal((await as("x", "GET", `/api/books/${id}/history/at?v=${v1}`)).doc.sheets[0].cells.length, 6);
});

test("history: an old database without inverses or kept snapshots is migrated in place", async () => {
  const dir = mkdtempSync(join(tmpdir(), "sheets-old-"));
  const { DatabaseSync } = await import("node:sqlite");
  const db = new DatabaseSync(join(dir, "sheets.db"));
  db.exec(`CREATE TABLE books (id TEXT PRIMARY KEY, name TEXT NOT NULL, created INTEGER NOT NULL, updated INTEGER NOT NULL,
             version INTEGER NOT NULL, snap_version INTEGER NOT NULL, snapshot TEXT NOT NULL, deleted INTEGER NOT NULL DEFAULT 0);
           CREATE TABLE changes (book TEXT NOT NULL, v INTEGER NOT NULL, ts INTEGER NOT NULL, actor TEXT NOT NULL, client TEXT NOT NULL, ops TEXT NOT NULL, seq INTEGER, PRIMARY KEY (book, v));`);
  const snap = { name: "Old", sheets: [{ id: "s1", name: "Runs", rows: 100, cols: 10, cells: [[0, 0, "a"], [1, 0, "b"], [2, 0, "c"]] }] };
  db.prepare("INSERT INTO books VALUES ('old1', 'Old', 1, 2, 3, 1, ?, 0)").run(JSON.stringify(snap));
  db.prepare("INSERT INTO books VALUES ('gone', 'Deleted', 1, 2, 0, 0, ?, 1)").run(JSON.stringify(snap));
  db.prepare("INSERT INTO changes VALUES ('old1', 1, 1, 'seed', '', ?, NULL)").run(JSON.stringify([{ t: "renameSheet", s: "s1", name: "Runs" }]));
  db.prepare("INSERT INTO changes VALUES ('old1', 2, 2, 'garrett', 'x', ?, NULL)").run(JSON.stringify([{ t: "delete", s: "s1", axis: "row", at: 1, n: 1 }]));
  db.prepare("INSERT INTO changes VALUES ('old1', 3, 3, 'garrett', 'x', ?, NULL)").run(JSON.stringify([{ t: "set", s: "s1", cells: [[0, 1, "x"]] }]));
  db.close();
  const old = startServer({ port: 0, host: "127.0.0.1", dataDir: dir, publicDir: dir });
  await new Promise((r) => old.server.once("listening", r));
  const url = `http://127.0.0.1:${(old.server.address() as AddressInfo).port}/api/books/old1`;
  try {
    // The row delete (after the kept snapshot) can be reverted: its inverse was rebuilt on load.
    const r = await fetch(`${url}/history/revert`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ versions: [2] }) }).then((x) => x.json() as Promise<any>);
    assert.equal(r.applied, true);
    const vals = await fetch(`${url}/range?ref=A1:A3`).then((x) => x.json() as Promise<any>);
    assert.deepEqual(vals.values, [["a"], ["b"], ["c"]]);
    // Before the kept snapshot, history is listed but not rebuildable.
    const h = await fetch(`${url}/history`).then((x) => x.json() as Promise<any>);
    assert.equal(h.oldest, 1);
    assert.equal(h.summaries[0].undoable, false);
    assert.equal((await fetch(`${url}/history/at?v=0`)).status, 410);
  } finally {
    await old.close();
  }
  // Deleted books get no kept snapshots; live ones keep their newest snapshot only in snaps.
  const after = new DatabaseSync(join(dir, "sheets.db"));
  assert.equal((after.prepare("SELECT COUNT(*) AS n FROM snaps WHERE book = 'gone'").get() as { n: number }).n, 0);
  assert.equal((after.prepare("SELECT snapshot FROM books WHERE id = 'old1'").get() as { snapshot: string }).snapshot, "");
  after.close();
  // And it still loads from there.
  const again = startServer({ port: 0, host: "127.0.0.1", dataDir: dir, publicDir: dir });
  await new Promise((r) => again.server.once("listening", r));
  try {
    const port = (again.server.address() as AddressInfo).port;
    const vals = await fetch(`http://127.0.0.1:${port}/api/books/old1/range?ref=A1:B1&show=inputs`).then((x) => x.json() as Promise<any>);
    assert.deepEqual(vals.values, [["a", "x"]]);
  } finally {
    await again.close();
    rmSync(dir, { recursive: true, force: true });
  }
});

test("purge erases a workbook and its history; a plain delete keeps it in the file", async () => {
  const a = await req("POST", "/api/books", { name: "Purge me" });
  await req("PUT", `/api/books/${a.data.id}/range?ref=A1`, { values: [["x"]] });
  const r = await req("DELETE", `/api/books/${a.data.id}?purge=1`);
  assert.equal(r.data.purged, true);
  assert.equal((await req("GET", `/api/books/${a.data.id}`)).status, 404);
  const { DatabaseSync } = await import("node:sqlite");
  const db = new DatabaseSync(join(dataDir, "sheets.db"), { readOnly: true });
  for (const t of ["books", "changes", "snaps"]) {
    const col = t === "books" ? "id" : "book";
    assert.equal((db.prepare(`SELECT COUNT(*) AS n FROM ${t} WHERE ${col} = ?`).get(a.data.id) as { n: number }).n, 0, t);
  }
  db.close();
});
