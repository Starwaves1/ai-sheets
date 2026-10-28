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
