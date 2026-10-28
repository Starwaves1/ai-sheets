// HTTP API. JSON in and out, A1 notation for anything an agent types by hand.
// Every route goes through identify() first: the seam where auth will go.

import type { IncomingMessage, ServerResponse } from "node:http";
import { clamp, MAX_COLS, MAX_ROWS, parseRange, rangeName, type SheetRange } from "../core/a1.ts";
import { parseDelimited, toDelimited } from "../core/csv.ts";
import { isFormula, moveRefs } from "../core/formula/rewrite.ts";
import { findSheet, newSheet, OpError, sheetToJSON, usedRange, type Sheet } from "../core/model.ts";
import type { FormatPatch, Op } from "../core/ops.ts";
import { Hub, newId, type LiveBook } from "./hub.ts";

export interface Actor {
  name: string;
}

/** Who is calling. LAN-only for now: trust the X-Actor header. Replace with real auth later. */
export function identify(req: IncomingMessage): Actor | null {
  const h = req.headers["x-actor"];
  const name = (Array.isArray(h) ? h[0] : h)?.trim().slice(0, 40);
  return { name: name || "browser" };
}

export class HttpError extends Error {
  constructor(readonly status: number, message: string) {
    super(message);
  }
}

const MAX_BODY = 32 * 1024 * 1024;

async function body(req: IncomingMessage): Promise<string> {
  const chunks: Buffer[] = [];
  let size = 0;
  for await (const chunk of req) {
    size += (chunk as Buffer).length;
    if (size > MAX_BODY) throw new HttpError(413, "body too large");
    chunks.push(chunk as Buffer);
  }
  return Buffer.concat(chunks).toString("utf8");
}

async function json(req: IncomingMessage): Promise<Record<string, unknown>> {
  const text = await body(req);
  if (!text) return {};
  try {
    const v = JSON.parse(text);
    if (typeof v !== "object" || v === null || Array.isArray(v)) throw new Error();
    return v;
  } catch {
    throw new HttpError(400, "body must be a JSON object");
  }
}

function send(res: ServerResponse, status: number, data: unknown): void {
  const text = JSON.stringify(data);
  res.writeHead(status, { "content-type": "application/json; charset=utf-8", "cache-control": "no-store" });
  res.end(text);
}

function sendText(res: ServerResponse, type: string, text: string, filename?: string): void {
  const headers: Record<string, string> = { "content-type": `${type}; charset=utf-8`, "cache-control": "no-store" };
  if (filename) headers["content-disposition"] = `attachment; filename="${filename.replace(/[^\w. -]/g, "_")}"`;
  res.writeHead(200, headers);
  res.end(text);
}

/** Resolve "Sheet!A1:B2" (sheet optional, defaults to the first sheet). Missing range = used range. */
function target(lb: LiveBook, ref: string | null): { sheet: Sheet; range: SheetRange; hadRange: boolean } {
  const doc = lb.wb.doc;
  let sheet = doc.sheets[0]!;
  if (!ref) {
    const used = usedRange(sheet) ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    return { sheet, range: used, hadRange: false };
  }
  let range = parseRange(ref);
  if (!range) {
    // A bare sheet name selects that sheet's used range.
    const s = findSheet(doc, ref.replace(/^'(.*)'!?$|!$/, "$1"));
    if (!s) throw new HttpError(400, `"${ref}" is not a range like A1:C3 or Sheet1!B2, or a sheet name`);
    return { sheet: s, range: usedRange(s) ?? { r1: 0, c1: 0, r2: 0, c2: 0 }, hadRange: false };
  }
  if (range.sheet !== undefined) {
    const s = findSheet(doc, range.sheet);
    if (!s) throw new HttpError(404, `no sheet named "${range.sheet}" (have: ${doc.sheets.map((x) => x.name).join(", ")})`);
    sheet = s;
  }
  // Whole rows/columns stop at the sheet's used extent.
  const used = usedRange(sheet);
  if (range.r2 >= MAX_ROWS - 1) range = { ...range, r2: Math.max(range.r1, used?.r2 ?? 0) };
  if (range.c2 >= MAX_COLS - 1) range = { ...range, c2: Math.max(range.c1, used?.c2 ?? 0) };
  return { sheet, range, hadRange: true };
}

function cellsFromBody(text: string, type: string | undefined): string[][] {
  if (type?.includes("json")) {
    let v: unknown;
    try {
      v = JSON.parse(text);
    } catch {
      throw new HttpError(400, "invalid JSON");
    }
    const values = (v as { values?: unknown }).values ?? v;
    if (!Array.isArray(values)) throw new HttpError(400, 'JSON body must be {"values": [[...], ...]}');
    return values.map((row) => {
      if (!Array.isArray(row)) return [row === null || row === undefined ? "" : String(row)];
      return row.map((x) => (x === null || x === undefined ? "" : typeof x === "boolean" ? (x ? "TRUE" : "FALSE") : String(x)));
    });
  }
  return parseDelimited(text, type?.includes("tab-separated") ? "\t" : undefined);
}

export function makeApi(hub: Hub) {
  const book = (ref: string): LiveBook => {
    const lb = hub.find(decodeURIComponent(ref));
    if (!lb) {
      const names = hub.list().map((b) => b.name).join(", ");
      throw new HttpError(404, `no workbook matching "${decodeURIComponent(ref)}" (have: ${names || "none"})`);
    }
    return lb;
  };

  const submit = (lb: LiveBook, ops: Op[], actor: Actor, extra: { base?: number; client?: string } = {}) => {
    try {
      return lb.submit(hub.store, ops, { actor: actor.name, client: extra.client ?? "", ...(extra.base !== undefined ? { base: extra.base } : {}) });
    } catch (e) {
      if (e instanceof OpError) throw new HttpError(400, e.message);
      throw e;
    }
  };

  return async function handle(req: IncomingMessage, res: ServerResponse, url: URL): Promise<void> {
    const actor = identify(req);
    if (!actor) throw new HttpError(401, "not signed in");
    const parts = url.pathname.split("/").filter(Boolean).slice(1); // drop "api"
    const m = req.method ?? "GET";
    const q = url.searchParams;

    // /api/books
    if (parts[0] === "books" && parts.length === 1) {
      if (m === "GET") return send(res, 200, hub.list().map((b) => b.summary()));
      if (m === "POST") {
        const name = q.get("name") ?? "";
        const type = req.headers["content-type"];
        if (type?.includes("json")) {
          const j = await json(req);
          const lb = hub.create(String(j.name ?? name ?? "").trim() || "Untitled");
          return send(res, 201, lb.summary());
        }
        // CSV/TSV body: import as a new workbook.
        const rows = cellsFromBody(await body(req), type);
        const sheet = newSheet(newId(6), q.get("sheet") || "Sheet1", Math.max(1000, rows.length + 100), Math.max(26, ...rows.map((r) => r.length)));
        rows.forEach((row, r) => row.forEach((v, c) => v !== "" && sheet.cells.set(r * 16384 + c, { v })));
        const lb = hub.create(name.trim() || "Imported", { name, sheets: [sheetToJSON(sheet)] });
        return send(res, 201, lb.summary());
      }
    }

    if (parts[0] !== "books" || !parts[1]) throw new HttpError(404, "not found");
    const lb = book(parts[1]);
    const sub = parts[2] ?? "";

    if (sub === "" && m === "GET") return send(res, 200, lb.summary());
    if (sub === "" && m === "DELETE") {
      hub.remove(lb);
      return send(res, 200, { deleted: lb.id });
    }
    if (sub === "doc" && m === "GET") return send(res, 200, { ...lb.summary(), doc: lb.wb.toJSON() });

    if (sub === "ops" && m === "POST") {
      const j = await json(req);
      const ops = j.ops as Op[];
      const base = typeof j.base === "number" ? j.base : undefined;
      const client = typeof j.client === "string" ? j.client : "";
      const c = submit(lb, ops, actor, { ...(base !== undefined ? { base } : {}), client });
      return send(res, 200, { v: c.v });
    }

    if (sub === "changes" && m === "GET") {
      const since = Number(q.get("since") ?? lb.version);
      return send(res, 200, lb.changesSince(hub.store, since));
    }

    if (sub === "events" && m === "GET") return events(lb, req, res, q);

    if (sub === "range") {
      const ref = q.get("ref");
      const { sheet, range, hadRange } = target(lb, ref);
      if (m === "GET") {
        const show = q.get("show") ?? "values";
        const as = q.get("as") ?? "json";
        const what = show === "inputs" ? "inputs" : show === "display" || as === "csv" || as === "tsv" ? "display" : "values";
        const values = lb.wb.read(sheet.id, range, what);
        if (as === "csv" || as === "tsv") {
          const rows = values.map((r) => r.map((v) => (v === null ? "" : String(v))));
          return sendText(res, as === "csv" ? "text/csv" : "text/tab-separated-values", toDelimited(rows, as === "csv" ? "," : "\t"));
        }
        const out: Record<string, unknown> = { book: lb.name, sheet: sheet.name, range: rangeName(range), version: lb.version, values };
        if (show === "both") out.inputs = lb.wb.read(sheet.id, range, "inputs");
        return send(res, 200, out);
      }
      if (m === "PUT" || m === "POST") {
        if (!hadRange) throw new HttpError(400, "give a target, e.g. ref=Sheet1!B2");
        const rows = cellsFromBody(await body(req), req.headers["content-type"]);
        const cells: Array<[number, number, string | null]> = [];
        const multi = range.r1 !== range.r2 || range.c1 !== range.c2;
        if (multi && rows.length === 1 && rows[0]!.length === 1) {
          // One value into a range fills it; formulas shift their relative refs like a fill-down.
          const v = rows[0]![0]! || null;
          for (let r = range.r1; r <= range.r2; r++) {
            for (let c = range.c1; c <= range.c2; c++) cells.push([r, c, isFormula(v ?? undefined) ? moveRefs(v!, r - range.r1, c - range.c1) : v]);
          }
        } else {
          rows.forEach((row, i) => row.forEach((v, j) => cells.push([range.r1 + i, range.c1 + j, v === "" ? null : v])));
        }
        const c = submit(lb, [{ t: "set", s: sheet.id, cells }], actor);
        const h = rows.length, w = Math.max(0, ...rows.map((r) => r.length));
        const written = multi && h === 1 && w === 1 ? range : { r1: range.r1, c1: range.c1, r2: range.r1 + Math.max(h, 1) - 1, c2: range.c1 + Math.max(w, 1) - 1 };
        return send(res, 200, { v: c.v, sheet: sheet.name, range: rangeName(written), cells: cells.length });
      }
      if (m === "DELETE") {
        if (!hadRange) throw new HttpError(400, "give a range to clear");
        const g = clamp(range, sheet.rows, sheet.cols);
        const cells: Array<[number, number, string | null]> = [];
        for (const k of sheet.cells.keys()) {
          const r = Math.floor(k / 16384), c = k % 16384;
          if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2 && sheet.cells.get(k)!.v !== undefined) cells.push([r, c, null]);
        }
        const ops: Op[] = cells.length ? [{ t: "set", s: sheet.id, cells }] : [];
        if (q.get("formats") === "1") ops.push({ t: "style", s: sheet.id, range: g, patch: { b: null, i: null, s: null, color: null, fill: null, align: null, nf: null, wrap: null } });
        if (!ops.length) return send(res, 200, { v: lb.version, cleared: 0 });
        const c = submit(lb, ops, actor);
        return send(res, 200, { v: c.v, cleared: cells.length });
      }
    }

    if (sub === "format" && (m === "POST" || m === "PUT")) {
      const { sheet, range, hadRange } = target(lb, q.get("ref"));
      if (!hadRange) throw new HttpError(400, "give a range to format");
      const patch = (await json(req)) as FormatPatch;
      const c = submit(lb, [{ t: "style", s: sheet.id, range: clamp(range, sheet.rows, sheet.cols), patch }], actor);
      return send(res, 200, { v: c.v, sheet: sheet.name, range: rangeName(range) });
    }

    if (sub === "csv" && m === "GET") {
      const s = q.get("sheet") ? findSheet(lb.wb.doc, q.get("sheet")!) : lb.wb.doc.sheets[0];
      if (!s) throw new HttpError(404, `no sheet named "${q.get("sheet")}"`);
      const used = usedRange(s);
      const rows = used ? lb.wb.read(s.id, used, "display").map((r) => r.map(String)) : [];
      return sendText(res, "text/csv", toDelimited(rows), `${lb.name} - ${s.name}.csv`);
    }

    if (sub === "sheets" && m === "POST") {
      // Add a sheet, optionally filled from a CSV/TSV body.
      const name = (q.get("name") ?? "").trim() || `Sheet${lb.wb.doc.sheets.length + 1}`;
      const text = await body(req);
      const rows = text ? cellsFromBody(text, req.headers["content-type"]) : [];
      const sheet = newSheet(newId(6), name, Math.max(1000, rows.length + 100), Math.max(26, ...rows.map((r) => r.length)));
      rows.forEach((row, r) => row.forEach((v, c) => v !== "" && sheet.cells.set(r * 16384 + c, { v })));
      const c = submit(lb, [{ t: "addSheet", sheet: sheetToJSON(sheet) }], actor);
      return send(res, 201, { v: c.v, sheet: { id: sheet.id, name } });
    }

    throw new HttpError(404, `no route ${m} ${url.pathname}`);
  };

  function events(lb: LiveBook, req: IncomingMessage, res: ServerResponse, q: URLSearchParams): void {
    res.writeHead(200, {
      "content-type": "text/event-stream",
      "cache-control": "no-store",
      connection: "keep-alive",
      "x-accel-buffering": "no",
    });
    const last = req.headers["last-event-id"] ?? q.get("since");
    const since = last === null || last === undefined ? lb.version : Number(last);
    res.write(`retry: 2000\nevent: hello\ndata: ${JSON.stringify({ v: lb.version })}\n\n`);
    if (Number.isFinite(since) && since < lb.version) {
      if (lb.version - since > 2000) res.write("event: reset\ndata: {}\n\n");
      else for (const c of lb.changesSince(hub.store, since)) res.write(`id: ${c.v}\nevent: change\ndata: ${JSON.stringify(c)}\n\n`);
    }
    lb.subs.add(res);
    // Comment line every 30 s keeps idle connections from being dropped; browsers do not wake JS for it.
    const ping = setInterval(() => res.write(": ping\n\n"), 30_000);
    req.on("close", () => {
      clearInterval(ping);
      lb.subs.delete(res);
    });
  }
}
