// History endpoints under /api/books/<book>/history: the timeline, any version rebuilt, compare,
// selective revert, restore, cell history and named versions.

import type { IncomingMessage, ServerResponse } from "node:http";
import { cellName, parseRange, type Range } from "../core/a1.ts";
import { carryCells, changedCells, groupSessions, rangeBefore, restoreOps, revertOps, summarize, touches, type Change, type Mark, type Meta } from "../core/history.ts";
import { docToJSON, findSheet, key } from "../core/model.ts";
import type { Op } from "../core/ops.ts";
import { HttpError, json, send, type Actor } from "./api.ts";
import type { Hub, LiveBook } from "./hub.ts";

const SCAN = 50_000; // most changes a history request reads

export async function history(
  hub: Hub, lb: LiveBook, req: IncomingMessage, res: ServerResponse, rest: string[], q: URLSearchParams, actor: Actor,
  submit: (ops: Op[], why: Meta) => Change,
): Promise<void> {
  const store = hub.store;
  const m = req.method ?? "GET";
  const what = rest[0] ?? "";
  const num = (name: string, fallback?: number) => {
    const raw = q.get(name);
    if (raw === null || raw === "") {
      if (fallback === undefined) throw new HttpError(400, `give ${name}`);
      return fallback;
    }
    const n = Number(raw.replace(/^v/, ""));
    if (!Number.isInteger(n)) throw new HttpError(400, `${name} must be a version number`);
    return n;
  };
  const docAt = (v: number) => {
    try {
      return lb.docAt(store, v);
    } catch (e) {
      if (e instanceof RangeError) throw new HttpError(410, e.message);
      throw e;
    }
  };
  /** A range in today's coordinates: "Runs!M:M", "B4", a sheet name. */
  const rangeArg = (ref: string): { sheet: string; range: Range } => {
    // A bare column ("M") or row ("5") means all of it, optionally after a sheet name.
    const g = parseRange(ref.replace(/(^|!)([A-Za-z]{1,3}|\d+)$/, (_, p: string, x: string) => `${p}${x}:${x}`));
    const doc = lb.wb.doc;
    if (!g) {
      const s = findSheet(doc, ref);
      if (!s) throw new HttpError(400, `"${ref}" is not a range or sheet name`);
      return { sheet: s.id, range: { r1: 0, c1: 0, r2: 1e7, c2: 1e5 } };
    }
    const s = g.sheet ? findSheet(doc, g.sheet) : doc.sheets[0];
    if (!s) throw new HttpError(404, `no sheet named "${g.sheet}"`);
    return { sheet: s.id, range: { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 } };
  };
  const log = (since: number) => store.changesBetween(lb.id, Math.max(since, lb.version - SCAN), lb.version);

  // GET history: one summary per change (newest last), named versions, what can be rebuilt.
  if (what === "" && m === "GET") {
    const limit = Math.min(num("limit", 2000), 20_000);
    const before = num("before", lb.version + 1);
    const by = q.get("by")?.toLowerCase().split(",").filter(Boolean);
    const since = q.get("since");
    const sinceTs = since && !/^v?\d+$/.test(since) ? parseSince(since) : null;
    const sinceV = since && /^v?\d+$/.test(since) ? Number(since.replace(/^v/, "")) : 0;
    let changes = log(sinceV).filter((c) => c.v < before);
    if (sinceTs !== null) changes = changes.filter((c) => c.ts >= sinceTs);
    if (by?.length) changes = changes.filter((c) => by.includes(c.actor.toLowerCase()));
    const ref = q.get("range");
    if (ref) {
      // Walk back from now, following the range through every structural change.
      const { sheet, range } = rangeArg(ref);
      const all = log(0);
      const keep = new Set<number>();
      let g: Range | null = range;
      for (let i = all.length - 1; i >= 0 && g; i--) {
        if (touches(all[i]!, sheet, g)) keep.add(all[i]!.v);
        g = rangeBefore(g, sheet, all[i]!.ops);
      }
      changes = changes.filter((c) => keep.has(c.v));
    }
    const more = changes.length > limit;
    changes = changes.slice(-limit);
    return send(res, 200, {
      version: lb.version,
      oldest: store.oldestSnap(lb.id),
      more,
      sheets: Object.fromEntries(lb.wb.doc.sheets.map((s) => [s.id, s.name])),
      summaries: changes.map((c) => summarize(c, lb.wb.doc)),
      checkpoints: store.checkpoints(lb.id),
    });
  }

  // GET history/at?v=N[&from=M][&only=v1,v2]: the workbook at N, and the cells that changed since M
  // (only those the listed changes touched, when given). N defaults to now.
  if (what === "at" && m === "GET") {
    const v = Math.min(num("v", lb.version), lb.version);
    const doc = docAt(v);
    let marks: Mark[] = [];
    if (q.get("from") !== null) {
      const from = num("from");
      if (from > v) throw new HttpError(400, "from must be at or before v");
      const between = store.changesBetween(lb.id, from, v);
      marks = changedCells(docAt(from), doc, between.flatMap((c) => c.ops));
      const only = q.get("only")?.split(",").map(Number).filter(Number.isInteger);
      if (only?.length) {
        const keep = touchedNow(between, new Set(only), doc);
        marks = marks.filter((mk) => keep.has(`${mk.s}:${key(mk.r, mk.c)}`));
        const who = new Map(between.map((c) => [c.v, c.actor]));
        for (const mk of marks) mk.by = who.get(only[only.length - 1]!) ?? "";
      }
    }
    return send(res, 200, { v, version: lb.version, doc: docToJSON(doc), marks });
  }

  // POST history/revert {versions | session, force, dryRun}: undo those changes as a new change.
  if (what === "revert" && m === "POST") {
    const j = await json(req);
    const all = log(0);
    let versions: number[] = Array.isArray(j.versions) ? (j.versions as unknown[]).map(Number).filter(Number.isInteger) : [];
    if (typeof j.session === "number" || typeof j.session === "string") {
      const sv = Number(String(j.session).replace(/^s/, ""));
      const ses = groupSessions(all.map((c) => summarize(c, lb.wb.doc))).find((s) => s.versions.includes(sv));
      if (!ses) throw new HttpError(404, `no session with v${sv}`);
      versions = ses.versions;
    }
    if (!versions.length) throw new HttpError(400, 'give {"versions": [..]} or {"session": <version>}');
    const targets = new Set(versions);
    const from = Math.min(...versions);
    const chosen = all.filter((c) => targets.has(c.v));
    if (chosen.length !== targets.size) throw new HttpError(404, `no change ${versions.find((v) => !chosen.some((c) => c.v === v))}`);
    let plan;
    try {
      plan = revertOps(all.filter((c) => c.v >= from), targets, lb.wb.doc, j.force === true);
    } catch (e) {
      throw new HttpError(409, (e as Error).message);
    }
    if (j.dryRun === true || !plan.ops.length) return send(res, 200, { v: lb.version, reverted: chosen.length, skipped: plan.skipped, applied: false });
    const why: Meta = { kind: "revert", of: versions, actors: [...new Set(chosen.map((c) => c.actor))], since: chosen[0]!.ts };
    const c = submit(plan.ops, why);
    return send(res, 200, { v: c.v, reverted: chosen.length, skipped: plan.skipped, applied: true });
  }

  // POST history/restore {v, sheet?}: make the workbook (or one sheet) what it was at v, as one change.
  if (what === "restore" && m === "POST") {
    const j = await json(req);
    const v = Number(String(j.v ?? "").replace(/^v/, ""));
    if (!Number.isInteger(v) || v < 0 || v > lb.version) throw new HttpError(400, "give a version to restore");
    const sheet = typeof j.sheet === "string" && j.sheet ? findSheet(lb.wb.doc, j.sheet) ?? null : null;
    if (typeof j.sheet === "string" && j.sheet && !sheet) throw new HttpError(404, `no sheet "${j.sheet}"`);
    let ops: Op[];
    try {
      ops = restoreOps(lb.wb.doc, docAt(v), sheet?.id);
    } catch (e) {
      if (e instanceof HttpError) throw e;
      throw new HttpError(409, (e as Error).message);
    }
    if (!ops.length) return send(res, 200, { v: lb.version, applied: false });
    const at = v === 0 ? lb.created : store.changesBetween(lb.id, v - 1, v)[0]?.ts ?? lb.created;
    const c = submit(ops, { kind: "restore", to: v, at, ...(sheet ? { sheet: sheet.name } : {}) });
    return send(res, 200, { v: c.v, applied: true });
  }

  // GET history/cell?ref=Runs!B4: every value the cell had, newest first.
  if (what === "cell" && m === "GET") {
    const ref = q.get("ref");
    if (!ref) throw new HttpError(400, "give ref, e.g. ref=Runs!B4");
    const { sheet, range } = rangeArg(ref);
    const all = log(0);
    const entries: Array<{ v: number; actor: string; ts: number; value: string | null; before?: string | null; meta?: Meta }> = [];
    let r = range.r1, c = range.c1;
    let born = false;
    for (let i = all.length - 1; i >= 0 && !born; i--) {
      const ch = all[i]!;
      let hit = false;
      for (let k = ch.ops.length - 1; k >= 0; k--) {
        const op = ch.ops[k]!;
        if (!hit && op.t === "set" && op.s === sheet) {
          const cell = op.cells.find(([cr, cc]) => cr === r && cc === c);
          if (cell) {
            hit = true;
            const old = ch.inv?.flatMap((x) => (x.t === "set" && x.s === sheet ? x.cells : [])).find(([cr, cc]) => cr === r && cc === c);
            entries.push({ v: ch.v, actor: ch.actor, ts: ch.ts, value: cell[2], ...(old ? { before: old[2] } : {}), ...(ch.meta ? { meta: ch.meta } : {}) });
          }
        }
        const g = rangeBefore({ r1: r, c1: c, r2: r, c2: c }, sheet, [op]);
        if (!g) {
          born = true;
          break;
        }
        [r, c] = [g.r1, g.c1];
      }
    }
    const s = lb.wb.doc.sheets.find((x) => x.id === sheet)!;
    return send(res, 200, {
      ref: `${s.name}!${cellName(range.r1, range.c1)}`, sheet, r: range.r1, c: range.c1,
      now: s.cells.get(key(range.r1, range.c1))?.v ?? null, entries, oldest: store.oldestSnap(lb.id),
    });
  }

  // Named versions.
  if (what === "checkpoints") {
    if (m === "GET") return send(res, 200, store.checkpoints(lb.id));
    if (m === "POST") {
      const j = await json(req);
      const name = String(j.name ?? "").trim().slice(0, 120);
      if (!name) throw new HttpError(400, "give the version a name");
      const v = j.v === undefined ? lb.version : Number(j.v);
      if (!Number.isInteger(v) || v < 0 || v > lb.version) throw new HttpError(400, `no version ${j.v}`);
      const cp = store.addCheckpoint(lb.id, v, name, actor.name, Date.now());
      lb.notify("checkpoints", store.checkpoints(lb.id));
      return send(res, 201, cp);
    }
    if (m === "DELETE" && rest[1]) {
      if (!store.removeCheckpoint(lb.id, Number(rest[1]))) throw new HttpError(404, "no such named version");
      lb.notify("checkpoints", store.checkpoints(lb.id));
      return send(res, 200, { deleted: Number(rest[1]) });
    }
  }

  throw new HttpError(404, `no history route ${m} ${what}`);
}

/** Cells (as "sheet:key" in the latest frame) touched by the listed changes in `between`. */
function touchedNow(between: Change[], only: Set<number>, doc: Parameters<typeof carryCells>[2]): Set<string> {
  const keep = new Set<string>();
  between.forEach((c, i) => {
    if (!only.has(c.v)) return;
    const later = between.slice(i + 1).flatMap((x) => x.ops);
    const cells: Array<{ s: string; r: number; c: number }> = [];
    c.ops.forEach((op, k) => {
      let list: Array<[number, number]> = [];
      if (op.t === "set" || op.t === "fmt" || op.t === "flag") list = op.cells.map(([r, cc]) => [r, cc]);
      else if (op.t === "style") for (let r = op.range.r1; r <= Math.min(op.range.r2, op.range.r1 + 999); r++) for (let cc = op.range.c1; cc <= Math.min(op.range.c2, op.range.c1 + 199); cc++) list.push([r, cc]);
      if (!list.length || !("s" in op)) return;
      const rest = c.ops.slice(k + 1).concat(later);
      for (const p of carryCells(list.map(([r, cc]) => ({ s: op.s, r, c: cc })), rest, doc)) if (p) keep.add(`${p.s}:${key(p.r, p.c)}`);
    });
  });
  return keep;
}

/** "2h", "30m", "3d", or anything Date.parse understands, to a timestamp. */
export function parseSince(text: string): number {
  const m = /^(\d+(?:\.\d+)?)\s*(m|min|h|d|w)$/i.exec(text.trim());
  if (m) return Date.now() - Number(m[1]) * { m: 60e3, min: 60e3, h: 3600e3, d: 86400e3, w: 604800e3 }[m[2]!.toLowerCase() as "m"]!;
  const t = Date.parse(text);
  if (Number.isNaN(t)) throw new HttpError(400, `since: use 2h, 30m, 3d, v120 or a date`);
  return t;
}
