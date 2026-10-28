// Live workbooks in memory. The hub orders every change (assigning versions),
// transforms late ops past concurrent structural edits, persists, and fans out over SSE.

import type { ServerResponse } from "node:http";
import { randomBytes } from "node:crypto";
import { newSheet, sheetToJSON, usedRange, type DocJSON } from "../core/model.ts";
import { transformAll, type Op } from "../core/ops.ts";
import { validateOps } from "../core/validate.ts";
import { Workbook } from "../core/workbook.ts";
import { rangeName } from "../core/a1.ts";
import type { Change, Store } from "./store.ts";

const SNAPSHOT_EVERY = 200;
/** Base URL people open in a browser (shown by the CLI); relative when unset. */
const PUBLIC_URL = (process.env.SHEETS_PUBLIC_URL ?? "").replace(/\/$/, "");
const RECENT = 1000;

export const newId = (n = 8): string => {
  const alphabet = "abcdefghijkmnpqrstuvwxyz23456789";
  return [...randomBytes(n)].map((b) => alphabet[b % alphabet.length]).join("");
};

export class LiveBook {
  readonly subs = new Set<ServerResponse>();
  recent: Change[] = [];

  constructor(
    readonly id: string,
    readonly wb: Workbook,
    public version: number,
    readonly created: number,
    public updated: number,
    private snapVersion: number,
  ) {}

  get name(): string {
    return this.wb.doc.name;
  }

  summary() {
    return {
      id: this.id,
      name: this.name,
      url: `${PUBLIC_URL}/b/${this.id}`,
      version: this.version,
      created: this.created,
      updated: this.updated,
      sheets: this.wb.doc.sheets.map((s) => {
        const used = usedRange(s);
        return {
          id: s.id, name: s.name, rows: s.rows, cols: s.cols,
          used: used ? rangeName(used) : null,
          frozenRows: s.frozenRows, frozenCols: s.frozenCols,
        };
      }),
    };
  }

  /** Changes after `since`, from memory when possible. */
  changesSince(store: Store, since: number): Change[] {
    const first = this.recent[0];
    if (first && first.v <= since + 1) return this.recent.filter((c) => c.v > since);
    return store.changesSince(this.id, since);
  }

  submit(store: Store, rawOps: unknown, meta: { base?: number; actor: string; client: string }): Change {
    let ops = validateOps(rawOps);
    const base = meta.base ?? this.version;
    if (base < this.version) {
      const prior = this.changesSince(store, base).filter((c) => c.client !== meta.client || !meta.client);
      ops = transformAll(ops, prior.flatMap((c) => c.ops), this.wb.doc);
    }
    this.wb.apply(ops); // throws OpError and leaves the doc untouched on invalid ops
    const change: Change = { v: this.version + 1, ops, actor: meta.actor, client: meta.client, ts: Date.now() };
    store.append(this.id, change, this.name);
    this.version = change.v;
    this.updated = change.ts;
    this.recent.push(change);
    if (this.recent.length > RECENT) this.recent.splice(0, this.recent.length - RECENT);
    if (this.version - this.snapVersion >= SNAPSHOT_EVERY) this.snapshot(store);
    this.broadcast(change);
    return change;
  }

  snapshot(store: Store): void {
    if (this.snapVersion === this.version) return;
    store.snapshot(this.id, this.version, this.wb.toJSON());
    this.snapVersion = this.version;
  }

  private broadcast(c: Change): void {
    const msg = `id: ${c.v}\nevent: change\ndata: ${JSON.stringify(c)}\n\n`;
    for (const res of this.subs) res.write(msg);
  }

  /** Tell subscribers the book is gone. */
  close(): void {
    for (const res of this.subs) {
      res.write("event: deleted\ndata: {}\n\n");
      res.end();
    }
    this.subs.clear();
  }
}

export class Hub {
  private books = new Map<string, LiveBook>();

  constructor(readonly store: Store) {
    for (const row of store.books()) {
      const wb = Workbook.fromJSON(row.snapshot);
      const changes = store.changesSince(row.id, row.snapVersion);
      for (const c of changes) wb.apply(c.ops);
      const lb = new LiveBook(row.id, wb, row.version, row.created, row.updated, row.snapVersion);
      lb.recent = changes.slice(-RECENT);
      this.books.set(row.id, lb);
    }
  }

  list(): LiveBook[] {
    return [...this.books.values()].sort((a, b) => b.updated - a.updated);
  }

  /** Find a book by id, exact name (case-insensitive), or a unique name prefix. */
  find(ref: string): LiveBook | undefined {
    const byId = this.books.get(ref);
    if (byId) return byId;
    const lower = ref.toLowerCase();
    const all = this.list();
    const exact = all.filter((b) => b.name.toLowerCase() === lower);
    if (exact.length === 1) return exact[0];
    const slug = (s: string) => s.toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "");
    const prefix = all.filter((b) => b.name.toLowerCase().startsWith(lower) || slug(b.name).startsWith(slug(ref)));
    return prefix.length === 1 ? prefix[0] : undefined;
  }

  create(name: string, doc?: DocJSON): LiveBook {
    const id = newId();
    const d: DocJSON = doc ?? { name, sheets: [sheetToJSON(newSheet(newId(6), "Sheet1"))] };
    d.name = name;
    const wb = Workbook.fromJSON(d);
    const now = Date.now();
    this.store.create(id, name, wb.toJSON(), now);
    const lb = new LiveBook(id, wb, 0, now, now, 0);
    this.books.set(id, lb);
    return lb;
  }

  remove(lb: LiveBook): void {
    lb.snapshot(this.store);
    this.store.remove(lb.id);
    this.books.delete(lb.id);
    lb.close();
  }

  /** Snapshot every book (on shutdown) so the next start replays nothing. */
  flush(): void {
    for (const b of this.books.values()) b.snapshot(this.store);
  }
}
