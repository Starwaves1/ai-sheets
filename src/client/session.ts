// Browser side of collaboration. Local edits apply immediately (optimistic). They are
// sent one change at a time, and the next is sent only after the server's SSE stream
// echoes the previous one back, so the server never has to rebase our own ops.
// When someone else's change arrives first, pending local ops are rebased over it.

import type { DocJSON } from "../core/model.ts";
import { touchedCells, transformAll, type Op } from "../core/ops.ts";
import { Workbook } from "../core/workbook.ts";

interface Change {
  v: number;
  ops: Op[];
  actor: string;
  client: string;
  ts: number;
  seq?: number;
}

interface Pending {
  ops: Op[];
  inverse: Op[];
  sent: boolean;
  seq: number;
}

export interface ChangeEvent {
  remote: boolean;
  structural: boolean; // rows/cols/sheets changed: layout must be rebuilt
  touched: Map<string, Array<[number, number]>>; // sheet id -> cells, for highlighting
  actor?: string;
  ops: Op[];
  ts?: number;
}

export type Status = "live" | "connecting" | "offline" | "deleted";

const STRUCTURAL = new Set(["insert", "delete", "size", "hide", "freeze", "view", "grid", "addSheet", "deleteSheet", "renameSheet", "moveSheet", "cond"]);

export class Session {
  wb!: Workbook;
  version = 0;
  readonly client = Math.random().toString(36).slice(2, 10);
  private pending: Pending[] = [];
  private undoStack: Op[][] = [];
  private redoStack: Op[][] = [];
  private seq = 0;
  private es: EventSource | null = null;
  status: Status = "connecting";
  onChange: (e: ChangeEvent) => void = () => {};
  onStatus: (s: Status) => void = () => {};
  onError: (msg: string) => void = () => {};
  onReset: () => void = () => {};

  constructor(readonly bookId: string) {}

  async load(): Promise<{ id: string; name: string }> {
    const res = await fetch(`/api/books/${encodeURIComponent(this.bookId)}/doc`);
    if (!res.ok) throw new Error(((await res.json().catch(() => ({}))) as { error?: string }).error ?? res.statusText);
    const data = (await res.json()) as { id: string; name: string; version: number; doc: DocJSON };
    this.wb = Workbook.fromJSON(data.doc);
    this.version = data.version;
    this.pending = [];
    return data;
  }

  connect(): void {
    this.es?.close();
    const es = new EventSource(`/api/books/${encodeURIComponent(this.bookId)}/events?since=${this.version}`);
    this.es = es;
    this.setStatus("connecting");
    es.addEventListener("hello", () => this.setStatus("live"));
    es.addEventListener("change", (e) => this.receive(JSON.parse((e as MessageEvent).data) as Change));
    es.addEventListener("reset", () => this.onReset());
    es.addEventListener("deleted", () => {
      es.close();
      this.setStatus("deleted");
    });
    es.onerror = () => {
      if (es.readyState !== EventSource.CLOSED) return this.setStatus("connecting");
      // The browser gave up (e.g. a proxy error page); keep trying, since saving waits on this stream.
      this.setStatus("offline");
      setTimeout(() => this.es === es && this.connect(), 3000);
    };
  }

  close(): void {
    this.es?.close();
    this.es = null;
  }

  get hasPending(): boolean {
    return this.pending.length > 0;
  }

  private setStatus(s: Status): void {
    if (this.status === s) return;
    this.status = s;
    this.onStatus(s);
  }

  /** Apply a local edit. Throws (and changes nothing) if the ops are invalid. */
  submit(ops: Op[], opts: { undoable?: boolean } = {}): void {
    if (!ops.length) return;
    const res = this.wb.apply(ops);
    this.pending.push({ ops, inverse: res.inverse, sent: false, seq: ++this.seq });
    if (opts.undoable !== false) {
      this.undoStack.push(res.inverse);
      if (this.undoStack.length > 200) this.undoStack.shift();
      this.redoStack = [];
    }
    this.emit(ops, false);
    this.flush();
  }

  undo(): Op[] | null {
    return this.step(this.undoStack, this.redoStack);
  }

  redo(): Op[] | null {
    return this.step(this.redoStack, this.undoStack);
  }

  get canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  get canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  private step(from: Op[][], to: Op[][]): Op[] | null {
    const ops = from.pop();
    if (!ops) return null;
    try {
      const res = this.wb.apply(ops);
      this.pending.push({ ops, inverse: res.inverse, sent: false, seq: ++this.seq });
      to.push(res.inverse);
      this.emit(ops, false);
      this.flush();
      return ops;
    } catch (e) {
      this.onError(`Couldn't undo: ${(e as Error).message}`);
      return null;
    }
  }

  private emit(ops: Op[], remote: boolean, c?: Change): void {
    const touched = new Map<string, Array<[number, number]>>();
    let structural = false;
    for (const op of ops) {
      if (STRUCTURAL.has(op.t)) structural = true;
      const t = touchedCells(op);
      if (!t) continue;
      let list = touched.get(t.s);
      if (!list) touched.set(t.s, (list = []));
      for (const cell of t.cells) list.push(cell);
    }
    this.onChange({ remote, structural, touched, ops, ...(c ? { actor: c.actor, ts: c.ts } : {}) });
  }

  private async flush(): Promise<void> {
    if (this.pending.some((p) => p.sent)) return; // wait for the echo of the one in flight
    const next = this.pending[0];
    if (!next) return;
    next.sent = true;
    const retry = (msg: string, ms: number) => {
      next.sent = false;
      this.onError(msg);
      setTimeout(() => this.flush(), ms);
    };
    try {
      const res = await fetch(`/api/books/${encodeURIComponent(this.bookId)}/ops`, {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ ops: next.ops, base: this.version, client: this.client, seq: next.seq }),
      });
      if (res.ok) return; // acknowledged through the event stream
      const msg = ((await res.json().catch(() => ({}))) as { error?: string }).error ?? res.statusText;
      if (res.status < 500) this.reject(next, msg);
      else retry(`Save failed (${msg}); retrying`, 2000);
    } catch {
      // The server deduplicates by seq, so resending a change that did arrive is safe.
      retry("Offline: edits will be saved when the server is back", 3000);
    }
  }

  /** The server refused a change: take it back out of the local state. */
  private reject(bad: Pending, msg: string): void {
    for (const p of [...this.pending].reverse()) this.wb.apply(p.inverse);
    this.pending = this.pending.filter((p) => p !== bad);
    for (const p of this.pending) p.inverse = this.wb.apply(p.ops).inverse;
    this.onError(`Change rejected: ${msg}`);
    this.emit(bad.ops, true);
    this.flush();
  }

  private receive(c: Change): void {
    if (c.v <= this.version) return;
    this.version = c.v;
    if (c.client === this.client) {
      // Our own change, acknowledged in order: send the next one.
      // Sends wait for this echo, so it is always the oldest pending change.
      this.pending.shift();
      this.flush();
      return;
    }
    // Someone else's change was ordered before our pending ones: rebase.
    for (const p of [...this.pending].reverse()) this.wb.apply(p.inverse);
    try {
      this.wb.apply(c.ops);
    } catch (e) {
      this.onError(`Out of sync (${(e as Error).message}); reloading`);
      this.onReset();
      return;
    }
    // Rebase in order: each pending change moves past the remote ops, and the remote
    // ops move past it before meeting the next pending change.
    const moves = c.ops.some((op) => op.t === "insert" || op.t === "delete" || op.t === "deleteSheet");
    let remote = c.ops;
    for (const p of this.pending) {
      const before = p.ops;
      p.ops = transformAll(before, remote, this.wb.doc);
      remote = transformAll(remote, before, this.wb.doc, false); // remote was ordered first
      try {
        p.inverse = this.wb.apply(p.ops).inverse;
      } catch {
        p.ops = [];
        p.inverse = [];
      }
    }
    if (moves) {
      this.undoStack = this.undoStack.map((b) => transformAll(b, c.ops, this.wb.doc)).filter((b) => b.length);
      this.redoStack = this.redoStack.map((b) => transformAll(b, c.ops, this.wb.doc)).filter((b) => b.length);
    }
    this.emit(c.ops, true, c);
  }
}
