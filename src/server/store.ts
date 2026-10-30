// Persistence: one SQLite file. Each workbook is a snapshot plus an append-only
// log of changes; loading replays the changes recorded after the snapshot.

import { DatabaseSync } from "node:sqlite";
import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import type { DocJSON } from "../core/model.ts";
import type { Change, Meta } from "../core/history.ts";
import type { Op } from "../core/ops.ts";

export type { Change };

export interface Checkpoint {
  id: number;
  v: number;
  name: string;
  actor: string;
  ts: number;
}

export interface BookRow {
  id: string;
  name: string;
  created: number;
  updated: number;
  version: number;
  snapVersion: number;
  snapshot: DocJSON;
}

export class Store {
  private db: DatabaseSync;

  constructor(file: string) {
    if (file !== ":memory:") mkdirSync(dirname(file), { recursive: true });
    this.db = new DatabaseSync(file);
    this.db.exec(`
      PRAGMA journal_mode = WAL;
      PRAGMA synchronous = NORMAL;
      CREATE TABLE IF NOT EXISTS books (
        id TEXT PRIMARY KEY,
        name TEXT NOT NULL,
        created INTEGER NOT NULL,
        updated INTEGER NOT NULL,
        version INTEGER NOT NULL,
        snap_version INTEGER NOT NULL,
        snapshot TEXT NOT NULL,
        deleted INTEGER NOT NULL DEFAULT 0
      );
      CREATE TABLE IF NOT EXISTS changes (
        book TEXT NOT NULL,
        v INTEGER NOT NULL,
        ts INTEGER NOT NULL,
        actor TEXT NOT NULL,
        client TEXT NOT NULL,
        ops TEXT NOT NULL,
        PRIMARY KEY (book, v)
      );
    `);
    const cols = this.db.prepare("PRAGMA table_info(changes)").all() as Array<{ name: string }>;
    if (!cols.some((c) => c.name === "seq")) this.db.exec("ALTER TABLE changes ADD COLUMN seq INTEGER");
    // History: the inverse of each change (to revert it), why it was made, every snapshot kept
    // (to rebuild any version from the nearest one), and named versions.
    if (!cols.some((c) => c.name === "inv")) this.db.exec("ALTER TABLE changes ADD COLUMN inv TEXT");
    if (!cols.some((c) => c.name === "meta")) this.db.exec("ALTER TABLE changes ADD COLUMN meta TEXT");
    this.db.exec(`
      CREATE TABLE IF NOT EXISTS snaps (
        book TEXT NOT NULL,
        v INTEGER NOT NULL,
        doc TEXT NOT NULL,
        PRIMARY KEY (book, v)
      );
      INSERT OR IGNORE INTO snaps (book, v, doc) SELECT id, snap_version, snapshot FROM books WHERE deleted = 0 AND snapshot != '';
      CREATE TABLE IF NOT EXISTS checkpoints (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        book TEXT NOT NULL,
        v INTEGER NOT NULL,
        name TEXT NOT NULL,
        actor TEXT NOT NULL,
        ts INTEGER NOT NULL
      );
    `);
    // The newest snapshot lives in snaps; books.snapshot no longer keeps a second copy.
    this.db.exec(`UPDATE books SET snapshot = '' WHERE snapshot != '' AND EXISTS (SELECT 1 FROM snaps s WHERE s.book = books.id AND s.v = books.snap_version)`);
  }

  books(): BookRow[] {
    const rows = this.db.prepare(`SELECT b.*, COALESCE(s.doc, NULLIF(b.snapshot, '')) AS doc FROM books b
      LEFT JOIN snaps s ON s.book = b.id AND s.v = b.snap_version WHERE b.deleted = 0 ORDER BY b.updated DESC`).all() as Array<Record<string, unknown>>;
    return rows.map((r) => ({
      id: r.id as string,
      name: r.name as string,
      created: r.created as number,
      updated: r.updated as number,
      version: r.version as number,
      snapVersion: r.snap_version as number,
      snapshot: JSON.parse(r.doc as string) as DocJSON,
    }));
  }

  create(id: string, name: string, doc: DocJSON, now: number): void {
    this.db
      .prepare("INSERT INTO books (id, name, created, updated, version, snap_version, snapshot) VALUES (?, ?, ?, ?, 0, 0, '')")
      .run(id, name, now, now);
    this.db.prepare("INSERT OR REPLACE INTO snaps (book, v, doc) VALUES (?, 0, ?)").run(id, JSON.stringify(doc));
  }

  /** Record a change and bump the book's version in one transaction. */
  append(book: string, c: Change, name: string): void {
    this.db.exec("BEGIN");
    try {
      this.db
        .prepare("INSERT INTO changes (book, v, ts, actor, client, ops, seq, inv, meta) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")
        .run(book, c.v, c.ts, c.actor, c.client, JSON.stringify(c.ops), c.seq ?? null, c.inv ? JSON.stringify(c.inv) : null, c.meta ? JSON.stringify(c.meta) : null);
      this.db.prepare("UPDATE books SET version = ?, updated = ?, name = ? WHERE id = ?").run(c.v, c.ts, name, book);
      this.db.exec("COMMIT");
    } catch (e) {
      this.db.exec("ROLLBACK");
      throw e;
    }
  }

  snapshot(book: string, version: number, doc: DocJSON): void {
    this.db.prepare("INSERT OR REPLACE INTO snaps (book, v, doc) VALUES (?, ?, ?)").run(book, version, JSON.stringify(doc));
    this.db.prepare("UPDATE books SET snap_version = ? WHERE id = ?").run(version, book);
  }

  /** The newest kept snapshot at or before version v. */
  snapBefore(book: string, v: number): { v: number; doc: DocJSON } | null {
    const r = this.db.prepare("SELECT v, doc FROM snaps WHERE book = ? AND v <= ? ORDER BY v DESC LIMIT 1").get(book, v) as { v: number; doc: string } | undefined;
    return r ? { v: r.v, doc: JSON.parse(r.doc) as DocJSON } : null;
  }

  /** The oldest version that can be rebuilt (history before it was not kept). */
  oldestSnap(book: string): number {
    const r = this.db.prepare("SELECT MIN(v) AS v FROM snaps WHERE book = ?").get(book) as { v: number | null };
    return r.v ?? 0;
  }

  changesSince(book: string, since: number, limit = 100_000): Change[] {
    return this.changesBetween(book, since, Number.MAX_SAFE_INTEGER, limit);
  }

  /** Changes with since < v <= upto, oldest first. */
  changesBetween(book: string, since: number, upto: number, limit = 100_000): Change[] {
    const rows = this.db
      .prepare("SELECT v, ts, actor, client, ops, seq, inv, meta FROM changes WHERE book = ? AND v > ? AND v <= ? ORDER BY v LIMIT ?")
      .all(book, since, upto, limit) as Array<Record<string, unknown>>;
    return rows.map((r) => ({
      ...(r.seq !== null ? { seq: r.seq as number } : {}),
      v: r.v as number,
      ts: r.ts as number,
      actor: r.actor as string,
      client: r.client as string,
      ops: JSON.parse(r.ops as string) as Op[],
      inv: r.inv ? (JSON.parse(r.inv as string) as Op[]) : null,
      ...(r.meta ? { meta: JSON.parse(r.meta as string) as Meta } : {}),
    }));
  }

  setInverse(book: string, v: number, inv: Op[]): void {
    this.db.prepare("UPDATE changes SET inv = ? WHERE book = ? AND v = ?").run(JSON.stringify(inv), book, v);
  }

  checkpoints(book: string): Checkpoint[] {
    return this.db.prepare("SELECT id, v, name, actor, ts FROM checkpoints WHERE book = ? ORDER BY v DESC, id DESC").all(book) as unknown as Checkpoint[];
  }

  addCheckpoint(book: string, v: number, name: string, actor: string, ts: number): Checkpoint {
    const r = this.db.prepare("INSERT INTO checkpoints (book, v, name, actor, ts) VALUES (?, ?, ?, ?, ?)").run(book, v, name, actor, ts);
    return { id: Number(r.lastInsertRowid), v, name, actor, ts };
  }

  removeCheckpoint(book: string, id: number): boolean {
    return this.db.prepare("DELETE FROM checkpoints WHERE book = ? AND id = ?").run(book, id).changes > 0;
  }

  /** Soft delete: the row and its history stay in the file. */
  remove(book: string): void {
    this.db.prepare("UPDATE books SET deleted = 1 WHERE id = ?").run(book);
  }

  /** Hard delete: the book and all its history leave the file (tests and scratch books). */
  purge(book: string): void {
    this.db.exec("BEGIN");
    for (const t of ["changes", "snaps", "checkpoints"]) this.db.prepare(`DELETE FROM ${t} WHERE book = ?`).run(book);
    this.db.prepare("DELETE FROM books WHERE id = ?").run(book);
    this.db.exec("COMMIT");
  }

  close(): void {
    this.db.close();
  }
}
