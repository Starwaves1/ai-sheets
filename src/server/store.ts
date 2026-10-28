// Persistence: one SQLite file. Each workbook is a snapshot plus an append-only
// log of changes; loading replays the changes recorded after the snapshot.

import { DatabaseSync } from "node:sqlite";
import { mkdirSync } from "node:fs";
import { dirname } from "node:path";
import type { DocJSON } from "../core/model.ts";
import type { Op } from "../core/ops.ts";

export interface Change {
  v: number;
  ops: Op[];
  actor: string;
  client: string;
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
  }

  books(): BookRow[] {
    const rows = this.db.prepare("SELECT * FROM books WHERE deleted = 0 ORDER BY updated DESC").all() as Array<Record<string, unknown>>;
    return rows.map((r) => ({
      id: r.id as string,
      name: r.name as string,
      created: r.created as number,
      updated: r.updated as number,
      version: r.version as number,
      snapVersion: r.snap_version as number,
      snapshot: JSON.parse(r.snapshot as string) as DocJSON,
    }));
  }

  create(id: string, name: string, doc: DocJSON, now: number): void {
    this.db
      .prepare("INSERT INTO books (id, name, created, updated, version, snap_version, snapshot) VALUES (?, ?, ?, ?, 0, 0, ?)")
      .run(id, name, now, now, JSON.stringify(doc));
  }

  /** Record a change and bump the book's version in one transaction. */
  append(book: string, c: Change, name: string): void {
    this.db.exec("BEGIN");
    try {
      this.db
        .prepare("INSERT INTO changes (book, v, ts, actor, client, ops) VALUES (?, ?, ?, ?, ?, ?)")
        .run(book, c.v, c.ts, c.actor, c.client, JSON.stringify(c.ops));
      this.db.prepare("UPDATE books SET version = ?, updated = ?, name = ? WHERE id = ?").run(c.v, c.ts, name, book);
      this.db.exec("COMMIT");
    } catch (e) {
      this.db.exec("ROLLBACK");
      throw e;
    }
  }

  snapshot(book: string, version: number, doc: DocJSON): void {
    this.db.prepare("UPDATE books SET snapshot = ?, snap_version = ? WHERE id = ?").run(JSON.stringify(doc), version, book);
  }

  changesSince(book: string, since: number, limit = 100_000): Change[] {
    const rows = this.db
      .prepare("SELECT v, ts, actor, client, ops FROM changes WHERE book = ? AND v > ? ORDER BY v LIMIT ?")
      .all(book, since, limit) as Array<Record<string, unknown>>;
    return rows.map((r) => ({
      v: r.v as number,
      ts: r.ts as number,
      actor: r.actor as string,
      client: r.client as string,
      ops: JSON.parse(r.ops as string) as Op[],
    }));
  }

  /** Soft delete: the row and its history stay in the file. */
  remove(book: string): void {
    this.db.prepare("UPDATE books SET deleted = 1 WHERE id = ?").run(book);
  }

  close(): void {
    this.db.close();
  }
}
