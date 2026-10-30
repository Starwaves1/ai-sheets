// The History panel (side panel; bottom sheet on phones) and the Cell history panel.
// Rows are sessions: one person's edits close together, described in plain words. Selecting one
// previews the sheet as it was right after it, with its cells marked; arrow keys scrub.

import { describe, groupSessions, summarize, type Change, type Meta, type Session, type Summary } from "../core/history.ts";
import type { Workbook } from "../core/workbook.ts";
import { clearChildren, h, icon, ICONS } from "./dom.ts";
import { myName, when } from "./flags.ts";
import { menu, type MenuItem } from "./popups.ts";

export interface Checkpoint {
  id: number;
  v: number;
  name: string;
  actor: string;
  ts: number;
}

export interface HistoryData {
  version: number;
  oldest: number;
  more: boolean;
  sheets: Record<string, string>;
  summaries: Summary[];
  checkpoints: Checkpoint[];
}

/** What the panel asks the page to do. */
export interface HistoryHost {
  bookId: string;
  wb(): Workbook; // live
  version(): number;
  selectionRef(): string; // "Runs!M:M" for the current selection
  preview(p: { v: number; from?: number; only?: number[]; title: string; actor?: string }): void;
  exitPreview(): void;
  previewing(): number | null; // version shown, or null when live
  revert(versions: number[], what: string): void;
  restore(v: number, ts: number): void;
  nameVersion(v: number, anchor: HTMLElement): void;
  removeName(cp: Checkpoint): void;
  askName(anchor: HTMLElement): void;
  close(): void;
}

const COLORS = ["blue", "violet", "green", "amber", "red", "grey"];
export const actorColor = (name: string) => {
  let x = 0;
  for (const ch of name) x = (x * 31 + ch.charCodeAt(0)) >>> 0;
  return COLORS[x % COLORS.length]!;
};

const dayOf = (ts: number) => {
  const d = new Date(ts), now = new Date();
  const y = new Date(now);
  y.setDate(y.getDate() - 1);
  if (d.toDateString() === now.toDateString()) return "Today";
  if (d.toDateString() === y.toDateString()) return "Yesterday";
  return d.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
};
const clock = (ts: number) => new Date(ts).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
const timeLabel = (ts: number) => {
  const d = new Date(ts);
  return d.toDateString() === new Date().toDateString() ? clock(ts) : `${d.toLocaleDateString([], { month: "short", day: "numeric" })}, ${clock(ts)}`;
};

type Row = { kind: "session"; s: Session } | { kind: "change"; s: Session; sum: Summary } | { kind: "name"; cp: Checkpoint };

export class HistoryPanel {
  readonly el: HTMLElement;
  private list: HTMLElement;
  private filters: HTMLElement;
  private foot: HTMLElement;
  private data: HistoryData | null = null;
  private by = new Set<string>();
  private range: string | null = null;
  private namedOnly = false;
  private expanded = new Set<string>();
  private selected: string | null = null; // row key
  private rows: Array<{ key: string; row: Row }> = [];
  private loading = 0;
  private actors = new Set<string>();

  constructor(private host: HistoryHost) {
    this.list = h("div", { class: "hist-list", tabindex: "0", role: "listbox", "aria-label": "Changes, newest first" });
    this.filters = h("div", { class: "hist-filters" });
    this.foot = h("div", { class: "hist-foot" });
    this.el = h("aside", { class: "panel hist-panel", "aria-label": "History" },
      h("header", {}, h("h2", {}, "History"), h("button", { class: "icon-btn", title: "Close", onclick: () => host.close() }, icon(ICONS.close))),
      this.filters, this.list, this.foot);
    this.list.addEventListener("keydown", (e) => this.key(e));
    // Relative times ("5 min ago") are refreshed when you come back to the panel, not on a timer.
    this.el.addEventListener("pointerenter", () => this.render());
    void this.load();
  }

  /** Fetch the timeline (again), with the current person/range filters. */
  async load(): Promise<void> {
    const n = ++this.loading;
    const qs = new URLSearchParams();
    if (this.by.size) qs.set("by", [...this.by].join(","));
    if (this.range) qs.set("range", this.range);
    const res = await fetch(`/api/books/${encodeURIComponent(this.host.bookId)}/history?${qs}`);
    if (n !== this.loading) return;
    if (!res.ok) {
      clearChildren(this.list);
      this.list.append(h("p", { class: "hint" }, `Couldn't load history (${res.status}).`));
      return;
    }
    const data = (await res.json()) as HistoryData;
    for (const x of data.summaries) this.actors.add(x.actor); // chips stay for everyone seen, even when filtered
    this.data = data;
    this.render();
  }

  /** A change just arrived over the live stream. */
  addChange(c: Change): void {
    if (!this.data) return;
    if (this.range || this.by.size) {
      void this.load(); // filters are applied by the server
      return;
    }
    if (this.data.summaries.some((s) => s.v === c.v)) return;
    this.data.summaries.push(summarize(c, this.host.wb().doc));
    this.actors.add(c.actor);
    this.data.version = c.v;
    this.render();
  }

  setCheckpoints(list: Checkpoint[]): void {
    if (!this.data) return;
    this.data.checkpoints = list;
    this.render();
  }

  /** Forget the highlighted row (the page left preview on its own). */
  clearSelection(): void {
    this.selected = null;
    this.render();
  }

  private sheetName = (id: string) => (this.data && Object.keys(this.data.sheets).length > 1 ? this.data.sheets[id] ?? null : null);

  render(): void {
    const d = this.data;
    this.renderFilters();
    clearChildren(this.list);
    clearChildren(this.foot);
    if (!d) {
      this.list.append(h("p", { class: "hint" }, "Loading…"));
      return;
    }
    const sessions = this.namedOnly ? [] : groupSessions(d.summaries);
    const names = [...d.checkpoints].sort((a, b) => b.v - a.v);
    // Merge sessions and named versions by version, newest first.
    const rows: Row[] = [];
    let ni = 0;
    for (const s of sessions) {
      while (ni < names.length && names[ni]!.v >= s.last) rows.push({ kind: "name", cp: names[ni++]! });
      rows.push({ kind: "session", s });
      if (this.expanded.has(s.id)) for (const sum of [...s.sums].reverse()) rows.push({ kind: "change", s, sum });
    }
    while (ni < names.length) rows.push({ kind: "name", cp: names[ni++]! });
    this.rows = [];

    const live = this.host.previewing() === null;
    const now = h("button", { class: `hist-now${live ? " on" : ""}`, role: "option", "aria-selected": String(live), onclick: () => this.goLive() },
      h("span", { class: "dot live-dot" }), h("span", { class: "what" }, "Now"), h("span", { class: "when" }, `v${d.version}`));
    this.list.append(now);
    let day = "";
    for (const row of rows) {
      const ts = row.kind === "name" ? row.cp.ts : row.kind === "session" ? row.s.end : row.sum.ts;
      if (row.kind !== "change" && dayOf(ts) !== day) {
        day = dayOf(ts);
        this.list.append(h("div", { class: "hist-day" }, day));
      }
      this.list.append(this.renderRow(row));
    }
    if (!rows.length) {
      this.list.append(h("p", { class: "hint" }, this.namedOnly ? "No named versions yet. Select a change and choose “Name this version”." :
        this.range || this.by.size ? "No changes match." : "No changes yet."));
    }
    if (d.more) this.list.append(h("p", { class: "hint" }, "Older changes are not shown here; `sheets history --since` reaches them."));
    if (d.summaries.length && d.summaries[0]!.v <= d.oldest) {
      this.list.append(h("p", { class: "hint" }, `Changes up to v${d.oldest} were made before history was kept: they are listed, but can't be previewed or reverted.`));
    }
    if (!myName()) {
      const btn = h("button", { class: "link" }, "Set your name");
      btn.addEventListener("click", () => this.host.askName(btn));
      this.foot.append(h("p", { class: "hint" }, "Your edits show up as “browser”. ", btn));
    }
  }

  private renderFilters(): void {
    clearChildren(this.filters);
    const actors = [...this.actors].sort();
    const chip = (label: string, on: boolean, run: () => void, color?: string) =>
      h("button", { class: `chip${on ? " on" : ""}`, "aria-pressed": String(on), onclick: run }, color ? h("span", { class: `dot ink-${color}` }) : null, label);
    this.filters.append(chip("Everyone", !this.by.size, () => {
      this.by.clear();
      void this.load();
    }));
    for (const a of actors) {
      this.filters.append(chip(a, this.by.has(a), () => {
        if (this.by.has(a)) this.by.delete(a);
        else this.by.add(a);
        void this.load();
      }, actorColor(a)));
    }
    const sel = this.range;
    this.filters.append(
      chip(sel ? `In ${sel.replace(/^[^!]*!/, "")}` : "In selection", !!sel, () => {
        this.range = this.range ? null : this.host.selectionRef();
        void this.load();
      }),
      chip("Named", this.namedOnly, () => {
        this.namedOnly = !this.namedOnly;
        this.render();
      }),
    );
  }

  private renderRow(row: Row): HTMLElement {
    const d = this.data!;
    if (row.kind === "name") {
      const key = `n${row.cp.id}`;
      this.rows.push({ key, row });
      const el = h("div", { class: `hist-row named${this.selected === key ? " on" : ""}`, role: "option", "data-key": key },
        h("span", { class: "flagmark", "aria-hidden": "true" }),
        h("button", { class: "body", onclick: () => this.pick(key) },
          h("span", { class: "what" }, row.cp.name),
          h("span", { class: "when" }, `${row.cp.actor} named v${row.cp.v} · ${when(row.cp.ts)}`)),
        this.moreButton(() => [
          { label: "Preview this version", run: () => this.pick(key) },
          { label: "Compare with now", run: () => this.host.preview({ v: d.version, from: row.cp.v, title: `Changes since “${row.cp.name}”` }) },
          { label: "Restore this version…", run: () => this.host.restore(row.cp.v, row.cp.ts) },
          "-",
          { label: "Remove the name", run: () => this.host.removeName(row.cp) },
        ]));
      return el;
    }
    const s = row.s;
    const sums = row.kind === "change" ? [row.sum] : s.sums;
    const key = row.kind === "change" ? `v${row.sum.v}` : s.id;
    this.rows.push({ key, row });
    const first = row.kind === "change" ? row.sum.v : s.first;
    const last = row.kind === "change" ? row.sum.v : s.last;
    const old = first - 1 < d.oldest; // the version before it can't be rebuilt
    const text = describe(sums, this.sheetName, timeLabel);
    const meta = sums.length === 1 && sums[0]!.meta ? sums[0]!.meta : undefined;
    const count = row.kind === "session" && s.versions.length > 1 ? ` · ${s.versions.length} edits` : "";
    const time = row.kind === "change" ? clock(row.sum.ts) : s.end - s.start > 60_000 ? `${clock(s.start)}–${clock(s.end)}` : when(s.end);
    const el = h("div", { class: `hist-row ${row.kind}${this.selected === key ? " on" : ""}${old ? " old" : ""}${meta ? ` meta-${meta.kind}` : ""}`, role: "option", "aria-selected": String(this.selected === key), "data-key": key },
      row.kind === "session" ? h("span", { class: `dot ink-${actorColor(s.actor)}`, "aria-hidden": "true" }) : h("span", { class: "tick", "aria-hidden": "true" }),
      h("button", { class: "body", title: old ? "Made before history was kept" : "Show the sheet right after this", onclick: (e: Event) => this.pick(key, (e as MouseEvent).shiftKey) },
        h("span", { class: "what" }, row.kind === "session" ? h("b", {}, s.actor) : null, row.kind === "session" ? " " : "", text),
        h("span", { class: "when" }, `${time}${count}${old ? " · before history" : ""}`)),
      this.moreButton(() => {
        const items: Array<MenuItem | "-"> = [];
        const versions = row.kind === "change" ? [row.sum.v] : s.versions;
        const who = s.actor === myName() ? "your" : `${s.actor}’s`;
        const canRevert = !old && sums.every((x) => x.undoable);
        items.push({ label: "Preview", disabled: old, run: () => this.pick(key) });
        if (row.kind === "session" && s.versions.length > 1) {
          items.push({ label: this.expanded.has(s.id) ? "Hide the separate edits" : `Show the ${s.versions.length} edits`, run: () => this.toggle(s.id) });
        }
        items.push("-",
          { label: row.kind === "change" ? "Revert this edit" : versions.length > 1 ? `Revert ${who} ${versions.length} edits` : "Revert this edit", disabled: !canRevert,
            run: () => this.host.revert(versions, row.kind === "change" ? `v${row.sum.v}` : `${who} edits from ${timeLabel(s.start)}`) },
          { label: "Compare with now", disabled: old, run: () => this.host.preview({ v: d.version, from: last, title: `Changes since ${timeLabel(row.kind === "change" ? row.sum.ts : s.end)}` }) },
          { label: "Restore this version…", disabled: last < d.oldest, run: () => this.host.restore(last, row.kind === "change" ? row.sum.ts : s.end) },
          { label: "Name this version…", run: () => this.host.nameVersion(last, el) });
        return items;
      }));
    if (row.kind === "session" && s.versions.length > 1) {
      el.append(h("button", { class: "expand", title: this.expanded.has(s.id) ? "Hide the separate edits" : "Show each edit", "aria-expanded": String(this.expanded.has(s.id)),
        onclick: () => this.toggle(s.id) }, icon(ICONS.caret)));
    }
    return el;
  }

  private moreButton(items: () => Array<MenuItem | "-">): HTMLElement {
    const b = h("button", { class: "icon-btn more", title: "More", "aria-label": "More actions" }, "⋯");
    b.addEventListener("click", (e) => {
      e.stopPropagation();
      const r = b.getBoundingClientRect();
      menu(items(), Math.max(8, r.right - 230), r.bottom + 2);
    });
    return b;
  }

  private toggle(id: string): void {
    if (this.expanded.has(id)) this.expanded.delete(id);
    else this.expanded.add(id);
    this.render();
  }

  private goLive(): void {
    this.selected = null;
    this.host.exitPreview();
    this.render();
  }

  /** Select a row and preview it; with Shift, compare it with the selected row instead. */
  pick(key: string, compare = false): void {
    const d = this.data;
    const entry = this.rows.find((x) => x.key === key);
    if (!d || !entry) return;
    const prev = this.rows.find((x) => x.key === this.selected);
    const at = (r: Row) => (r.kind === "name" ? r.cp.v : r.kind === "change" ? r.sum.v : r.s.last);
    if (compare && prev && prev.key !== key) {
      const [a, b] = [at(prev.row), at(entry.row)].sort((x, y) => x - y);
      this.host.preview({ v: b!, from: a!, title: `Changes from v${a} to v${b}` });
      return;
    }
    this.selected = key;
    const r = entry.row;
    if (r.kind === "name") this.host.preview({ v: r.cp.v, title: `“${r.cp.name}”` });
    else {
      const versions = r.kind === "change" ? [r.sum.v] : r.s.versions;
      const first = r.kind === "change" ? r.sum.v : r.s.first;
      if (first - 1 < d.oldest) return this.render();
      const sums = r.kind === "change" ? [r.sum] : r.s.sums;
      this.host.preview({ v: at(r), from: first - 1, only: versions, actor: r.s.actor, title: `${r.s.actor} ${describe(sums, this.sheetName, timeLabel)}` });
    }
    this.render();
    this.list.querySelector<HTMLElement>(`[data-key="${key}"]`)?.scrollIntoView({ block: "nearest" });
  }

  /** Up/Down scrub through the timeline; Enter previews; Escape goes back to live. */
  private key(e: KeyboardEvent): void {
    if (e.key === "Escape") {
      this.goLive();
      return;
    }
    if (e.key !== "ArrowDown" && e.key !== "ArrowUp") return;
    e.preventDefault();
    const i = this.rows.findIndex((x) => x.key === this.selected);
    const next = this.rows[e.key === "ArrowDown" ? i + 1 : i - 1];
    if (next) this.pick(next.key);
    else if (e.key === "ArrowUp") this.goLive();
  }
}

// ---- cell history --------------------------------------------------------------------------

export interface CellHistory {
  ref: string;
  sheet: string;
  r: number;
  c: number;
  now: string | null;
  oldest: number;
  entries: Array<{ v: number; actor: string; ts: number; value: string | null; before?: string | null; meta?: Meta }>;
}

export class CellHistoryPanel {
  readonly el: HTMLElement;

  constructor(data: CellHistory, actions: { restore: (value: string | null, ts: number) => void; preview: (v: number) => void; close: () => void }) {
    const list = h("ol", { class: "cell-hist" });
    const val = (v: string | null | undefined) => (v === null || v === undefined || v === "" ? h("span", { class: "empty" }, "empty") : h("code", {}, v));
    data.entries.forEach((e, i) => {
      const current = i === 0 && e.value === data.now;
      list.append(h("li", { class: current ? "current" : "" },
        h("button", { class: "body", title: "Show the sheet at this point", onclick: () => actions.preview(e.v) },
          val(e.value),
          h("span", { class: "meta" }, h("span", { class: `dot ink-${actorColor(e.actor)}` }), `${e.actor} · ${timeLabel(e.ts)}${e.meta?.kind === "revert" ? " · revert" : e.meta?.kind === "restore" ? " · restore" : ""}`)),
        current ? h("span", { class: "now" }, "now") : h("button", { class: "btn restore", onclick: () => actions.restore(e.value, e.ts) }, "Restore")));
    });
    const last = data.entries.at(-1);
    if (last && last.before !== undefined) {
      list.append(h("li", { class: "origin" },
        h("div", { class: "body" }, val(last.before), h("span", { class: "meta" }, "before that")),
        data.now === last.before ? h("span", { class: "now" }, "now") : h("button", { class: "btn restore", onclick: () => actions.restore(last.before ?? null, last.ts - 1) }, "Restore")));
    }
    if (!data.entries.length) list.append(h("li", { class: "empty-note" }, `No edits to this cell since history began (v${data.oldest}).`));
    this.el = h("aside", { class: "panel cell-hist-panel", "aria-label": `History of ${data.ref}` },
      h("header", {}, h("h2", {}, "Cell history ", h("span", { class: "ref" }, data.ref)), h("button", { class: "icon-btn", title: "Close", onclick: actions.close }, icon(ICONS.close))),
      h("p", { class: "hint" }, "Every value this cell had, newest first. Restore puts one back as a new edit, so Undo takes it back."),
      list);
  }
}
