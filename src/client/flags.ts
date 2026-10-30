// Flags in the browser: who you are, the flag popover, the card on a flagged cell, the Flags panel.

import { flagGroups, type FlagGroup } from "../core/flags.ts";
import type { Flag } from "../core/model.ts";
import type { Workbook } from "../core/workbook.ts";
import { clearChildren, h, icon, ICONS } from "./dom.ts";
import { popup } from "./popups.ts";

// ---- your name ------------------------------------------------------------------

const NAME_KEY = "sheets.name";

/** The name others see on your edits and flags; "" until you set one. */
export function myName(): string {
  try {
    return localStorage.getItem(NAME_KEY) ?? "";
  } catch {
    return "";
  }
}

export function setMyName(name: string): void {
  try {
    localStorage.setItem(NAME_KEY, name.trim().slice(0, 40));
  } catch {}
}

/** Small popover asking for a name. */
export function namePopover(anchor: { getBoundingClientRect(): DOMRect }, done: (name: string) => void): void {
  const input = h("input", { class: "name-input", value: myName(), placeholder: "Your name", "aria-label": "Your name", maxlength: "40", autocomplete: "name" });
  const form = h("form", { class: "name-pop" },
    h("label", {}, "Your name", input),
    h("p", { class: "hint" }, "Shown on your edits and flags, here and to agents."),
    h("div", { class: "actions" }, h("button", { type: "submit", class: "btn primary" }, "Save")));
  let close = () => {};
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    const name = input.value.trim();
    if (!name) return input.focus();
    setMyName(name);
    close();
    done(name);
  });
  const r = anchor.getBoundingClientRect();
  close = popup(form, Math.max(8, r.right - 260), r.bottom + 6);
  input.focus();
  input.select();
}

// ---- time ------------------------------------------------------------------------

export function when(ts: number): string {
  const d = new Date(ts), now = new Date();
  const s = (Date.now() - ts) / 1000;
  const time = d.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
  if (s < 45) return "just now";
  if (s < 3600) return `${Math.round(s / 60)} min ago`;
  if (d.toDateString() === now.toDateString()) return time;
  const y = new Date(now);
  y.setDate(y.getDate() - 1);
  if (d.toDateString() === y.toDateString()) return `yesterday ${time}`;
  return `${d.toLocaleDateString([], { month: "short", day: "numeric" })} ${time}`;
}

// ---- flag popover (toolbar button) ------------------------------------------------

export function flagPopover(opts: {
  anchor: { getBoundingClientRect(): DOMRect };
  where: string; // "B4" or "B4:C9"
  current: Flag | null;
  openCount: number;
  save: (note: string) => void;
  remove: () => void;
  showAll: () => void;
}): void {
  const note = h("textarea", { class: "flag-note", rows: "3", placeholder: "What should Claude look at?", "aria-label": "Note" });
  note.value = opts.current && !opts.current.done ? opts.current.note : "";
  const needName = !myName();
  const name = needName ? h("input", { class: "name-input", placeholder: "Your name", "aria-label": "Your name", maxlength: "40", autocomplete: "name" }) : null;
  const form = h("form", { class: "flag-pop" },
    h("div", { class: "title" }, h("span", { class: "flag-dot" }), opts.current && !opts.current.done ? `Flag on ${opts.where}` : `Flag ${opts.where} for Claude`),
    note,
    name ? h("label", { class: "inline" }, h("span", {}, "Your name"), name) : null,
    h("div", { class: "actions" },
      opts.current ? h("button", { type: "button", class: "btn", onclick: () => { close(); opts.remove(); } }, opts.current.done ? "Dismiss" : "Remove") : null,
      h("button", { type: "submit", class: "btn primary flag-save" }, opts.current && !opts.current.done ? "Save" : "Flag")),
    h("button", { type: "button", class: "link all-flags", onclick: () => { close(); opts.showAll(); } },
      opts.openCount ? `All open flags (${opts.openCount})` : "All flags"));
  let close = () => {};
  const submit = () => {
    const text = note.value.trim();
    if (!text) return note.focus();
    if (name) {
      if (!name.value.trim()) return name.focus();
      setMyName(name.value);
    }
    close();
    opts.save(text);
  };
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    submit();
  });
  note.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      submit();
    }
  });
  const r = opts.anchor.getBoundingClientRect();
  close = popup(form, Math.max(8, r.right - 300), r.bottom + 6);
  note.focus();
}

// ---- card on a flagged cell ---------------------------------------------------------

export function flagCard(flag: Flag, actions: { edit: () => void; remove: () => void }): HTMLElement {
  const card = h("div", { class: `flag-card${flag.done ? " done" : ""}`, role: "note" },
    h("p", { class: "note" }, flag.note),
    h("p", { class: "meta" }, h("span", { class: "flag-dot" }), `${flag.by} · ${when(flag.ts)}`));
  if (flag.done) {
    card.append(h("div", { class: "reply" },
      h("p", { class: "meta" }, `✓ Resolved by ${flag.done.by} · ${when(flag.done.ts)}`),
      flag.done.reply ? h("p", { class: "note" }, flag.done.reply) : null));
  }
  card.append(h("div", { class: "actions" },
    flag.done ? null : h("button", { class: "link", onclick: actions.edit }, "Edit"),
    h("button", { class: "link", onclick: actions.remove }, flag.done ? "Dismiss" : "Remove")));
  return card;
}

// ---- Flags panel ------------------------------------------------------------------------

export class FlagsPanel {
  readonly el: HTMLElement;
  private list: HTMLElement;

  constructor(private wb: () => Workbook, private jump: (g: FlagGroup) => void, private dismiss: (g: FlagGroup) => void, close: () => void) {
    this.list = h("ol", { class: "flag-list" });
    this.el = h("aside", { class: "panel flags-panel", "aria-label": "Flags" },
      h("header", {}, h("h2", {}, "Flags"), h("button", { class: "icon-btn", title: "Close", onclick: close }, icon(ICONS.close))),
      h("p", { class: "hint" }, "Notes on cells for Claude. Tell Claude “check my flags”; it reads them with ", h("code", {}, "sheets flags"), " and resolves them with a reply."),
      this.list);
    this.render();
  }

  render(): void {
    clearChildren(this.list);
    const wb = this.wb();
    const groups = flagGroups(wb.doc).sort((a, b) => Number(!!a.flag.done) - Number(!!b.flag.done) || b.flag.ts - a.flag.ts);
    for (const g of groups) {
      const sheet = wb.doc.sheets.find((s) => s.id === g.sheet)!;
      const value = wb.display(sheet, g.range.r1, g.range.c1);
      const item = h("li", { class: g.flag.done ? "done" : "" },
        h("button", { class: "flag-item", onclick: () => this.jump(g) },
          h("span", { class: "where" }, g.ref.replace(/^.*!/, wb.doc.sheets.length > 1 ? g.ref.split("!")[0] + "!" : "")),
          value ? h("span", { class: "value" }, value) : null,
          h("span", { class: "note" }, g.flag.note),
          h("span", { class: "meta" }, `${g.flag.by} · ${when(g.flag.ts)}`),
          g.flag.done ? h("span", { class: "reply" }, `✓ ${g.flag.done.by}${g.flag.done.reply ? `: ${g.flag.done.reply}` : ""}`) : null),
        g.flag.done ? h("button", { class: "link dismiss", onclick: () => this.dismiss(g) }, "Dismiss") : null);
      this.list.append(item);
    }
    if (!groups.length) this.list.append(h("li", { class: "empty" }, "No flags. Select cells and press the orange tag in the toolbar to leave Claude a note."));
  }
}
