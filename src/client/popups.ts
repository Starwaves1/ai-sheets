// Menus, popovers and the conditional-fill panel. One popup open at a time.

import { parseRange, rangeName, type Range } from "../core/a1.ts";
import { COND_OPS, PALETTE, type ColFilter, type CondOp, type CondRule } from "../core/model.ts";
import { clearChildren, h, icon, ICONS } from "./dom.ts";

let open: { el: HTMLElement; close: () => void } | null = null;

export function closePopup(): void {
  open?.close();
}

/** Show an element as a popup at screen coordinates; closes on outside click or Escape. */
export function popup(el: HTMLElement, x: number, y: number, onClose?: () => void): () => void {
  closePopup();
  el.classList.add("popup");
  document.body.append(el);
  const { innerWidth: W, innerHeight: H } = window;
  const r = el.getBoundingClientRect();
  el.style.left = `${Math.max(4, Math.min(x, W - r.width - 4))}px`;
  el.style.top = `${y + r.height > H - 4 ? Math.max(4, y - r.height) : y}px`;
  const outside = (e: MouseEvent) => {
    if (!el.contains(e.target as Node)) close();
  };
  const key = (e: KeyboardEvent) => {
    if (e.key === "Escape") {
      e.stopPropagation();
      close();
    }
  };
  const close = () => {
    if (open?.el !== el) return;
    open = null;
    el.remove();
    document.removeEventListener("mousedown", outside, true);
    document.removeEventListener("keydown", key, true);
    onClose?.();
  };
  setTimeout(() => document.addEventListener("mousedown", outside, true));
  document.addEventListener("keydown", key, true);
  open = { el, close };
  return close;
}

export interface MenuItem {
  label: string;
  hint?: string;
  run?: () => void;
  disabled?: boolean;
  checked?: boolean;
}

export function menu(items: Array<MenuItem | "-">, x: number, y: number): void {
  const el = h("div", { class: "menu", role: "menu" });
  let close = () => {};
  for (const it of items) {
    if (it === "-") {
      el.append(h("div", { class: "sep", role: "separator" }));
      continue;
    }
    const b = h("button", { class: `item${it.checked ? " checked" : ""}`, role: "menuitem", disabled: it.disabled },
      h("span", {}, it.label), it.hint ? h("kbd", {}, it.hint) : null);
    b.addEventListener("click", () => {
      close();
      it.run?.();
    });
    el.append(b);
  }
  close = popup(el, x, y);
  (el.querySelector("button:not([disabled])") as HTMLElement | null)?.focus();
  el.addEventListener("keydown", (e) => {
    const btns = [...el.querySelectorAll<HTMLButtonElement>("button:not([disabled])")];
    const i = btns.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === "ArrowDown") btns[(i + 1) % btns.length]?.focus();
    else if (e.key === "ArrowUp") btns[(i - 1 + btns.length) % btns.length]?.focus();
    else return;
    e.preventDefault();
  });
}

export function menuBelow(anchor: HTMLElement, items: Array<MenuItem | "-">): void {
  const r = anchor.getBoundingClientRect();
  menu(items, r.left, r.bottom + 4);
}

// ---- colors ------------------------------------------------------------------

const COLOR_LABEL: Record<string, string> = { green: "Green", red: "Red", amber: "Amber", blue: "Blue", violet: "Violet", grey: "Grey" };

/** Palette picker: theme-aware named colors, none, or a custom hex. */
export function colorPicker(anchor: HTMLElement, kind: "fill" | "ink", current: string | undefined, pick: (c: string | null) => void): void {
  const el = h("div", { class: "swatches" });
  let close = () => {};
  const choose = (c: string | null) => {
    close();
    pick(c);
  };
  el.append(h("button", { class: `swatch none${current ? "" : " on"}`, title: "None", onclick: () => choose(null) }));
  for (const name of PALETTE) {
    el.append(h("button", {
      class: `swatch ${kind}-${name}${current === name ? " on" : ""}`,
      title: COLOR_LABEL[name]!,
      "aria-label": COLOR_LABEL[name]!,
      onclick: () => choose(name),
    }, kind === "ink" ? "A" : ""));
  }
  const custom = h("input", { type: "color", title: "Custom color", value: current?.startsWith("#") ? current : "#888888" });
  custom.addEventListener("change", () => choose(custom.value));
  el.append(h("label", { class: "custom" }, custom, "Custom"));
  const r = anchor.getBoundingClientRect();
  close = popup(el, r.left, r.bottom + 4);
}

// ---- filter popover ------------------------------------------------------------

export function filterPopover(opts: {
  x: number;
  y: number;
  column: string;
  values: Array<[string, number]>;
  filter: ColFilter | undefined;
  sort: "asc" | "desc" | null;
  onSort: (dir: "asc" | "desc" | null) => void;
  onApply: (f: ColFilter | null) => void;
}): void {
  const hidden = new Set(opts.filter?.hide ?? []);
  const el = h("div", { class: "filter-pop" });
  let close = () => {};
  const sortBtn = (dir: "asc" | "desc", label: string, ic: string) =>
    h("button", { class: `item${opts.sort === dir ? " checked" : ""}`, onclick: () => { close(); opts.onSort(opts.sort === dir ? null : dir); } },
      icon(ic), h("span", {}, label));
  const search = h("input", { type: "search", placeholder: "Search values", "aria-label": "Search values" });
  const list = h("div", { class: "values" });
  const render = () => {
    clearChildren(list);
    const q = search.value.toLowerCase();
    for (const [v, n] of opts.values) {
      if (q && !v.toLowerCase().includes(q)) continue;
      const cb = h("input", { type: "checkbox" });
      cb.checked = !hidden.has(v);
      cb.addEventListener("change", () => (cb.checked ? hidden.delete(v) : hidden.add(v)));
      list.append(h("label", {}, cb, h("span", { class: "v" }, v === "" ? "(blanks)" : v), h("span", { class: "n" }, String(n))));
    }
    if (!list.firstChild) list.append(h("div", { class: "empty" }, "No matching values"));
  };
  search.addEventListener("input", render);
  render();
  const setAll = (on: boolean) => {
    const q = search.value.toLowerCase();
    for (const [v] of opts.values) if (!q || v.toLowerCase().includes(q)) (on ? hidden.delete(v) : hidden.add(v));
    render();
  };
  const cond = opts.filter?.op ? h("p", { class: "note" }, `Also filtered by condition: ${opts.filter.op} ${opts.filter.value ?? ""}`) : null;
  el.append(
    h("div", { class: "title" }, `Column ${opts.column}`),
    sortBtn("asc", "Sort A to Z", ICONS.sortAsc),
    sortBtn("desc", "Sort Z to A", ICONS.sortDesc),
    h("div", { class: "sep" }),
    search,
    h("div", { class: "links" },
      h("button", { class: "link", onclick: () => setAll(true) }, "Select all"),
      h("button", { class: "link", onclick: () => setAll(false) }, "Clear")),
    list,
    ...(cond ? [cond] : []),
    h("div", { class: "actions" },
      h("button", { class: "btn", onclick: () => close() }, "Cancel"),
      h("button", { class: "btn primary", onclick: () => {
        close();
        const f: ColFilter = { ...opts.filter };
        if (hidden.size) f.hide = [...hidden];
        else delete f.hide;
        opts.onApply(f.hide || f.op ? f : null);
      } }, "Apply")),
  );
  close = popup(el, opts.x, opts.y);
  search.focus();
}

// ---- conditional fill panel ------------------------------------------------------

const OP_LABEL: Record<CondOp, string> = {
  eq: "is equal to", ne: "is not equal to", gt: "is greater than", gte: "is at least", lt: "is less than",
  lte: "is at most", between: "is between", contains: "contains", empty: "is empty", notEmpty: "is not empty",
};

export class CondPanel {
  readonly el: HTMLElement;
  private list: HTMLElement;
  private form: HTMLFormElement;
  private editing = -1;
  private fill: string | null = "green";
  rules: CondRule[] = [];

  constructor(private save: (rules: CondRule[]) => void, private close: () => void, private selection: () => Range) {
    this.list = h("ol", { class: "rules" });
    this.form = h("form", { class: "rule-form" });
    this.el = h("aside", { class: "panel", "aria-label": "Conditional fill" },
      h("header", {}, h("h2", {}, "Conditional fill"),
        h("button", { class: "icon-btn", title: "Close", onclick: () => this.close() }, icon(ICONS.close))),
      h("p", { class: "hint" }, "Rules color cells by their value. The first rule that matches a cell wins."),
      this.list, this.form);
    this.renderForm(null);
  }

  setRules(rules: CondRule[]): void {
    this.rules = rules;
    clearChildren(this.list);
    rules.forEach((r, i) => {
      const li = h("li", { class: i === this.editing ? "editing" : "" },
        h("span", { class: `chip fill-${r.style.fill ?? "none"}`, style: r.style.fill?.startsWith("#") ? `background:${r.style.fill}` : null }, "1"),
        h("button", { class: "rule-text", title: "Edit rule", onclick: () => this.renderForm(i) },
          h("b", {}, rangeName(r.range)), ` ${OP_LABEL[r.op]}${r.value !== undefined && r.op !== "empty" && r.op !== "notEmpty" ? ` ${r.value}` : ""}${r.op === "between" ? ` and ${r.value2 ?? ""}` : ""}`),
        h("button", { class: "icon-btn", title: "Delete rule", onclick: () => this.save(rules.filter((_, k) => k !== i)) }, icon(ICONS.close)));
      this.list.append(li);
    });
    if (!rules.length) this.list.append(h("li", { class: "empty" }, "No rules on this sheet yet."));
  }

  private renderForm(index: number | null): void {
    this.editing = index ?? -1;
    const rule = index !== null ? this.rules[index] : undefined;
    this.fill = rule?.style.fill ?? "green";
    const f = this.form;
    clearChildren(f);
    const rangeIn = h("input", { name: "range", value: rangeName(rule?.range ?? this.selection()), spellcheck: "false", "aria-label": "Apply to range" });
    const opSel = h("select", { name: "op", "aria-label": "Condition" }, ...COND_OPS.map((o) => h("option", { value: o, selected: rule?.op === o }, OP_LABEL[o])));
    const v1 = h("input", { name: "value", value: rule?.value ?? "", placeholder: "Value", "aria-label": "Value" });
    const v2 = h("input", { name: "value2", value: rule?.value2 ?? "", placeholder: "and", "aria-label": "Second value" });
    const sync = () => {
      v1.hidden = opSel.value === "empty" || opSel.value === "notEmpty";
      v2.hidden = opSel.value !== "between";
    };
    opSel.addEventListener("change", sync);
    sync();
    const swatches = h("div", { class: "swatch-row", role: "radiogroup", "aria-label": "Fill" });
    const paintSwatches = () => {
      clearChildren(swatches);
      for (const name of PALETTE) {
        swatches.append(h("button", {
          type: "button", class: `swatch fill-${name}${this.fill === name ? " on" : ""}`, title: name,
          role: "radio", "aria-checked": String(this.fill === name),
          onclick: () => { this.fill = name; paintSwatches(); },
        }));
      }
    };
    paintSwatches();
    const err = h("p", { class: "error", hidden: true });
    f.append(
      h("label", {}, h("span", {}, "Apply to"), rangeIn),
      h("label", {}, h("span", {}, "When the cell"), opSel),
      h("div", { class: "pair" }, v1, v2),
      h("div", { class: "label" }, "Fill"), swatches, err,
      h("div", { class: "actions" },
        index !== null ? h("button", { type: "button", class: "btn", onclick: () => { this.renderForm(null); this.setRules(this.rules); } }, "Cancel") : null,
        h("button", { type: "submit", class: "btn primary" }, index !== null ? "Save rule" : "Add rule")),
    );
    f.onsubmit = (e) => {
      e.preventDefault();
      const g = parseRange(rangeIn.value);
      if (!g || g.sheet) {
        err.hidden = false;
        err.textContent = "Use a range on this sheet, like B2:D20 or C:C.";
        return;
      }
      const next: CondRule = { range: { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, op: opSel.value as CondOp, style: { fill: this.fill ?? "green" } };
      if (!v1.hidden) next.value = v1.value;
      if (!v2.hidden) next.value2 = v2.value;
      const rules = [...this.rules];
      if (index !== null) rules[index] = next;
      else rules.push(next);
      this.editing = -1;
      this.save(rules);
      this.renderForm(null);
    };
    this.setRules(this.rules);
  }
}
