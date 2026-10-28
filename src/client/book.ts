// The workbook page: toolbar, formula bar, grid, tabs, and every command they trigger.

import { cellName, colName, parseRange, rangeName, type Range } from "../core/a1.ts";
import { parseDelimited, toDelimited } from "../core/csv.ts";
import { describeOp } from "../core/describe.ts";
import { displayValue } from "../core/display.ts";
import { isFormula, moveRefs } from "../core/formula/rewrite.ts";
import { key, newSheet, sheetToJSON, type CondRule, type Format, type Sheet, type ViewState } from "../core/model.ts";
import type { FormatPatch, Op } from "../core/ops.ts";
import { columnValues } from "../core/view.ts";
import { clearChildren, h, icon, ICONS } from "./dom.ts";
import { Grid, type Pos } from "./grid.ts";
import { colorPicker, CondPanel, filterPopover, menu, menuBelow, type MenuItem } from "./popups.ts";
import { Session, type ChangeEvent } from "./session.ts";

const isMac = navigator.platform.startsWith("Mac");
const MOD = isMac ? "⌘" : "Ctrl+";

interface Clip {
  text: string;
  sheet: string;
  rows: number[];
  cols: number[];
  block: Array<Array<{ v: string | null; f: Format | null }>>;
}

export async function openBook(root: HTMLElement, bookId: string): Promise<void> {
  const session = new Session(bookId);
  let meta: { id: string; name: string };
  try {
    meta = await session.load();
  } catch (e) {
    root.replaceChildren(h("main", { class: "missing" },
      h("h1", {}, "Workbook not found"),
      h("p", {}, (e as Error).message),
      h("a", { href: "/", class: "btn primary" }, "Back to all workbooks")));
    return;
  }
  new BookPage(root, session, meta.id);
}

class BookPage {
  private grid: Grid;
  private sheet: Sheet;
  private editor: HTMLTextAreaElement;
  private editing: null | { pos: Pos; quick: boolean; original: string } = null;
  private addr: HTMLInputElement;
  private formula: HTMLInputElement;
  private title: HTMLInputElement;
  private activity: HTMLButtonElement;
  private live: HTMLElement;
  private stats: HTMLElement;
  private tabs: HTMLElement;
  private toast: HTMLElement;
  private tools: Record<string, HTMLButtonElement> = {};
  private clip: Clip | null = null;
  private panel: CondPanel | null = null;
  private panelHost: HTMLElement;
  private lastRemote: { sheet: string; range: Range } | null = null;

  constructor(private root: HTMLElement, private session: Session, private bookId: string) {
    const wb = session.wb;
    this.sheet = this.initialSheet();
    this.grid = new Grid(wb, {
      select: () => this.onSelect(),
      edit: (initial) => this.startEdit(initial, initial === null ? false : true),
      menu: (e, kind) => this.contextMenu(e, kind),
      filterMenu: (col, x, y) => this.openFilter(col, x, y),
      resize: (axis, idx, size) => this.submit([{ t: "size", s: this.sheet.id, axis, sizes: idx.map((i) => [i, size]) }]),
      unhide: (axis, idx) => this.submit([{ t: "hide", s: this.sheet.id, axis, idx, hidden: false }]),
    });

    this.title = h("input", { class: "title", value: wb.doc.name, "aria-label": "Workbook name", spellcheck: "false" });
    this.title.addEventListener("change", () => {
      const name = this.title.value.trim();
      if (name && name !== wb.doc.name) this.submit([{ t: "rename", name }]);
      else this.title.value = wb.doc.name;
    });
    this.title.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === "Escape") {
        if (e.key === "Escape") this.title.value = wb.doc.name;
        this.grid.focusGrid();
      }
    });
    this.activity = h("button", { class: "activity", hidden: true, title: "Show the change" });
    this.activity.addEventListener("click", () => this.jumpToRemote());
    this.live = h("span", { class: "live connecting", role: "status" }, "Connecting");

    this.addr = h("input", { class: "addr", "aria-label": "Cell reference", spellcheck: "false" });
    this.addr.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        this.goTo(this.addr.value);
        e.preventDefault();
      } else if (e.key === "Escape") this.grid.focusGrid();
    });
    this.addr.addEventListener("focus", () => this.addr.select());
    this.formula = h("input", { class: "formula", "aria-label": "Cell contents", spellcheck: "false" });
    this.formula.addEventListener("focus", () => {
      if (!this.editing) this.startEdit(null, false, true);
    });
    this.formula.addEventListener("input", () => {
      if (this.editing) {
        this.editor.value = this.formula.value;
        this.sizeEditor();
      }
    });
    this.formula.addEventListener("keydown", (e) => this.editorKey(e));

    this.editor = h("textarea", { class: "editor", spellcheck: "false", "aria-label": "Edit cell", hidden: true });
    this.editor.addEventListener("input", () => {
      this.formula.value = this.editor.value;
      this.sizeEditor();
    });
    this.editor.addEventListener("keydown", (e) => this.editorKey(e));
    this.editor.addEventListener("blur", (e) => {
      if (e.relatedTarget !== this.formula && this.editing) this.commitEdit(null, false);
    });
    this.grid.root.append(this.editor);
    this.grid.onPaint = () => this.placeEditor();

    this.stats = h("div", { class: "stats", "aria-live": "polite" });
    this.tabs = h("div", { class: "tabs", role: "tablist" });
    this.toast = h("div", { class: "toast", role: "alert", hidden: true });
    this.panelHost = h("div", { class: "panel-host" });

    root.replaceChildren(h("div", { class: "book" },
      h("header", { class: "bar" },
        h("a", { class: "mark", href: "/", title: "All workbooks" }, h("span", { class: "glyph", "aria-hidden": "true" }), "Sheets"),
        this.title, h("div", { class: "spacer" }), this.activity, this.live),
      this.toolbar(),
      h("div", { class: "fbar" }, this.addr, h("span", { class: "fx", "aria-hidden": "true" }, "fx"), this.formula),
      h("div", { class: "work" }, this.grid.root, this.panelHost),
      h("footer", { class: "foot" },
        h("button", { class: "icon-btn add-sheet", title: "Add sheet", onclick: () => this.addSheet() }, icon(ICONS.plus)),
        this.tabs, this.stats),
      this.toast));

    this.bindKeys();
    session.onChange = (e) => this.onChange(e);
    session.onStatus = (s) => {
      this.live.className = `live ${s}`;
      this.live.textContent = { live: "Live", connecting: "Connecting", offline: "Offline", deleted: "Deleted" }[s];
      this.live.title = s === "live" ? "Changes from others appear as they happen" : "";
    };
    session.onError = (msg) => this.showToast(msg);
    session.onReset = () => location.reload();
    session.connect();
    document.title = `${wb.doc.name} · Sheets`;
    this.renderTabs();
    this.grid.setSheet(this.sheet);
    this.onSelect();
    this.grid.focusGrid();
    window.addEventListener("hashchange", () => {
      const s = this.initialSheet();
      if (s !== this.sheet) this.switchSheet(s);
    });
    window.addEventListener("beforeunload", (e) => {
      if (session.hasPending) e.preventDefault();
    });
  }

  private initialSheet(): Sheet {
    const id = decodeURIComponent(location.hash.slice(1));
    const doc = this.session.wb.doc;
    return doc.sheets.find((s) => s.id === id || s.name === id) ?? doc.sheets[0]!;
  }

  // ---- toolbar ---------------------------------------------------------------------

  private toolbar(): HTMLElement {
    const btn = (name: string, title: string, content: Node | string, run: (b: HTMLButtonElement) => void, cls = "") => {
      const b = h("button", { class: `tool ${cls}`, title, "aria-label": title, type: "button" }, content);
      b.addEventListener("mousedown", (e) => e.preventDefault()); // keep grid focus
      b.addEventListener("click", () => run(b));
      this.tools[name] = b;
      return b;
    };
    const sep = () => h("span", { class: "tsep", "aria-hidden": "true" });
    return h("div", { class: "toolbar", role: "toolbar", "aria-label": "Formatting" },
      btn("undo", `Undo (${MOD}Z)`, icon(ICONS.undo), () => this.undo(false)),
      btn("redo", `Redo (${MOD}Shift+Z)`, icon(ICONS.redo), () => this.undo(true)),
      sep(),
      btn("nf", "Number format", h("span", { class: "nf-label" }, "123", icon(ICONS.caret)), (b) => this.numberMenu(b), "wide"),
      btn("pct", "Format as percent", "%", () => this.style({ nf: "0%" }), "glyph-btn"),
      btn("dec-", "Fewer decimal places", h("span", { class: "dec" }, ".0", h("small", {}, "←")), () => this.decimals(-1), "glyph-btn"),
      btn("dec+", "More decimal places", h("span", { class: "dec" }, ".00", h("small", {}, "→")), () => this.decimals(1), "glyph-btn"),
      sep(),
      btn("b", `Bold (${MOD}B)`, h("b", {}, "B"), () => this.toggle("b"), "glyph-btn"),
      btn("i", `Italic (${MOD}I)`, h("i", {}, "I"), () => this.toggle("i"), "glyph-btn"),
      btn("s", "Strikethrough", h("s", {}, "S"), () => this.toggle("s"), "glyph-btn"),
      btn("color", "Text color", h("span", { class: "color-tool" }, icon(ICONS.textColor), h("i", { class: "bar-ink" })), (b) =>
        colorPicker(b, "ink", this.activeFormat()?.color, (c) => this.style({ color: c }))),
      btn("fill", "Fill color", h("span", { class: "color-tool" }, icon(ICONS.fillColor), h("i", { class: "bar-fill" })), (b) =>
        colorPicker(b, "fill", this.activeFormat()?.fill, (c) => this.style({ fill: c }))),
      sep(),
      btn("left", "Align left", icon(ICONS.alignLeft), () => this.style({ align: "left" })),
      btn("center", "Align center", icon(ICONS.alignCenter), () => this.style({ align: "center" })),
      btn("right", "Align right", icon(ICONS.alignRight), () => this.style({ align: "right" })),
      btn("wrap", "Wrap text", icon(ICONS.wrap), () => this.toggle("wrap")),
      btn("rot", "Rotate text up", icon(ICONS.rotate), () => this.style({ rot: this.activeFormat()?.rot ? null : 90 })),
      sep(),
      btn("freeze", "Freeze rows and columns", h("span", { class: "nf-label" }, icon(ICONS.freeze), icon(ICONS.caret)), (b) => this.freezeMenu(b), "wide"),
      btn("filter", "Filter and sort (shared with everyone)", icon(ICONS.filter), () => this.toggleFilter()),
      btn("cond", "Conditional fill", icon(ICONS.cond), () => this.togglePanel()),
    );
  }

  private numberMenu(anchor: HTMLElement): void {
    const cur = this.activeFormat()?.nf;
    const opt = (label: string, nf: string | null, hint: string): MenuItem => ({
      label, hint, checked: (cur ?? null) === nf, run: () => this.style({ nf }),
    });
    menuBelow(anchor, [
      opt("Automatic", null, "1000.5"),
      opt("Number", "#,##0.00", "1,000.50"),
      opt("Whole number", "0", "1001"),
      opt("Percent", "0.0%", "10.5%"),
      opt("Currency", "$#,##0.00", "$1,000.50"),
      opt("Scientific", "0.00E+00", "1.00E+03"),
    ]);
  }

  private freezeMenu(anchor: HTMLElement): void {
    const s = this.sheet;
    const a = this.grid.anchor;
    const f = (rows: number, cols: number) => () => this.submit([{ t: "freeze", s: s.id, rows, cols }]);
    menuBelow(anchor, [
      { label: "No rows", checked: s.frozenRows === 0, run: f(0, s.frozenCols) },
      { label: "1 row", checked: s.frozenRows === 1, run: f(1, s.frozenCols) },
      { label: "2 rows", checked: s.frozenRows === 2, run: f(2, s.frozenCols) },
      { label: `Up to row ${a.r + 1}`, checked: s.frozenRows === a.r + 1 && a.r > 1, run: f(a.r + 1, s.frozenCols) },
      "-",
      { label: "No columns", checked: s.frozenCols === 0, run: f(s.frozenRows, 0) },
      { label: "1 column", checked: s.frozenCols === 1, run: f(s.frozenRows, 1) },
      { label: "2 columns", checked: s.frozenCols === 2, run: f(s.frozenRows, 2) },
      { label: `Up to column ${colName(a.c)}`, checked: s.frozenCols === a.c + 1 && a.c > 1, run: f(s.frozenRows, a.c + 1) },
    ]);
  }

  // ---- selection helpers -----------------------------------------------------------------

  /** Data ranges covering the selection; one per contiguous run of stored rows and columns. */
  private selectionRuns(): Range[] {
    const r = this.grid.rect;
    const runs = (list: number[], a: number, b: number) => {
      const out: Array<[number, number]> = [];
      for (let k = a; k <= b; k++) {
        const x = list[k]!;
        const last = out[out.length - 1];
        if (last && last[1] === x - 1) last[1] = x;
        else out.push([x, x]);
      }
      return out;
    };
    const rr = runs(this.grid.rows, r.i1, r.i2), cr = runs(this.grid.cols, r.j1, r.j2);
    const out: Range[] = [];
    for (const [r1, r2] of rr) for (const [c1, c2] of cr) out.push({ r1, r2, c1, c2 });
    return out;
  }

  /** Every selected data cell, in display order. */
  private selectedCells(): Array<[number, number]> {
    const r = this.grid.rect;
    const out: Array<[number, number]> = [];
    for (let i = r.i1; i <= r.i2; i++) for (let j = r.j1; j <= r.j2; j++) out.push([this.grid.rows[i]!, this.grid.cols[j]!]);
    return out;
  }

  private activeFormat(): Format | undefined {
    const a = this.grid.anchor;
    return this.sheet.cells.get(key(a.r, a.c))?.f;
  }

  private selectionLabel(): string {
    const r = this.grid.rect;
    const a = this.grid.dataOf(r.i1, r.j1), b = this.grid.dataOf(r.i2, r.j2);
    if (r.i1 === 0 && r.i2 === this.grid.rowCount - 1 && !(r.j1 === 0 && r.j2 === this.grid.colCount - 1)) return `${colName(a.c)}:${colName(b.c)}`;
    if (r.j1 === 0 && r.j2 === this.grid.colCount - 1 && !(r.i1 === 0 && r.i2 === this.grid.rowCount - 1)) return `${a.r + 1}:${b.r + 1}`;
    return rangeName({ r1: a.r, c1: a.c, r2: b.r, c2: b.c });
  }

  private onSelect(): void {
    const a = this.grid.anchor;
    const r = this.grid.rect;
    this.addr.value = r.i1 === r.i2 && r.j1 === r.j2 ? cellName(a.r, a.c) : this.selectionLabel();
    if (!this.editing) this.formula.value = this.sheet.cells.get(key(a.r, a.c))?.v ?? "";
    const f = this.activeFormat();
    const v = this.session.wb.value(this.sheet.id, a.r, a.c);
    const align = f?.align ?? (typeof v === "number" ? "right" : "left");
    const on = (name: string, state: boolean) => this.tools[name]?.classList.toggle("on", state);
    on("b", !!f?.b);
    on("i", !!f?.i);
    on("s", !!f?.s);
    on("wrap", !!f?.wrap);
    on("rot", !!f?.rot);
    on("left", align === "left" && !!f?.align);
    on("center", align === "center");
    on("right", align === "right" && !!f?.align);
    on("filter", !!this.sheet.view);
    on("cond", !!this.panel);
    this.tools.undo!.disabled = !this.session.canUndo;
    this.tools.redo!.disabled = !this.session.canRedo;
    this.updateColorBars(f);
    this.updateStats();
  }

  private updateColorBars(f: Format | undefined): void {
    const ink = this.tools.color?.querySelector<HTMLElement>(".bar-ink");
    const fill = this.tools.fill?.querySelector<HTMLElement>(".bar-fill");
    const paint = (el: HTMLElement | null | undefined, c: string | undefined, kind: string) => {
      if (!el) return;
      el.className = el.className.split(" ")[0]! + (c && !c.startsWith("#") ? ` ${kind}-${c}` : "");
      el.style.background = c?.startsWith("#") ? c : "";
    };
    paint(ink, f?.color, "ink");
    paint(fill, f?.fill, "fill");
  }

  private updateStats(): void {
    const cells = this.selectedCells();
    clearChildren(this.stats);
    if (cells.length < 2 || cells.length > 200_000) return;
    let sum = 0, nums = 0, count = 0;
    for (const [r, c] of cells) {
      const v = this.session.wb.value(this.sheet.id, r, c);
      if (v === null || v === "") continue;
      count++;
      if (typeof v === "number") {
        sum += v;
        nums++;
      }
    }
    const stat = (label: string, v: string) => h("span", {}, h("span", { class: "k" }, label), " ", v);
    if (nums) this.stats.append(stat("Sum", displayValue(sum)), stat("Average", displayValue(Math.round((sum / nums) * 1e6) / 1e6)));
    this.stats.append(stat("Count", String(count)));
  }

  // ---- edits ------------------------------------------------------------------------------

  private submit(ops: Op[], opts?: { undoable?: boolean }): boolean {
    try {
      this.session.submit(ops, opts);
      return true;
    } catch (e) {
      this.showToast((e as Error).message);
      return false;
    }
  }

  private style(patch: FormatPatch): void {
    this.submit(this.selectionRuns().map((range) => ({ t: "style", s: this.sheet.id, range, patch }) as Op));
  }

  private toggle(k: "b" | "i" | "s" | "wrap"): void {
    this.style({ [k]: this.activeFormat()?.[k] ? null : true });
  }

  private decimals(delta: number): void {
    const a = this.grid.anchor;
    const cur = this.activeFormat()?.nf;
    const v = this.session.wb.value(this.sheet.id, a.r, a.c);
    let nf = cur;
    if (!nf || nf === "general") {
      const text = typeof v === "number" ? displayValue(v) : "";
      const d = text.includes(".") ? text.split(".")[1]!.length : 0;
      nf = d ? `0.${"0".repeat(d)}` : "0";
    }
    const m = /\.(0+)/.exec(nf);
    const d = Math.max(0, Math.min(10, (m ? m[1]!.length : 0) + delta));
    const zeros = d ? `.${"0".repeat(d)}` : "";
    const next = m ? nf.replace(/\.0+/, zeros) : nf.replace(/0(?![\s\S]*0)/, `0${zeros}`);
    this.style({ nf: next });
  }

  private clearSelection(formats = false): void {
    const cells = this.selectedCells().filter(([r, c]) => this.sheet.cells.get(key(r, c))?.v !== undefined);
    const ops: Op[] = [];
    if (cells.length) ops.push({ t: "set", s: this.sheet.id, cells: cells.map(([r, c]) => [r, c, null]) });
    if (formats) {
      const all = { b: null, i: null, s: null, color: null, fill: null, align: null, nf: null, wrap: null, rot: null };
      for (const range of this.selectionRuns()) ops.push({ t: "style", s: this.sheet.id, range, patch: all });
    }
    this.submit(ops);
  }

  private undo(redo: boolean): void {
    if (this.editing) return;
    const ops = redo ? this.session.redo() : this.session.undo();
    if (!ops) return;
    // Put the selection where the change happened.
    for (const op of ops) {
      if ((op.t === "set" || op.t === "fmt") && op.s === this.sheet.id && op.cells.length) {
        const rs = op.cells.map((c) => c[0]), cs = op.cells.map((c) => c[1]);
        this.grid.select({ r: Math.min(...rs), c: Math.min(...cs) }, { r: Math.max(...rs), c: Math.max(...cs) });
        break;
      }
    }
  }

  // ---- editing ------------------------------------------------------------------------------

  private startEdit(initial: string | null, quick: boolean, fromFormulaBar = false): void {
    const pos = { ...this.grid.anchor };
    this.grid.select(pos, pos);
    const original = this.sheet.cells.get(key(pos.r, pos.c))?.v ?? "";
    this.editing = { pos, quick, original };
    const text = initial ?? original;
    this.editor.value = text;
    this.formula.value = text;
    this.editor.hidden = false;
    this.root.classList.add("is-editing");
    this.placeEditor();
    if (!fromFormulaBar) {
      this.editor.focus();
      this.editor.setSelectionRange(text.length, text.length);
    }
  }

  private placeEditor(): void {
    if (!this.editing) return;
    const { i, j } = this.grid.displayOf(this.editing.pos);
    const box = this.grid.cellBox(i, j);
    const f = this.sheet.cells.get(key(this.editing.pos.r, this.editing.pos.c))?.f;
    const st = this.editor.style;
    st.left = `${box.x}px`;
    st.top = `${box.y}px`;
    st.minWidth = `${box.w + 1}px`;
    st.minHeight = `${Math.min(box.h, 200) + 1}px`;
    this.editor.classList.toggle("b", !!f?.b);
    this.editor.classList.toggle("i", !!f?.i);
    this.sizeEditor();
  }

  private sizeEditor(): void {
    const st = this.editor.style;
    st.width = "0px";
    st.height = "0px";
    st.width = `${Math.min(this.editor.scrollWidth + 4, 640)}px`;
    st.height = `${Math.min(this.editor.scrollHeight, 240)}px`;
  }

  private commitEdit(move: "down" | "up" | "right" | "left" | null, refocus = true): void {
    const ed = this.editing;
    if (!ed) return;
    this.editing = null;
    this.editor.hidden = true;
    this.root.classList.remove("is-editing");
    const value = this.editor.value;
    if (value !== ed.original) {
      this.submit([{ t: "set", s: this.sheet.id, cells: [[ed.pos.r, ed.pos.c, value === "" ? null : value]] }]);
    }
    if (move) this.move(move === "down" ? 1 : move === "up" ? -1 : 0, move === "right" ? 1 : move === "left" ? -1 : 0, false);
    else this.onSelect();
    if (refocus) this.grid.focusGrid();
  }

  private cancelEdit(): void {
    if (!this.editing) return;
    this.editing = null;
    this.editor.hidden = true;
    this.root.classList.remove("is-editing");
    this.onSelect();
    this.grid.focusGrid();
  }

  private editorKey(e: KeyboardEvent): void {
    if (!this.editing) {
      if (e.key === "Escape") this.grid.focusGrid();
      return;
    }
    const k = e.key;
    if (k === "Enter" && !e.altKey && !e.ctrlKey && !e.metaKey) {
      e.preventDefault();
      this.commitEdit(e.shiftKey ? "up" : "down");
    } else if (k === "Enter") {
      // Alt/Ctrl+Enter: newline inside the cell.
      e.preventDefault();
      const t = this.editor;
      const at = t.selectionStart;
      t.setRangeText("\n", at, t.selectionEnd, "end");
      this.formula.value = t.value;
      this.sizeEditor();
    } else if (k === "Tab") {
      e.preventDefault();
      this.commitEdit(e.shiftKey ? "left" : "right");
    } else if (k === "Escape") {
      e.preventDefault();
      this.cancelEdit();
    } else if (this.editing.quick && !isFormula(this.editor.value) && (k === "ArrowUp" || k === "ArrowDown" || k === "ArrowLeft" || k === "ArrowRight")) {
      e.preventDefault();
      this.commitEdit(k === "ArrowUp" ? "up" : k === "ArrowDown" ? "down" : k === "ArrowLeft" ? "left" : "right");
    }
  }

  // ---- navigation ------------------------------------------------------------------------------

  private move(di: number, dj: number, extend: boolean, jump = false): void {
    const g = this.grid;
    const from = extend ? g.displayOf(g.focus) : g.displayOf(g.anchor);
    let i = from.i, j = from.j;
    if (jump) [i, j] = this.jumpTarget(i, j, di, dj);
    else {
      i = Math.max(0, Math.min(g.rowCount - 1, i + di));
      j = Math.max(0, Math.min(g.colCount - 1, j + dj));
    }
    if (extend) g.select(g.anchor, g.dataOf(i, j));
    else g.selectDisplay(i, j);
  }

  /** Ctrl+arrow: to the edge of the current block of data, or the next block. */
  private jumpTarget(i: number, j: number, di: number, dj: number): [number, number] {
    const g = this.grid;
    const filled = (a: number, b: number) => {
      const p = g.dataOf(a, b);
      return this.sheet.cells.get(key(p.r, p.c))?.v !== undefined;
    };
    const inside = (a: number, b: number) => a >= 0 && b >= 0 && a < g.rowCount && b < g.colCount;
    let a = i, b = j;
    if (filled(a, b) && inside(a + di, b + dj) && filled(a + di, b + dj)) {
      while (inside(a + di, b + dj) && filled(a + di, b + dj)) { a += di; b += dj; }
      return [a, b];
    }
    a += di; b += dj;
    while (inside(a, b) && !filled(a, b)) { a += di; b += dj; }
    if (!inside(a, b)) return [Math.max(0, Math.min(g.rowCount - 1, a)), Math.max(0, Math.min(g.colCount - 1, b))];
    return [a, b];
  }

  private goTo(text: string): void {
    const g = parseRange(text.trim());
    if (!g) {
      this.showToast(`"${text}" is not a cell or range`);
      return;
    }
    if (g.sheet) {
      const s = this.session.wb.doc.sheets.find((x) => x.name.toLowerCase() === g.sheet!.toLowerCase());
      if (!s) return this.showToast(`No sheet named ${g.sheet}`);
      this.switchSheet(s);
    }
    this.grid.select({ r: g.r1, c: g.c1 }, { r: Math.min(g.r2, this.sheet.rows - 1), c: Math.min(g.c2, this.sheet.cols - 1) });
    this.grid.focusGrid();
  }

  private bindKeys(): void {
    const target = this.grid.focusTarget;
    target.addEventListener("keydown", (e) => {
      if (this.editing) return;
      const mod = e.ctrlKey || e.metaKey;
      const k = e.key;
      const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (arrows[k]) {
        e.preventDefault();
        this.move(arrows[k][0], arrows[k][1], e.shiftKey, mod);
        return;
      }
      const g = this.grid;
      switch (true) {
        case k === "Enter":
          e.preventDefault();
          if (e.shiftKey) this.move(-1, 0, false);
          else this.startEdit(null, false);
          return;
        case k === "Tab":
          e.preventDefault();
          this.move(0, e.shiftKey ? -1 : 1, false);
          return;
        case k === "F2":
          e.preventDefault();
          this.startEdit(null, false);
          return;
        case k === "Delete" || k === "Backspace":
          e.preventDefault();
          this.clearSelection();
          return;
        case k === "Home": {
          e.preventDefault();
          const { i } = g.displayOf(g.anchor);
          if (mod) g.selectDisplay(0, 0);
          else g.selectDisplay(i, 0);
          return;
        }
        case k === "End": {
          e.preventDefault();
          const used = this.usedDisplayEnd();
          const { i } = g.displayOf(g.anchor);
          if (mod) g.selectDisplay(used.i, used.j);
          else g.selectDisplay(i, used.j);
          return;
        }
        case k === "PageDown" || k === "PageUp":
          e.preventDefault();
          this.move((k === "PageDown" ? 1 : -1) * g.pageRows, 0, e.shiftKey);
          return;
        case k === "Escape":
          g.select(g.anchor, g.anchor, false);
          return;
      }
      if (mod && !e.altKey) {
        const lower = k.toLowerCase();
        const run: Record<string, () => void> = {
          z: () => this.undo(e.shiftKey),
          y: () => this.undo(true),
          b: () => this.toggle("b"),
          i: () => this.toggle("i"),
          a: () => g.selectDisplay(0, 0, g.rowCount - 1, g.colCount - 1, false),
          d: () => this.fill("down"),
          r: () => this.fill("right"),
        };
        if (run[lower]) {
          e.preventDefault();
          run[lower]!();
        }
        return;
      }
      if (k.length === 1 && !e.altKey) {
        e.preventDefault();
        this.startEdit(k, true);
      }
    });
    target.addEventListener("copy", (e) => this.copy(e, false));
    target.addEventListener("cut", (e) => this.copy(e, true));
    target.addEventListener("paste", (e) => this.paste(e));
  }

  private usedDisplayEnd(): { i: number; j: number } {
    let r = 0, c = 0;
    for (const [k, cell] of this.sheet.cells) {
      if (cell.v === undefined) continue;
      r = Math.max(r, Math.floor(k / 16384));
      c = Math.max(c, k % 16384);
    }
    return this.grid.displayOf({ r, c });
  }

  // ---- clipboard ----------------------------------------------------------------------------------

  private copy(e: ClipboardEvent, cut: boolean): void {
    if (this.editing) return;
    e.preventDefault();
    const r = this.grid.rect;
    const rows = this.grid.rows.slice(r.i1, r.i2 + 1), cols = this.grid.cols.slice(r.j1, r.j2 + 1);
    const s = this.sheet;
    const block = rows.map((row) => cols.map((c) => {
      const cell = s.cells.get(key(row, c));
      return { v: cell?.v ?? null, f: cell?.f ?? null };
    }));
    const text = toDelimited(rows.map((row) => cols.map((c) => displayValue(this.session.wb.value(s.id, row, c), s.cells.get(key(row, c))?.f?.nf))), "\t").replace(/\n$/, "");
    const esc = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;");
    const html = `<table>${rows.map((row) => `<tr>${cols.map((c) => `<td>${esc(displayValue(this.session.wb.value(s.id, row, c), s.cells.get(key(row, c))?.f?.nf))}</td>`).join("")}</tr>`).join("")}</table>`;
    e.clipboardData?.setData("text/plain", text);
    e.clipboardData?.setData("text/html", html);
    this.clip = { text, sheet: s.id, rows, cols, block };
    if (cut) this.clearSelection();
  }

  private paste(e: ClipboardEvent): void {
    if (this.editing) return;
    e.preventDefault();
    const text = e.clipboardData?.getData("text/plain") ?? "";
    const g = this.grid;
    const r = g.rect;
    const s = this.sheet;
    const internal = this.clip && this.clip.text === text.replace(/\r?\n$/, "") ? this.clip : null;
    const block: Array<Array<{ v: string | null; f?: Format | null; src?: Pos }>> = internal
      ? internal.block.map((row, i) => row.map((cell, j) => ({ ...cell, src: { r: internal.rows[i]!, c: internal.cols[j]! } })))
      : parseDelimited(text.replace(/\r?\n$/, ""), text.includes("\t") ? "\t" : ",").map((row) => row.map((v) => ({ v: v === "" ? null : v })));
    if (!block.length) return;
    const h1 = block.length, w1 = Math.max(...block.map((x) => x.length));
    // A single copied cell fills the whole selection; otherwise paste at the top-left.
    const fill = h1 === 1 && w1 === 1;
    const rows = fill ? r.i2 - r.i1 + 1 : h1, cols = fill ? r.j2 - r.j1 + 1 : w1;
    const vals: Array<[number, number, string | null]> = [];
    const fmts: Array<[number, number, Format | null]> = [];
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const cell = fill ? block[0]![0]! : block[i]?.[j];
        if (!cell) continue;
        const di = r.i1 + i, dj = r.j1 + j;
        if (di >= g.rowCount || dj >= g.colCount) continue;
        const t = g.dataOf(di, dj);
        let v = cell.v;
        if (v && cell.src && isFormula(v)) v = moveRefs(v, t.r - cell.src.r, t.c - cell.src.c);
        vals.push([t.r, t.c, v]);
        if (cell.f !== undefined) fmts.push([t.r, t.c, cell.f]);
      }
    }
    const ops: Op[] = [{ t: "set", s: s.id, cells: vals }];
    if (fmts.length) ops.push({ t: "fmt", s: s.id, cells: fmts });
    // Grow the grid if the paste runs off the end.
    const maxR = r.i1 + rows - 1, maxC = r.j1 + cols - 1;
    if (maxR >= g.rowCount || maxC >= g.colCount) {
      ops.unshift({ t: "grid", s: s.id, rows: Math.max(s.rows, s.rows + maxR - g.rowCount + 1), cols: Math.max(s.cols, s.cols + maxC - g.colCount + 1) });
      if (this.submit(ops)) this.grid.select(g.dataOf(r.i1, r.j1), g.dataOf(maxR, maxC));
      return;
    }
    if (this.submit(ops)) g.select(g.dataOf(r.i1, r.j1), g.dataOf(r.i1 + rows - 1, r.j1 + cols - 1), false);
  }

  /** Ctrl+D / Ctrl+R: copy the first row (or column) of the selection across it. */
  private fill(dir: "down" | "right"): void {
    const g = this.grid;
    const r = g.rect;
    const s = this.sheet;
    const vals: Array<[number, number, string | null]> = [];
    const fmts: Array<[number, number, Format | null]> = [];
    const i1 = dir === "down" && r.i1 === r.i2 ? Math.max(0, r.i1 - 1) : r.i1;
    const j1 = dir === "right" && r.j1 === r.j2 ? Math.max(0, r.j1 - 1) : r.j1;
    for (let i = i1; i <= r.i2; i++) {
      for (let j = j1; j <= r.j2; j++) {
        if (dir === "down" ? i === i1 : j === j1) continue;
        const src = g.dataOf(dir === "down" ? i1 : i, dir === "right" ? j1 : j);
        const t = g.dataOf(i, j);
        const cell = s.cells.get(key(src.r, src.c));
        const v = cell?.v ?? null;
        vals.push([t.r, t.c, v && isFormula(v) ? moveRefs(v, t.r - src.r, t.c - src.c) : v]);
        fmts.push([t.r, t.c, cell?.f ?? null]);
      }
    }
    if (vals.length) this.submit([{ t: "set", s: s.id, cells: vals }, { t: "fmt", s: s.id, cells: fmts }]);
  }

  // ---- rows / columns ----------------------------------------------------------------------------

  private contextMenu(e: MouseEvent, kind: "cell" | "col" | "row"): void {
    const g = this.grid;
    const r = g.rect;
    const s = this.sheet;
    const rowsSel = g.rows.slice(r.i1, r.i2 + 1), colsSel = g.cols.slice(r.j1, r.j2 + 1);
    const nr = rowsSel.length, nc = colsSel.length;
    const top = rowsSel[0]!, bottom = rowsSel[nr - 1]!, left = colsSel[0]!, right = colsSel[nc - 1]!;
    const plural = (n: number, w: string) => (n === 1 ? `1 ${w}` : `${n} ${w}s`);
    const ins = (axis: "row" | "col", at: number, n: number) => () => this.submit([{ t: "insert", s: s.id, axis, at, n }]);
    const del = (axis: "row" | "col") => () => {
      // Delete each contiguous run, bottom-up so indices stay valid.
      const list = (axis === "row" ? rowsSel : colsSel).slice().sort((a, b) => b - a);
      const ops: Op[] = [];
      let i = 0;
      while (i < list.length) {
        let j = i;
        while (j + 1 < list.length && list[j + 1] === list[j]! - 1) j++;
        ops.push({ t: "delete", s: s.id, axis, at: list[j]!, n: j - i + 1 });
        i = j + 1;
      }
      this.submit(ops);
    };
    const hiddenIn = (axis: "row" | "col") => {
      const set = axis === "row" ? s.hiddenRows : s.hiddenCols;
      const [a, b] = axis === "row" ? [top, bottom] : [left, right];
      return [...set].filter((x) => x >= a && x <= b);
    };
    const items: Array<MenuItem | "-"> = [];
    if (kind === "cell") {
      items.push(
        { label: "Cut", hint: `${MOD}X`, run: () => { g.focusGrid(); document.execCommand("cut"); } },
        { label: "Copy", hint: `${MOD}C`, run: () => { g.focusGrid(); document.execCommand("copy"); } },
        { label: "Paste", hint: `${MOD}V`, disabled: true },
        "-",
      );
    }
    if (kind !== "col") {
      items.push(
        { label: `Insert ${plural(nr, "row")} above`, run: ins("row", top, nr) },
        { label: `Insert ${plural(nr, "row")} below`, run: ins("row", bottom + 1, nr) },
      );
    }
    if (kind !== "row") {
      items.push(
        { label: `Insert ${plural(nc, "column")} left`, run: ins("col", left, nc) },
        { label: `Insert ${plural(nc, "column")} right`, run: ins("col", right + 1, nc) },
      );
    }
    items.push("-");
    if (kind !== "col") items.push({ label: `Delete ${plural(nr, "row")}`, run: del("row") });
    if (kind !== "row") items.push({ label: `Delete ${plural(nc, "column")}`, run: del("col") });
    if (kind === "col") {
      const hid = hiddenIn("col");
      items.push(
        { label: `Hide ${plural(nc, "column")}`, run: () => this.submit([{ t: "hide", s: s.id, axis: "col", idx: colsSel, hidden: true }]) },
        ...(hid.length ? [{ label: `Unhide ${plural(hid.length, "column")}`, run: () => this.submit([{ t: "hide", s: s.id, axis: "col", idx: hid, hidden: false }]) }] : []),
        { label: "Fit to data", run: () => this.submit([{ t: "size", s: s.id, axis: "col", sizes: colsSel.map((c) => [c, g.fitWidth(c)] as [number, number]) }]) },
        "-",
        { label: "Sort A to Z", run: () => this.sortBy(left, false) },
        { label: "Sort Z to A", run: () => this.sortBy(left, true) },
      );
      if (s.view?.sort) items.push({ label: "Clear sort", run: () => this.sortBy(null, false) });
    }
    if (kind === "row") {
      const hid = hiddenIn("row");
      items.push(
        { label: `Hide ${plural(nr, "row")}`, run: () => this.submit([{ t: "hide", s: s.id, axis: "row", idx: rowsSel, hidden: true }]) },
        ...(hid.length ? [{ label: `Unhide ${plural(hid.length, "row")}`, run: () => this.submit([{ t: "hide", s: s.id, axis: "row", idx: hid, hidden: false }]) }] : []),
        { label: "Reset row height", run: () => this.submit([{ t: "size", s: s.id, axis: "row", sizes: rowsSel.map((x) => [x, null] as [number, null]) }]) },
      );
    }
    if (kind === "cell") {
      items.push("-",
        { label: "Clear contents", hint: "Delete", run: () => this.clearSelection() },
        { label: "Clear formatting", run: () => this.clearSelection(true) },
      );
    }
    menu(items, e.clientX, e.clientY);
  }

  // ---- sort & filter ---------------------------------------------------------------------------

  private currentView(): ViewState {
    return structuredClone(this.sheet.view ?? { headerRow: Math.max(0, this.sheet.frozenRows - 1), filters: {} });
  }

  private sortBy(col: number | null, desc: boolean): void {
    const v = this.currentView();
    if (col === null) delete v.sort;
    else v.sort = desc ? { col, desc } : { col };
    this.submit([{ t: "view", s: this.sheet.id, view: v }]);
  }

  private toggleFilter(): void {
    const s = this.sheet;
    if (s.view) return void this.submit([{ t: "view", s: s.id, view: null }]);
    const top = Math.min(this.grid.anchor.r, this.grid.focus.r);
    const headerRow = s.frozenRows > 0 ? s.frozenRows - 1 : top;
    this.submit([{ t: "view", s: s.id, view: { headerRow, filters: {} } }]);
  }

  private openFilter(col: number, x: number, y: number): void {
    const s = this.sheet;
    const view = s.view;
    if (!view) return;
    filterPopover({
      x, y,
      column: `${colName(col)}${s.cells.get(key(view.headerRow, col))?.v ? `: ${this.session.wb.display(s, view.headerRow, col)}` : ""}`,
      values: columnValues(s, col, (r, c) => this.session.wb.value(s.id, r, c)),
      filter: view.filters[col],
      sort: view.sort?.col === col ? (view.sort.desc ? "desc" : "asc") : null,
      onSort: (dir) => this.sortBy(dir ? col : null, dir === "desc"),
      onApply: (f) => {
        const v = this.currentView();
        if (f) v.filters[col] = f;
        else delete v.filters[col];
        this.submit([{ t: "view", s: s.id, view: v }]);
      },
    });
  }

  // ---- conditional fill panel ----------------------------------------------------------------

  private togglePanel(): void {
    if (this.panel) {
      this.panel = null;
      clearChildren(this.panelHost);
      this.onSelect();
      requestAnimationFrame(() => this.grid.render());
      return;
    }
    this.panel = new CondPanel(
      (rules: CondRule[]) => this.submit([{ t: "cond", s: this.sheet.id, rules }]),
      () => this.togglePanel(),
      () => {
        const r = this.selectionRuns()[0]!;
        return r;
      },
    );
    this.panel.setRules(this.sheet.cond);
    this.panelHost.append(this.panel.el);
    this.onSelect();
  }

  // ---- sheets & tabs ---------------------------------------------------------------------------

  private renderTabs(): void {
    clearChildren(this.tabs);
    for (const s of this.session.wb.doc.sheets) {
      const tab = h("button", { class: `tab${s === this.sheet ? " on" : ""}`, role: "tab", "aria-selected": String(s === this.sheet) }, s.name);
      tab.addEventListener("click", () => this.switchSheet(s));
      tab.addEventListener("dblclick", () => this.renameTab(tab, s));
      tab.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        const doc = this.session.wb.doc;
        const i = doc.sheets.indexOf(s);
        menu([
          { label: "Rename", run: () => this.renameTab(tab, s) },
          { label: "Move left", disabled: i === 0, run: () => this.submit([{ t: "moveSheet", s: s.id, index: i - 1 }]) },
          { label: "Move right", disabled: i === doc.sheets.length - 1, run: () => this.submit([{ t: "moveSheet", s: s.id, index: i + 1 }]) },
          "-",
          { label: "Delete", disabled: doc.sheets.length === 1, run: () => {
            if (confirm(`Delete sheet "${s.name}"? You can undo this.`)) this.submit([{ t: "deleteSheet", s: s.id }]);
          } },
        ], e.clientX, e.clientY - 8);
      });
      this.tabs.append(tab);
    }
  }

  private renameTab(tab: HTMLElement, s: Sheet): void {
    const input = h("input", { class: "tab-input", value: s.name, "aria-label": "Sheet name" });
    tab.replaceWith(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (save: boolean) => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      if (save && name && name !== s.name) this.submit([{ t: "renameSheet", s: s.id, name }]);
      this.renderTabs();
      this.grid.focusGrid();
    };
    input.addEventListener("keydown", (e) => {
      if (e.key === "Enter") finish(true);
      if (e.key === "Escape") finish(false);
    });
    input.addEventListener("blur", () => finish(true));
  }

  private addSheet(): void {
    const doc = this.session.wb.doc;
    let n = doc.sheets.length + 1;
    while (doc.sheets.some((s) => s.name.toLowerCase() === `sheet${n}`)) n++;
    const id = Math.random().toString(36).slice(2, 8);
    if (this.submit([{ t: "addSheet", sheet: sheetToJSON(newSheet(id, `Sheet${n}`)) }])) {
      this.switchSheet(doc.sheets.find((s) => s.id === id)!);
    }
  }

  private switchSheet(s: Sheet): void {
    if (this.editing) this.commitEdit(null, false);
    this.sheet = s;
    history.replaceState(null, "", `#${encodeURIComponent(s.id)}`);
    this.grid.setSheet(s);
    this.renderTabs();
    this.panel?.setRules(s.cond);
    this.onSelect();
    this.grid.focusGrid();
  }

  // ---- incoming changes ----------------------------------------------------------------------------

  private onChange(e: ChangeEvent): void {
    const doc = this.session.wb.doc;
    if (!doc.sheets.includes(this.sheet)) {
      // Our sheet was deleted (or replaced by undo of a delete).
      this.sheet = doc.sheets.find((s) => s.id === this.sheet.id) ?? doc.sheets[0]!;
      this.grid.setSheet(this.sheet);
    } else if (e.structural || this.sheet.view) {
      this.grid.layout();
    } else this.grid.refresh();
    if (e.ops.some((op) => op.t === "addSheet" || op.t === "deleteSheet" || op.t === "renameSheet" || op.t === "moveSheet")) this.renderTabs();
    if (e.ops.some((op) => op.t === "rename")) {
      if (document.activeElement !== this.title) this.title.value = doc.name;
      document.title = `${doc.name} · Sheets`;
    }
    if (this.panel && e.ops.some((op) => op.t === "cond" || op.t === "insert" || op.t === "delete")) this.panel.setRules(this.sheet.cond);
    if (e.remote && e.actor) {
      for (const [sid, cells] of e.touched) this.grid.flash(sid, cells);
      this.noteActivity(e);
    }
    this.onSelect();
  }

  private noteActivity(e: ChangeEvent): void {
    const names = (id: string) => this.session.wb.doc.sheets.find((s) => s.id === id)?.name ?? "?";
    const op = e.ops[0];
    if (!op) return;
    let text = describeOp(op, names);
    if (text.length > 70) text = text.slice(0, 68) + "…";
    const more = e.ops.length > 1 ? ` and ${e.ops.length - 1} more` : "";
    const time = new Date(e.ts ?? Date.now()).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    this.activity.hidden = false;
    clearChildren(this.activity);
    this.activity.append(h("b", {}, e.actor ?? "someone"), ` ${text}${more}`, h("time", {}, time));
    const first = [...e.touched][0];
    if (first && first[1].length) {
      const rs = first[1].map((c) => c[0]), cs = first[1].map((c) => c[1]);
      this.lastRemote = { sheet: first[0], range: { r1: Math.min(...rs), c1: Math.min(...cs), r2: Math.max(...rs), c2: Math.max(...cs) } };
    } else this.lastRemote = null;
  }

  private jumpToRemote(): void {
    const t = this.lastRemote;
    if (!t) return;
    const s = this.session.wb.doc.sheets.find((x) => x.id === t.sheet);
    if (!s) return;
    if (s !== this.sheet) this.switchSheet(s);
    this.grid.select({ r: t.range.r1, c: t.range.c1 }, { r: t.range.r2, c: t.range.c2 });
    this.grid.focusGrid();
  }

  private toastTimer = 0;
  private showToast(msg: string): void {
    this.toast.textContent = msg;
    this.toast.hidden = false;
    clearTimeout(this.toastTimer);
    this.toastTimer = window.setTimeout(() => (this.toast.hidden = true), 4000);
  }
}

