// The workbook page: toolbar, formula bar, grid, tabs, and every command they trigger.

import { cellName, colName, parseRange, rangeName, type Range } from "../core/a1.ts";
import { parseDelimited, toDelimited } from "../core/csv.ts";
import { describeOp } from "../core/describe.ts";
import { displayValue } from "../core/display.ts";
import { isFormula, moveRefs } from "../core/formula/rewrite.ts";
import { key, newSheet, sheetToJSON, type CondRule, type Flag, type Format, type Sheet, type ViewState } from "../core/model.ts";
import type { FormatPatch, Op } from "../core/ops.ts";
import { columnValues } from "../core/view.ts";
import { clearChildren, h, icon, ICONS, TAG_ICON } from "./dom.ts";
import { flagGroups, groupAt, type FlagGroup } from "../core/flags.ts";
import { flagCard, FlagsPanel, flagPopover, myName, namePopover } from "./flags.ts";
import { autofillOps } from "../core/fill.ts";
import { Grid, union, type DisplayRect, type Pos } from "./grid.ts";
import { colorPicker, CondPanel, filterPopover, menu, menuBelow, type MenuItem } from "./popups.ts";
import { Session, type ChangeEvent } from "./session.ts";

// Shortcuts use the platform's command key: Cmd on Apple devices, Ctrl elsewhere. Anything not
// listed in bindKeysOn (reload, tabs, zoom, find, print, bookmarks, devtools...) stays the browser's.
const isMac = /Mac|iPhone|iPad|iPod/.test(navigator.platform);
const MOD = isMac ? "⌘" : "Ctrl+";
const REDO_HINT = isMac ? "⌘⇧Z" : "Ctrl+Y";

interface Clip {
  text: string;
  sheet: string;
  rows: number[];
  cols: number[];
  block: Array<Array<{ v: string | null; f: Format | null; flag: Flag | null }>>;
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
  private flagsPanel: FlagsPanel | null = null;
  private me: HTMLButtonElement;
  private card: { el: HTMLElement; pos: Pos; flag: Flag } | null = null;
  private hoverPos: Pos | null = null;
  private panelHost: HTMLElement;
  private lastRemote: { sheet: string; range: Range } | null = null;

  constructor(private root: HTMLElement, private session: Session, private bookId: string) {
    const wb = session.wb;
    this.sheet = this.initialSheet();
    this.grid = new Grid(wb, {
      select: () => this.onSelect(),
      edit: (initial) => this.startEdit(initial, initial === null ? false : true),
      menu: (x, y, kind) => this.contextMenu(x, y, kind),
      filterMenu: (col, x, y) => this.openFilter(col, x, y),
      resize: (axis, idx, size) => this.submit([{ t: "size", s: this.sheet.id, axis, sizes: idx.map((i) => [i, size]) }]),
      unhide: (axis, idx) => this.submit([{ t: "hide", s: this.sheet.id, axis, idx, hidden: false }]),
      pick: (p) => this.pickReference(p),
      move: (axis, lo, hi, gap) => this.moveBlock(axis, lo, hi, gap),
      hoverCell: (p) => {
        this.hoverPos = p;
        this.updateFlagCard();
      },
      fill: (source, target, alt) => this.autofill(source, target, alt),
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
    // Who you are to others. Unset, it reads "Set your name" until you pick one (once per browser).
    this.me = h("button", { class: "me", title: "Your name, shown on your edits and flags" });
    const showMe = () => {
      this.me.textContent = myName() || "Your name";
      this.me.classList.toggle("unset", !myName());
    };
    showMe();
    this.me.addEventListener("click", () => namePopover(this.me, showMe));

    this.addr = h("input", { class: "addr", "aria-label": "Cell reference", spellcheck: "false" });
    this.addr.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        this.goTo(this.addr.value);
        e.preventDefault();
      } else if (e.key === "Escape") this.grid.focusGrid();
    });
    // Select the address on focus; Safari would clear that selection on the mouseup that focused it.
    let justFocused = false;
    this.addr.addEventListener("focus", () => {
      this.addr.select();
      justFocused = true;
    });
    this.addr.addEventListener("blur", () => this.onSelect());
    this.addr.addEventListener("mouseup", (e) => {
      if (justFocused) e.preventDefault();
      justFocused = false;
    });
    this.formula = h("input", { class: "formula", "aria-label": "Cell contents", spellcheck: "false" });
    this.formula.addEventListener("focus", () => {
      if (!this.editing) this.startEdit(null, false, true);
      else this.editing.quick = false; // arrows move the caret here, they never commit
    });
    // Leaving the formula bar (say, by clicking a cell) commits, as leaving the cell editor does.
    this.formula.addEventListener("blur", (e) => {
      if (e.relatedTarget !== this.editor && this.editing) this.commitEdit(null, false);
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
    this.grid.onPaint = () => {
      this.placeEditor();
      this.placeFlagCard();
    };

    this.stats = h("div", { class: "stats", "aria-live": "polite" });
    this.tabs = h("div", { class: "tabs", role: "tablist" });
    this.toast = h("div", { class: "toast", role: "alert", hidden: true });
    this.panelHost = h("div", { class: "panel-host" });

    root.replaceChildren(h("div", { class: "book" },
      h("header", { class: "bar" },
        h("a", { class: "mark", href: "/", title: "All workbooks" }, h("span", { class: "glyph", "aria-hidden": "true" }), "Sheets"),
        this.title, h("div", { class: "spacer" }), this.activity, this.me, this.live),
      this.toolbar(),
      h("div", { class: "fbar" }, this.addr, h("span", { class: "fx", "aria-hidden": "true" }, "fx"), this.formula),
      h("div", { class: "work" }, this.grid.root, this.panelHost),
      h("footer", { class: "foot" },
        h("button", { class: "icon-btn add-sheet", title: "Add sheet", onclick: () => this.addSheet() }, icon(ICONS.plus)),
        this.tabs, this.stats),
      this.toast));

    this.bindKeys();
    this.followKeyboard();
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

  /**
   * Phones: the on-screen keyboard covers the bottom of the page without resizing it.
   * Shrink the app to the visible area so the formula bar and the edited cell stay in view.
   */
  private followKeyboard(): void {
    const vv = window.visualViewport;
    if (!vv) return;
    const book = this.root.firstElementChild as HTMLElement;
    const fit = () => {
      const covered = vv.height < document.documentElement.clientHeight - 40;
      book.style.height = covered ? `${vv.height}px` : "";
      book.style.transform = covered && vv.offsetTop ? `translateY(${vv.offsetTop}px)` : "";
      if (this.editing) {
        this.grid.scrollIntoView(this.editing.pos);
        this.grid.render();
      }
    };
    vv.addEventListener("resize", fit);
    vv.addEventListener("scroll", fit);
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
      btn("redo", `Redo (${REDO_HINT})`, icon(ICONS.redo), () => this.undo(true)),
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
      btn("flag", "Flag for Claude: leave a note on the selected cells", icon(TAG_ICON), (b) => this.flagSelection(b), "tag-tool"),
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
    // Not while someone is typing a reference there (a remote change also lands here).
    if (document.activeElement !== this.addr) this.addr.value = r.i1 === r.i2 && r.j1 === r.j2 ? cellName(a.r, a.c) : this.selectionLabel();
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
    on("flag", !!this.sheet.cells.get(key(a.r, a.c))?.flag);
    this.tools.undo!.disabled = !this.session.canUndo;
    this.tools.redo!.disabled = !this.session.canRedo;
    this.updateColorBars(f);
    this.updateStats();
    this.updateFlagCard();
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

  private clearSelection(formats = false, flags = false): void {
    const cells = this.selectedCells().filter(([r, c]) => this.sheet.cells.get(key(r, c))?.v !== undefined);
    const ops: Op[] = [];
    if (cells.length) ops.push({ t: "set", s: this.sheet.id, cells: cells.map(([r, c]) => [r, c, null]) });
    const flagged = flags ? this.selectedCells().filter(([r, c]) => this.sheet.cells.get(key(r, c))?.flag) : [];
    if (flagged.length) ops.push({ t: "flag", s: this.sheet.id, cells: flagged.map(([r, c]) => [r, c, null]) });
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

  /**
   * Clicking a cell while typing a formula, with the caret right after an operator, "(", "," or "=",
   * inserts that cell's reference (clicking again swaps it) instead of leaving the editor.
   */
  private lastPick: { input: HTMLInputElement | HTMLTextAreaElement; start: number; end: number; text: string } | null = null;
  private pickReference(p: Pos): boolean {
    if (!this.editing) return false;
    const input = document.activeElement === this.formula ? this.formula : document.activeElement === this.editor ? this.editor : null;
    if (!input || !isFormula(input.value)) return false;
    const ref = cellName(p.r, p.c);
    const lp = this.lastPick;
    let start = input.selectionStart ?? input.value.length;
    let end = input.selectionEnd ?? start;
    if (lp && lp.input === input && lp.text === input.value && start === lp.end && end === lp.end) start = lp.start;
    else if (start !== end || !/[=(,+\-*/^&<>:;]\s*$/.test(input.value.slice(0, start))) return false;
    input.setRangeText(ref, start, end, "end");
    this.lastPick = { input, start, end: start + ref.length, text: input.value };
    (input === this.formula ? this.editor : this.formula).value = input.value;
    this.sizeEditor();
    return true;
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
    if (e.isComposing || e.keyCode === 229) return; // Enter or Esc while an IME is composing is the IME's
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
    } else if (k === "Tab" && !e.ctrlKey && !e.metaKey && !e.altKey) {
      e.preventDefault();
      this.commitEdit(e.shiftKey ? "left" : "right");
    } else if (k === "Escape") {
      e.preventDefault();
      this.cancelEdit();
    } else if (this.editing.quick && !isFormula(this.editor.value) && !e.shiftKey && !e.altKey && !e.ctrlKey && !e.metaKey && (k === "ArrowUp" || k === "ArrowDown" || k === "ArrowLeft" || k === "ArrowRight")) {
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
    for (const target of this.grid.focusTargets) this.bindKeysOn(target);
  }

  private bindKeysOn(target: HTMLElement): void {
    target.addEventListener("keydown", (e) => {
      if (this.editing || e.isComposing) return;
      const mod = isMac ? e.metaKey : e.ctrlKey;
      // Browser or system shortcuts: Ctrl on a Mac, the Windows key, Alt with anything but a typed
      // character (Mac Option and Windows AltGr = Ctrl+Alt type characters; plain Alt+letter is a menu key).
      const foreign = (isMac ? e.ctrlKey : e.metaKey) || (e.altKey && (e.key.length !== 1 || (!isMac && !e.ctrlKey)));
      const k = e.key;
      const arrows: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
      if (foreign) return; // Alt+Left is Back, Ctrl+Tab and Ctrl+PageDown switch tabs, Mac Ctrl+arrows switch spaces
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
        case k === "Tab" && !mod:
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
        case (k === "PageDown" || k === "PageUp") && !mod:
          e.preventDefault();
          this.move((k === "PageDown" ? 1 : -1) * g.pageRows, 0, e.shiftKey);
          return;
        case k === "Escape":
          g.select(g.anchor, g.anchor, false);
          return;
      }
      const altGr = !isMac && e.ctrlKey && e.altKey && k.length === 1; // AltGr types characters on Windows
      if (mod && !altGr) {
        // Only these; Cmd/Ctrl+C, X and V arrive as clipboard events.
        const lower = e.shiftKey || e.altKey ? (e.code === "KeyZ" && e.shiftKey && !e.altKey ? "Z" : "") : k.toLowerCase();
        const run: Record<string, () => void> = {
          z: () => this.undo(false),
          Z: () => this.undo(true),
          b: () => this.toggle("b"),
          i: () => this.toggle("i"),
          a: () => g.selectDisplay(0, 0, g.rowCount - 1, g.colCount - 1, false),
        };
        if (!isMac) run.y = () => this.undo(true); // on a Mac, Cmd+Y opens History
        if (run[lower]) {
          e.preventDefault();
          run[lower]!();
        }
        return;
      }
      if (k.length === 1) { // includes AltGr and Option characters
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
      return { v: cell?.v ?? null, f: cell?.f ?? null, flag: cell?.flag ?? null };
    }));
    const text = toDelimited(rows.map((row) => cols.map((c) => displayValue(this.session.wb.value(s.id, row, c), s.cells.get(key(row, c))?.f?.nf))), "\t").replace(/\n$/, "");
    const esc = (t: string) => t.replace(/&/g, "&amp;").replace(/</g, "&lt;");
    const html = `<table>${rows.map((row) => `<tr>${cols.map((c) => `<td>${esc(displayValue(this.session.wb.value(s.id, row, c), s.cells.get(key(row, c))?.f?.nf))}</td>`).join("")}</tr>`).join("")}</table>`;
    e.clipboardData?.setData("text/plain", text);
    e.clipboardData?.setData("text/html", html);
    this.clip = { text, sheet: s.id, rows, cols, block };
    if (cut) this.clearSelection(false, true); // a cut moves flags along with values
  }

  private paste(e: ClipboardEvent): void {
    if (this.editing) return;
    e.preventDefault();
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!text) return;
    const g = this.grid;
    const r = g.rect;
    const s = this.sheet;
    const internal = this.clip && this.clip.text === text.replace(/\r?\n$/, "") ? this.clip : null;
    const block: Array<Array<{ v: string | null; f?: Format | null; flag?: Flag | null; src?: Pos }>> = internal
      ? internal.block.map((row, i) => row.map((cell, j) => ({ ...cell, src: { r: internal.rows[i]!, c: internal.cols[j]! } })))
      : parseDelimited(text.replace(/\r?\n$/, ""), text.includes("\t") ? "\t" : ",").map((row) => row.map((v) => ({ v: v === "" ? null : v })));
    if (!block.length) return;
    const h1 = block.length, w1 = Math.max(...block.map((x) => x.length));
    // A single copied cell fills the whole selection; otherwise paste at the top-left.
    const fill = h1 === 1 && w1 === 1;
    const rows = fill ? r.i2 - r.i1 + 1 : h1, cols = fill ? r.j2 - r.j1 + 1 : w1;
    const vals: Array<[number, number, string | null]> = [];
    const fmts: Array<[number, number, Format | null]> = [];
    const flags: Array<[number, number, Flag | null]> = [];
    let flagsTouched = false;
    for (let i = 0; i < rows; i++) {
      for (let j = 0; j < cols; j++) {
        const cell = fill ? block[0]![0]! : block[i]?.[j];
        if (!cell) continue;
        const di = r.i1 + i, dj = r.j1 + j;
        // Past the last row/column the paste grows the sheet (a set beyond the edge extends it).
        const t = {
          r: di < g.rowCount ? g.rows[di]! : s.rows + di - g.rowCount,
          c: dj < g.colCount ? g.cols[dj]! : s.cols + dj - g.colCount,
        };
        let v = cell.v;
        if (v && cell.src && isFormula(v)) v = moveRefs(v, t.r - cell.src.r, t.c - cell.src.c);
        vals.push([t.r, t.c, v]);
        if (cell.f !== undefined) fmts.push([t.r, t.c, cell.f]);
        if (cell.flag !== undefined) {
          flags.push([t.r, t.c, cell.flag]);
          if (cell.flag || s.cells.get(key(t.r, t.c))?.flag) flagsTouched = true;
        }
      }
    }
    const ops: Op[] = [{ t: "set", s: s.id, cells: vals }];
    if (fmts.length) ops.push({ t: "fmt", s: s.id, cells: fmts });
    if (flagsTouched) ops.push({ t: "flag", s: s.id, cells: flags });
    const end = vals[vals.length - 1];
    if (end && this.submit(ops)) g.select(g.dataOf(r.i1, r.j1), { r: end[0], c: end[1] }, false);
  }

  /** Fill `target` (display positions) from the pattern in `source`; one undoable change. */
  private autofill(source: DisplayRect, target: DisplayRect, alt: boolean): void {
    const g = this.grid;
    const box = (r: DisplayRect) => ({ r1: r.i1, c1: r.j1, r2: r.i2, c2: r.j2 });
    let ops: Op[];
    try {
      ops = autofillOps(this.sheet, box(source), box(target), alt, (i) => g.rows[i]!, (j) => g.cols[j]!);
    } catch (e) {
      return this.showToast((e as Error).message);
    }
    if (this.submit(ops)) {
      const u = union(source, target);
      g.selectDisplay(u.i1, u.j1, u.i2, u.j2, false);
    }
  }

  /**
   * The menu's Autofill: the selection's filled leading rows (or columns) are the pattern,
   * the empty ones after them get filled. Null when the selection has no such split.
   */
  private autofillSplit(): { source: DisplayRect; target: DisplayRect; label: string } | null {
    const g = this.grid, r = g.rect, s = this.sheet;
    const filled = (i: number, j: number) => s.cells.get(key(g.rows[i]!, g.cols[j]!))?.v !== undefined;
    const lastRow = () => {
      for (let i = r.i2; i >= r.i1; i--) for (let j = r.j1; j <= r.j2; j++) if (filled(i, j)) return i;
      return -1;
    };
    const lastCol = () => {
      for (let j = r.j2; j >= r.j1; j--) for (let i = r.i1; i <= r.i2; i++) if (filled(i, j)) return j;
      return -1;
    };
    const i = lastRow();
    if (r.i2 > r.i1 && i >= r.i1 && i < r.i2) return { source: { ...r, i2: i }, target: { ...r, i1: i + 1 }, label: "Autofill down" };
    const j = lastCol();
    if (r.j2 > r.j1 && j >= r.j1 && j < r.j2) return { source: { ...r, j2: j }, target: { ...r, j1: j + 1 }, label: "Autofill right" };
    return null;
  }

  // ---- rows / columns ----------------------------------------------------------------------------

  private contextMenu(x: number, y: number, kind: "cell" | "col" | "row"): void {
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
        { label: "Cut", hint: `${MOD}X`, run: () => { g.focusProxy(); document.execCommand("cut"); g.focusGrid(); } },
        { label: "Copy", hint: `${MOD}C`, run: () => { g.focusProxy(); document.execCommand("copy"); g.focusGrid(); } },
        // Paste needs the keyboard shortcut (browsers block menu paste on plain HTTP); hide it on touch screens.
        ...(matchMedia("(pointer: coarse)").matches ? [] : [{ label: "Paste", hint: `${MOD}V`, disabled: true }]),
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
    // Moving by one step, for touch screens (a mouse can drag the header instead).
    const step = (axis: "row" | "col", dir: -1 | 1) => () => {
      const r2 = g.rect;
      const [lo, hi] = axis === "row" ? [r2.i1, r2.i2] : [r2.j1, r2.j2];
      const n = axis === "row" ? g.rowCount : g.colCount;
      if (dir < 0 ? lo > 0 : hi < n - 1) this.moveBlock(axis, lo, hi, dir < 0 ? lo - 1 : hi + 2);
    };
    if (kind === "row") {
      items.push(
        { label: nr === 1 ? "Move row up" : "Move rows up", disabled: g.rect.i1 === 0, run: step("row", -1) },
        { label: nr === 1 ? "Move row down" : "Move rows down", disabled: g.rect.i2 >= g.rowCount - 1, run: step("row", 1) },
      );
    }
    if (kind === "col") {
      items.push(
        { label: nc === 1 ? "Move column left" : "Move columns left", disabled: g.rect.j1 === 0, run: step("col", -1) },
        { label: nc === 1 ? "Move column right" : "Move columns right", disabled: g.rect.j2 >= g.colCount - 1, run: step("col", 1) },
      );
    }
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
      const split = this.autofillSplit();
      const flagged = !!s.cells.get(key(g.anchor.r, g.anchor.c))?.flag;
      // Anchor the popover on the active cell: on a phone the toolbar's tag may be scrolled away.
      const cellRect = () => {
        const b = g.cellBox(g.displayOf(g.anchor).i, g.displayOf(g.anchor).j), root = g.root.getBoundingClientRect();
        return new DOMRect(root.left + b.x, root.top + b.y, b.w, b.h);
      };
      items.push("-",
        { label: flagged ? "Edit flag…" : "Flag for Claude…", run: () => this.flagSelection({ getBoundingClientRect: cellRect }) },
        ...(split ? [{ label: split.label, run: () => this.autofill(split.source, split.target, false) }] : []),
        { label: "Clear contents", hint: "Delete", run: () => this.clearSelection() },
        { label: "Clear formatting", run: () => this.clearSelection(true) },
      );
    }
    menu(items, x, y);
  }

  /**
   * Move display rows/cols [lo, hi] to the gap before display index `gap` (one undoable change),
   * then keep them selected. Needs a contiguous block of stored rows/cols and no sort on rows.
   */
  private moveBlock(axis: "row" | "col", lo: number, hi: number, gap: number): void {
    const g = this.grid, s = this.sheet;
    const list = axis === "row" ? g.rows : g.cols;
    if (axis === "row" && s.view?.sort) return this.showToast("Rows can't be moved while the sheet is sorted; clear the sort first");
    if (list[hi]! - list[lo]! !== hi - lo) return this.showToast("Only rows or columns next to each other can move together");
    const from = list[lo]!, n = hi - lo + 1;
    const count = axis === "row" ? s.rows : s.cols;
    const to = gap < list.length ? list[gap]! : count;
    if (to >= from && to <= from + n) return;
    if (!this.submit([{ t: "move", s: s.id, axis, from, n, to }])) return;
    const start = to <= from ? to : to - n;
    const a = g.displayOf(axis === "row" ? { r: start, c: 0 } : { r: 0, c: start });
    const b = g.displayOf(axis === "row" ? { r: start + n - 1, c: 0 } : { r: 0, c: start + n - 1 });
    if (axis === "row") g.selectDisplay(a.i, 0, b.i, g.colCount - 1, false);
    else g.selectDisplay(0, a.j, g.rowCount - 1, b.j, false);
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
    if (this.flagsPanel) this.toggleFlagsPanel();
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

  // ---- flags ------------------------------------------------------------------------------------

  /** Toolbar tag: flag the selection (or edit the flag on it) with a note for Claude. */
  private flagSelection(anchor: { getBoundingClientRect(): DOMRect }): void {
    if (this.editing) this.commitEdit(null, false);
    const a = this.grid.anchor;
    const r = this.grid.rect;
    const single = r.i1 === r.i2 && r.j1 === r.j2;
    const existing = this.sheet.cells.get(key(a.r, a.c))?.flag ?? null;
    const group = existing && single ? groupAt(this.session.wb.doc, this.sheet.id, a.r, a.c) : null;
    const cells: Array<[number, number]> = group ? group.cells : this.selectedCells();
    if (cells.length > 10_000) return this.showToast("Flag at most 10,000 cells at once");
    flagPopover({
      anchor,
      where: group ? group.ref.replace(/^.*!/, "") : this.selectionLabel(),
      current: existing,
      openCount: flagGroups(this.session.wb.doc).filter((g) => !g.flag.done).length,
      save: (note) => {
        const flag: Flag = { note, by: myName() || "browser", ts: Date.now() };
        this.submit([{ t: "flag", s: this.sheet.id, cells: cells.map(([r, c]) => [r, c, flag]) }]);
        this.me.textContent = myName() || "Your name";
        this.me.classList.toggle("unset", !myName());
        this.grid.focusGrid();
      },
      remove: () => this.removeFlag(this.sheet.id, a),
      showAll: () => {
        if (!this.flagsPanel) this.toggleFlagsPanel();
      },
    });
  }

  /** Remove (or, for a resolved one, dismiss) the flag at p and every cell flagged with it. */
  private removeFlag(sheet: string, p: Pos): void {
    const g = groupAt(this.session.wb.doc, sheet, p.r, p.c);
    if (g) this.submit([{ t: "flag", s: sheet, cells: g.cells.map(([r, c]) => [r, c, null]) }]);
  }

  /** The note on the hovered cell, or else on the selected cell, shown beside it. */
  private updateFlagCard(): void {
    const p = this.hoverPos ?? this.grid.anchor;
    const flag = this.editing ? undefined : this.sheet.cells.get(key(p.r, p.c))?.flag;
    if (!flag) {
      this.card?.el.remove();
      this.card = null;
      return;
    }
    if (!this.card || this.card.flag !== flag || this.card.pos.r !== p.r || this.card.pos.c !== p.c) {
      this.card?.el.remove();
      const sheet = this.sheet.id;
      const el = flagCard(flag, {
        edit: () => {
          this.grid.select(p, p, false);
          this.flagSelection(this.tools.flag!);
        },
        remove: () => this.removeFlag(sheet, p),
      });
      this.grid.root.append(el);
      this.card = { el, pos: { ...p }, flag };
    }
    this.placeFlagCard();
  }

  private placeFlagCard(): void {
    if (!this.card) return;
    const { i, j } = this.grid.displayOf(this.card.pos);
    const box = this.grid.cellBox(i, j);
    const el = this.card.el;
    // Beside the cell if there is room, else below it (above near the bottom): never over it.
    const W = this.grid.root.clientWidth, H = this.grid.root.clientHeight;
    const w = el.offsetWidth || 260, hgt = el.offsetHeight || 80;
    let x = box.x + box.w + 6, y = box.y;
    if (x + w > W - 4) x = box.x - w - 6;
    if (x < 4) {
      x = Math.max(4, Math.min(box.x, W - w - 4));
      y = box.y + box.h + hgt + 6 <= H ? box.y + box.h + 6 : box.y - hgt - 6;
    }
    el.style.left = `${x}px`;
    el.style.top = `${Math.max(4, y)}px`;
  }

  private toggleFlagsPanel(): void {
    if (this.flagsPanel) {
      this.flagsPanel = null;
      clearChildren(this.panelHost);
      requestAnimationFrame(() => this.grid.render());
      return;
    }
    if (this.panel) this.togglePanel();
    this.flagsPanel = new FlagsPanel(
      () => this.session.wb,
      (g) => this.jumpToFlag(g),
      (g) => this.submit([{ t: "flag", s: g.sheet, cells: g.cells.map(([r, c]) => [r, c, null]) }]),
      () => this.toggleFlagsPanel(),
    );
    this.panelHost.append(this.flagsPanel.el);
  }

  private jumpToFlag(g: FlagGroup): void {
    const s = this.session.wb.doc.sheets.find((x) => x.id === g.sheet);
    if (!s) return;
    if (s !== this.sheet) this.switchSheet(s);
    this.grid.select({ r: g.range.r1, c: g.range.c1 }, { r: g.range.r2, c: g.range.c2 });
    this.grid.focusGrid();
  }

  // ---- sheets & tabs ---------------------------------------------------------------------------

  private renderTabs(): void {
    clearChildren(this.tabs);
    for (const s of this.session.wb.doc.sheets) {
      const tab = h("button", { class: `tab${s === this.sheet ? " on" : ""}`, role: "tab", "aria-selected": String(s === this.sheet) }, s.name);
      tab.addEventListener("click", () => this.switchSheet(s));
      tab.addEventListener("dblclick", () => this.renameTab(tab, s));
      tab.addEventListener("contextmenu", (e) => {
        if (e.shiftKey) return; // Shift+right-click: the browser's own menu
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
    this.flagsPanel?.render();
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

