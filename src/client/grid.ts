// Virtualized grid. Only cells inside the viewport exist in the DOM. A native
// scroller provides scrollbars and wheel/touch input; the painted panes sit on top
// with pointer-events off, and every mouse interaction is hit-tested by arithmetic.
// Nothing runs unless the user scrolls or data changes: no timers, no animation loops.

import { colName } from "../core/a1.ts";
import { displayValue, effectiveFormat } from "../core/display.ts";
import { Err } from "../core/formula/values.ts";
import { DEFAULT_COL_W, DEFAULT_ROW_H, key, type Format, type Sheet } from "../core/model.ts";
import { visibleCols, visibleRows } from "../core/view.ts";
import type { Workbook } from "../core/workbook.ts";
import { h } from "./dom.ts";

export interface Pos {
  r: number;
  c: number;
}

/** Selection rectangle in display positions (indices into the visible row/col lists). */
export interface DisplayRect {
  i1: number;
  j1: number;
  i2: number;
  j2: number;
}

export interface GridEvents {
  select(): void;
  edit(initial: string | null): void;
  menu(x: number, y: number, kind: "cell" | "col" | "row"): void;
  filterMenu(col: number, x: number, y: number): void;
  resize(axis: "row" | "col", indices: number[], size: number | null): void;
  unhide(axis: "row" | "col", indices: number[]): void;
  /** A click on a cell while a formula is being typed: return true if it was taken as a reference. */
  pick(p: Pos): boolean;
}

interface Pane {
  clip: HTMLElement;
  content: HTMLElement;
  cells: Map<number, HTMLElement & { _sig?: string }>;
  range: HTMLElement;
  active: HTMLElement;
  mx: boolean; // scrolls horizontally
  my: boolean; // scrolls vertically
}

const HEAD_W = 46;
const HEAD_H = 24;
const FLASH_MS = 1800;

const PALETTE_FILL = new Set(["green", "red", "amber", "blue", "violet", "grey"]);

function lowerBound(arr: Float64Array, n: number, x: number): number {
  // Largest index i in [0, n) with arr[i] <= x.
  let lo = 0, hi = n - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (arr[mid]! <= x) lo = mid;
    else hi = mid - 1;
  }
  return lo;
}

export class Grid {
  readonly root: HTMLElement;
  private scroller: HTMLElement;
  private sizer: HTMLElement;
  private panes: Pane[];
  private colHeads: Array<{ clip: HTMLElement; content: HTMLElement; mx: boolean }>;
  private rowHeads: Array<{ clip: HTMLElement; content: HTMLElement; my: boolean }>;
  private frozenLineV: HTMLElement;
  private frozenLineH: HTMLElement;
  private guide: HTMLElement;
  private handle: HTMLElement; // touch: drag to extend the selection
  // Keyboard focus sits in this hidden textarea with a one-space selection, not on the
  // scroller: Safari only fires copy/cut/paste on an editable element with a selection.
  private proxy: HTMLTextAreaElement;

  sheet!: Sheet;
  rows: number[] = [];
  cols: number[] = [];
  private rowPos = new Int32Array(0);
  private colPos = new Int32Array(0);
  private rowTop = new Float64Array(1);
  private colLeft = new Float64Array(1);
  private fr = 0; // frozen display rows
  private fc = 0;

  anchor: Pos = { r: 0, c: 0 };
  focus: Pos = { r: 0, c: 0 };
  private flashes = new Map<number, number>(); // cell key -> start time
  private flashSheet = "";
  private raf = 0;
  private dragging: null | { kind: "cells" | "rows" | "cols" } = null;

  constructor(private wb: Workbook, private ev: GridEvents) {
    this.sizer = h("div", { class: "sizer" });
    this.scroller = h("div", { class: "scroller", tabindex: "0", "aria-label": "Spreadsheet grid" }, this.sizer);
    const mkPane = (mx: boolean, my: boolean): Pane => {
      const range = h("div", { class: "sel-range" });
      const active = h("div", { class: "sel-active" });
      // Selection boxes live in content coordinates, so they scroll with the cells.
      const content = h("div", { class: "content" }, range, active);
      const clip = h("div", { class: "pane" }, content);
      return { clip, content, cells: new Map(), range, active, mx, my };
    };
    this.panes = [mkPane(true, true), mkPane(true, false), mkPane(false, true), mkPane(false, false)];
    const mkHead = (cls: string) => {
      const content = h("div", { class: "content" });
      return { clip: h("div", { class: `pane ${cls}` }, content), content };
    };
    this.colHeads = [{ ...mkHead("colhead"), mx: true }, { ...mkHead("colhead"), mx: false }];
    this.rowHeads = [{ ...mkHead("rowhead"), my: true }, { ...mkHead("rowhead"), my: false }];
    const corner = h("div", { class: "corner-head", title: "Select all" });
    this.frozenLineV = h("div", { class: "frozen-line v" });
    this.frozenLineH = h("div", { class: "frozen-line h" });
    this.guide = h("div", { class: "resize-guide", hidden: true });
    this.handle = h("div", { class: "sel-handle", "aria-hidden": "true" });
    const layer = h("div", { class: "layer" },
      ...this.panes.map((p) => p.clip), ...this.colHeads.map((x) => x.clip), ...this.rowHeads.map((x) => x.clip),
      corner, this.frozenLineV, this.frozenLineH, this.guide);
    this.proxy = h("textarea", {
      class: "kbd-proxy", "aria-label": "Spreadsheet grid", inputmode: "none", tabindex: "-1",
      autocomplete: "off", autocorrect: "off", autocapitalize: "off", spellcheck: "false",
    });
    this.proxy.value = " ";
    this.root = h("div", { class: "grid" }, this.scroller, layer, this.handle, this.proxy);

    this.scroller.addEventListener("scroll", () => this.render(), { passive: true });
    this.scroller.addEventListener("focus", () => {
      if (!matchMedia("(pointer: coarse)").matches) this.focusProxy();
    });
    this.scroller.addEventListener("mousedown", (e) => this.mouseDown(e));
    this.scroller.addEventListener("dblclick", (e) => this.dblClick(e));
    this.scroller.addEventListener("contextmenu", (e) => this.contextMenu(e));
    this.scroller.addEventListener("mousemove", (e) => this.hover(e));
    // Frozen-pane limits depend on the viewport, so a resize re-lays out.
    new ResizeObserver(() => this.sheet && this.layout()).observe(this.scroller);
    this.bindTouch();
  }

  /**
   * Give the grid keyboard focus. With a mouse that is the clipboard proxy; on touch screens
   * it is the scroller, so a tap never brings up the on-screen keyboard.
   */
  focusGrid(): void {
    if (matchMedia("(pointer: coarse)").matches) this.scroller.focus({ preventScroll: true });
    else this.focusProxy();
  }

  /** Focus the proxy and select its text, so the next copy/cut/paste fires in every browser. */
  focusProxy(): void {
    this.proxy.focus({ preventScroll: true });
    this.proxy.value = " ";
    this.proxy.select();
  }

  /** Elements that receive grid keystrokes and clipboard events. */
  get focusTargets(): HTMLElement[] {
    return [this.proxy, this.scroller];
  }

  // ---- layout ---------------------------------------------------------------

  setSheet(sheet: Sheet): void {
    if (this.sheet !== sheet) {
      this.sheet = sheet;
      this.anchor = { r: 0, c: 0 };
      this.focus = { r: 0, c: 0 };
      this.flashes.clear();
      this.scroller.scrollTo(0, 0);
      for (const p of this.panes) {
        p.cells.forEach((el) => el.remove());
        p.cells.clear();
      }
    }
    this.layout();
  }

  /** Recompute visible rows/cols and pixel offsets; call after structural changes. */
  layout(): void {
    const s = this.sheet;
    const valueOf = (r: number, c: number) => this.wb.value(s.id, r, c);
    this.rows = visibleRows(s, valueOf);
    this.cols = visibleCols(s);
    this.rowPos = new Int32Array(s.rows).fill(-1);
    this.colPos = new Int32Array(s.cols).fill(-1);
    this.rowTop = new Float64Array(this.rows.length + 1);
    this.colLeft = new Float64Array(this.cols.length + 1);
    this.rows.forEach((r, i) => {
      this.rowPos[r] = i;
      this.rowTop[i + 1] = this.rowTop[i]! + (s.rowH.get(r) ?? DEFAULT_ROW_H);
    });
    this.cols.forEach((c, j) => {
      this.colPos[c] = j;
      this.colLeft[j + 1] = this.colLeft[j]! + (s.colW.get(c) ?? DEFAULT_COL_W);
    });
    this.fr = 0;
    while (this.fr < this.rows.length && this.rows[this.fr]! < s.frozenRows) this.fr++;
    this.fc = 0;
    while (this.fc < this.cols.length && this.cols[this.fc]! < s.frozenCols) this.fc++;
    // Frozen panes never take more than a share of the screen (a phone would otherwise show
    // nothing but the frozen name column): shrink frozen sizes proportionally to fit.
    const W = this.scroller.clientWidth - HEAD_W, H = this.scroller.clientHeight - HEAD_H;
    const small = W < 700;
    fitPrefix(this.colLeft, this.fc, Math.max(60, W * (small ? 0.42 : 0.5)));
    fitPrefix(this.rowTop, this.fr, Math.max(48, H * (small ? 0.25 : 0.4)));
    this.sizer.style.width = `${HEAD_W + this.colLeft[this.cols.length]! + 120}px`;
    this.sizer.style.height = `${HEAD_H + this.rowTop[this.rows.length]! + 200}px`;
    for (const p of this.panes) {
      p.cells.forEach((el) => el.remove());
      p.cells.clear();
    }
    for (const x of [...this.colHeads, ...this.rowHeads]) x.content.textContent = "";
    this.render();
  }

  // ---- coordinates --------------------------------------------------------------

  /** Display position of a data cell, clamped to the nearest visible row/col. */
  displayOf(p: Pos): { i: number; j: number } {
    const near = (pos: Int32Array, list: number[], x: number) => {
      if (x < pos.length && pos[x]! >= 0) return pos[x]!;
      let best = 0;
      for (let k = 0; k < list.length; k++) if (list[k]! <= x) best = k;
      return best;
    };
    return { i: near(this.rowPos, this.rows, p.r), j: near(this.colPos, this.cols, p.c) };
  }

  dataOf(i: number, j: number): Pos {
    const ci = Math.max(0, Math.min(i, this.rows.length - 1));
    const cj = Math.max(0, Math.min(j, this.cols.length - 1));
    return { r: this.rows[ci] ?? 0, c: this.cols[cj] ?? 0 };
  }

  get rect(): DisplayRect {
    const a = this.displayOf(this.anchor), b = this.displayOf(this.focus);
    return { i1: Math.min(a.i, b.i), j1: Math.min(a.j, b.j), i2: Math.max(a.i, b.i), j2: Math.max(a.j, b.j) };
  }

  get rowCount(): number {
    return this.rows.length;
  }

  get colCount(): number {
    return this.cols.length;
  }

  /** Screen box (relative to grid root) of a display cell. */
  cellBox(i: number, j: number): { x: number; y: number; w: number; h: number } {
    const sl = this.scroller.scrollLeft, st = this.scroller.scrollTop;
    const X = this.colLeft[j]!, Y = this.rowTop[i]!;
    return {
      x: HEAD_W + (j < this.fc ? X : X - sl),
      y: HEAD_H + (i < this.fr ? Y : Y - st),
      w: this.colLeft[j + 1]! - X,
      h: this.rowTop[i + 1]! - Y,
    };
  }

  private hit(e: { clientX: number; clientY: number }): { zone: "corner" | "col" | "row" | "cell"; i: number; j: number; x: number; y: number; sx: number; sy: number } {
    const box = this.root.getBoundingClientRect();
    const x = e.clientX - box.left, y = e.clientY - box.top;
    const fw = this.colLeft[this.fc]!, fh = this.rowTop[this.fr]!;
    const sx = x - HEAD_W < fw ? x - HEAD_W : x - HEAD_W + this.scroller.scrollLeft;
    const sy = y - HEAD_H < fh ? y - HEAD_H : y - HEAD_H + this.scroller.scrollTop;
    const j = lowerBound(this.colLeft, this.cols.length, Math.max(0, sx));
    const i = lowerBound(this.rowTop, this.rows.length, Math.max(0, sy));
    const zone = x < HEAD_W && y < HEAD_H ? "corner" : y < HEAD_H ? "col" : x < HEAD_W ? "row" : "cell";
    return { zone, i, j, x, y, sx, sy };
  }

  /** Column/row boundary under the pointer in a header (for resizing), or -1. */
  private edgeAt(zone: string, sx: number, sy: number): number {
    const tol = 4;
    if (zone === "col") {
      const j = lowerBound(this.colLeft, this.cols.length + 1, sx + tol);
      if (j > 0 && Math.abs(this.colLeft[j]! - sx) <= tol) return j - 1;
    } else if (zone === "row") {
      const i = lowerBound(this.rowTop, this.rows.length + 1, sy + tol);
      if (i > 0 && Math.abs(this.rowTop[i]! - sy) <= tol) return i - 1;
    }
    return -1;
  }

  // ---- painting -----------------------------------------------------------------

  /** Re-render now (coalesced to one frame when called repeatedly). */
  render(): void {
    if (this.raf) return;
    this.raf = requestAnimationFrame(() => {
      this.raf = 0;
      this.paint();
    });
  }

  /** Repaint cell contents after a data change (layout unchanged). */
  refresh(): void {
    // Writing past the edge grows the sheet; that needs a new layout.
    if (this.sheet.rows !== this.rowPos.length || this.sheet.cols !== this.colPos.length) this.layout();
    else this.render();
  }

  flash(sheetId: string, cells: Array<[number, number]>): void {
    if (sheetId !== this.sheet.id) return;
    const now = performance.now();
    if (this.flashSheet !== sheetId) this.flashes.clear();
    this.flashSheet = sheetId;
    for (const [k, t] of this.flashes) if (now - t > FLASH_MS) this.flashes.delete(k);
    for (const [r, c] of cells.slice(0, 5000)) this.flashes.set(key(r, c), now);
    // Force those cells to restart their fade on the next paint.
    for (const p of this.panes) for (const [r, c] of cells.slice(0, 5000)) {
      const el = p.cells.get(key(r, c));
      if (el) el._sig = "";
    }
    this.render();
  }

  private paint(): void {
    if (!this.sheet) return;
    const W = this.scroller.clientWidth, H = this.scroller.clientHeight;
    const sl = this.scroller.scrollLeft, st = this.scroller.scrollTop;
    const fw = this.colLeft[this.fc]!, fh = this.rowTop[this.fr]!;
    const mainW = Math.max(0, W - HEAD_W - fw), mainH = Math.max(0, H - HEAD_H - fh);
    const now = performance.now();

    for (const p of this.panes) {
      const left = HEAD_W + (p.mx ? fw : 0), top = HEAD_H + (p.my ? fh : 0);
      const width = p.mx ? mainW : Math.min(fw, W - HEAD_W), height = p.my ? mainH : Math.min(fh, H - HEAD_H);
      const st0 = p.clip.style;
      st0.left = `${left}px`; st0.top = `${top}px`; st0.width = `${width}px`; st0.height = `${height}px`;
      p.clip.hidden = width <= 0 || height <= 0;
      const ox = p.mx ? fw + sl : 0, oy = p.my ? fh + st : 0;
      p.content.style.transform = `translate(${-ox}px, ${-oy}px)`;
      if (p.clip.hidden) {
        p.cells.forEach((el) => el.remove());
        p.cells.clear();
        continue;
      }
      const j0 = p.mx ? Math.max(this.fc, lowerBound(this.colLeft, this.cols.length, ox)) : 0;
      const j1 = p.mx ? lowerBound(this.colLeft, this.cols.length, ox + width) : this.fc - 1;
      const i0 = p.my ? Math.max(this.fr, lowerBound(this.rowTop, this.rows.length, oy)) : 0;
      const i1 = p.my ? lowerBound(this.rowTop, this.rows.length, oy + height) : this.fr - 1;
      this.paintCells(p, i0, i1, j0, j1, now);
      this.paintSelection(p);
    }

    // Headers
    for (const hd of this.colHeads) {
      const left = HEAD_W + (hd.mx ? fw : 0), width = hd.mx ? mainW : Math.min(fw, W - HEAD_W);
      Object.assign(hd.clip.style, { left: `${left}px`, top: "0px", width: `${width}px`, height: `${HEAD_H}px` });
      hd.clip.hidden = width <= 0;
      const ox = hd.mx ? fw + sl : 0;
      hd.content.style.transform = `translate(${-ox}px, 0)`;
      const j0 = hd.mx ? Math.max(this.fc, lowerBound(this.colLeft, this.cols.length, ox)) : 0;
      const j1 = hd.mx ? lowerBound(this.colLeft, this.cols.length, ox + width) : this.fc - 1;
      this.paintColHeads(hd.content, j0, j1);
    }
    for (const hd of this.rowHeads) {
      const top = HEAD_H + (hd.my ? fh : 0), height = hd.my ? mainH : Math.min(fh, H - HEAD_H);
      Object.assign(hd.clip.style, { left: "0px", top: `${top}px`, width: `${HEAD_W}px`, height: `${height}px` });
      hd.clip.hidden = height <= 0;
      const oy = hd.my ? fh + st : 0;
      hd.content.style.transform = `translate(0, ${-oy}px)`;
      const i0 = hd.my ? Math.max(this.fr, lowerBound(this.rowTop, this.rows.length, oy)) : 0;
      const i1 = hd.my ? lowerBound(this.rowTop, this.rows.length, oy + height) : this.fr - 1;
      this.paintRowHeads(hd.content, i0, i1);
    }
    this.frozenLineV.hidden = this.fc === 0;
    this.frozenLineV.style.left = `${HEAD_W + fw - 1}px`;
    this.frozenLineH.hidden = this.fr === 0;
    this.frozenLineH.style.top = `${HEAD_H + fh - 1}px`;
    this.placeHandle(W, H);
    this.onPaint?.();
  }

  /** Hook for overlays (the cell editor) that must follow scrolling. */
  onPaint?: () => void;

  private paintCells(p: Pane, i0: number, i1: number, j0: number, j1: number, now: number): void {
    const s = this.sheet;
    const seen = new Set<number>();
    const view = s.view;
    for (let i = i0; i <= i1; i++) {
      const r = this.rows[i]!;
      const y = this.rowTop[i]!, hgt = this.rowTop[i + 1]! - y;
      for (let j = j0; j <= j1; j++) {
        const c = this.cols[j]!;
        const k = key(r, c);
        seen.add(k);
        const v = this.wb.value(s.id, r, c);
        const f = effectiveFormat(s, r, c, v);
        const text = displayValue(v, f?.nf);
        const x = this.colLeft[j]!;
        let w = this.colLeft[j + 1]! - x;
        // Left-aligned text spills into empty neighbours, as in other spreadsheets.
        let spill = false;
        const align = f?.align ?? (typeof v === "number" ? "right" : typeof v === "boolean" || v instanceof Err ? "center" : "left");
        if (text && align === "left" && !f?.wrap && !f?.fill && !f?.rot && text.length * 6.5 > w - 8) {
          for (let jj = j + 1; jj < this.cols.length && jj <= j + 12; jj++) {
            const c2 = this.cols[jj]!;
            // Stop at any cell with a value or its own formatting: formatted cells are "taken".
            if (s.cells.has(key(r, c2))) break;
            w += this.colLeft[jj + 1]! - this.colLeft[jj]!;
            spill = true;
            if (text.length * 7 < w) break;
          }
        }
        const filterBtn = view && r === view.headerRow && (view.filters[c] !== undefined || view.sort?.col === c || this.inUsedCols(c));
        const flashT = this.flashes.get(k);
        const flashing = flashT !== undefined && now - flashT < FLASH_MS;
        const cls = cellClass(f, align, v instanceof Err, spill, flashing, filterBtn ? (view!.filters[c] || view!.sort?.col === c ? "filter on" : "filter") : "");
        const style = cellStyle(f);
        const sig = `${text}\u0001${cls}\u0001${style}\u0001${x},${y},${w},${hgt}`;
        let el = p.cells.get(k);
        if (!el) {
          el = h("div") as HTMLElement & { _sig?: string };
          p.cells.set(k, el);
          p.content.insertBefore(el, p.range);
        }
        if (el._sig === sig) continue;
        el._sig = sig;
        el.className = cls;
        el.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${hgt}px;${style}${flashing ? `--flash-delay:${-(now - flashT!)}ms;` : ""}`;
        if (f?.rot) {
          el.textContent = "";
          el.append(h("span", { class: "rot" }, text));
        } else el.textContent = text;
        if (v instanceof Err && v.msg) el.title = `${v.code} ${v.msg}`;
        else if (el.title) el.removeAttribute("title");
      }
    }
    for (const [k, el] of p.cells) {
      if (!seen.has(k)) {
        el.remove();
        p.cells.delete(k);
      }
    }
  }

  private usedColsCache: { sheet: Sheet; size: number; max: number } | null = null;
  private inUsedCols(c: number): boolean {
    const s = this.sheet;
    if (!this.usedColsCache || this.usedColsCache.sheet !== s || this.usedColsCache.size !== s.cells.size) {
      let max = -1;
      for (const k of s.cells.keys()) max = Math.max(max, k % 16384);
      this.usedColsCache = { sheet: s, size: s.cells.size, max };
    }
    return c <= this.usedColsCache.max;
  }

  private paintSelection(p: Pane): void {
    const r = this.rect;
    const x = this.colLeft[r.j1]!, y = this.rowTop[r.i1]!;
    const w = this.colLeft[r.j2 + 1]! - x, hh = this.rowTop[r.i2 + 1]! - y;
    const multi = r.i1 !== r.i2 || r.j1 !== r.j2;
    p.range.hidden = !multi;
    if (multi) p.range.style.cssText = `left:${x}px;top:${y}px;width:${w}px;height:${hh}px`;
    const a = this.displayOf(this.anchor);
    const ax = this.colLeft[a.j]!, ay = this.rowTop[a.i]!;
    p.active.style.cssText = `left:${ax}px;top:${ay}px;width:${this.colLeft[a.j + 1]! - ax}px;height:${this.rowTop[a.i + 1]! - ay}px`;
  }

  private paintColHeads(content: HTMLElement, j0: number, j1: number): void {
    const r = this.rect;
    const all = r.i1 === 0 && r.i2 === this.rows.length - 1;
    let html = "";
    for (let j = j0; j <= j1; j++) {
      const c = this.cols[j]!;
      const x = this.colLeft[j]!, w = this.colLeft[j + 1]! - x;
      const sel = j >= r.j1 && j <= r.j2 ? (all ? " full" : " sel") : "";
      const gapBefore = j > 0 ? this.cols[j - 1]! !== c - 1 : c > 0;
      const sorted = this.sheet.view?.sort?.col === c ? (this.sheet.view.sort.desc ? " ↓" : " ↑") : "";
      html += `<div class="h${sel}" style="left:${x}px;width:${w}px">${colName(c)}${sorted}${gapBefore ? '<i class="gap" title="Show hidden columns"></i>' : ""}</div>`;
    }
    if (content.innerHTML !== html) content.innerHTML = html;
  }

  private paintRowHeads(content: HTMLElement, i0: number, i1: number): void {
    const r = this.rect;
    const all = r.j1 === 0 && r.j2 === this.cols.length - 1;
    let html = "";
    for (let i = i0; i <= i1; i++) {
      const row = this.rows[i]!;
      const y = this.rowTop[i]!, hh = this.rowTop[i + 1]! - y;
      const sel = i >= r.i1 && i <= r.i2 ? (all ? " full" : " sel") : "";
      const prev = i > 0 ? this.rows[i - 1]! : -1;
      const hiddenGap = row - prev > 1 && this.sheet.hiddenRows.has(row - 1);
      html += `<div class="h${sel}" style="top:${y}px;height:${hh}px">${row + 1}${hiddenGap ? '<i class="gap" title="Show hidden rows"></i>' : ""}</div>`;
    }
    if (content.innerHTML !== html) content.innerHTML = html;
  }

  // ---- selection -------------------------------------------------------------------

  select(anchor: Pos, focus: Pos = anchor, scroll = true): void {
    this.anchor = anchor;
    this.focus = focus;
    if (scroll) this.scrollIntoView(focus);
    this.render();
    this.ev.select();
  }

  selectDisplay(i1: number, j1: number, i2 = i1, j2 = j1, scroll = true): void {
    this.select(this.dataOf(i1, j1), this.dataOf(i2, j2), scroll);
  }

  scrollIntoView(p: Pos): void {
    const { i, j } = this.displayOf(p);
    const fw = this.colLeft[this.fc]!, fh = this.rowTop[this.fr]!;
    const W = this.scroller.clientWidth - HEAD_W - fw, H = this.scroller.clientHeight - HEAD_H - fh;
    if (j >= this.fc) {
      const x = this.colLeft[j]! - fw, x2 = this.colLeft[j + 1]! - fw;
      const sl = this.scroller.scrollLeft;
      if (x < sl) this.scroller.scrollLeft = x;
      else if (x2 > sl + W) this.scroller.scrollLeft = Math.min(x, x2 - W);
    }
    if (i >= this.fr) {
      const y = this.rowTop[i]! - fh, y2 = this.rowTop[i + 1]! - fh;
      const st = this.scroller.scrollTop;
      if (y < st) this.scroller.scrollTop = y;
      else if (y2 > st + H) this.scroller.scrollTop = Math.min(y, y2 - H);
    }
  }

  /** Visible rows on one screen (for PageUp/PageDown). */
  get pageRows(): number {
    return Math.max(1, Math.floor((this.scroller.clientHeight - HEAD_H) / DEFAULT_ROW_H) - 1);
  }

  // ---- mouse ---------------------------------------------------------------------------

  private hover(e: MouseEvent): void {
    if (this.dragging) return;
    const t = this.hit(e);
    const edge = this.edgeAt(t.zone, t.sx, t.sy);
    this.scroller.style.cursor = edge >= 0 ? (t.zone === "col" ? "col-resize" : "row-resize") : "";
  }

  private mouseDown(e: MouseEvent): void {
    if (e.button !== 0 && e.button !== 2) return;
    // Clicks on the scrollbars themselves.
    if (e.offsetX > this.scroller.clientWidth || e.offsetY > this.scroller.clientHeight) return;
    const t = this.hit(e);
    if (e.button === 2) {
      // Right-click keeps a selection that contains the clicked cell.
      const r = this.rect;
      const inside = t.zone === "cell" ? t.i >= r.i1 && t.i <= r.i2 && t.j >= r.j1 && t.j <= r.j2
        : t.zone === "col" ? t.j >= r.j1 && t.j <= r.j2 && r.i1 === 0 && r.i2 === this.rows.length - 1
        : t.zone === "row" ? t.i >= r.i1 && t.i <= r.i2 && r.j1 === 0 && r.j2 === this.cols.length - 1 : true;
      if (inside) return;
    }
    e.preventDefault();
    if (e.button === 0 && t.zone === "cell" && this.ev.pick(this.dataOf(t.i, t.j))) return;
    this.focusGrid();
    const edge = this.edgeAt(t.zone, t.sx, t.sy);
    if (edge >= 0 && e.button === 0) return this.startResize(t.zone === "col" ? "col" : "row", edge, e);

    // Hidden-column / hidden-row markers in headers.
    if (t.zone === "col") {
      const x0 = this.colLeft[t.j]!;
      const c = this.cols[t.j]!;
      if (t.sx - x0 < 7 && (t.j === 0 ? c > 0 : this.cols[t.j - 1]! !== c - 1)) {
        const from = t.j === 0 ? 0 : this.cols[t.j - 1]! + 1;
        return this.ev.unhide("col", range(from, c - 1));
      }
    }
    if (t.zone === "row") {
      const y0 = this.rowTop[t.i]!;
      const r = this.rows[t.i]!;
      if (t.sy - y0 < 6 && r > 0 && this.sheet.hiddenRows.has(r - 1)) {
        let from = r - 1;
        while (from > 0 && this.sheet.hiddenRows.has(from - 1)) from--;
        return this.ev.unhide("row", range(from, r - 1));
      }
    }

    const last = { i: this.rows.length - 1, j: this.cols.length - 1 };
    if (t.zone === "corner") return this.selectDisplay(0, 0, last.i, last.j, false);

    // Filter button in the header row.
    if (t.zone === "cell" && e.button === 0) {
      const view = this.sheet.view;
      const r = this.rows[t.i]!, c = this.cols[t.j]!;
      if (view && r === view.headerRow && this.colLeft[t.j + 1]! - t.sx < 22) {
        this.selectDisplay(t.i, t.j, t.i, t.j, false);
        const box = this.cellBox(t.i, t.j);
        const rootBox = this.root.getBoundingClientRect();
        return this.ev.filterMenu(c, rootBox.left + box.x + box.w - 20, rootBox.top + box.y + box.h);
      }
    }

    const a = this.displayOf(this.anchor);
    const kind = t.zone === "col" ? "cols" : t.zone === "row" ? "rows" : "cells";
    const apply = (i: number, j: number, extend: boolean) => {
      if (kind === "cols") {
        if (extend) this.selectDisplay(0, a.j, last.i, j, false);
        else this.selectDisplay(0, j, last.i, j, false);
        // Keep the anchor on the clicked column's first row.
        this.anchor = this.dataOf(0, extend ? a.j : j);
      } else if (kind === "rows") {
        if (extend) this.selectDisplay(a.i, 0, i, last.j, false);
        else this.selectDisplay(i, 0, i, last.j, false);
      } else if (extend) this.select(this.anchor, this.dataOf(i, j), false);
      else this.selectDisplay(i, j, i, j, false);
    };
    apply(t.i, t.j, e.shiftKey);
    if (e.button !== 0) return;
    if (kind === "cols" && !e.shiftKey) this.anchor = this.dataOf(0, t.j);
    const startA = this.displayOf(this.anchor);
    this.dragging = { kind };
    const move = (ev: MouseEvent) => {
      const m = this.hit(ev);
      // Nudge-scroll when dragging past the edge.
      const W = this.scroller.clientWidth, H = this.scroller.clientHeight;
      if (m.x > W) this.scroller.scrollLeft += m.x - W;
      if (m.y > H) this.scroller.scrollTop += m.y - H;
      if (m.x < HEAD_W && kind !== "rows") this.scroller.scrollLeft -= HEAD_W - m.x;
      if (m.y < HEAD_H && kind !== "cols") this.scroller.scrollTop -= HEAD_H - m.y;
      if (kind === "cols") this.select(this.dataOf(0, startA.j), this.dataOf(last.i, m.j), false);
      else if (kind === "rows") this.select(this.dataOf(startA.i, 0), this.dataOf(m.i, last.j), false);
      else this.select(this.anchor, this.dataOf(m.i, m.j), false);
    };
    const up = () => {
      this.dragging = null;
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  private startResize(axis: "row" | "col", idx: number, e: MouseEvent): void {
    const startPos = axis === "col" ? e.clientX : e.clientY;
    const startSize = axis === "col" ? this.colLeft[idx + 1]! - this.colLeft[idx]! : this.rowTop[idx + 1]! - this.rowTop[idx]!;
    const box = this.cellBox(axis === "col" ? 0 : idx, axis === "col" ? idx : 0);
    const g = this.guide;
    g.hidden = false;
    g.className = `resize-guide ${axis}`;
    let size = startSize;
    const place = () => {
      if (axis === "col") g.style.cssText = `left:${box.x + size - 1}px`;
      else g.style.cssText = `top:${box.y + size - 1}px`;
    };
    place();
    this.dragging = { kind: "cells" };
    const move = (ev: MouseEvent) => {
      size = Math.max(axis === "col" ? 16 : 14, startSize + (axis === "col" ? ev.clientX : ev.clientY) - startPos);
      place();
    };
    const up = () => {
      g.hidden = true;
      this.dragging = null;
      window.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      if (Math.round(size) === Math.round(startSize)) return;
      // Resizing one of several selected columns resizes them all.
      const r = this.rect;
      const all = axis === "col" ? r.i1 === 0 && r.i2 === this.rows.length - 1 : r.j1 === 0 && r.j2 === this.cols.length - 1;
      const [lo, hi] = axis === "col" ? [r.j1, r.j2] : [r.i1, r.i2];
      const list = all && idx >= lo && idx <= hi ? range(lo, hi).map((k) => (axis === "col" ? this.cols[k]! : this.rows[k]!))
        : [axis === "col" ? this.cols[idx]! : this.rows[idx]!];
      this.ev.resize(axis, list, Math.round(size));
    };
    window.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
  }

  private dblClick(e: MouseEvent): void {
    const t = this.hit(e);
    const edge = this.edgeAt(t.zone, t.sx, t.sy);
    if (edge >= 0 && t.zone === "col") {
      const c = this.cols[edge]!;
      this.ev.resize("col", [c], this.fitWidth(c));
      return;
    }
    if (edge >= 0 && t.zone === "row") {
      this.ev.resize("row", [this.rows[edge]!], null);
      return;
    }
    if (t.zone === "cell") this.ev.edit(null);
  }

  private contextMenu(e: MouseEvent): void {
    e.preventDefault();
    if (performance.now() - this.lastLongPress < 1000) return; // Android also fires this after our long-press
    const t = this.hit(e);
    this.ev.menu(e.clientX, e.clientY, t.zone === "col" ? "col" : t.zone === "row" ? "row" : "cell");
  }

  // ---- touch ------------------------------------------------------------------------
  // One-finger panning stays native (passive listeners). A tap selects, tapping the
  // selected cell edits, a long press opens the menu, and the corner handle extends
  // the selection. The only timer is the long-press one, alive while a finger is down.

  private lastLongPress = 0;

  private bindTouch(): void {
    const sc = this.scroller;
    let start: { x: number; y: number } | null = null;
    let moved = false, long = false, timer = 0;
    sc.addEventListener("touchstart", (e) => {
      clearTimeout(timer);
      start = null;
      if (e.touches.length !== 1) return;
      const p = e.touches[0]!;
      const s = (start = { x: p.clientX, y: p.clientY });
      moved = long = false;
      timer = window.setTimeout(() => {
        if (moved || start !== s) return;
        long = true;
        this.lastLongPress = performance.now();
        this.longPress(s.x, s.y);
      }, 500);
    }, { passive: true });
    sc.addEventListener("touchmove", (e) => {
      const p = e.touches[0];
      if (start && p && Math.hypot(p.clientX - start.x, p.clientY - start.y) > 10) {
        moved = true;
        clearTimeout(timer);
      }
    }, { passive: true });
    sc.addEventListener("touchend", (e) => {
      clearTimeout(timer);
      const s = start;
      start = null;
      if (!s || moved) return;
      e.preventDefault(); // no synthetic mouse events or click after our own handling
      if (!long) this.tap(s.x, s.y);
    }, { passive: false });
    sc.addEventListener("touchcancel", () => {
      clearTimeout(timer);
      start = null;
    });

    // Selection handle: its own element with touch-action: none, so dragging it never scrolls.
    this.handle.addEventListener("pointerdown", (e) => {
      e.preventDefault();
      this.handle.setPointerCapture(e.pointerId);
      const move = (ev: PointerEvent) => {
        const m = this.hit(ev);
        const W = this.scroller.clientWidth, H = this.scroller.clientHeight;
        if (m.x > W - 16) this.scroller.scrollLeft += 24;
        if (m.y > H - 16) this.scroller.scrollTop += 24;
        this.select(this.anchor, this.dataOf(m.i, m.j), false);
      };
      const up = () => {
        this.handle.removeEventListener("pointermove", move);
        this.handle.removeEventListener("pointerup", up);
        this.handle.removeEventListener("pointercancel", up);
      };
      this.handle.addEventListener("pointermove", move);
      this.handle.addEventListener("pointerup", up);
      this.handle.addEventListener("pointercancel", up);
    });
  }

  private tap(x: number, y: number): void {
    const t = this.hit({ clientX: x, clientY: y });
    const r = this.rect;
    const a = this.displayOf(this.anchor);
    const onActive = t.zone === "cell" && r.i1 === r.i2 && r.j1 === r.j2 && a.i === t.i && a.j === t.j;
    this.focusGrid(); // also commits an open editor (it blurs)
    const last = { i: this.rows.length - 1, j: this.cols.length - 1 };
    if (t.zone === "corner") return this.selectDisplay(0, 0, last.i, last.j, false);
    if (t.zone === "col") return this.selectDisplay(0, t.j, last.i, t.j, false);
    if (t.zone === "row") return this.selectDisplay(t.i, 0, t.i, last.j, false);
    const view = this.sheet.view;
    if (view && this.rows[t.i] === view.headerRow && this.colLeft[t.j + 1]! - t.sx < 28) {
      this.selectDisplay(t.i, t.j, t.i, t.j, false);
      return this.ev.filterMenu(this.cols[t.j]!, x - 120, y + 12);
    }
    if (onActive) return this.ev.edit(null);
    this.selectDisplay(t.i, t.j, t.i, t.j, false);
  }

  private longPress(x: number, y: number): void {
    const t = this.hit({ clientX: x, clientY: y });
    const r = this.rect;
    const last = { i: this.rows.length - 1, j: this.cols.length - 1 };
    const inside = t.i >= r.i1 && t.i <= r.i2 && t.j >= r.j1 && t.j <= r.j2;
    if (t.zone === "col" && !(inside && r.i1 === 0 && r.i2 === last.i)) this.selectDisplay(0, t.j, last.i, t.j, false);
    else if (t.zone === "row" && !(inside && r.j1 === 0 && r.j2 === last.j)) this.selectDisplay(t.i, 0, t.i, last.j, false);
    else if (t.zone === "cell" && !inside) this.selectDisplay(t.i, t.j, t.i, t.j, false);
    this.ev.menu(x, y, t.zone === "col" ? "col" : t.zone === "row" ? "row" : "cell");
  }

  /** Put the touch handle on the selection's bottom-right corner, if that corner is on screen. */
  private placeHandle(W: number, H: number): void {
    const r = this.rect;
    const b = this.cellBox(r.i2, r.j2);
    const x = b.x + b.w, y = b.y + b.h;
    const ok = x > HEAD_W && y > HEAD_H && x < W && y < H;
    this.handle.hidden = !ok;
    if (ok) this.handle.style.transform = `translate(${x}px, ${y}px)`;
  }

  /** Width that fits the longest text in a column. */
  fitWidth(c: number): number {
    const ctx = document.createElement("canvas").getContext("2d")!;
    const font = getComputedStyle(this.root).getPropertyValue("--cell-font").trim() || "13px sans-serif";
    let max = 30;
    const s = this.sheet;
    for (const r of this.rows) {
      const v = this.wb.value(s.id, r, c);
      if (v === null) continue;
      const f = s.cells.get(key(r, c))?.f;
      if (f?.rot) continue;
      ctx.font = (f?.b ? "600 " : "") + font;
      max = Math.max(max, ctx.measureText(displayValue(v, f?.nf)).width + 14);
    }
    return Math.min(600, Math.ceil(max));
  }
}

/** Scale offsets[0..n] so offsets[n] <= max, shifting everything after by the same amount. */
function fitPrefix(offsets: Float64Array, n: number, max: number): void {
  const total = offsets[n]!;
  if (n === 0 || total <= max) return;
  const k = max / total;
  for (let i = 1; i <= n; i++) offsets[i] = offsets[i]! * k;
  const shift = total - max;
  for (let i = n + 1; i < offsets.length; i++) offsets[i] = offsets[i]! - shift;
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let i = a; i <= b; i++) out.push(i);
  return out;
}

function cellClass(f: Format | undefined, align: string, err: boolean, spill: boolean, flash: boolean, filter: string): string {
  let c = "c";
  if (align !== "left") c += align === "right" ? " ar" : " ac";
  if (f) {
    if (f.b) c += " b";
    if (f.i) c += " i";
    if (f.s) c += " s";
    if (f.wrap) c += " wrap";
    if (f.rot) c += " vt";
    if (f.fill && PALETTE_FILL.has(f.fill)) c += ` fill-${f.fill}`;
    if (f.color && PALETTE_FILL.has(f.color)) c += ` ink-${f.color}`;
  }
  if (err) c += " err";
  if (spill) c += " spill";
  if (flash) c += " flash";
  if (filter) c += ` ${filter}`;
  return c;
}

function cellStyle(f: Format | undefined): string {
  if (!f) return "";
  let s = "";
  if (f.fill && !PALETTE_FILL.has(f.fill)) s += `background:${cssColor(f.fill)};`;
  if (f.color && !PALETTE_FILL.has(f.color)) s += `color:${cssColor(f.color)};`;
  return s;
}

function cssColor(c: string): string {
  return /^#[0-9a-f]{3,8}$/i.test(c) || /^[a-z]+$/i.test(c) ? c : "inherit";
}
