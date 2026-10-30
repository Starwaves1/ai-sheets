// Tiny DOM helpers. No framework: the grid needs exact control over what touches the DOM.

type Attrs = Record<string, string | number | boolean | null | undefined | EventListener>;
type Child = Node | string | null | undefined | false;

export function h<K extends keyof HTMLElementTagNameMap>(tag: K, attrs: Attrs = {}, ...children: Child[]): HTMLElementTagNameMap[K] {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === null || v === undefined || v === false) continue;
    if (k.startsWith("on") && typeof v === "function") el.addEventListener(k.slice(2), v as EventListener);
    else if (k === "class") el.className = String(v);
    else if (v === true) el.setAttribute(k, "");
    else el.setAttribute(k, String(v));
  }
  for (const c of children) if (c !== null && c !== undefined && c !== false) el.append(c);
  return el;
}

/** Inline SVG icon from a path list on a 16×16 grid. */
export function icon(paths: string, title?: string): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.setAttribute("aria-hidden", "true");
  svg.classList.add("icon");
  svg.innerHTML = paths;
  if (title) svg.setAttribute("aria-label", title);
  return svg;
}

export const ICONS = {
  undo: '<path d="M5.5 4 2.5 7l3 3M3 7h6.5a3.5 3.5 0 0 1 0 7H8" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  redo: '<path d="M10.5 4l3 3-3 3M13 7H6.5a3.5 3.5 0 0 0 0 7H8" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  alignLeft: '<path d="M2.5 3.5h11M2.5 6.5h7M2.5 9.5h11M2.5 12.5h7" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>',
  alignCenter: '<path d="M2.5 3.5h11M4.5 6.5h7M2.5 9.5h11M4.5 12.5h7" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>',
  alignRight: '<path d="M2.5 3.5h11M6.5 6.5h7M2.5 9.5h11M6.5 12.5h7" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>',
  wrap: '<path d="M2.5 3.5h11M2.5 7.5h9a2 2 0 0 1 0 4H8m1.5-1.5L8 11.5 9.5 13M2.5 11.5h3" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  rotate: '<path d="M4 13.5V4m-2 2 2-2.5L6 6M8.5 13.5l2.25-8 2.25 8M9.4 10.5h2.7" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  filter: '<path d="M2.5 3.5h11l-4.25 5v4l-2.5 1.25V8.5z" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/>',
  freeze: '<path d="M2.5 2.5h11v11h-11zM2.5 6h11M6 2.5v11" fill="none" stroke="currentColor" stroke-width="1.4"/>',
  cond: '<path d="M2.5 2.5h5v5h-5zM8.5 8.5h5v5h-5z" fill="currentColor" opacity=".35"/><path d="M2.5 2.5h11v11h-11zM8 2.5v11M2.5 8h11" fill="none" stroke="currentColor" stroke-width="1.2"/>',
  caret: '<path d="M4.5 6.5 8 10l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  plus: '<path d="M8 3v10M3 8h10" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>',
  sortAsc: '<path d="M4.5 3v10m-2-2 2 2 2-2M9 4.5h4.5M9 8h3.25M9 11.5h2" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  sortDesc: '<path d="M4.5 13V3m-2 2 2-2 2 2M9 4.5h2M9 8h3.25M9 11.5h4.5" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  close: '<path d="m4 4 8 8m0-8-8 8" stroke="currentColor" stroke-width="1.4" stroke-linecap="round"/>',
  textColor: '<path d="M4.5 11 8 2.5 11.5 11M5.6 8.3h4.8" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round"/>',
  fillColor: '<path d="M3 8.5 7.5 4l4 4-4.5 4.5zM7.5 4 5.5 2" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><path d="M13 10.5s1.25 1.5 1.25 2.25a1.25 1.25 0 0 1-2.5 0c0-.75 1.25-2.25 1.25-2.25z" fill="currentColor"/>',
};

export const TAG_ICON = '<path d="M2.5 2.5h5.6l5.4 5.4-5.6 5.6-5.4-5.4z" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linejoin="round"/><circle cx="5.4" cy="5.4" r="1.1" fill="currentColor"/>';

export function clearChildren(el: Element): void {
  while (el.firstChild) el.firstChild.remove();
}
