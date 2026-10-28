// Home page: every workbook, newest edit first.

import { clearChildren, h, icon, ICONS } from "./dom.ts";
import { menu } from "./popups.ts";

interface BookSummary {
  id: string;
  name: string;
  version: number;
  updated: number;
  sheets: Array<{ name: string; used: string | null }>;
}

const when = (ts: number) => {
  const d = new Date(ts);
  const today = new Date();
  const time = d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
  if (d.toDateString() === today.toDateString()) return `Today, ${time}`;
  return `${d.toLocaleDateString([], { month: "short", day: "numeric", year: d.getFullYear() === today.getFullYear() ? undefined : "numeric" })}, ${time}`;
};

export async function openHome(root: HTMLElement): Promise<void> {
  document.title = "Sheets";
  const list = h("ol", { class: "books" });
  const fileIn = h("input", { type: "file", accept: ".csv,.tsv,text/csv,text/tab-separated-values", hidden: true });
  const error = h("p", { class: "error", hidden: true });

  const load = async () => {
    const res = await fetch("/api/books");
    const books = (await res.json()) as BookSummary[];
    clearChildren(list);
    if (!books.length) {
      list.append(h("li", { class: "empty" },
        h("p", {}, "No workbooks yet."),
        h("p", {}, "Create one, import a CSV, or ask an agent to run ", h("code", {}, "sheets import file.csv"), ".")));
    }
    for (const b of books) {
      const more = h("button", { class: "icon-btn more", title: `Actions for ${b.name}`, "aria-label": `Actions for ${b.name}` }, "⋯");
      more.addEventListener("click", (e) => {
        e.preventDefault();
        const r = more.getBoundingClientRect();
        menu([
          { label: "Open", run: () => (location.href = `/b/${b.id}`) },
          { label: "Download CSV (first sheet)", run: () => (location.href = `/api/books/${b.id}/csv`) },
          { label: "Rename", run: () => rename(b) },
          "-",
          { label: "Delete", run: () => remove(b) },
        ], r.right - 200, r.bottom + 4);
      });
      const tabs = b.sheets.map((s) => s.name).join(", ");
      const size = b.sheets[0]?.used ?? "empty";
      list.append(h("li", {},
        h("a", { href: `/b/${b.id}`, class: "book-link" },
          h("span", { class: "name" }, b.name),
          h("span", { class: "tabs-meta" }, `${b.sheets.length === 1 ? "1 sheet" : `${b.sheets.length} sheets`}: ${tabs}`),
          h("span", { class: "size" }, size),
          h("span", { class: "when" }, when(b.updated))),
        more));
    }
  };

  const create = async () => {
    const res = await fetch("/api/books", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ name: "Untitled workbook" }) });
    const b = (await res.json()) as BookSummary;
    location.href = `/b/${b.id}`;
  };

  const rename = async (b: BookSummary) => {
    const name = prompt("Rename workbook", b.name)?.trim();
    if (!name || name === b.name) return;
    await fetch(`/api/books/${b.id}/ops`, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ ops: [{ t: "rename", name }] }) });
    load();
  };

  const remove = async (b: BookSummary) => {
    if (!confirm(`Delete "${b.name}"? It disappears for everyone.`)) return;
    await fetch(`/api/books/${b.id}`, { method: "DELETE" });
    load();
  };

  fileIn.addEventListener("change", async () => {
    const f = fileIn.files?.[0];
    if (!f) return;
    const name = f.name.replace(/\.[^.]+$/, "");
    const res = await fetch(`/api/books?name=${encodeURIComponent(name)}`, {
      method: "POST",
      headers: { "content-type": f.name.endsWith(".tsv") ? "text/tab-separated-values" : "text/csv" },
      body: await f.text(),
    });
    if (!res.ok) {
      error.hidden = false;
      error.textContent = `Import failed: ${((await res.json()) as { error: string }).error}`;
      return;
    }
    location.href = `/b/${((await res.json()) as BookSummary).id}`;
  });

  root.replaceChildren(h("div", { class: "home" },
    h("header", { class: "bar" },
      h("a", { class: "mark", href: "/" }, h("span", { class: "glyph", "aria-hidden": "true" }), "Sheets"),
      h("div", { class: "spacer" }),
      h("button", { class: "btn", onclick: () => fileIn.click() }, "Import CSV"),
      h("button", { class: "btn primary", onclick: create }, icon(ICONS.plus), "New workbook"),
      fileIn),
    h("main", {},
      h("div", { class: "list-head" }, h("span", {}, "Name"), h("span", {}, "Sheets"), h("span", {}, "Size"), h("span", {}, "Last edit")),
      error, list,
      h("p", { class: "agents" }, "Agents edit these too. On this machine run ", h("code", {}, "sheets help"), "; changes show up in open tabs within a second."))));
  await load();
  // Refresh when coming back to the tab rather than polling.
  document.addEventListener("visibilitychange", () => document.visibilityState === "visible" && load());
}
