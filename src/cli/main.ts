// `sheets` command-line client. Talks to the HTTP API; see AGENTS.md for usage.

import { readFileSync, writeFileSync } from "node:fs";
import { colName, parseRange, rangeName, MAX_COLS, MAX_ROWS, type SheetRange } from "../core/a1.ts";
import { parseDelimited } from "../core/csv.ts";
import { describeOp } from "../core/describe.ts";
import { autofillOps } from "../core/fill.ts";
import { groupAt } from "../core/flags.ts";
import { COND_OPS, sheetFromJSON, type CondOp, type CondRule, type SheetJSON, type ViewState } from "../core/model.ts";
import type { FormatPatch, Op } from "../core/ops.ts";

const SERVER = (process.env.SHEETS_URL ?? "http://127.0.0.1:8790").replace(/\/$/, "");
const ACTOR = process.env.SHEETS_ACTOR ?? "agent";

interface SheetInfo {
  id: string;
  name: string;
  rows: number;
  cols: number;
  used: string | null;
  frozenRows: number;
  frozenCols: number;
}
interface BookInfo {
  id: string;
  name: string;
  url: string;
  version: number;
  updated: number;
  sheets: SheetInfo[];
}

class CliError extends Error {}

// ---- argument parsing ------------------------------------------------------

interface Args {
  pos: string[];
  flags: Map<string, string | true>;
}

const BOOL_FLAGS = new Set(["csv", "tsv", "json", "formulas", "desc", "bold", "no-bold", "italic", "no-italic", "strike", "no-strike", "wrap", "no-wrap", "rotate", "no-rotate", "reset", "formats", "copy", "clear", "all", "remove", "off", "yes", "help", "values"]);

function parseArgs(argv: string[]): Args {
  const pos: string[] = [];
  const flags = new Map<string, string | true>();
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i]!;
    if (a === "-o" && i + 1 < argv.length) flags.set("o", argv[++i]!);
    else if (a.startsWith("--") && a.length > 2) {
      const eq = a.indexOf("=");
      if (eq > 0) flags.set(a.slice(2, eq), a.slice(eq + 1));
      else {
        const name = a.slice(2);
        if (BOOL_FLAGS.has(name) || i + 1 >= argv.length) flags.set(name, true);
        else flags.set(name, argv[++i]!);
      }
    } else pos.push(a);
  }
  return { pos, flags };
}

const flag = (a: Args, name: string): string | undefined => {
  const v = a.flags.get(name);
  return typeof v === "string" ? v : undefined;
};

// ---- HTTP -------------------------------------------------------------------

async function call(method: string, path: string, body?: unknown, type = "application/json"): Promise<Response> {
  let res: Response;
  try {
    res = await fetch(SERVER + path, {
      method,
      headers: { "x-actor": ACTOR, ...(body !== undefined ? { "content-type": type } : {}) },
      ...(body !== undefined ? { body: typeof body === "string" ? body : JSON.stringify(body) } : {}),
    });
  } catch {
    throw new CliError(`cannot reach the sheets server at ${SERVER} (set SHEETS_URL?)`);
  }
  if (!res.ok) {
    let msg = res.statusText;
    try {
      msg = ((await res.json()) as { error?: string }).error ?? msg;
    } catch {}
    throw new CliError(msg);
  }
  return res;
}

const api = async <T>(method: string, path: string, body?: unknown, type?: string): Promise<T> =>
  (await (await call(method, path, body, type)).json()) as T;

const enc = encodeURIComponent;

async function info(book: string): Promise<BookInfo> {
  return api<BookInfo>("GET", `/api/books/${enc(book)}`);
}

function sheetOf(b: BookInfo, name?: string): SheetInfo {
  if (!name) return b.sheets[0]!;
  const s = b.sheets.find((x) => x.name.toLowerCase() === name.toLowerCase() || x.id === name);
  if (!s) throw new CliError(`no sheet "${name}" in ${b.name} (have: ${b.sheets.map((x) => x.name).join(", ")})`);
  return s;
}

/** Parse a range argument that may carry a sheet prefix. */
function rangeArg(b: BookInfo, text: string, sheetFlag?: string): { sheet: SheetInfo; range: SheetRange } {
  const g = parseRange(text);
  if (!g) throw new CliError(`"${text}" is not a range (try A1, B2:D9, C:C, 3:5, or Sheet1!A1)`);
  return { sheet: sheetOf(b, g.sheet ?? sheetFlag), range: g };
}

async function ops(book: BookInfo, list: Op[]): Promise<number> {
  const r = await api<{ v: number }>("POST", `/api/books/${enc(book.id)}/ops`, { ops: list });
  return r.v;
}

function readInput(src: string): string {
  return readFileSync(src === "-" ? 0 : src, "utf8");
}

// ---- output -----------------------------------------------------------------

function cellText(v: unknown): string {
  if (v === null || v === undefined) return "";
  return String(v).replace(/\t/g, " ").replace(/\r?\n/g, "⏎");
}

function printGrid(range: SheetRange, values: unknown[][]): void {
  const lines = ["\t" + Array.from({ length: values[0]?.length ?? 0 }, (_, j) => colName(range.c1 + j)).join("\t")];
  values.forEach((row, i) => lines.push(`${range.r1 + i + 1}\t` + row.map(cellText).join("\t")));
  console.log(lines.join("\n"));
}

const ago = (ts: number) => {
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 60) return `${s}s ago`;
  if (s < 3600) return `${Math.round(s / 60)}m ago`;
  if (s < 86400) return `${Math.round(s / 3600)}h ago`;
  return `${Math.round(s / 86400)}d ago`;
};

// ---- commands -----------------------------------------------------------------

const HELP = `sheets: read and edit workbooks on the sheets server (${SERVER})

  sheets ls                                   list workbooks
  sheets show <book>                          tabs, sizes, used ranges, URL
  sheets get <book> [range] [--csv|--tsv|--json] [--formulas]
  sheets set <book> <range> <value...>        one value; fills the range if it is larger
  sheets set <book> <anchor> --csv <file|-> | --tsv <file|-> | --json '[[1,2],[3,4]]'
  sheets clear <book> <range> [--formats]
  sheets fill <book> <source> <target> [--copy]   continue the pattern in source into target, like
                                              dragging the fill handle: 1,2,3 → 4,5,6; Item 1 → Item 2;
                                              Mon → Tue; formulas shift; formats follow. --copy repeats
                                              instead (and counts up from a single number)
  sheets format <book> <range> [--bold] [--italic] [--strike] [--color C] [--fill C]
                [--align left|center|right] [--nf 0.00|0%|#,##0|$#,##0.00] [--wrap] [--rotate] [--reset]
  sheets cond <book> [--sheet S]              list conditional fill rules
  sheets cond <book> add <range> <op> [value] [value2] --fill C [--color C] [--bold]
                ops: ${COND_OPS.join(" ")}
  sheets cond <book> rm <n> | clear [--sheet S]
  sheets freeze <book> <rows> <cols> [--sheet S]
  sheets width <book> <A:C> <px>  |  sheets height <book> <2:5> <px>
  sheets hide <book> <C:D|3:5>    |  sheets unhide <book> <C:D|3:5>
  sheets move <book> <5:6|C:D> --to <2|H>     move rows/columns so they start at row 2 / column H;
                                              formulas, formats, flags and sizes go with them
  sheets insert <book> <5:7|C:D>  (insert before)  |  sheets delete <book> <5:7|C:D>
  sheets sort <book> <col> [--desc] [--header N] [--sheet S]  |  sheets sort <book> --off
  sheets filter <book> <col> --hide v1,v2 | <col> <op> [value]  |  sheets filter <book> --clear
  sheets tab <book> add <name> | rename <old> <new> | rm <name> | move <name> <index>
  sheets new <name>  |  sheets rename <book> <name>  |  sheets rm <book> --yes
  sheets import <file.csv|-> [--name N] [--book B [--sheet S]]
  sheets export <book> [sheet] [-o file]
  sheets flags [book] [--all]                 flags people left for you, across all workbooks (--all: resolved too)
  sheets flag <book> <range> "note"           flag cells for someone to look at (--remove to take it off)
  sheets resolve <book> <cell> ["reply"]      mark a flag done; the reply shows on it until they dismiss it
  sheets watch <book>                         print changes as they happen
  sheets ops <book> '<json ops>'              raw ops (see AGENTS.md)

<book> is a workbook id, its name, or a unique name prefix. Ranges use A1
notation with an optional sheet prefix: B2, A1:C9, C:C, 3:5, 'Run log'!A1:B4.
Colors: green red amber blue violet grey (theme-aware) or #rrggbb.
Env: SHEETS_URL (server), SHEETS_ACTOR (name shown to others, default "agent").`;

async function main(argv: string[]): Promise<void> {
  const a = parseArgs(argv);
  const [cmd, ...rest] = a.pos;
  if (!cmd || cmd === "help" || a.flags.has("help")) {
    console.log(HELP);
    return;
  }
  const need = (n: number, usage: string) => {
    if (rest.length < n) throw new CliError(`usage: sheets ${usage}`);
  };

  switch (cmd) {
    case "ls": {
      const books = await api<BookInfo[]>("GET", "/api/books");
      if (!books.length) return console.log("no workbooks yet (sheets new <name> or sheets import <file.csv>)");
      for (const b of books) {
        const tabs = b.sheets.map((s) => `${s.name}${s.used ? ` ${s.used}` : ""}`).join(", ");
        console.log(`${b.id}  ${b.name}  [${tabs}]  v${b.version}  ${ago(b.updated)}`);
      }
      return;
    }
    case "show": {
      need(1, "show <book>");
      const b = await info(rest[0]!);
      console.log(`${b.name}  (id ${b.id}, v${b.version}, edited ${ago(b.updated)})\n${b.url.startsWith("http") ? b.url : SERVER + b.url}`);
      for (const s of b.sheets) {
        const frozen = s.frozenRows || s.frozenCols ? `, frozen ${s.frozenRows}r/${s.frozenCols}c` : "";
        console.log(`  ${s.name}: used ${s.used ?? "(empty)"}, grid ${s.rows}×${s.cols}${frozen}`);
      }
      return;
    }
    case "new": {
      need(1, "new <name>");
      const b = await api<BookInfo>("POST", "/api/books", { name: rest.join(" ") });
      console.log(`created ${b.name} (id ${b.id})`);
      return;
    }
    case "rename": {
      need(2, "rename <book> <new name>");
      const b = await info(rest[0]!);
      const v = await ops(b, [{ t: "rename", name: rest.slice(1).join(" ") }]);
      console.log(`renamed (v${v})`);
      return;
    }
    case "rm": {
      need(1, "rm <book> --yes");
      const b = await info(rest[0]!);
      if (!a.flags.has("yes")) throw new CliError(`this deletes "${b.name}"; repeat with --yes`);
      await api("DELETE", `/api/books/${enc(b.id)}`);
      console.log(`deleted ${b.name}`);
      return;
    }
    case "get": {
      need(1, "get <book> [range]");
      const ref = rest[1];
      const as = a.flags.has("csv") ? "csv" : a.flags.has("tsv") ? "tsv" : "json";
      const show = a.flags.has("formulas") ? "inputs" : as === "json" && !a.flags.has("json") ? "display" : "values";
      const qs = `?as=${as}&show=${show}${ref ? `&ref=${enc(ref)}` : ""}`;
      if (as !== "json") {
        process.stdout.write(await (await call("GET", `/api/books/${enc(rest[0]!)}/range${qs}`)).text());
        return;
      }
      const r = await api<{ book: string; sheet: string; range: string; version: number; values: unknown[][] }>(
        "GET", `/api/books/${enc(rest[0]!)}/range${qs}`);
      if (a.flags.has("json")) return console.log(JSON.stringify(r));
      console.log(`# ${r.book} / ${r.sheet}!${r.range} (v${r.version})`);
      printGrid(parseRange(r.range)!, r.values);
      return;
    }
    case "set": {
      need(2, "set <book> <range> <value> | --csv file | --tsv file | --json '[[...]]'");
      const [book, ref, ...vals] = rest;
      const path = `/api/books/${enc(book!)}/range?ref=${enc(ref!)}`;
      let r: { v: number; sheet: string; range: string; cells: number };
      if (a.flags.has("csv") || a.flags.has("tsv")) {
        const src = flag(a, "csv") ?? flag(a, "tsv") ?? vals[0] ?? "-";
        const tsv = a.flags.has("tsv");
        r = await api("PUT", path, readInput(src), tsv ? "text/tab-separated-values" : "text/csv");
      } else if (a.flags.has("json")) {
        const text = flag(a, "json") ?? vals.join(" ");
        r = await api("PUT", path, { values: JSON.parse(text) });
      } else {
        if (!vals.length) throw new CliError("give a value (use \"\" to clear, or sheets clear)");
        r = await api("PUT", path, { values: [[vals.join(" ")]] });
      }
      console.log(`set ${r.sheet}!${r.range} (${r.cells} cell${r.cells === 1 ? "" : "s"}, v${r.v})`);
      return;
    }
    case "clear": {
      need(2, "clear <book> <range> [--formats]");
      const r = await api<{ v: number; cleared: number }>(
        "DELETE", `/api/books/${enc(rest[0]!)}/range?ref=${enc(rest[1]!)}${a.flags.has("formats") ? "&formats=1" : ""}`);
      console.log(`cleared ${r.cleared} cell${r.cleared === 1 ? "" : "s"} (v${r.v})`);
      return;
    }
    case "format": {
      need(2, "format <book> <range> [--bold] [--fill C] ...");
      const patch: FormatPatch = {};
      if (a.flags.has("reset")) Object.assign(patch, { b: null, i: null, s: null, color: null, fill: null, align: null, nf: null, wrap: null, rot: null });
      for (const [f, k] of [["bold", "b"], ["italic", "i"], ["strike", "s"], ["wrap", "wrap"]] as const) {
        if (a.flags.has(f)) patch[k] = true;
        if (a.flags.has("no-" + f)) patch[k] = null;
      }
      if (a.flags.has("rotate")) patch.rot = 90;
      if (a.flags.has("no-rotate")) patch.rot = null;
      for (const k of ["color", "fill", "align", "nf"] as const) {
        const v = flag(a, k);
        if (v !== undefined) (patch as Record<string, unknown>)[k] = v === "none" || v === "" ? null : v;
      }
      if (patch.align && !["left", "center", "right"].includes(patch.align)) throw new CliError("--align is left, center or right");
      if (!Object.keys(patch).length) throw new CliError("nothing to change; see sheets help");
      const r = await api<{ v: number; sheet: string; range: string }>(
        "POST", `/api/books/${enc(rest[0]!)}/format?ref=${enc(rest[1]!)}`, patch);
      console.log(`formatted ${r.sheet}!${r.range} (v${r.v})`);
      return;
    }
    case "cond": return cond(a, rest);
    case "fill": {
      need(3, "fill <book> <source range> <target range> [--copy]");
      const b = await api<BookInfo & { doc: { sheets: SheetJSON[] } }>("GET", `/api/books/${enc(rest[0]!)}/doc`);
      const src = rangeArg(b, rest[1]!);
      const dst = rangeArg(b, rest[2]!, src.sheet.name);
      if (dst.sheet.id !== src.sheet.id) throw new CliError("source and target must be on the same sheet");
      // A target that contains the source means "fill the rest of it".
      let t = { r1: dst.range.r1, c1: dst.range.c1, r2: dst.range.r2, c2: dst.range.c2 };
      const g = src.range;
      if (t.c1 === g.c1 && t.c2 === g.c2 && t.r1 === g.r1 && t.r2 > g.r2) t = { ...t, r1: g.r2 + 1 };
      else if (t.c1 === g.c1 && t.c2 === g.c2 && t.r2 === g.r2 && t.r1 < g.r1) t = { ...t, r2: g.r1 - 1 };
      else if (t.r1 === g.r1 && t.r2 === g.r2 && t.c1 === g.c1 && t.c2 > g.c2) t = { ...t, c1: g.c2 + 1 };
      else if (t.r1 === g.r1 && t.r2 === g.r2 && t.c2 === g.c2 && t.c1 < g.c1) t = { ...t, c2: g.c1 - 1 };
      const sheet = sheetFromJSON(b.doc.sheets.find((x) => x.id === src.sheet.id)!);
      let list: Op[];
      try {
        list = autofillOps(sheet, g, t, a.flags.has("copy"));
      } catch (e) {
        throw new CliError(`${(e as Error).message} (e.g. fill <book> A2:A4 A5:A20)`);
      }
      const r = await api<{ v: number }>("POST", `/api/books/${enc(b.id)}/ops`, { ops: list, base: b.version });
      console.log(`filled ${src.sheet.name}!${rangeName(t)} from ${rangeName(g)} (v${r.v})`);
      return;
    }
    case "freeze": {
      need(3, "freeze <book> <rows> <cols>");
      const b = await info(rest[0]!);
      const s = sheetOf(b, flag(a, "sheet"));
      const v = await ops(b, [{ t: "freeze", s: s.id, rows: Number(rest[1]), cols: Number(rest[2]) }]);
      console.log(`froze ${rest[1]} rows, ${rest[2]} cols in ${s.name} (v${v})`);
      return;
    }
    case "width":
    case "height": {
      need(3, `${cmd} <book> <range> <px|auto>`);
      const b = await info(rest[0]!);
      const { sheet, range } = rangeArg(b, rest[1]!, flag(a, "sheet"));
      const axis = cmd === "width" ? "col" : "row";
      const [lo, hi] = axis === "col" ? [range.c1, range.c2] : [range.r1, range.r2];
      const px = rest[2] === "auto" ? null : Number(rest[2]);
      const sizes: Array<[number, number | null]> = [];
      for (let i = lo; i <= hi; i++) sizes.push([i, px]);
      const v = await ops(b, [{ t: "size", s: sheet.id, axis, sizes }]);
      console.log(`${axis === "col" ? "widths" : "heights"} set for ${rangeName(range)} (v${v})`);
      return;
    }
    case "hide":
    case "unhide":
    case "insert":
    case "delete": {
      need(2, `${cmd} <book> <C:D | 3:5>`);
      const b = await info(rest[0]!);
      const { sheet, range } = rangeArg(b, rest[1]!.includes(":") || rest[1]!.includes("!") ? rest[1]! : `${rest[1]}:${rest[1]}`, flag(a, "sheet"));
      const wholeCols = range.r1 === 0 && range.r2 === MAX_ROWS - 1;
      const wholeRows = range.c1 === 0 && range.c2 === MAX_COLS - 1;
      if (wholeCols === wholeRows) throw new CliError(`give whole rows (3:5) or whole columns (C:D), not ${rest[1]}`);
      const axis = wholeRows ? "row" : "col";
      const [lo, hi] = axis === "row" ? [range.r1, range.r2] : [range.c1, range.c2];
      let op: Op;
      if (cmd === "hide" || cmd === "unhide") {
        op = { t: "hide", s: sheet.id, axis, idx: Array.from({ length: hi - lo + 1 }, (_, i) => lo + i), hidden: cmd === "hide" };
      } else op = { t: cmd, s: sheet.id, axis, at: lo, n: hi - lo + 1 };
      const v = await ops(b, [op]);
      console.log(`${describeOp(op, () => sheet.name)} (v${v})`);
      return;
    }
    case "move": {
      need(2, "move <book> <5:6|C:D> --to <row|column>");
      const b = await info(rest[0]!);
      const { sheet, range } = rangeArg(b, rest[1]!);
      const dest = flag(a, "to");
      const rows = range.c1 === 0 && range.c2 >= MAX_COLS - 1, cols = range.r1 === 0 && range.r2 >= MAX_ROWS - 1;
      if (rows === cols || !dest) throw new CliError("give whole rows or columns and a destination: move <book> 5:6 --to 2, move <book> C:D --to H");
      const axis = rows ? "row" : "col";
      const from = rows ? range.r1 : range.c1, n = rows ? range.r2 - range.r1 + 1 : range.c2 - range.c1 + 1;
      const target = parseRange(rows ? `A${dest}` : `${dest}1`);
      if (!target) throw new CliError(`--to takes a ${rows ? "row number" : "column letter"}`);
      const start = rows ? target.r1 : target.c1;
      const count = rows ? sheet.rows : sheet.cols;
      if (start === from) throw new CliError("they are already there");
      if (start + n > count) throw new CliError(`the sheet has ${count} ${axis}s`);
      const op: Op = { t: "move", s: sheet.id, axis, from, n, to: start < from ? start : start + n };
      const v = await ops(b, [op]);
      console.log(`${describeOp(op, () => sheet.name)} (v${v})`);
      return;
    }
    case "sort":
    case "filter": return view(cmd, a, rest);
    case "tab":
    case "tabs": return tab(a, rest);
    case "import": {
      need(1, "import <file.csv|-> [--name N] [--book B [--sheet S]]");
      const text = readInput(rest[0]!);
      const type = rest[0]!.endsWith(".tsv") ? "text/tab-separated-values" : "text/csv";
      const book = flag(a, "book");
      if (book) {
        const sheetName = flag(a, "sheet") ?? rest[0]!.replace(/^.*\//, "").replace(/\.[^.]+$/, "");
        const r = await api<{ sheet: { name: string } }>("POST", `/api/books/${enc(book)}/sheets?name=${enc(sheetName)}`, text, type);
        console.log(`imported ${rest[0]} as sheet ${r.sheet.name}`);
      } else {
        const name = flag(a, "name") ?? rest[0]!.replace(/^.*\//, "").replace(/\.[^.]+$/, "");
        const b = await api<BookInfo>("POST", `/api/books?name=${enc(name)}`, text, type);
        console.log(`imported ${rest[0]} as ${b.name} (id ${b.id}, ${b.sheets[0]!.used ?? "empty"})`);
      }
      return;
    }
    case "export": {
      need(1, "export <book> [sheet] [-o file]");
      const text = await (await call("GET", `/api/books/${enc(rest[0]!)}/csv${rest[1] ? `?sheet=${enc(rest[1])}` : ""}`)).text();
      const out = flag(a, "o") ?? flag(a, "out");
      if (out) writeFileSync(out, text);
      else process.stdout.write(text);
      return;
    }
    case "ops": {
      need(2, "ops <book> '<json array of ops>'");
      const b = await info(rest[0]!);
      const list = JSON.parse(rest.slice(1).join(" ")) as Op[];
      const v = await ops(b, Array.isArray(list) ? list : [list]);
      console.log(`applied ${Array.isArray(list) ? list.length : 1} op(s) (v${v})`);
      return;
    }
    case "watch": return watch(rest);
    case "flags": return flags(a, rest);
    case "flag": {
      need(2, 'flag <book> <range> "note" | flag <book> <range> --remove');
      const b = await info(rest[0]!);
      const { sheet, range } = rangeArg(b, rest[1]!);
      const g = { r1: range.r1, c1: range.c1, r2: Math.min(range.r2, sheet.rows - 1), c2: Math.min(range.c2, sheet.cols - 1) };
      if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 10_000) throw new CliError("flag at most 10000 cells at once");
      const note = rest.slice(2).join(" ").trim();
      const remove = a.flags.has("remove");
      if (!remove && !note) throw new CliError('give a note: sheets flag <book> <range> "what to look at"');
      const flag = remove ? null : { note, by: ACTOR, ts: Date.now() };
      const cells: Array<[number, number, typeof flag]> = [];
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) cells.push([r, c, flag]);
      const v = await ops(b, [{ t: "flag", s: sheet.id, cells }]);
      console.log(`${remove ? "removed flags on" : "flagged"} ${sheet.name}!${rangeName(g)} (v${v})`);
      return;
    }
    case "resolve": {
      need(2, 'resolve <book> <cell> ["reply"]');
      const b = await info(rest[0]!);
      const { sheet, range } = rangeArg(b, rest[1]!);
      const doc = await api<{ doc: { sheets: SheetJSON[] } }>("GET", `/api/books/${enc(b.id)}/doc`);
      const sh = sheetFromJSON(doc.doc.sheets.find((x) => x.id === sheet.id)!);
      const here = sh.cells.get(range.r1 * 16384 + range.c1)?.flag;
      if (!here) throw new CliError(`no flag on ${sheet.name}!${rangeName({ ...range, r2: range.r1, c2: range.c1 })} (sheets flags ${JSON.stringify(b.name)} lists them)`);
      // Resolve the whole block of cells flagged together with this one.
      const group = groupAt({ name: b.name, sheets: [sh] }, sh.id, range.r1, range.c1)!;
      const reply = rest.slice(2).join(" ").trim();
      const done = { ...here, done: { by: ACTOR, ts: Date.now(), ...(reply ? { reply } : {}) } };
      const cells = group.cells.map(([r, c]) => [r, c, done] as [number, number, typeof done]);
      const v = await ops(b, [{ t: "flag", s: sheet.id, cells }]);
      console.log(`resolved the flag "${here.note}" on ${cells.length} cell${cells.length === 1 ? "" : "s"} (v${v})`);
      return;
    }
    default:
      throw new CliError(`unknown command "${cmd}"; run sheets help`);
  }
}

async function cond(a: Args, rest: string[]): Promise<void> {
  if (!rest[0]) throw new CliError("usage: sheets cond <book> [add|rm|clear] ...");
  const b = await api<BookInfo & { doc: { sheets: Array<{ id: string; name: string; cond?: CondRule[] }> } }>(
    "GET", `/api/books/${enc(rest[0])}/doc`);
  const sub = rest[1];
  const sheetFor = (name?: string) => {
    const s = sheetOf(b, name);
    return { s, rules: b.doc.sheets.find((x) => x.id === s.id)!.cond ?? [] };
  };
  if (!sub || sub === "ls") {
    const { s, rules } = sheetFor(flag(a, "sheet"));
    if (!rules.length) return console.log(`no conditional rules in ${s.name}`);
    rules.forEach((r, i) => console.log(`${i + 1}. ${rangeName(r.range)} ${r.op}${r.value !== undefined ? " " + r.value : ""}${r.value2 !== undefined ? " " + r.value2 : ""} → ${JSON.stringify(r.style)}`));
    return;
  }
  if (sub === "add") {
    const [, , ref, op, value, value2] = rest;
    if (!ref || !op) throw new CliError("usage: sheets cond <book> add <range> <op> [value] [value2] --fill C");
    if (!COND_OPS.includes(op as CondOp)) throw new CliError(`op is one of: ${COND_OPS.join(", ")}`);
    const { sheet, range } = rangeArg(b, ref, flag(a, "sheet"));
    const style: CondRule["style"] = {};
    if (flag(a, "fill")) style.fill = flag(a, "fill")!;
    if (flag(a, "color")) style.color = flag(a, "color")!;
    if (a.flags.has("bold")) style.b = true;
    if (!Object.keys(style).length) throw new CliError("give a style: --fill C, --color C, or --bold");
    const rule: CondRule = { range: { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2 }, op: op as CondOp, style };
    if (value !== undefined) rule.value = value;
    if (value2 !== undefined) rule.value2 = value2;
    const { rules } = sheetFor(sheet.name);
    const v = await ops(b, [{ t: "cond", s: sheet.id, rules: [...rules, rule] }]);
    console.log(`added rule ${rules.length + 1} to ${sheet.name} (v${v})`);
    return;
  }
  if (sub === "rm" || sub === "clear") {
    const { s, rules } = sheetFor(flag(a, "sheet"));
    const n = Number(rest[2]);
    if (sub === "rm" && !(n >= 1 && n <= rules.length)) throw new CliError(`rule number 1..${rules.length}`);
    const next = sub === "clear" ? [] : rules.filter((_, i) => i !== n - 1);
    const v = await ops(b, [{ t: "cond", s: s.id, rules: next }]);
    console.log(`${sub === "clear" ? "cleared rules" : `removed rule ${n}`} in ${s.name} (v${v})`);
    return;
  }
  throw new CliError(`unknown cond command "${sub}"`);
}

async function view(cmd: "sort" | "filter", a: Args, rest: string[]): Promise<void> {
  if (!rest[0]) throw new CliError(`usage: sheets ${cmd} <book> ...`);
  const b = await api<BookInfo & { doc: { sheets: Array<{ id: string; view?: ViewState | null }> } }>(
    "GET", `/api/books/${enc(rest[0])}/doc`);
  const s = sheetOf(b, flag(a, "sheet"));
  const cur: ViewState = structuredClone(b.doc.sheets.find((x) => x.id === s.id)!.view ?? { headerRow: 0, filters: {} });
  if (flag(a, "header")) cur.headerRow = Number(flag(a, "header")) - 1;
  const col = (t: string | undefined) => {
    const g = t ? parseRange(/^[A-Za-z]+$/.test(t) ? `${t}:${t}` : t) : null;
    if (!g) throw new CliError(`give a column letter, e.g. ${cmd === "sort" ? "sort <book> M" : "filter <book> B --hide x"}`);
    return g.c1;
  };
  let next: ViewState | null = cur;
  if (cmd === "sort") {
    if (a.flags.has("off")) delete cur.sort;
    else cur.sort = { col: col(rest[1]), ...(a.flags.has("desc") ? { desc: true } : {}) };
  } else if (a.flags.has("clear")) {
    cur.filters = {};
  } else {
    const c = col(rest[1]);
    const hide = flag(a, "hide");
    if (hide !== undefined) cur.filters[c] = { ...cur.filters[c], hide: hide.split(",") };
    else if (rest[2]) {
      if (!COND_OPS.includes(rest[2] as CondOp)) throw new CliError(`op is one of: ${COND_OPS.join(", ")}`);
      cur.filters[c] = { ...cur.filters[c], op: rest[2] as CondOp, ...(rest[3] !== undefined ? { value: rest[3] } : {}), ...(rest[4] !== undefined ? { value2: rest[4] } : {}) };
    } else delete cur.filters[c];
  }
  if (!cur.sort && !Object.keys(cur.filters).length && !flag(a, "header")) next = null;
  const op: Op = { t: "view", s: s.id, view: next };
  const v = await ops(b, [op]);
  console.log(`${describeOp(op, () => s.name)} (v${v})`);
}

async function tab(a: Args, rest: string[]): Promise<void> {
  if (!rest[0]) throw new CliError("usage: sheets tab <book> add|rename|rm|move ...");
  const b = await info(rest[0]);
  const [, sub, x, y] = rest;
  if (!sub) {
    b.sheets.forEach((s, i) => console.log(`${i}. ${s.name}  used ${s.used ?? "(empty)"}`));
    return;
  }
  let op: Op;
  if (sub === "add") {
    if (!x) throw new CliError("usage: sheets tab <book> add <name>");
    const r = await api<{ v: number }>("POST", `/api/books/${enc(b.id)}/sheets?name=${enc(rest.slice(2).join(" "))}`);
    console.log(`added sheet (v${r.v})`);
    return;
  } else if (sub === "rename") {
    if (!x || !y) throw new CliError("usage: sheets tab <book> rename <old> <new>");
    op = { t: "renameSheet", s: sheetOf(b, x).id, name: rest.slice(3).join(" ") };
  } else if (sub === "rm") {
    if (!a.flags.has("yes")) throw new CliError(`this deletes sheet "${x}"; repeat with --yes`);
    op = { t: "deleteSheet", s: sheetOf(b, x).id };
  } else if (sub === "move") {
    op = { t: "moveSheet", s: sheetOf(b, x).id, index: Number(y) };
  } else throw new CliError(`unknown tab command "${sub}"`);
  const v = await ops(b, [op]);
  console.log(`${describeOp(op, (id) => b.sheets.find((s) => s.id === id)?.name ?? id)} (v${v})`);
}

interface FlagRow {
  book: string; ref: string; cells: number; value: string; input: string; url: string;
  note: string; by: string; ts: number; done?: { by: string; ts: number; reply?: string };
}

async function flags(a: Args, rest: string[]): Promise<void> {
  const all = a.flags.has("all");
  const list = rest[0]
    ? (await api<FlagRow[]>("GET", `/api/books/${enc((await info(rest[0])).id)}/flags${all ? "?all=1" : ""}`))
    : await api<FlagRow[]>("GET", `/api/flags${all ? "?all=1" : ""}`);
  const open = list.filter((f) => !f.done).length;
  if (!list.length) return console.log(all ? "no flags" : "no open flags");
  console.log(`# ${open} open flag${open === 1 ? "" : "s"}${all ? `, ${list.length - open} resolved` : ""}. When done: sheets resolve <book> <cell> "what you did"`);
  const when = (ts: number) => {
    const d = new Date(ts);
    const p = (n: number) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())} (${ago(ts)})`;
  };
  for (const f of list) {
    const value = f.input.startsWith("=") ? `${JSON.stringify(f.value)} from ${f.input}` : JSON.stringify(f.value);
    const many = f.cells > 1 ? ` (${f.cells} cells, first shown)` : "";
    console.log(`${f.done ? "[done] " : ""}${f.book}  ${f.ref}  value ${value}${many}\n  note: ${JSON.stringify(f.note)}  by ${f.by}, ${when(f.ts)}` +
      (f.done ? `\n  resolved by ${f.done.by}, ${when(f.done.ts)}${f.done.reply ? `: ${JSON.stringify(f.done.reply)}` : ""}` : ""));
  }
}

async function watch(rest: string[]): Promise<void> {
  if (!rest[0]) throw new CliError("usage: sheets watch <book>");
  const b = await info(rest[0]);
  const names = new Map(b.sheets.map((s) => [s.id, s.name]));
  console.log(`watching ${b.name} from v${b.version} (Ctrl-C to stop)`);
  const res = await call("GET", `/api/books/${enc(b.id)}/events?since=${b.version}`);
  const decoder = new TextDecoder();
  let buf = "";
  for await (const chunk of res.body as unknown as AsyncIterable<Uint8Array>) {
    buf += decoder.decode(chunk, { stream: true });
    let i;
    while ((i = buf.indexOf("\n\n")) >= 0) {
      const block = buf.slice(0, i);
      buf = buf.slice(i + 2);
      const event = /^event: (.*)$/m.exec(block)?.[1];
      const data = /^data: (.*)$/m.exec(block)?.[1];
      if (event === "deleted") return console.log("workbook deleted");
      if (event !== "change" || !data) continue;
      const c = JSON.parse(data) as { v: number; actor: string; ts: number; ops: Op[] };
      for (const op of c.ops) {
        if (op.t === "addSheet") names.set(op.sheet.id, op.sheet.name);
        if (op.t === "renameSheet") names.set(op.s, op.name);
      }
      const time = new Date(c.ts).toTimeString().slice(0, 8);
      for (const op of c.ops) console.log(`v${c.v} ${time} ${c.actor}: ${describeOp(op, (id) => names.get(id) ?? id, true)}`);
    }
  }
}

main(process.argv.slice(2)).catch((e) => {
  console.error(`sheets: ${e instanceof CliError || e instanceof SyntaxError ? e.message : (e as Error).stack}`);
  process.exit(1);
});
