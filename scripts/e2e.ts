// End-to-end check against a running server, in a real (headless) browser:
//   1. an agent edit through the CLI appears in an open tab without reload
//   2. an edit typed in the browser is readable through the CLI (and undo works)
//   3. the page is idle when nothing changes (no rAF, timers, layouts or animations)
//   4. a 50k-cell sheet scrolls with only the visible cells in the DOM
// Usage: SHEETS_URL=http://192.168.1.5:8790 node dist/e2e.js [--shots dir]

import { execFile } from "node:child_process";
import { join } from "node:path";
import { promisify } from "node:util";
import { BROWSER, launch } from "./browser.ts";

const BASE = (process.env.SHEETS_URL ?? "http://192.168.1.5:8790").replace(/\/$/, "");
const shots = process.argv.includes("--shots") ? process.argv[process.argv.indexOf("--shots") + 1] : undefined;
const run = promisify(execFile);
const CLI = join(import.meta.dirname, "..", "bin", "sheets");
const cli = async (...args: string[]) => (await run(CLI, args, { env: { ...process.env, SHEETS_URL: BASE, SHEETS_ACTOR: "e2e-agent" } })).stdout;

let failed = 0;
function check(ok: boolean, what: string, detail = ""): void {
  console.log(`${ok ? "PASS" : "FAIL"}  ${what}${detail ? `  (${detail})` : ""}`);
  if (!ok) failed++;
}

async function api(method: string, path: string, body?: unknown, type = "application/json") {
  const res = await fetch(BASE + path, {
    method,
    headers: body !== undefined ? { "content-type": type, "x-actor": "e2e" } : { "x-actor": "e2e" },
    ...(body !== undefined ? { body: typeof body === "string" ? body : JSON.stringify(body) } : {}),
  });
  return res.json() as Promise<any>;
}

const INSTRUMENT = `
  window.__counts = { raf: 0, timeout: 0, interval: 0 };
  const raf = window.requestAnimationFrame.bind(window);
  window.requestAnimationFrame = (cb) => { window.__counts.raf++; return raf(cb); };
  const st = window.setTimeout.bind(window);
  window.setTimeout = (...a) => { window.__counts.timeout++; return st(...a); };
  const si = window.setInterval.bind(window);
  window.setInterval = (...a) => { window.__counts.interval++; return si(...a); };
`;

const browser = await launch();
const created: string[] = [];
try {
  const book = await api("POST", "/api/books", { name: `e2e ${Date.now()}` });
  created.push(book.id);
  await api("PUT", `/api/books/${book.id}/range?ref=A1`, { values: [["item", "qty"], ["tea", 3], ["cake", 4]] });

  const ctx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  await ctx.addInitScript(INSTRUMENT);
  const page = await ctx.newPage();
  await page.goto(`${BASE}/b/${book.id}`, { waitUntil: "networkidle" });
  await page.waitForSelector(".live.live");
  const loads = await page.evaluate(() => performance.getEntriesByType("navigation").length);

  // 1. CLI -> browser
  const token = `from-cli-${Math.random().toString(36).slice(2, 7)}`;
  const t0 = Date.now();
  await cli("set", book.id, "B5", token);
  const cliMs = Date.now() - t0;
  await page.waitForFunction((t) => [...document.querySelectorAll(".c")].some((e) => e.textContent === t), token, { timeout: 5000 });
  const seenMs = Date.now() - t0;
  const flashed = await page.evaluate((t) => [...document.querySelectorAll(".c")].find((e) => e.textContent === t)?.classList.contains("flash"), token);
  const activity = await page.textContent(".activity");
  const sameLoad = await page.evaluate(() => performance.getEntriesByType("navigation").length);
  check(seenMs < 1500, "CLI edit shows in the open tab", `${seenMs} ms after CLI start, CLI itself ${cliMs} ms`);
  check(sameLoad === loads, "no page reload");
  check(!!flashed, "remote edit is highlighted");
  check(!!activity?.includes("e2e-agent"), "activity line names the agent", activity ?? "");
  if (shots) await page.screenshot({ path: join(shots, "e2e-live-edit.png") });

  // 2. browser -> CLI
  await page.click(".addr");
  await page.keyboard.type("C2");
  await page.keyboard.press("Enter");
  await page.keyboard.type("typed in browser");
  await page.keyboard.press("Enter");
  await page.keyboard.type("=B2*10");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  const got = await cli("get", book.id, "C2:C3", "--csv");
  check(got === "typed in browser\n30\n", "browser edits read back through the CLI", JSON.stringify(got));

  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  const afterUndo = await cli("get", book.id, "C3", "--csv");
  check(afterUndo.trim() === "", "undo in the browser reaches the server", JSON.stringify(afterUndo));

  // Real keyboard copy and paste (Safari and Firefox only fire clipboard events on an editable target).
  const MODK = process.platform === "darwin" ? "Meta" : "Control";
  const goto = async (ref: string) => {
    await page.click(".addr");
    await page.keyboard.type(ref);
    await page.keyboard.press("Enter");
  };
  await goto("A2:B3");
  await page.keyboard.press(`${MODK}+C`);
  await goto("H2");
  await page.keyboard.press(`${MODK}+V`);
  await page.waitForTimeout(400);
  const kbPaste = await cli("get", book.id, "H2:I3", "--csv");
  check(kbPaste === "tea,3\ncake,4\n", "keyboard copy and paste a block", JSON.stringify(kbPaste));
  // From a frozen pane: freeze the first row and column, copy A1 (frozen corner), paste it with a click-selected target.
  const sid0 = (await api("GET", `/api/books/${book.id}`)).sheets[0].id;
  await api("POST", `/api/books/${book.id}/ops`, { ops: [{ t: "freeze", s: sid0, rows: 1, cols: 1 }] });
  await page.waitForTimeout(300);
  const a1 = await page.evaluate(() => {
    const el = [...document.querySelectorAll(".c")].find((e) => e.textContent === "item")!.getBoundingClientRect();
    return { x: el.x + el.width / 2, y: el.y + el.height / 2 };
  });
  await page.mouse.click(a1.x, a1.y);
  await page.keyboard.press(`${MODK}+C`);
  await goto("J5");
  await page.keyboard.press(`${MODK}+V`);
  await page.waitForTimeout(400);
  const frozenPaste = (await cli("get", book.id, "J5", "--csv")).trim();
  check(frozenPaste === "item", "copy from a frozen-pane cell (mouse-selected) pastes", JSON.stringify(frozenPaste));
  await api("POST", `/api/books/${book.id}/ops`, { ops: [{ t: "freeze", s: sid0, rows: 0, cols: 0 }] });
  // Pasting past the last column grows the sheet instead of dropping cells.
  await api("PUT", `/api/books/${book.id}/range?ref=A20`, { values: [["p", "q", "r", "s"]] });
  await goto("A20:D20");
  await page.keyboard.press(`${MODK}+C`);
  await goto("Y1");
  await page.keyboard.press(`${MODK}+V`);
  await page.waitForTimeout(400);
  const pasted = await cli("get", book.id, "Y1:AB1", "--csv");
  check(pasted === "p,q,r,s\n", "paste past the last column grows the sheet", JSON.stringify(pasted));
  await cli("clear", book.id, "A20:D20");
  // Typing still starts an edit after the focus moved to the clipboard proxy.
  await goto("K2");
  await page.keyboard.type("typed after copy");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  check((await cli("get", book.id, "K2", "--csv")).trim() === "typed after copy", "typing over a selected cell still edits it");

  // Formula bar: Enter commits and Esc cancels; arrows move the caret; clicking a cell commits and
  // selects it (or, mid-formula, inserts its reference); afterwards typing and the clipboard still work.
  const inputs = async (ref: string) => (await api("GET", `/api/books/${book.id}/range?ref=${ref}&show=inputs`)).values.map((r: string[]) => r.join("|")).join("/");
  const cellCenter = (col: string, row: string) => page.evaluate(([c, r]) => {
    const f = (sel: string, t: string) => [...document.querySelectorAll(sel)].find((e) => e.textContent === t)!.getBoundingClientRect();
    const a = f(".colhead .h", c), b = f(".rowhead .h", r);
    return { x: a.x + a.width / 2, y: b.y + b.height / 2 };
  }, [col, row] as const);
  const fbar = async (ref: string, text: string) => {
    await goto(ref);
    await page.click(".formula");
    await page.keyboard.type(text);
  };
  await goto("A1");
  await fbar("C5", "hello");
  await page.keyboard.press("Enter");
  await page.keyboard.type("next");
  await page.keyboard.press("Enter");
  await fbar("C7", "nope");
  await page.keyboard.press("Escape");
  await fbar("C8", "abc");
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.press("ArrowLeft");
  await page.keyboard.type("X");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  check(await inputs("C5:C8") === "hello/next//aXbc", "formula bar: Enter commits and focus returns to the grid, Esc cancels, arrows move the caret", await inputs("C5:C8"));
  await fbar("C9", "clicked");
  const d9 = await cellCenter("D", "9");
  await page.mouse.click(d9.x, d9.y);
  const afterClick = await page.evaluate(() => [(document.querySelector(".addr") as HTMLInputElement).value, (document.querySelector(".editor") as HTMLElement).hidden]);
  await page.keyboard.type("after");
  await page.keyboard.press("Enter");
  await fbar("C10", "=1+");
  for (const row of ["2", "3"]) {
    const p = await cellCenter("B", row);
    await page.mouse.click(p.x, p.y);
  }
  await page.keyboard.type("*2");
  await page.keyboard.press("Enter");
  await goto("C9");
  await page.keyboard.press(`${MODK}+C`);
  await goto("E10");
  await page.keyboard.press(`${MODK}+V`);
  await page.waitForTimeout(400);
  const fb = await inputs("C9:E10");
  check(fb === "clicked|after|/=1+B3*2||clicked" && afterClick[0] === "D9" && afterClick[1] === true,
    "formula bar: clicking a cell commits and selects it; mid-formula it inserts the reference; then typing and copy/paste work", `${fb}, after click ${JSON.stringify(afterClick)}`);
  await cli("clear", book.id, "C5:E10");
  // Fill handle: a real mouse drag continues the series; Ctrl while dragging counts up from a single number.
  await api("PUT", `/api/books/${book.id}/range?ref=B12`, { values: [[1, "Item 1", "", 5], [2, "Item 2"], [3, "Item 3"]] });
  const drag = async (to: { x: number; y: number }, mod?: string) => {
    await page.waitForTimeout(100); // the handle moves on the next paint
    const hb = (await page.locator(".fill-handle").boundingBox())!;
    await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
    await page.mouse.down();
    for (let k = 1; k <= 6; k++) await page.mouse.move(hb.x + (to.x - hb.x) * k / 6, hb.y + (to.y - hb.y) * k / 6);
    const preview = await page.evaluate(() => !(document.querySelector(".fill-preview") as HTMLElement).hidden);
    if (mod) await page.keyboard.down(mod);
    await page.mouse.up();
    if (mod) await page.keyboard.up(mod);
    await page.waitForTimeout(400);
    return preview;
  };
  await goto("A1");
  await goto("B12:C14");
  const previewed = await drag(await cellCenter("B", "18"));
  const series = await cli("get", book.id, "B15:C18", "--csv");
  check(previewed && series === "4,Item 4\n5,Item 5\n6,Item 6\n7,Item 7\n", "dragging the fill handle continues the series (dashed preview while dragging)", JSON.stringify(series));
  await goto("E12");
  await drag(await cellCenter("H", "12"), "Control");
  const counted = await cli("get", book.id, "F12:H12", "--csv");
  check(counted === "6,7,8\n", "Ctrl while dragging counts up from a single number", JSON.stringify(counted));
  await page.keyboard.press(`${MODK}+Z`);
  await page.waitForTimeout(300);
  check((await cli("get", book.id, "F12:H12", "--csv")) === ",,\n", "a fill is one undoable change");
  await cli("clear", book.id, "B12:H18");

  // Browser and OS shortcuts pass through: with the grid focused, none of these keydowns is
  // default-prevented (synthetic events: they never really reload or close anything). Ours still are.
  const KEYS = (mac: boolean) => {
    const M = mac ? "metaKey" : "ctrlKey";
    const k = (key: string, mods: string[] = [], code?: string) => ({ key, code: code ?? (key.length === 1 ? `Key${key.toUpperCase()}` : key), mods });
    const essentials = [
      k("r", [M]), k("R", [M, "shiftKey"]), k("F5"), k("t", [M]), k("w", [M]), k("n", [M]), k("T", [M, "shiftKey"]),
      k("Tab", ["ctrlKey"]), k("Tab", ["ctrlKey", "shiftKey"]), k("1", [M], "Digit1"), k("9", [M], "Digit9"), k("l", [M]), k("d", [M]),
      k("=", [M], "Equal"), k("-", [M], "Minus"), k("0", [M], "Digit0"), k("p", [M]), k("f", [M]), k("g", [M]), k("h", [M]),
      k("I", [M, mac ? "altKey" : "shiftKey"]), k("J", [M, mac ? "altKey" : "shiftKey"]), k("F12"), k("F11"), k("F6"), k("F3"),
      ...(mac
        ? [k("y", [M]), k("[", [M], "BracketLeft"), k("]", [M], "BracketRight"), k("q", [M]), k("m", [M]), k(",", [M], "Comma"),
          k("a", ["ctrlKey"]), k("e", ["ctrlKey"]), k("k", ["ctrlKey"]), k("ArrowLeft", ["ctrlKey"]), k("ArrowRight", ["ctrlKey"])]
        : [k("PageDown", ["ctrlKey"]), k("PageUp", ["ctrlKey"]), k("ArrowLeft", ["altKey"]), k("ArrowRight", ["altKey"]), k("Home", ["altKey"]),
          k("d", ["altKey"]), k("u", [M]), k("h", [M]), k("j", [M]), k("k", [M]), k("o", [M]), k("s", [M]), k("e", [M]), k("ArrowUp", ["metaKey"])]),
    ];
    const ours = [k("ArrowDown"), k("b", [M]), k("b", [M]), k("i", [M]), k("i", [M]), k("z", [M]), k("Z", [M, "shiftKey"]), k("a", [M]), ...(mac ? [] : [k("y", [M])])];
    return { essentials, ours };
  };
  const pressAll = (p: typeof page, mac: boolean) => p.evaluate((sets) => {
    const t = document.querySelector(".kbd-proxy")!;
    const send = (x: { key: string; code: string; mods: string[] }) =>
      !t.dispatchEvent(new KeyboardEvent("keydown", { key: x.key, code: x.code, bubbles: true, cancelable: true, ...Object.fromEntries(x.mods.map((m) => [m, true])) }));
    const name = (x: { key: string; mods: string[] }) => [...x.mods.map((m) => m.replace("Key", "")), x.key].join("+");
    return { taken: sets.essentials.filter(send).map(name), missed: sets.ours.filter((x) => !send(x)).map(name) };
  }, KEYS(mac));
  await goto("H20");
  const linux = await pressAll(page, false);
  await page.keyboard.press("Escape");
  const macCtx = await browser.newContext({ viewport: { width: 1280, height: 800 } });
  await macCtx.addInitScript(() => Object.defineProperty(Navigator.prototype, "platform", { get: () => "MacIntel" }));
  const macPage = await macCtx.newPage();
  await macPage.goto(`${BASE}/b/${book.id}`, { waitUntil: "networkidle" });
  await macPage.waitForSelector(".live.live");
  const macHint = await macPage.getAttribute(".tool[aria-label^='Redo']", "title");
  const mac = await pressAll(macPage, true);
  await macCtx.close();
  check(!linux.taken.length && !linux.missed.length, "Windows/Linux: browser shortcuts pass through; ours still work",
    `taken: ${linux.taken.join(" ") || "none"}; not handled: ${linux.missed.join(" ") || "none"}`);
  check(!mac.taken.length && !mac.missed.length && macHint === "Redo (⌘⇧Z)", "Mac: browser shortcuts pass through; ours use ⌘",
    `taken: ${mac.taken.join(" ") || "none"}; not handled: ${mac.missed.join(" ") || "none"}; redo hint ${macHint}`);

  // Native text editing inside the cell editor: word jumps and select-all act on the text.
  await goto("H21");
  await page.keyboard.type("hello world");
  await page.keyboard.press(`${MODK}+ArrowLeft`);
  const caret = await page.evaluate(() => { const e = document.querySelector(".editor") as HTMLTextAreaElement; return [e.hidden, e.selectionStart, e.value]; });
  await page.keyboard.press(`${MODK}+A`);
  const sel = await page.evaluate(() => { const e = document.querySelector(".editor") as HTMLTextAreaElement; return [e.selectionStart, e.selectionEnd]; });
  await page.keyboard.press("Escape");
  check(caret[0] === false && caret[1] !== 11 && caret[2] === "hello world" && sel[0] === 0 && sel[1] === 11,
    "in the cell editor, modifier+arrow moves the caret and select-all selects the text", `${JSON.stringify(caret)} ${JSON.stringify(sel)}`);
  // Right-click opens our menu; Shift+right-click leaves the browser's own.
  const h21 = await cellCenter("H", "21");
  await page.mouse.click(h21.x, h21.y, { button: "right" });
  const ourMenu = await page.evaluate(() => !!document.querySelector(".popup.menu"));
  await page.keyboard.press("Escape");
  const shiftPrevented = await page.evaluate(([x, y]) => !document.querySelector(".scroller")!.dispatchEvent(
    new MouseEvent("contextmenu", { clientX: x, clientY: y, shiftKey: true, bubbles: true, cancelable: true })), [h21.x, h21.y] as const);
  check(ourMenu && !shiftPrevented && !(await page.evaluate(() => !!document.querySelector(".popup.menu"))), "right-click opens the sheet menu; Shift+right-click is left to the browser");

  // 3. idle cost
  await page.waitForTimeout(2200); // let the highlight fade finish
  // Layout and style-recalc counters come from Chromium's DevTools protocol; other browsers check timers and animations only.
  const cdp = BROWSER === "chromium" ? await ctx.newCDPSession(page) : null;
  await cdp?.send("Performance.enable");
  const metric = async (): Promise<Record<string, number>> => (cdp ? Object.fromEntries(((await cdp.send("Performance.getMetrics")).metrics as Array<{ name: string; value: number }>).map((m) => [m.name, m.value])) : {});
  await page.evaluate(() => Object.assign((window as any).__counts, { raf: 0, timeout: 0, interval: 0 }));
  const m0 = await metric();
  await page.waitForTimeout(5000);
  const m1 = await metric();
  const counts = await page.evaluate(() => (window as any).__counts);
  const anims = await page.evaluate(() => document.getAnimations().filter((a) => a.playState === "running").length);
  const d = (k: string) => Math.round(((m1[k] ?? 0) - (m0[k] ?? 0)) * 1000) / 1000;
  const idle = counts.raf === 0 && counts.timeout === 0 && counts.interval === 0 && anims === 0 && d("LayoutCount") === 0 && d("RecalcStyleCount") === 0;
  check(idle, "idle page does no work for 5 s",
    `rAF ${counts.raf}, timeouts ${counts.timeout}, intervals ${counts.interval}, running animations ${anims}, layouts ${d("LayoutCount")}, style recalcs ${d("RecalcStyleCount")}, script ${d("ScriptDuration")} s, tasks ${d("TaskDuration")} s`);

  // 4. two browser tabs writing at once, one of them inserting a row
  // Its own browser, as a second person would have: in Firefox and WebKit, pages in one browser
  // process share window focus, so typing into both at once drops keys a real user could never send.
  const browser2 = await launch();
  const ctx2 = await browser2.newContext({ viewport: { width: 1280, height: 800 } });
  const other = await ctx2.newPage();
  await other.goto(`${BASE}/b/${book.id}`, { waitUntil: "networkidle" });
  await other.waitForSelector(".live.live");
  const typeAt = async (p: typeof page, ref: string, text: string) => {
    await p.click(".addr");
    await p.keyboard.type(ref);
    await p.keyboard.press("Enter");
    await p.keyboard.type(text);
    await p.keyboard.press("Enter");
  };
  await Promise.all([
    (async () => { for (let i = 0; i < 8; i++) await typeAt(page, `E${i + 2}`, `a${i}`); })(),
    (async () => { for (let i = 0; i < 8; i++) await typeAt(other, `F${i + 2}`, `b${i}`); })(),
  ]);
  const typed = await cli("get", book.id, "E2:F9", "--csv");
  const want = Array.from({ length: 8 }, (_, i) => `a${i},b${i}`).join("\n") + "\n";
  check(typed === want, "two tabs typing at once never clobber each other", JSON.stringify(typed));
  // An agent inserts a row above row 2: both tabs and the server agree on where the data went.
  const sid = (await api("GET", `/api/books/${book.id}`)).sheets[0].id;
  await api("POST", `/api/books/${book.id}/ops`, { ops: [{ t: "insert", s: sid, axis: "row", at: 1, n: 1 }] });
  await page.waitForTimeout(500);
  const shifted = await cli("get", book.id, "E2:E3", "--csv");
  const rowOf = (p: typeof page) => p.evaluate(() => {
    const el = [...document.querySelectorAll(".c")].find((e) => e.textContent === "a0") as HTMLElement | undefined;
    return el ? Math.round(el.getBoundingClientRect().top) : -1;
  });
  const [ya, yb] = [await rowOf(page), await rowOf(other)];
  check(shifted === "\na0\n" && ya === yb && ya > 0, "a row inserted by an agent moves data in both tabs", `server ${JSON.stringify(shifted)}, a0 at y=${ya} and y=${yb}`);
  const inSync = await Promise.all([page, other].map((p) => p.evaluate(() => [...document.querySelectorAll(".c")].map((e) => e.textContent).filter((t) => /^(a|b)\d$/.test(t!)).sort().join(","))));
  check(inSync[0] === inSync[1] && inSync[0]!.split(",").length === 16, "both tabs show the same cells as the server", `${inSync[0]!.split(",").length} cells each`);
  await browser2.close();

  // 5. 50k cells
  const rows: string[] = [];
  for (let r = 0; r < 1000; r++) rows.push(Array.from({ length: 50 }, (_, c) => (c === 0 ? `row ${r + 1}` : String((r * 50 + c) % 997))).join(","));
  const big = await api("POST", `/api/books?name=${encodeURIComponent(`e2e big ${Date.now()}`)}`, rows.join("\n"), "text/csv");
  created.push(big.id);
  const p2 = await ctx.newPage();
  await p2.goto(`${BASE}/b/${big.id}`, { waitUntil: "networkidle" });
  await p2.waitForSelector(".c");
  const perf = await p2.evaluate(async () => {
    const sc = document.querySelector(".scroller") as HTMLElement;
    const frame = () => new Promise((r) => requestAnimationFrame(() => r(null)));
    const times: number[] = [];
    let maxCells = 0;
    for (let i = 0; i < 120; i++) {
      const t = performance.now();
      sc.scrollTop += 233;
      sc.scrollLeft = (i % 20) * 90;
      await frame();
      await frame();
      times.push(performance.now() - t);
      maxCells = Math.max(maxCells, document.querySelectorAll(".c").length);
    }
    times.sort((a, b) => a - b);
    return { median: times[60]!, p95: times[114]!, maxCells, total: document.querySelectorAll("*").length };
  });
  check(perf.maxCells < 2500, "50k-cell sheet keeps only visible cells in the DOM", `max ${perf.maxCells} cell elements, ${perf.total} elements total`);
  check(perf.p95 < 40, "scrolling 50k cells stays smooth", `two frames per step: median ${perf.median.toFixed(1)} ms, p95 ${perf.p95.toFixed(1)} ms`);
} finally {
  for (const id of created) await api("DELETE", `/api/books/${id}`);
  await browser.close();
}
console.log(`${BROWSER}: ${failed ? `${failed} check(s) failed` : "all checks passed"}`);
process.exit(failed ? 1 : 0);
