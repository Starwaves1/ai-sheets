// End-to-end check against a running server, in a real (headless) browser:
//   1. an agent edit through the CLI appears in an open tab without reload
//   2. an edit typed in the browser is readable through the CLI (and undo works)
//   3. the page is idle when nothing changes (no rAF, timers, layouts or animations)
//   4. a 50k-cell sheet scrolls with only the visible cells in the DOM
// Usage: SHEETS_URL=http://192.168.1.5:8790 node dist/e2e.js [--shots dir]

import { execFile } from "node:child_process";
import { join } from "node:path";
import { promisify } from "node:util";
import { launch } from "./browser.ts";

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

  // 3. idle cost
  await page.waitForTimeout(2200); // let the highlight fade finish
  const cdp = await ctx.newCDPSession(page);
  await cdp.send("Performance.enable");
  const metric = async () => Object.fromEntries(((await cdp.send("Performance.getMetrics")).metrics as Array<{ name: string; value: number }>).map((m) => [m.name, m.value]));
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
  const other = await ctx.newPage();
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
  await other.close();

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
console.log(failed ? `${failed} check(s) failed` : "all checks passed");
process.exit(failed ? 1 : 0);
