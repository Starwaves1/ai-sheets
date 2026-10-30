// Phone and tablet check in emulated devices (touch, mobile viewport):
//   screenshots of the home page and a workbook, frozen panes leave room for data,
//   tap to select, tap again to edit, CLI read-back, long-press menu, panning to the far columns.
// Usage: SHEETS_URL=http://192.168.1.5:8790 node dist/mobile.js [--book "DeepSWE runs"] [--shots dir]

import { execFile } from "node:child_process";
import { join } from "node:path";
import { promisify } from "node:util";
import { devices, type Page } from "playwright-core";
import { launch } from "./browser.ts";

const BASE = (process.env.SHEETS_URL ?? "http://192.168.1.5:8790").replace(/\/$/, "");
const arg = (name: string) => (process.argv.includes(name) ? process.argv[process.argv.indexOf(name) + 1] : undefined);
const shots = arg("--shots");
const showcase = arg("--book") ?? "DeepSWE runs";
const run = promisify(execFile);
const CLI = join(import.meta.dirname, "..", "bin", "sheets");
const cli = async (...args: string[]) => (await run(CLI, args, { env: { ...process.env, SHEETS_URL: BASE, SHEETS_ACTOR: "mobile-check" } })).stdout;

let failed = 0;
const check = (ok: boolean, what: string, detail = "") => {
  console.log(`${ok ? "PASS" : "FAIL"}  ${what}${detail ? `  (${detail})` : ""}`);
  if (!ok) failed++;
};

const DEVICES: Array<[string, (typeof devices)[string]]> = [
  ["iphone-15", devices["iPhone 15"]!],
  ["iphone-15-landscape", devices["iPhone 15 landscape"]!],
  ["android-360", devices["Galaxy S8"]!],
  ["ipad", devices["iPad (gen 7)"]!],
];

/** Center of the on-screen cell showing `text`, or null. */
const cellAt = (page: Page, text: string) =>
  page.evaluate((t) => {
    const el = [...document.querySelectorAll(".c")].find((e) => e.textContent === t);
    if (!el) return null;
    const r = el.getBoundingClientRect();
    return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
  }, text);

const books = (await (await fetch(`${BASE}/api/books`)).json()) as Array<{ id: string; name: string }>;
const show = books.find((b) => b.name === showcase);
const scratch = (await (await fetch(`${BASE}/api/books`, {
  method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ name: `mobile check ${Date.now()}` }),
})).json()) as { id: string };
await fetch(`${BASE}/api/books/${scratch.id}/range?ref=A1`, {
  method: "PUT", headers: { "content-type": "application/json" }, body: JSON.stringify({ values: [["fruit", "qty"], ["fig", 3], ["kiwi", 5]] }),
});

const browser = await launch();
try {
  for (const [name, dev] of DEVICES) {
    const ctx = await browser.newContext({ ...dev, colorScheme: name === "iphone-15" ? "dark" : "light" });
    const page = await ctx.newPage();

    await page.goto(`${BASE}/`, { waitUntil: "networkidle" });
    const homeOverflow = await page.evaluate(() => document.documentElement.scrollWidth - window.innerWidth);
    check(homeOverflow <= 0, `${name}: home page fits the width`, `${homeOverflow}px overflow`);
    if (shots) await page.screenshot({ path: join(shots, `mobile-home-${name}.png`) });

    if (show) {
      await page.goto(`${BASE}/b/${show.id}`, { waitUntil: "networkidle" });
      await page.evaluate(() => document.fonts.ready);
      const m = await page.evaluate(() => {
        const grid = document.querySelector(".grid")!.getBoundingClientRect();
        const lines = [...document.querySelectorAll<HTMLElement>(".frozen-line")].filter((e) => !e.hidden);
        const v = lines.find((e) => e.classList.contains("v"))?.getBoundingClientRect();
        const h = lines.find((e) => e.classList.contains("h"))?.getBoundingClientRect();
        const tb = document.querySelector(".toolbar")!.getBoundingClientRect();
        return { gw: grid.width, gh: grid.height, fw: v ? v.left - grid.left : 0, fh: h ? h.top - grid.top : 0, toolbarH: tb.height };
      });
      check(m.fw < m.gw * 0.6 && m.fh < m.gh * 0.45, `${name}: frozen panes leave room for data`,
        `frozen ${Math.round(m.fw)}×${Math.round(m.fh)} of grid ${Math.round(m.gw)}×${Math.round(m.gh)}`);
      check(m.toolbarH <= 52, `${name}: toolbar is one row`, `${Math.round(m.toolbarH)}px tall`);
      if (shots) await page.screenshot({ path: join(shots, `mobile-book-${name}.png`) });

      // Pan to the far right with one-finger swipes (raw touch events, as a phone sends them).
      const g = await page.locator(".grid").boundingBox();
      for (let k = 0; k < 12; k++) await swipe(page, g!.x + g!.width * 0.85, g!.y + g!.height * 0.7, -g!.width * 0.6, 0);
      await page.waitForTimeout(400);
      const far = await page.evaluate(() => {
        const sc = document.querySelector(".scroller") as HTMLElement;
        const heads = [...document.querySelector(".colhead")!.querySelectorAll(".h")].map((e) => e.textContent ?? ""); // scrolling header pane
        return { left: sc.scrollLeft, max: sc.scrollWidth - sc.clientWidth, last: heads[heads.length - 1] };
      });
      check(far.left >= far.max - 2, `${name}: one-finger pan reaches the far columns`, `scrollLeft ${Math.round(far.left)}/${far.max}, last header ${far.last}`);
      if (shots && name === "iphone-15") await page.screenshot({ path: join(shots, `mobile-book-${name}-panned.png`) });
    }

    // Tap a cell, tap it again to edit, type, and read it back through the CLI.
    await page.goto(`${BASE}/b/${scratch.id}`, { waitUntil: "networkidle" });
    await page.waitForSelector(".live.live");
    const fig = await cellAt(page, "fig");
    await page.touchscreen.tap(fig!.x + 90, fig!.y); // a cell right of "fig"; the address box says which
    const target = await page.evaluate(() => (document.querySelector(".addr") as HTMLInputElement).value);
    await page.touchscreen.tap(fig!.x + 90, fig!.y);
    const editing = await page.evaluate(() => !(document.querySelector(".editor") as HTMLElement).hidden && document.activeElement?.className === "editor");
    check(editing, `${name}: tap selects (${target}), tapping it again opens the editor`);
    const word = `${name}-ok`;
    await page.keyboard.press("Control+a"); // the editor opens with the current text; replace it
    await page.keyboard.type(word);
    await page.keyboard.press("Enter");
    await page.waitForTimeout(400);
    const back = (await cli("get", scratch.id, target, "--csv")).trim();
    check(back === word, `${name}: edit reads back through the CLI`, JSON.stringify(back));

    // Long press opens the context menu.
    const kiwi = await cellAt(page, "kiwi");
    await cdpLongPress(page, kiwi!.x, kiwi!.y);
    const menuOpen = await page.evaluate(() => !!document.querySelector(".popup.menu"));
    check(menuOpen, `${name}: long press opens the row/column menu`);
    if (shots && name === "iphone-15") await page.screenshot({ path: join(shots, `mobile-menu-${name}.png`) });
    await page.keyboard.press("Escape");

    // Flags on touch: tapping a flagged cell shows its note; the long-press menu offers "Flag for Claude…".
    await cli("flag", scratch.id, "B3", "is 5 kiwis right?");
    await page.waitForTimeout(400);
    const five = await cellAt(page, "5");
    await page.touchscreen.tap(five!.x, five!.y);
    await page.waitForTimeout(200);
    const card = await page.evaluate(() => {
      const el = document.querySelector(".flag-card") as HTMLElement | null;
      if (!el) return null;
      const r = el.getBoundingClientRect();
      return { note: el.querySelector(".note")?.textContent, inside: r.left >= 0 && r.right <= window.innerWidth };
    });
    check(card?.note === "is 5 kiwis right?" && card.inside, `${name}: tapping a flagged cell shows its note on screen`, JSON.stringify(card));
    if (shots && name === "iphone-15") await page.screenshot({ path: join(shots, `mobile-flag-${name}.png`) });
    await page.waitForTimeout(1000);
    await cdpLongPress(page, five!.x, five!.y);
    const flagItem = await page.locator(".menu .item", { hasText: "Edit flag…" }).count();
    check(flagItem === 1, `${name}: long-press menu offers the flag`);
    await page.keyboard.press("Escape");
    await cli("flag", scratch.id, "B3", "--remove");

    // Moving rows on touch: long-press a row header, "Move row down".
    if (name === "iphone-15") {
      await page.waitForTimeout(1000);
      const head = await page.evaluate(() => {
        const el = [...document.querySelectorAll(".rowhead .h")].find((e) => e.textContent === "2")!.getBoundingClientRect();
        return { x: el.x + el.width / 2, y: el.y + el.height / 2 };
      });
      await cdpLongPress(page, head.x, head.y);
      const down = page.locator(".menu .item", { hasText: "Move row down" });
      const offered = await down.count();
      if (offered) await down.tap();
      await page.waitForTimeout(400);
      const order = (await cli("get", scratch.id, "A1:A3", "--csv")).trim().split("\n").join(",");
      check(offered === 1 && order === "fruit,kiwi,fig", `${name}: long-press a row header, "Move row down"`, order);
      await cli("move", scratch.id, "3:3", "--to", "2");
    }

    // Autofill on touch: select C2:C6 with 1, 2 at the top, long-press inside, choose "Autofill down".
    if (name === "iphone-15" || name === "ipad") {
      await page.waitForTimeout(1000); // the menu ignores a contextmenu within 1 s of a long press
      await cli("set", scratch.id, "C2", "1");
      await cli("set", scratch.id, "C3", "2");
      await page.waitForTimeout(300);
      await page.touchscreen.tap(...(Object.values((await page.locator(".addr").boundingBox())!).slice(0, 2).map((v) => v + 5) as [number, number]));
      await page.keyboard.press("Control+a");
      await page.keyboard.type("C2:C6");
      await page.keyboard.press("Enter");
      const two = await cellAt(page, "2");
      await cdpLongPress(page, two!.x, two!.y);
      const item = page.locator(".menu .item", { hasText: "Autofill down" });
      const offered = await item.count();
      if (offered) await item.tap();
      await page.waitForTimeout(400);
      const filled = await cli("get", scratch.id, "C2:C6", "--csv");
      check(offered === 1 && filled === "1\n2\n3\n4\n5\n", `${name}: long-press "Autofill down" fills the selection's empty cells`, JSON.stringify(filled));
      await cli("clear", scratch.id, "C2:C6");
    }
    await ctx.close();
  }
} finally {
  await fetch(`${BASE}/api/books/${scratch.id}`, { method: "DELETE" });
  await browser.close();
}

async function swipe(page: Page, x: number, y: number, dx: number, dy: number): Promise<void> {
  const cdp = await page.context().newCDPSession(page);
  const at = (f: number) => [{ x: Math.round(x + dx * f), y: Math.round(y + dy * f) }];
  await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: at(0) });
  for (let k = 1; k <= 8; k++) await cdp.send("Input.dispatchTouchEvent", { type: "touchMove", touchPoints: at(k / 8) });
  await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
  await cdp.detach();
}

async function cdpLongPress(page: Page, x: number, y: number): Promise<void> {
  const cdp = await page.context().newCDPSession(page);
  const pt = [{ x: Math.round(x), y: Math.round(y) }];
  await cdp.send("Input.dispatchTouchEvent", { type: "touchStart", touchPoints: pt });
  await page.waitForTimeout(650);
  await cdp.send("Input.dispatchTouchEvent", { type: "touchEnd", touchPoints: [] });
  await page.waitForTimeout(150);
  await cdp.detach();
}

console.log(failed ? `${failed} check(s) failed` : "all checks passed");
process.exit(failed ? 1 : 0);
