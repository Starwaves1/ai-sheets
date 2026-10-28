// Screenshot a page: node dist/screenshot.js <url> <out.png> [--dark] [--size 1440x900] [--wait-ms 600] [--eval "js"]
import { launch } from "./browser.ts";

const [url, out, ...rest] = process.argv.slice(2);
if (!url || !out) {
  console.error("usage: screenshot <url> <out.png> [--dark] [--size WxH] [--wait-ms N] [--eval js]");
  process.exit(2);
}
const opt = (name: string) => {
  const i = rest.indexOf(name);
  return i >= 0 ? rest[i + 1] : undefined;
};
const [w, h] = (opt("--size") ?? "1440x900").split("x").map(Number);
const browser = await launch();
const page = await browser.newPage({ viewport: { width: w!, height: h! }, colorScheme: rest.includes("--dark") ? "dark" : "light", deviceScaleFactor: 1 });
await page.goto(url, { waitUntil: "networkidle" });
await page.evaluate(() => document.fonts.ready);
const js = opt("--eval");
if (js) await page.evaluate(js);
await page.waitForTimeout(Number(opt("--wait-ms") ?? 600));
await page.screenshot({ path: out });
await browser.close();
console.log(out);
