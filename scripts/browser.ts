// Shared by the e2e test and the screenshot tool: launch a headless browser via playwright-core.
// SHEETS_BROWSER=chromium (default) | firefox | webkit. Chromium uses $CHROME_PATH, else the newest
// Playwright headless shell in ~/.cache/ms-playwright; Firefox and WebKit use Playwright's own builds
// (`npx playwright-core install firefox webkit`).
import { existsSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { chromium, firefox, webkit, type Browser } from "playwright-core";

export const BROWSER = (process.env.SHEETS_BROWSER ?? "chromium") as "chromium" | "firefox" | "webkit";

export function chromePath(): string {
  if (process.env.CHROME_PATH) return process.env.CHROME_PATH;
  const base = join(homedir(), ".cache", "ms-playwright");
  const dirs = existsSync(base) ? readdirSync(base).filter((d) => d.startsWith("chromium_headless_shell-")).sort().reverse() : [];
  for (const d of dirs) {
    const p = join(base, d, "chrome-headless-shell-linux64", "chrome-headless-shell");
    if (existsSync(p)) return p;
  }
  throw new Error("No Chromium found: set CHROME_PATH or run `npx playwright install --only-shell chromium`");
}

export function launch(): Promise<Browser> {
  if (BROWSER === "firefox") return firefox.launch();
  if (BROWSER === "webkit") return webkit.launch();
  return chromium.launch({ executablePath: chromePath(), args: ["--no-sandbox"] });
}
