// Shared by the e2e test and the screenshot tool: launch headless Chromium via playwright-core.
// Uses $CHROME_PATH, else the newest Playwright headless shell in ~/.cache/ms-playwright.
import { existsSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { chromium, type Browser } from "playwright-core";

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
  return chromium.launch({ executablePath: chromePath(), args: ["--no-sandbox"] });
}
