// One build step: bundle server, CLI, browser app (and tests) with esbuild.
import { build, context } from "esbuild";
import { cpSync, mkdirSync, readdirSync, rmSync } from "node:fs";

const args = new Set(process.argv.slice(2));
const node = { platform: "node", format: "esm", target: "node22", bundle: true, sourcemap: "linked", logLevel: "warning",
  banner: { js: "import { createRequire as __cr } from 'node:module'; const require = __cr(import.meta.url);" } };

const targets = [
  { ...node, entryPoints: ["src/server/run.ts"], outfile: "dist/server.js" },
  { ...node, entryPoints: ["src/cli/main.ts"], outfile: "dist/cli.js" },
  { ...node, entryPoints: ["scripts/e2e.ts"], outfile: "dist/e2e.js", external: ["playwright-core"] },
  { ...node, entryPoints: ["scripts/screenshot.ts"], outfile: "dist/screenshot.js", external: ["playwright-core"] },
  { entryPoints: ["src/client/main.ts"], outfile: "dist/public/app.js", bundle: true, format: "esm", platform: "browser",
    target: "es2022", minify: !args.has("--watch"), sourcemap: "linked", logLevel: "warning" },
];
if (args.has("--tests")) {
  rmSync("dist/test", { recursive: true, force: true });
  const tests = readdirSync("test").filter((f) => f.endsWith(".test.ts")).map((f) => `test/${f}`);
  targets.push({ ...node, entryPoints: tests, outdir: "dist/test", sourcemap: "inline" });
}

mkdirSync("dist/public", { recursive: true });
cpSync("public", "dist/public", { recursive: true });

if (args.has("--watch")) {
  for (const t of targets) await (await context(t)).watch();
  console.log("watching…");
} else {
  await Promise.all(targets.map((t) => build(t)));
}
