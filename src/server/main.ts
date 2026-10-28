// HTTP server for the API and the browser app. Started by run.ts.

import { createServer, type IncomingMessage, type ServerResponse } from "node:http";
import { readFile, stat } from "node:fs/promises";
import { extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { HttpError, makeApi } from "./api.ts";
import { Hub } from "./hub.ts";
import { Store } from "./store.ts";

const here = fileURLToPath(new URL(".", import.meta.url));

export function startServer(opts: { port: number; host: string; dataDir: string; publicDir?: string }) {
  const store = new Store(opts.dataDir === ":memory:" ? ":memory:" : join(opts.dataDir, "sheets.db"));
  const hub = new Hub(store);
  const api = makeApi(hub);
  const publicDir = opts.publicDir ?? join(here, "public");

  const TYPES: Record<string, string> = {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".woff2": "font/woff2",
    ".svg": "image/svg+xml",
    ".map": "application/json",
    ".png": "image/png",
    ".txt": "text/plain; charset=utf-8",
  };

  async function serveStatic(req: IncomingMessage, res: ServerResponse, path: string): Promise<void> {
    // App routes ("/", "/b/<id>") all load the single page.
    const file = path === "/" || path.startsWith("/b/") ? "index.html" : normalize(path).replace(/^([/\\])+/, "");
    const full = resolve(publicDir, file);
    if (!full.startsWith(resolve(publicDir))) throw new HttpError(403, "forbidden");
    let info;
    try {
      info = await stat(full);
    } catch {
      throw new HttpError(404, "not found");
    }
    const etag = `"${info.size.toString(36)}-${info.mtimeMs.toString(36)}"`;
    if (req.headers["if-none-match"] === etag) {
      res.writeHead(304);
      res.end();
      return;
    }
    const fonts = file.startsWith("fonts/");
    res.writeHead(200, {
      "content-type": TYPES[extname(full)] ?? "application/octet-stream",
      etag,
      "cache-control": fonts ? "public, max-age=31536000, immutable" : "no-cache",
    });
    res.end(await readFile(full));
  }

  const server = createServer(async (req, res) => {
    const url = new URL(req.url ?? "/", "http://local");
    try {
      if (url.pathname.startsWith("/api/")) await api(req, res, url);
      else if (req.method === "GET" || req.method === "HEAD") await serveStatic(req, res, url.pathname);
      else throw new HttpError(405, "method not allowed");
    } catch (e) {
      const status = e instanceof HttpError ? e.status : 500;
      if (status === 500) console.error(e);
      if (!res.headersSent) {
        res.writeHead(status, { "content-type": "application/json; charset=utf-8" });
        res.end(JSON.stringify({ error: (e as Error).message }));
      } else res.end();
    }
  });
  server.keepAliveTimeout = 65_000;
  server.listen(opts.port, opts.host);

  const close = () =>
    new Promise<void>((done) => {
      hub.flush();
      for (const b of hub.list()) b.close();
      server.close(() => {
        store.close();
        done();
      });
      server.closeAllConnections();
    });
  return { server, hub, close };
}
