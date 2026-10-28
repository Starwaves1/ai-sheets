// Process entry point.
//   SHEETS_PORT (8790)  SHEETS_HOST (0.0.0.0)  SHEETS_DATA (<repo>/data)

import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { startServer } from "./main.ts";

const here = fileURLToPath(new URL(".", import.meta.url));
const port = Number(process.env.SHEETS_PORT ?? 8790);
const host = process.env.SHEETS_HOST ?? "0.0.0.0";
const dataDir = process.env.SHEETS_DATA ?? join(here, "..", "data");
const { server, close } = startServer({ port, host, dataDir });
server.on("listening", () => console.log(`sheets listening on http://${host}:${port} (data: ${dataDir})`));
const stop = () => close().then(() => process.exit(0));
process.on("SIGTERM", stop);
process.on("SIGINT", stop);
