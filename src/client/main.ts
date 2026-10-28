// Router: "/" lists workbooks, "/b/<id>" opens one.

import { openBook } from "./book.ts";
import { openHome } from "./home.ts";

const root = document.getElementById("app")!;
const m = /^\/b\/([^/]+)/.exec(location.pathname);
if (m) openBook(root, decodeURIComponent(m[1]!));
else openHome(root);
