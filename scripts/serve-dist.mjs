// Minimal zero-dependency static file server for the built add-in (dist/).
// Used by the Playwright e2e suite. Not for production serving.
import http from "node:http";
import { readFile, stat } from "node:fs/promises";
import path from "node:path";

const ROOT = path.resolve("dist");
const PORT = Number(process.env.PORT || 8214);

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".map": "application/json; charset=utf-8",
};

const server = http.createServer(async (req, res) => {
  try {
    let urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
    if (urlPath === "/") {
      urlPath = "/taskpane.html";
    }
    const filePath = path.join(ROOT, urlPath);
    // Reject path traversal: the resolved path must be ROOT itself or live
    // directly under it (avoids the startsWith-prefix-confusion pitfall).
    if (filePath !== ROOT && !filePath.startsWith(ROOT + path.sep)) {
      res.statusCode = 403;
      res.end("forbidden");
      return;
    }
    const info = await stat(filePath);
    if (info.isDirectory()) {
      res.statusCode = 403;
      res.end("directory");
      return;
    }
    const data = await readFile(filePath);
    res.setHeader("Content-Type", TYPES[path.extname(filePath)] || "application/octet-stream");
    res.end(data);
  } catch {
    res.statusCode = 404;
    res.end("not found");
  }
});

server.listen(PORT, () => {
  console.log(`[serve-dist] http://localhost:${PORT} -> ${ROOT}`);
});
