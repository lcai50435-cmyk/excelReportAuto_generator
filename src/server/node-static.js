import fs from "node:fs";
import path from "node:path";
import { MIME_TYPES } from "./constants.js";

function serveStatic(pathname, request, response) {
  const rootDir = arguments.length > 3 ? arguments[3] : process.cwd();
  const safePathname = pathname === "/" ? "/index.html" : pathname;
  const decoded = decodeURIComponent(safePathname);
  const filePath = path.normalize(path.join(rootDir, decoded));

  if (!filePath.startsWith(rootDir) || filePath.includes(`${path.sep}.env`)) {
    response.writeHead(403);
    response.end("Forbidden");
    return;
  }

  fs.stat(filePath, (statError, stats) => {
    if (statError || !stats.isFile()) {
      response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Not found");
      return;
    }

    const headers = {
      "Content-Type": MIME_TYPES[path.extname(filePath).toLowerCase()] || "application/octet-stream",
      "Content-Length": stats.size,
    };
    if ([".html", ".js", ".css"].includes(path.extname(filePath).toLowerCase())) {
      headers["Cache-Control"] = "no-store";
    }
    response.writeHead(200, headers);
    if (request.method === "HEAD") {
      response.end();
      return;
    }
    fs.createReadStream(filePath).pipe(response);
  });
}

export { serveStatic };
