import http from "node:http";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { createRuntimeConfig } from "./src/server/config.js";
import { fetch as edgeFetch } from "./src/server/edge-handler.js";
import { EXPORT_PATH_PREFIX, SKILL_PATH } from "./src/server/constants.js";
import { toWebRequest, writeWebResponse } from "./src/server/node-adapter.js";
import { loadEnvFile } from "./src/server/node-env.js";
import { serveStatic } from "./src/server/node-static.js";

const ROOT_DIR = path.dirname(fileURLToPath(import.meta.url));
loadEnvFile(path.join(ROOT_DIR, ".env"));

const config = createRuntimeConfig(process.env);

const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url || "/", `http://${request.headers.host || `${config.host}:${config.port}`}`);

    if (shouldHandleWithEdge(request, url)) {
      const edgeRequest = toWebRequest(request, config.host, config.port);
      const edgeResponse = await edgeFetch(edgeRequest, process.env);
      await writeWebResponse(edgeResponse, response);
      return;
    }

    if (request.method === "GET" || request.method === "HEAD") {
      serveStatic(url.pathname, request, response, ROOT_DIR);
      return;
    }

    const edgeRequest = toWebRequest(request, config.host, config.port);
    const edgeResponse = await edgeFetch(edgeRequest, process.env);
    await writeWebResponse(edgeResponse, response);
  } catch (error) {
    const body = JSON.stringify({ error: error.message || "服务器内部错误" });
    response.writeHead(500, {
      "Content-Type": "application/json; charset=utf-8",
      "Content-Length": Buffer.byteLength(body),
    });
    response.end(body);
  }
});

server.listen(config.port, config.host, () => {
  const urls = getAccessUrls(config.port, config.host);
  console.log(`Excel 智能填表工具已启动：${urls.join("  ")}`);
});

function shouldHandleWithEdge(request, url) {
  if (request.method === "OPTIONS") {
    return true;
  }
  if (request.method === "POST" && url.pathname === SKILL_PATH) {
    return true;
  }
  if (request.method === "POST" && url.pathname === EXPORT_PATH_PREFIX) {
    return true;
  }
  return (request.method === "GET" || request.method === "HEAD") && url.pathname.startsWith(`${EXPORT_PATH_PREFIX}/`);
}

function getAccessUrls(port, host) {
  const urls = [`http://127.0.0.1:${port}`];
  if (host !== "127.0.0.1" && host !== "localhost") {
    Object.values(os.networkInterfaces()).forEach((items) => {
      (items || []).forEach((item) => {
        if (item.family === "IPv4" && !item.internal) {
          urls.push(`http://${item.address}:${port}`);
        }
      });
    });
  }
  return Array.from(new Set(urls));
}
