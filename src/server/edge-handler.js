import { createRuntimeConfig } from "./config.js";
import { EXPORT_PATH_PREFIX, LLM_CONFIG_API_PATH, SKILL_PATH } from "./constants.js";
import { createCorsHeaders, isRequestOriginAllowed } from "./cors.js";
import { createExportStore } from "./export-store.js";
import { handleExportCreate, handleExportDownload, isExportDownloadPath } from "./exports.js";
import { handleLlmConfigGet, handleLlmConfigSave } from "./llm-config.js";
import { handleNaturalFillExtract } from "./natural-fill.js";
import { createRateLimiter } from "./rate-limit.js";
import { emptyResponse, jsonResponse } from "./response.js";

const defaultRateLimiter = createRateLimiter();
const defaultExportStore = createExportStore();

function createFetchHandler(options = {}) {
  const rateLimiter = options.rateLimiter || defaultRateLimiter;
  const exportStore = options.exportStore || defaultExportStore;

  return async function fetchHandler(request, env = {}) {
    const config = createRuntimeConfig(env);
    const context = { config, exportStore, rateLimiter };
    const url = new URL(request.url);
    const cors = createCorsHeaders(request, config);

    try {
      if (request.method === "OPTIONS") {
        return withCors(emptyResponse(cors.allowed ? 204 : 403), cors.headers);
      }

      if (request.method === "POST" && url.pathname === SKILL_PATH) {
        if (!cors.allowed || !isRequestOriginAllowed(request, config)) {
          return withCors(jsonResponse(403, { error: "Origin is not allowed" }), cors.headers);
        }
        return withCors(await handleNaturalFillExtract(request, context), cors.headers);
      }

      if ((request.method === "GET" || request.method === "POST") && url.pathname === LLM_CONFIG_API_PATH) {
        if (!cors.allowed || !isRequestOriginAllowed(request, config)) {
          return withCors(jsonResponse(403, { error: "Origin is not allowed" }), cors.headers);
        }
        const response = request.method === "GET"
          ? await handleLlmConfigGet(request, context)
          : await handleLlmConfigSave(request, context);
        return withCors(response, cors.headers);
      }

      if (request.method === "POST" && url.pathname === EXPORT_PATH_PREFIX) {
        if (!cors.allowed || !isRequestOriginAllowed(request, config)) {
          return withCors(jsonResponse(403, { error: "Origin is not allowed" }), cors.headers);
        }
        return withCors(await handleExportCreate(request, url, context), cors.headers);
      }

      if ((request.method === "GET" || request.method === "HEAD") && isExportDownloadPath(url.pathname)) {
        return withCors(handleExportDownload(request, url, context), cors.headers);
      }

      if (request.method === "GET" || request.method === "HEAD") {
        return withCors(jsonResponse(404, { error: "Not found" }), cors.headers);
      }

      return withCors(jsonResponse(405, { error: "请求方式不支持" }), cors.headers);
    } catch (error) {
      return withCors(jsonResponse(500, { error: error.message || "服务器内部错误" }), cors.headers);
    }
  };
}

const fetch = createFetchHandler();

function withCors(response, corsHeaders) {
  const headers = new Headers(response.headers);
  corsHeaders.forEach((value, key) => {
    headers.set(key, value);
  });
  return new Response(response.body, {
    status: response.status,
    statusText: response.statusText,
    headers,
  });
}

export { createFetchHandler, fetch };
