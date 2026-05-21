import { normalizeOrigin } from "./config.js";

function createCorsHeaders(request, config) {
  const headers = new Headers();
  const allowedOrigin = getAllowedOrigin(request, config);
  const origin = request.headers.get("origin");
  if (allowedOrigin) {
    headers.set("Access-Control-Allow-Origin", allowedOrigin);
    headers.set("Vary", "Origin");
  } else if (!config.isProduction) {
    headers.set("Access-Control-Allow-Origin", "*");
  }
  headers.set("Access-Control-Allow-Methods", "GET, HEAD, POST, OPTIONS");
  headers.set("Access-Control-Allow-Headers", "Content-Type, X-App-Password");
  return {
    allowed: !origin || Boolean(allowedOrigin) || !config.isProduction,
    headers,
  };
}

function getAllowedOrigin(request, config) {
  const origin = request.headers.get("origin");
  if (!origin) {
    return "";
  }
  if (isSameHostOrigin(origin, request)) {
    return origin;
  }
  if (!config.isProduction) {
    return origin;
  }
  return config.allowedOrigins.has(normalizeOrigin(origin)) ? origin : "";
}

function isRequestOriginAllowed(request, config) {
  const origin = request.headers.get("origin");
  const normalizedOrigin = normalizeOrigin(origin);
  return !config.isProduction || !origin || isSameHostOrigin(origin, request) || config.allowedOrigins.has(normalizedOrigin);
}

function getRequestOrigin(request) {
  const forwardedHost = request.headers.get("x-forwarded-host");
  const host = forwardedHost || request.headers.get("host");
  if (!host) {
    return "";
  }
  const forwardedProto = String(request.headers.get("x-forwarded-proto") || "").split(",")[0].trim();
  const url = new URL(request.url);
  const proto = forwardedProto || url.protocol.replace(/:$/, "") || "http";
  return normalizeOrigin(`${proto}://${host}`);
}

function isSameHostOrigin(origin, request) {
  try {
    const originUrl = new URL(origin);
    const requestOrigin = getRequestOrigin(request);
    if (!requestOrigin) {
      return false;
    }
    const requestUrl = new URL(requestOrigin);
    return originUrl.host === requestUrl.host;
  } catch (error) {
    return false;
  }
}

export {
  createCorsHeaders,
  getAllowedOrigin,
  getRequestOrigin,
  isRequestOriginAllowed,
  isSameHostOrigin,
};
