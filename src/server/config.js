import { DEFAULT_DOUBAO_BASE_URL } from "./constants.js";

function createRuntimeConfig(env = {}) {
  const nodeEnv = env.NODE_ENV || "development";
  return {
    appPassword: env.APP_PASSWORD || "",
    allowedOrigins: parseAllowedOrigins(env.PUBLIC_ORIGIN, env.ALLOWED_ORIGINS),
    doubaoApiKey: env.DOUBAO_API_KEY || "",
    doubaoBaseUrl: (env.DOUBAO_BASE_URL || DEFAULT_DOUBAO_BASE_URL).replace(/\/+$/, ""),
    doubaoModel: env.DOUBAO_MODEL || "",
    doubaoTimeoutMs: Number(env.DOUBAO_TIMEOUT_MS || 25000),
    exportMaxBytes: Number(env.MAX_EXPORT_BYTES || 30 * 1024 * 1024),
    exportMaxItems: Number(env.EXPORT_MAX_ITEMS || 50),
    exportTtlMs: Number(env.EXPORT_TTL_MS || 10 * 60 * 1000),
    host: env.HOST || "0.0.0.0",
    isProduction: nodeEnv === "production",
    maxBodyBytes: 1024 * 1024,
    nodeEnv,
    port: Number(env.PORT || 4173),
    rateLimitMax: Number(env.RATE_LIMIT_MAX || 20),
    rateLimitWindowMs: Number(env.RATE_LIMIT_WINDOW_MS || 60 * 1000),
  };
}

function parseAllowedOrigins(publicOrigin, allowedOrigins) {
  return new Set([publicOrigin, ...(allowedOrigins || "").split(",")]
    .map(normalizeOrigin)
    .filter(Boolean));
}

function normalizeOrigin(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }
  try {
    const url = new URL(text);
    return url.origin;
  } catch (error) {
    return text.replace(/\/+$/, "");
  }
}

export { createRuntimeConfig, normalizeOrigin, parseAllowedOrigins };
