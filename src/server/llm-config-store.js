import fs from "node:fs/promises";
import path from "node:path";

import { DEFAULT_DOUBAO_BASE_URL } from "./constants.js";

const DEFAULT_LLM_TIMEOUT_MS = 60000;
const MIN_LLM_TIMEOUT_MS = 5000;
const MAX_LLM_TIMEOUT_MS = 180000;
const DEFAULT_CONFIG_PATH = path.join(process.cwd(), "data", "llm-config.json");

function getConfigPath() {
  return process.env.LLM_CONFIG_PATH || DEFAULT_CONFIG_PATH;
}

async function readStoredLlmConfig() {
  try {
    const text = await fs.readFile(getConfigPath(), "utf8");
    const parsed = JSON.parse(text);
    return normalizeStoredLlmConfig(parsed);
  } catch (error) {
    if (error && error.code === "ENOENT") {
      return {};
    }
    throw error;
  }
}

async function writeStoredLlmConfig(config) {
  const normalized = normalizeStoredLlmConfig(config);
  const filePath = getConfigPath();
  const dir = path.dirname(filePath);
  await fs.mkdir(dir, { recursive: true });
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await fs.writeFile(tempPath, `${JSON.stringify(normalized, null, 2)}\n`, { mode: 0o600 });
  await fs.rename(tempPath, filePath);
  try {
    await fs.chmod(filePath, 0o600);
  } catch (error) {
    // Some filesystems do not support chmod. The file is still usable.
  }
  return normalized;
}

async function getEffectiveLlmConfig(runtimeConfig = {}) {
  const stored = await readStoredLlmConfig();
  return normalizeStoredLlmConfig({
    baseUrl: stored.baseUrl || runtimeConfig.doubaoBaseUrl || DEFAULT_DOUBAO_BASE_URL,
    model: stored.model || runtimeConfig.doubaoModel || "",
    apiKey: stored.apiKey || runtimeConfig.doubaoApiKey || "",
    timeoutMs: stored.timeoutMs || runtimeConfig.doubaoTimeoutMs || DEFAULT_LLM_TIMEOUT_MS,
  });
}

async function getPublicLlmConfig(runtimeConfig = {}) {
  const effective = await getEffectiveLlmConfig(runtimeConfig);
  return {
    baseUrl: effective.baseUrl,
    model: effective.model,
    timeoutMs: effective.timeoutMs,
    hasApiKey: Boolean(effective.apiKey),
  };
}

async function saveLlmConfigUpdate(update, runtimeConfig = {}) {
  const current = await getEffectiveLlmConfig(runtimeConfig);
  const normalizedUpdate = normalizeLlmConfigUpdate(update);
  const next = normalizeStoredLlmConfig({
    baseUrl: normalizedUpdate.baseUrl || current.baseUrl || DEFAULT_DOUBAO_BASE_URL,
    model: normalizedUpdate.model || current.model || "",
    apiKey: normalizedUpdate.clearApiKey ? "" : normalizedUpdate.apiKey || current.apiKey || "",
    timeoutMs: normalizedUpdate.timeoutMs || current.timeoutMs || DEFAULT_LLM_TIMEOUT_MS,
  });

  validateCompleteLlmConfig(next);
  await writeStoredLlmConfig(next);
  return {
    baseUrl: next.baseUrl,
    model: next.model,
    timeoutMs: next.timeoutMs,
    hasApiKey: Boolean(next.apiKey),
  };
}

function normalizeStoredLlmConfig(value) {
  const source = value && typeof value === "object" ? value : {};
  return {
    baseUrl: normalizeBaseUrl(source.baseUrl || source.doubaoBaseUrl || ""),
    model: String(source.model || source.doubaoModel || "").trim(),
    apiKey: String(source.apiKey || source.doubaoApiKey || "").trim(),
    timeoutMs: normalizeTimeoutMs(source.timeoutMs || source.doubaoTimeoutMs || DEFAULT_LLM_TIMEOUT_MS),
  };
}

function normalizeLlmConfigUpdate(value) {
  const source = value && typeof value === "object" ? value : {};
  const hasApiKey = Object.prototype.hasOwnProperty.call(source, "apiKey");
  return {
    baseUrl: normalizeBaseUrl(source.baseUrl || ""),
    model: String(source.model || "").trim(),
    apiKey: hasApiKey ? String(source.apiKey || "").trim() : "",
    clearApiKey: Boolean(source.clearApiKey),
    timeoutMs: normalizeTimeoutMs(source.timeoutMs || 0, 0),
  };
}

function normalizeBaseUrl(value) {
  let text = String(value || "").trim().replace(/\/+$/, "");
  if (text === "https://cloud.siliconflow.cn") {
    text = "https://api.siliconflow.cn/v1";
  }
  if (!text) {
    return "";
  }
  try {
    const url = new URL(text);
    if (!/^https?:$/.test(url.protocol)) {
      return "";
    }
    const normalized = url.toString().replace(/\/+$/, "");
    if (normalized === "https://cloud.siliconflow.cn") {
      return "https://api.siliconflow.cn/v1";
    }
    return normalized;
  } catch (error) {
    return "";
  }
}

function normalizeTimeoutMs(value, fallback = DEFAULT_LLM_TIMEOUT_MS) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) {
    return fallback;
  }
  return Math.min(Math.max(Math.round(number), MIN_LLM_TIMEOUT_MS), MAX_LLM_TIMEOUT_MS);
}

function validateCompleteLlmConfig(config) {
  if (!config.baseUrl) {
    throw new Error("请填写有效的大模型 Base URL");
  }
  if (!config.model) {
    throw new Error("请填写大模型名称");
  }
  if (!config.apiKey) {
    throw new Error("请填写大模型 API Key");
  }
}

export {
  DEFAULT_LLM_TIMEOUT_MS,
  getEffectiveLlmConfig,
  getPublicLlmConfig,
  readStoredLlmConfig,
  saveLlmConfigUpdate,
  writeStoredLlmConfig,
};
