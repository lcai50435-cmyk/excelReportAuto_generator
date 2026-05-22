import { NATURAL_FILL_ENDPOINT } from "./constants.js";
import { fetchWithAppPassword } from "./auth-config.js";

const root = typeof window !== "undefined" ? window : globalThis;

async function requestNaturalFillExtraction(text, fields, calculationRules) {
  const body = JSON.stringify({ text, fields, calculationRules: calculationRules || [] });
  let response = await fetchNaturalFillEndpoint(body);

  if (response.status === 404 && shouldTryLocalNaturalFillEndpoint()) {
    response = await fetchLocalNaturalFillEndpoint(body);
  }

  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok) {
    throw new Error((payload && payload.error) || getNaturalFillHttpError(response.status));
  }

  if (!payload || !Array.isArray(payload.rows)) {
    throw new Error("智能填行服务返回格式不正确");
  }

  return payload;
}

async function fetchNaturalFillEndpoint(body) {
  try {
    return await fetchWithAppPassword(NATURAL_FILL_ENDPOINT, {
      method: "POST",
      body,
    });
  } catch (error) {
    return await fetchLocalNaturalFillEndpoint(body, error);
  }
}

async function fetchLocalNaturalFillEndpoint(body, originalError) {
  if (!shouldTryLocalNaturalFillEndpoint()) {
    throw new Error(getNaturalFillConnectionError(originalError));
  }

  try {
    return await fetchWithAppPassword(`${getNaturalFillFallbackOrigin()}${NATURAL_FILL_ENDPOINT}`, {
      method: "POST",
      body,
    });
  } catch (error) {
    throw new Error(getNaturalFillConnectionError(originalError || error));
  }
}

function shouldTryLocalNaturalFillEndpoint() {
  if (!root.location) {
    return false;
  }

  if (root.location.protocol === "file:") {
    return true;
  }

  const hostname = root.location.hostname;
  return isLocalNaturalFillHostname(hostname) && root.location.port !== "4173";
}

function getNaturalFillFallbackOrigin() {
  if (!root.location || root.location.protocol === "file:") {
    return "http://127.0.0.1:4173";
  }

  const hostname = root.location.hostname || "127.0.0.1";
  return `http://${hostname === "::1" ? "[::1]" : hostname}:4173`;
}

function isLocalNaturalFillHostname(hostname) {
  return hostname === "127.0.0.1" || hostname === "localhost" || hostname === "::1";
}

function getNaturalFillConnectionError(error) {
  if (root.location && root.location.protocol === "file:") {
    return "请先启动本地服务，并通过服务启动后显示的地址打开页面";
  }

  return `无法连接智能填行服务，请确认服务端已运行，并用服务启动后显示的地址打开页面（${error && error.message || "网络错误"}）`;
}

function getNaturalFillHttpError(status) {
  if (status === 504 || status === 524) {
    return "智能填行服务响应超时，请稍后重试；删除字段、取消规则等简单指令请确认线上后端已更新";
  }
  if (status === 404) {
    return "当前页面不是由智能填行服务打开，请使用服务启动后显示的地址访问后再智能填行";
  }

  return `本地智能填行服务返回 ${status}`;
}

export { requestNaturalFillExtraction };
