import { LLM_CONFIG_ENDPOINT, NATURAL_FILL_PASSWORD_STORAGE_KEY } from "./constants.js";

const root = typeof window !== "undefined" ? window : globalThis;
const APP_PASSWORD_EXPIRES_AT_STORAGE_KEY = `${NATURAL_FILL_PASSWORD_STORAGE_KEY}:expiresAt`;
const APP_PASSWORD_TTL_MS = 7 * 24 * 60 * 60 * 1000;

function readAppPassword() {
  try {
    if (!root.localStorage) {
      return "";
    }
    const expiresAt = Number(root.localStorage.getItem(APP_PASSWORD_EXPIRES_AT_STORAGE_KEY) || 0);
    if (expiresAt && Date.now() > expiresAt) {
      clearAppPassword();
      return "";
    }
    return root.localStorage.getItem(NATURAL_FILL_PASSWORD_STORAGE_KEY) || "";
  } catch (error) {
    return "";
  }
}

function saveAppPassword(password) {
  try {
    if (root.localStorage) {
      root.localStorage.setItem(NATURAL_FILL_PASSWORD_STORAGE_KEY, String(password || ""));
      root.localStorage.setItem(APP_PASSWORD_EXPIRES_AT_STORAGE_KEY, String(Date.now() + APP_PASSWORD_TTL_MS));
    }
  } catch (error) {
    // Ignore private browsing/storage restrictions; the password can be re-entered.
  }
}

function clearAppPassword() {
  try {
    if (root.localStorage) {
      root.localStorage.removeItem(NATURAL_FILL_PASSWORD_STORAGE_KEY);
      root.localStorage.removeItem(APP_PASSWORD_EXPIRES_AT_STORAGE_KEY);
    }
  } catch (error) {
    // Ignore private browsing/storage restrictions.
  }
}

function buildAuthHeaders(password) {
  const headers = {
    "Content-Type": "application/json",
  };
  const appPassword = String(password || readAppPassword()).trim();
  if (appPassword) {
    headers["X-App-Password"] = appPassword;
  }
  return headers;
}

async function fetchWithAppPassword(url, options = {}) {
  return fetch(url, {
    ...options,
    headers: {
      ...buildAuthHeaders(),
      ...(options.headers || {}),
    },
  });
}

async function verifyAppPassword(password) {
  const response = await fetch(LLM_CONFIG_ENDPOINT, {
    method: "GET",
    headers: buildAuthHeaders(password),
  });
  const payload = await parseJson(response);
  if (!response.ok) {
    throw new Error((payload && payload.error) || "访问口令不正确");
  }
  saveAppPassword(password);
  return payload;
}

async function fetchLlmConfig() {
  const response = await fetchWithAppPassword(LLM_CONFIG_ENDPOINT, { method: "GET" });
  const payload = await parseJson(response);
  if (!response.ok) {
    if (response.status === 401) {
      clearAppPassword();
    }
    throw new Error((payload && payload.error) || `配置接口返回 ${response.status}`);
  }
  return payload;
}

async function saveLlmConfig(config) {
  const response = await fetchWithAppPassword(LLM_CONFIG_ENDPOINT, {
    method: "POST",
    body: JSON.stringify(config || {}),
  });
  const payload = await parseJson(response);
  if (!response.ok) {
    if (response.status === 401) {
      clearAppPassword();
    }
    throw new Error((payload && payload.error) || `配置接口返回 ${response.status}`);
  }
  return payload;
}

async function parseJson(response) {
  try {
    return await response.json();
  } catch (error) {
    return null;
  }
}

export {
  buildAuthHeaders,
  APP_PASSWORD_TTL_MS,
  clearAppPassword,
  fetchLlmConfig,
  fetchWithAppPassword,
  readAppPassword,
  saveAppPassword,
  saveLlmConfig,
  verifyAppPassword,
};
