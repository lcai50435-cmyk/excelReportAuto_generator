"use strict";

const http = require("node:http");
const crypto = require("node:crypto");
const fs = require("node:fs");
const path = require("node:path");
const { URL } = require("node:url");
const os = require("node:os");

const ROOT_DIR = __dirname;
loadEnvFile(path.join(ROOT_DIR, ".env"));

const NODE_ENV = process.env.NODE_ENV || "development";
const IS_PRODUCTION = NODE_ENV === "production";
const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const MAX_BODY_BYTES = 1024 * 1024;
const DOUBAO_TIMEOUT_MS = Number(process.env.DOUBAO_TIMEOUT_MS || 25000);
const MAX_EXPORT_BYTES = Number(process.env.MAX_EXPORT_BYTES || 30 * 1024 * 1024);
const EXPORT_TTL_MS = Number(process.env.EXPORT_TTL_MS || 10 * 60 * 1000);
const EXPORT_MAX_ITEMS = Number(process.env.EXPORT_MAX_ITEMS || 50);
const DEFAULT_DOUBAO_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3";
const SKILL_PATH = "/api/agent-skills/doubao-excel-natural-fill/extract";
const EXPORT_PATH_PREFIX = "/api/exports/xlsx";
const APP_PASSWORD = process.env.APP_PASSWORD || "";
const RATE_LIMIT_WINDOW_MS = Number(process.env.RATE_LIMIT_WINDOW_MS || 60 * 1000);
const RATE_LIMIT_MAX = Number(process.env.RATE_LIMIT_MAX || 20);
const ALLOWED_ORIGINS = parseAllowedOrigins(process.env.PUBLIC_ORIGIN, process.env.ALLOWED_ORIGINS);
const naturalFillRateLimits = new Map();
const exportFiles = new Map();

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".log": "text/plain; charset=utf-8",
};

const server = http.createServer(async (request, response) => {
  try {
    const url = new URL(request.url || "/", `http://${request.headers.host || `${HOST}:${PORT}`}`);
    const corsAllowed = setCorsHeaders(request, response);

    if (request.method === "OPTIONS") {
      response.writeHead(corsAllowed ? 204 : 403);
      response.end();
      return;
    }

    if (request.method === "POST" && url.pathname === SKILL_PATH) {
      if (!corsAllowed || !isRequestOriginAllowed(request)) {
        sendJson(response, 403, { error: "Origin is not allowed" });
        return;
      }
      await handleNaturalFillExtract(request, response);
      return;
    }

    if (request.method === "POST" && url.pathname === EXPORT_PATH_PREFIX) {
      if (!corsAllowed || !isRequestOriginAllowed(request)) {
        sendJson(response, 403, { error: "Origin is not allowed" });
        return;
      }
      await handleExportCreate(request, response, url);
      return;
    }

    if ((request.method === "GET" || request.method === "HEAD") && url.pathname.startsWith(`${EXPORT_PATH_PREFIX}/`)) {
      handleExportDownload(request, response, url);
      return;
    }

    if (request.method === "GET" || request.method === "HEAD") {
      serveStatic(url.pathname, request, response);
      return;
    }

    sendJson(response, 405, { error: "不支持的请求方法" });
  } catch (error) {
    sendJson(response, 500, { error: error.message || "服务异常" });
  }
});

server.listen(PORT, HOST, () => {
  const urls = getAccessUrls(PORT, HOST);
  console.log(`Excel 自动生成工具已启动：${urls.join("  ")}`);
});

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

async function handleNaturalFillExtract(request, response) {
  const auth = validateAppPassword(request);
  if (!auth.ok) {
    sendJson(response, auth.statusCode, { error: auth.error });
    return;
  }

  const rateLimit = checkNaturalFillRateLimit(request);
  if (!rateLimit.ok) {
    response.setHeader("Retry-After", String(Math.ceil(rateLimit.retryAfterMs / 1000)));
    sendJson(response, 429, { error: "Too many AI requests. Please try again later." });
    return;
  }

  const apiKey = process.env.DOUBAO_API_KEY;
  const model = process.env.DOUBAO_MODEL;
  const baseUrl = (process.env.DOUBAO_BASE_URL || DEFAULT_DOUBAO_BASE_URL).replace(/\/+$/, "");

  const payload = await readJsonBody(request);
  const validation = validateExtractPayload(payload);
  if (!validation.ok) {
    sendJson(response, 400, { error: validation.error });
    return;
  }

  if (!apiKey) {
    sendJson(response, 500, { error: "请先设置 DOUBAO_API_KEY 环境变量" });
    return;
  }

  if (!model) {
    sendJson(response, 500, { error: "请先设置 DOUBAO_MODEL 环境变量" });
    return;
  }

  const fields = normalizeRequestFields(payload.fields);
  if (!fields.length) {
    sendJson(response, 400, { error: "缺少可填充字段" });
    return;
  }

  const calculationRules = normalizeCalculationRules(payload.calculationRules, fields);
  const localStructuralResult = extractLocalStructuralChanges(payload.text, fields, calculationRules);
  if (hasLocalStructuralResult(localStructuralResult)) {
    sendJson(response, 200, localStructuralResult);
    return;
  }

  const prompt = buildEnhancedExtractionPrompt(fields, calculationRules, payload.text);
  const doubaoPayload = {
    model,
    temperature: 0,
    response_format: { type: "json_object" },
    messages: [
      {
        role: "user",
        content: prompt,
      },
    ],
  };
  if (shouldDisableThinking(baseUrl, model)) {
    doubaoPayload.thinking = { type: "disabled" };
  }

  let upstreamResponse;
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timeout = controller ? setTimeout(() => controller.abort(), DOUBAO_TIMEOUT_MS) : 0;
  try {
    upstreamResponse = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(doubaoPayload),
      signal: controller ? controller.signal : undefined,
    });
  } catch (error) {
    if (error && (error.name === "AbortError" || /abort|timeout/i.test(error.message || ""))) {
      sendJson(response, 504, { error: "豆包服务响应超时，请稍后重试或先使用更简单的表头/规则指令" });
      return;
    }
    sendJson(response, 502, { error: `无法连接豆包服务：${error.message || "网络错误"}` });
    return;
  } finally {
    if (timeout) {
      clearTimeout(timeout);
    }
  }

  const upstreamText = await upstreamResponse.text();
  if (!upstreamResponse.ok) {
    sendJson(response, upstreamResponse.status >= 500 ? 502 : upstreamResponse.status, {
      error: normalizeUpstreamError(upstreamText) || `豆包服务返回 ${upstreamResponse.status}`,
    });
    return;
  }

  const rawContent = extractMessageContent(upstreamText);
  if (!rawContent) {
    sendJson(response, 502, { error: "豆包服务没有返回可解析内容" });
    return;
  }

  let extracted;
  try {
    extracted = JSON.parse(extractJsonText(rawContent));
  } catch (error) {
    sendJson(response, 502, { error: "豆包返回内容不是有效 JSON" });
    return;
  }

  const result = normalizeExtractionResult(extracted, fields, calculationRules);
  sendJson(response, 200, result);
}

function shouldDisableThinking(baseUrl, model) {
  const normalizedBaseUrl = String(baseUrl || "").toLowerCase();
  const normalizedModel = String(model || "").toLowerCase();
  return normalizedBaseUrl.includes("deepseek") && normalizedModel.includes("deepseek");
}

function buildExtractionPrompt(fields, text) {
  return [
    "你是用于 Excel 自动填表的自然语言填行 AgentSkill。",
    "根据用户输入的自然语言描述，按给定 Excel 表头字段生成一行或多行数据。",
    "只返回 JSON，不要 Markdown、解释、代码块或额外文字。",
    "返回格式必须是：{\"rows\":[{\"values\":{\"field_key\":\"value\"}}],\"warnings\":[]}",
    "规则：",
    "1. 描述中的每个项目、商品、房间、安装位置、安装明细或换行分隔明细对应 rows 中的一行。",
    "2. values 的键必须来自字段清单的 key；不要创造新键。",
    "3. 无法确定的字段填空字符串或省略。",
    "4. 数字字段去掉货币符号、中文单位和空格，只保留非负数字与小数点。",
    "5. 日期字段尽量返回 YYYY-MM-DD。",
    "6. 下拉字段优先使用 options 中最接近的原文。",
    "7. 如果描述不足以填写任何字段，返回 rows 为空，并在 warnings 说明。",
    "字段清单：",
    JSON.stringify(fields, null, 2),
    "用户描述：",
    String(text || ""),
  ].join("\n");
}

function buildEnhancedExtractionPrompt(fields, calculationRules, text) {
  return [
    "You are an AgentSkill for filling an Excel table from natural-language Chinese input.",
    "Use the provided field list to generate one or more data rows. If, and only if, the user explicitly asks to change current headers or calculation rules, also return temporary table changes.",
    "Return JSON only. Do not return Markdown, explanations, code fences, or any extra text.",
    "The JSON shape must be: {\"rows\":[{\"values\":{\"field_key\":\"value\"}}],\"fieldChanges\":[],\"calculationRuleChanges\":[],\"warnings\":[]}",
    "Rules:",
    "1. Each item/product/room/install location/detail/newline usually maps to one row.",
    "2. Row value keys must come from the supplied field list. Never invent row value keys. Do not put image fields in row values.",
    "3. Unknown row fields should be omitted or set to an empty string.",
    "4. Number fields should remove currency symbols, Chinese units, and spaces; keep only non-negative numbers and decimal points.",
    "5. Date fields should preferably use YYYY-MM-DD.",
    "6. Select fields should prefer the closest original option text.",
    "7. Return fieldChanges only when the user explicitly asks to add, rename, modify, or delete headers. Header changes are temporary for the current table.",
    "8. fieldChanges items must use one of these shapes:",
    "   - add/update: {\"action\":\"add|update\",\"key\":\"existing_key_or_empty\",\"label\":\"header label\",\"group\":\"parent header\",\"type\":\"text|number|date|select\",\"options\":[],\"required\":false}",
    "   - delete: {\"action\":\"delete\",\"key\":\"existing_key\"}",
    "9. For add, omit key when unsure; the client will generate it. For update/delete, key must be an existing key from the field list. Never invent field keys.",
    "10. Return calculationRuleChanges only when the user explicitly asks to set, replace, or delete a calculation rule.",
    "11. calculationRuleChanges items must use one of these shapes:",
    "   - set: {\"action\":\"set\",\"targetKey\":\"existing_key\",\"sourceKeys\":[\"existing_key_1\",\"existing_key_2\"],\"operator\":\"add|subtract|multiply|divide\"}",
    "   - delete: {\"action\":\"delete\",\"targetKey\":\"existing_key\"}",
    "12. Calculation v1 supports exactly two source fields. If a requested rule needs more than two fields or conditions, do not invent it; return a warning.",
    "13. Example: if the user says material/usage should be width times height, target the material/meters field, source width and height fields, operator multiply.",
    "14. If there are no fillable rows and no valid changes, return rows as [] and explain briefly in warnings.",
    "Field list:",
    JSON.stringify(fields, null, 2),
    "Current calculation rules:",
    JSON.stringify(calculationRules || [], null, 2),
    "User description:",
    String(text || ""),
  ].join("\n");
}

function validateExtractPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return { ok: false, error: "请求体必须是 JSON 对象" };
  }

  const text = String(payload.text || "").trim();
  if (!text) {
    return { ok: false, error: "请输入要填入表格的自然语言内容" };
  }

  if (text.length > 20000) {
    return { ok: false, error: "自然语言内容过长，请分段填写" };
  }

  if (!Array.isArray(payload.fields) || !payload.fields.length) {
    return { ok: false, error: "缺少可填充字段" };
  }

  return { ok: true };
}

function normalizeRequestFields(fields) {
  return fields
    .filter((field) => field && typeof field === "object" && field.key)
    .map((field) => ({
      key: String(field.key),
      label: String(field.label || ""),
      group: String(field.group || ""),
      type: ["text", "number", "date", "select", "image"].includes(field.type) ? field.type : "text",
      options: Array.isArray(field.options) ? field.options.map(String).filter(Boolean) : [],
      required: Boolean(field.required),
    }));
}

function normalizeExtractionResult(extracted, fields, currentCalculationRules) {
  const allowedFields = new Map(fields.map((field) => [field.key, field]));
  const rows = Array.isArray(extracted.rows) ? extracted.rows : [];
  const normalizedRows = rows
    .map((row) => {
      const sourceValues = row && row.values && typeof row.values === "object" ? row.values : row;
      const values = {};

      if (!sourceValues || typeof sourceValues !== "object") {
        return { values };
      }

      Object.keys(sourceValues).forEach((key) => {
        const field = allowedFields.get(key);
        if (!field || field.type === "image") {
          return;
        }

        const value = normalizeFieldValue(sourceValues[key], field);
        if (value !== "") {
          values[key] = value;
        }
      });

      return { values };
    })
    .filter((row) => Object.keys(row.values).length);

  const warnings = Array.isArray(extracted.warnings)
    ? extracted.warnings.map(String).filter(Boolean).slice(0, 5)
    : [];

  const fieldChanges = normalizeFieldChanges(extracted.fieldChanges, fields);
  const calculationRuleChanges = normalizeCalculationRuleChanges(
    extracted.calculationRuleChanges,
    extracted.calculationRules,
    fields,
    currentCalculationRules
  );
  const fieldChangeCount = Array.isArray(extracted.fieldChanges) ? extracted.fieldChanges.length : 0;
  const rawRuleChanges = Array.isArray(extracted.calculationRuleChanges)
    ? extracted.calculationRuleChanges
    : Array.isArray(extracted.calculationRules)
      ? extracted.calculationRules
      : [];
  const calculationRuleCount = rawRuleChanges.length;
  if (fieldChangeCount > fieldChanges.length) {
    warnings.push("Some invalid header changes were ignored.");
  }
  if (calculationRuleCount > calculationRuleChanges.length) {
    warnings.push("Some invalid calculation rules were ignored.");
  }

  return {
    rows: normalizedRows,
    fieldChanges,
    calculationRuleChanges,
    calculationRules: calculationRuleChanges
      .filter((change) => change.action === "set")
      .map(({ action, ...rule }) => rule),
    warnings: warnings.slice(0, 5),
  };
}

function normalizeFieldChanges(changes, fields) {
  const allowedKeys = new Set(fields.map((field) => field.key));
  return (Array.isArray(changes) ? changes : [])
    .map((change) => {
      if (!change || typeof change !== "object") {
        return null;
      }

      const action = ["add", "update", "delete"].includes(change.action) ? change.action : "";
      if (!action) {
        return null;
      }

      const key = String(change.key || "").trim();
      if ((action === "update" || action === "delete") && !allowedKeys.has(key)) {
        return null;
      }
      if (action === "delete") {
        return { action, key };
      }

      const label = String(change.label || "").trim();
      const group = String(change.group || "");
      const type = ["text", "number", "date", "select"].includes(change.type) ? change.type : "";
      const options = Array.isArray(change.options)
        ? change.options.map(String).map((item) => item.trim()).filter(Boolean).slice(0, 50)
        : [];
      const normalized = { action };

      if (key) {
        normalized.key = key;
      }
      if (label) {
        normalized.label = label;
      }
      if (Object.prototype.hasOwnProperty.call(change, "group")) {
        normalized.group = group;
      }
      if (type) {
        normalized.type = type;
      }
      if (type === "select" || options.length) {
        normalized.options = options;
      }
      if (Object.prototype.hasOwnProperty.call(change, "required")) {
        normalized.required = Boolean(change.required);
      }

      return action === "add" && !normalized.label ? null : normalized;
    })
    .filter(Boolean)
    .slice(0, 20);
}

function normalizeCalculationRuleChanges(changes, legacyRules, fields, currentRules) {
  const allowedKeys = new Set(fields.map((field) => field.key));
  const rawChanges = Array.isArray(changes)
    ? changes
    : Array.isArray(legacyRules)
      ? legacyRules.map((rule) => ({ ...rule, action: "set" }))
      : [];

  return rawChanges
    .map((change) => {
      if (!change || typeof change !== "object") {
        return null;
      }

      const action = change.action === "delete" ? "delete" : change.action === "set" ? "set" : "";
      if (!action) {
        return null;
      }

      const targetKey = String(change.targetKey || "").trim();
      if (!allowedKeys.has(targetKey)) {
        return null;
      }

      if (action === "delete") {
        return { action, targetKey };
      }

      const normalizedRule = normalizeCalculationRules([change], fields)[0];
      return normalizedRule ? { action, ...normalizedRule } : null;
    })
    .filter(Boolean)
    .slice(0, 20);
}

function extractLocalStructuralChanges(text, fields, calculationRules) {
  const sourceText = String(text || "").trim();
  const result = {
    rows: [],
    fieldChanges: [],
    calculationRuleChanges: [],
    calculationRules: [],
    warnings: [],
  };

  if (!sourceText) {
    return result;
  }

  const matchedFields = findFieldsMentionedInText(sourceText, fields);
  if (isDeleteHeaderRequest(sourceText, matchedFields) && !isDeleteCalculationRequest(sourceText)) {
    matchedFields.forEach((field) => {
      result.fieldChanges.push({ action: "delete", key: field.key });
    });
    if (!result.fieldChanges.length && hasHeaderIntent(sourceText)) {
      result.warnings.push("未找到要删除的表头，请尽量使用当前表头名称。");
    }
  }

  if (isDeleteCalculationRequest(sourceText)) {
    const targets = matchedFields.length
      ? matchedFields
      : normalizeCalculationRules(calculationRules, fields)
          .map((rule) => fields.find((field) => field.key === rule.targetKey))
          .filter(Boolean);
    targets.forEach((field) => {
      result.calculationRuleChanges.push({ action: "delete", targetKey: field.key });
    });
    if (!result.calculationRuleChanges.length) {
      result.warnings.push("未找到要取消的计算规则，请尽量使用目标表头名称。");
    }
  }

  result.fieldChanges = dedupeChangesByKey(result.fieldChanges, "key");
  result.calculationRuleChanges = dedupeChangesByKey(result.calculationRuleChanges, "targetKey");
  return result;
}

function hasLocalStructuralResult(result) {
  return Boolean(result && (
    (Array.isArray(result.fieldChanges) && result.fieldChanges.length) ||
    (Array.isArray(result.calculationRuleChanges) && result.calculationRuleChanges.length)
  ));
}

function findFieldsMentionedInText(text, fields) {
  const normalizedText = normalizeText(text);
  if (!normalizedText) {
    return [];
  }

  return fields.filter((field) => {
    const label = normalizeText(field.label);
    const group = normalizeText(field.group);
    const key = normalizeText(field.key);
    const fullLabel = normalizeText(`${field.group || ""}${field.label || ""}`);
    return (label && normalizedText.includes(label)) ||
      (fullLabel && normalizedText.includes(fullLabel)) ||
      (group && label && normalizedText.includes(`${group}${label}`)) ||
      (key && normalizedText.includes(key));
  });
}

function isDeleteHeaderRequest(text, matchedFields) {
  const normalized = normalizeText(text);
  const hasDeleteWord = /(删除|删掉|移除|去掉|不要|取消)/.test(text) || /(delete|remove|drop)/i.test(normalized);
  return hasDeleteWord && ((Array.isArray(matchedFields) && matchedFields.length > 0) || /(字段|表头|列|栏)/.test(text)) ||
    /(delete|remove|drop).*(field|header|column)/i.test(normalized);
}

function hasHeaderIntent(text) {
  return /(字段|表头|列|栏|field|header|column)/i.test(text);
}

function isDeleteCalculationRequest(text) {
  const normalized = normalizeText(text);
  return /(取消|删除|删掉|移除|去掉|不要|关闭|禁用)/.test(text) && /(计算|规则|公式|自动计算)/.test(text) ||
    /(disable|cancel|delete|remove).*(calculation|formula|rule)/i.test(normalized);
}

function dedupeChangesByKey(changes, keyName) {
  const seen = new Set();
  return changes.filter((change) => {
    const key = change && change[keyName];
    if (!key || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function normalizeCalculationRules(rules, fields) {
  const allowedKeys = new Set(fields.map((field) => field.key));
  const allowedOperators = new Set(["add", "subtract", "multiply", "divide"]);
  return (Array.isArray(rules) ? rules : [])
    .map((rule) => {
      if (!rule || typeof rule !== "object") {
        return null;
      }

      const targetKey = String(rule.targetKey || "").trim();
      const sourceKeys = Array.isArray(rule.sourceKeys)
        ? rule.sourceKeys.map((key) => String(key || "").trim()).filter(Boolean)
        : [];
      const operator = String(rule.operator || "").trim();
      if (!allowedKeys.has(targetKey) || sourceKeys.length !== 2 || !allowedOperators.has(operator)) {
        return null;
      }
      if (!sourceKeys.every((key) => allowedKeys.has(key)) || sourceKeys.includes(targetKey)) {
        return null;
      }

      return { targetKey, sourceKeys, operator };
    })
    .filter(Boolean)
    .slice(0, 20);
}

function normalizeFieldValue(value, field) {
  const text = String(value == null ? "" : value).trim();
  if (!text) {
    return "";
  }

  if (field.type === "number") {
    const cleaned = text
      .replace(/,/g, "")
      .replace(/[¥￥元米个件套平方㎡mM]/g, "")
      .replace(/[^\d.]/g, "");
    const numericValue = Number(cleaned);
    if (!Number.isFinite(numericValue)) {
      return "";
    }
    return numericValue < 0 ? "0" : cleaned;
  }

  if (field.type === "date") {
    const match = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
    if (match) {
      return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
    }
  }

  if (field.type === "select" && field.options.length) {
    const exact = field.options.find((option) => option === text);
    if (exact) {
      return exact;
    }
    const normalizedText = normalizeText(text);
    const fuzzy = field.options.find((option) => normalizeText(option) === normalizedText);
    return fuzzy || text;
  }

  return text;
}

function extractMessageContent(text) {
  try {
    const parsed = JSON.parse(text);
    return parsed && parsed.choices && parsed.choices[0] && parsed.choices[0].message
      ? String(parsed.choices[0].message.content || "")
      : "";
  } catch (error) {
    return "";
  }
}

function extractJsonText(text) {
  const trimmed = String(text || "").trim();
  const fenced = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  if (fenced) {
    return fenced[1].trim();
  }

  const firstBrace = trimmed.indexOf("{");
  const lastBrace = trimmed.lastIndexOf("}");
  if (firstBrace >= 0 && lastBrace > firstBrace) {
    return trimmed.slice(firstBrace, lastBrace + 1);
  }

  return trimmed;
}

function normalizeUpstreamError(text) {
  try {
    const parsed = JSON.parse(text);
    if (parsed && parsed.error) {
      return typeof parsed.error === "string" ? parsed.error : parsed.error.message;
    }
    if (parsed && parsed.message) {
      return parsed.message;
    }
  } catch (error) {
    return String(text || "").slice(0, 180);
  }
  return "";
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, "").replace(/[()（）]/g, "").toLowerCase();
}

function readJsonBody(request) {
  return new Promise((resolve, reject) => {
    let total = 0;
    const chunks = [];

    request.on("data", (chunk) => {
      total += chunk.length;
      if (total > MAX_BODY_BYTES) {
        request.destroy();
        reject(new Error("请求内容过大，请分段填写"));
        return;
      }
      chunks.push(chunk);
    });

    request.on("end", () => {
      try {
        const text = Buffer.concat(chunks).toString("utf8");
        resolve(text ? JSON.parse(text) : {});
      } catch (error) {
        reject(new Error("请求体不是有效 JSON"));
      }
    });
    request.on("error", reject);
  });
}

function readBinaryBody(request, maxBytes) {
  return new Promise((resolve, reject) => {
    let total = 0;
    const chunks = [];

    request.on("data", (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        request.destroy();
        reject(new Error("导出文件过大，请减少图片或行数后重试"));
        return;
      }
      chunks.push(chunk);
    });

    request.on("end", () => {
      resolve(Buffer.concat(chunks));
    });
    request.on("error", reject);
  });
}

async function handleExportCreate(request, response, url) {
  const contentType = String(request.headers["content-type"] || "").split(";")[0].trim().toLowerCase();
  if (contentType && contentType !== MIME_TYPES[".xlsx"]) {
    sendJson(response, 415, { error: "只支持上传 xlsx 文件" });
    return;
  }

  let buffer;
  try {
    buffer = await readBinaryBody(request, MAX_EXPORT_BYTES);
  } catch (error) {
    sendJson(response, 413, { error: error.message || "导出文件过大，请减少图片或行数后重试" });
    return;
  }

  if (!buffer.length) {
    sendJson(response, 400, { error: "导出文件为空" });
    return;
  }

  pruneExportFiles();
  const id = crypto.randomBytes(16).toString("hex");
  const requestedName = url.searchParams.get("filename") || "自动生成表格.xlsx";
  const filename = sanitizeDownloadFilename(requestedName);
  exportFiles.set(id, {
    buffer,
    filename,
    expiresAt: Date.now() + EXPORT_TTL_MS,
  });
  pruneExportFiles();

  sendJson(response, 200, {
    url: `${EXPORT_PATH_PREFIX}/${id}/${encodeURIComponent(filename)}`,
    expiresInSeconds: Math.max(1, Math.floor(EXPORT_TTL_MS / 1000)),
  });
}

function handleExportDownload(request, response, url) {
  pruneExportFiles();
  const parts = url.pathname.split("/").filter(Boolean);
  const id = parts[3] || "";
  const item = exportFiles.get(id);
  if (!item || item.expiresAt <= Date.now()) {
    exportFiles.delete(id);
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Download expired");
    return;
  }

  const filename = item.filename || "自动生成表格.xlsx";
  response.writeHead(200, {
    "Content-Type": MIME_TYPES[".xlsx"],
    "Content-Length": item.buffer.length,
    "Content-Disposition": `attachment; filename*=UTF-8''${encodeRFC5987ValueChars(filename)}`,
    "Cache-Control": "private, no-store",
  });
  if (request.method === "HEAD") {
    response.end();
    return;
  }
  response.end(item.buffer);
}

function pruneExportFiles() {
  const now = Date.now();
  exportFiles.forEach((item, id) => {
    if (!item || item.expiresAt <= now) {
      exportFiles.delete(id);
    }
  });

  if (exportFiles.size <= EXPORT_MAX_ITEMS) {
    return;
  }

  const oldest = Array.from(exportFiles.entries()).sort((left, right) => left[1].expiresAt - right[1].expiresAt);
  oldest.slice(0, Math.max(0, exportFiles.size - EXPORT_MAX_ITEMS)).forEach(([id]) => {
    exportFiles.delete(id);
  });
}

function sanitizeDownloadFilename(value) {
  const text = String(value || "")
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
    .trim()
    .slice(0, 120);
  const filename = text || "自动生成表格.xlsx";
  return /\.xlsx$/i.test(filename) ? filename : `${filename}.xlsx`;
}

function encodeRFC5987ValueChars(value) {
  return encodeURIComponent(value)
    .replace(/['()]/g, escape)
    .replace(/\*/g, "%2A");
}

function serveStatic(pathname, request, response) {
  const safePathname = pathname === "/" ? "/index.html" : pathname;
  const decoded = decodeURIComponent(safePathname);
  const filePath = path.normalize(path.join(ROOT_DIR, decoded));

  if (!filePath.startsWith(ROOT_DIR) || filePath.includes(`${path.sep}.env`)) {
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

function sendJson(response, statusCode, payload) {
  const body = JSON.stringify(payload);
  response.writeHead(statusCode, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
  });
  response.end(body);
}

function setCorsHeaders(request, response) {
  const origin = request.headers.origin;
  const allowedOrigin = getAllowedOrigin(request);
  if (allowedOrigin) {
    response.setHeader("Access-Control-Allow-Origin", allowedOrigin);
    response.setHeader("Vary", "Origin");
  } else if (!IS_PRODUCTION) {
    response.setHeader("Access-Control-Allow-Origin", "*");
  }
  response.setHeader("Access-Control-Allow-Methods", "GET, HEAD, POST, OPTIONS");
  response.setHeader("Access-Control-Allow-Headers", "Content-Type, X-App-Password");
  return !origin || Boolean(allowedOrigin) || !IS_PRODUCTION;
}

function getAllowedOrigin(request) {
  const origin = request.headers.origin;
  if (!origin) {
    return "";
  }
  if (isSameHostOrigin(origin, request)) {
    return origin;
  }
  if (!IS_PRODUCTION) {
    return origin;
  }
  return ALLOWED_ORIGINS.has(normalizeOrigin(origin)) ? origin : "";
}

function isRequestOriginAllowed(request) {
  const origin = request.headers.origin;
  const normalizedOrigin = normalizeOrigin(origin);
  return !IS_PRODUCTION || !origin || isSameHostOrigin(origin, request) || ALLOWED_ORIGINS.has(normalizedOrigin);
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

function getRequestOrigin(request) {
  const host = request.headers["x-forwarded-host"] || request.headers.host;
  if (!host) {
    return "";
  }
  const forwardedProto = String(request.headers["x-forwarded-proto"] || "").split(",")[0].trim();
  const proto = forwardedProto || (request.socket.encrypted ? "https" : "http");
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

function validateAppPassword(request) {
  if (!APP_PASSWORD) {
    if (IS_PRODUCTION) {
      return { ok: false, statusCode: 500, error: "APP_PASSWORD is not configured" };
    }
    return { ok: true };
  }

  const provided = String(request.headers["x-app-password"] || "");
  if (!provided || !timingSafeEqualText(provided, APP_PASSWORD)) {
    return { ok: false, statusCode: 401, error: "Invalid app password" };
  }

  return { ok: true };
}

function timingSafeEqualText(a, b) {
  const left = Buffer.from(String(a));
  const right = Buffer.from(String(b));
  return left.length === right.length && crypto.timingSafeEqual(left, right);
}

function checkNaturalFillRateLimit(request) {
  if (!RATE_LIMIT_MAX || !RATE_LIMIT_WINDOW_MS) {
    return { ok: true };
  }

  const now = Date.now();
  const key = getClientIp(request);
  let bucket = naturalFillRateLimits.get(key);
  if (!bucket || now >= bucket.resetAt) {
    bucket = { count: 0, resetAt: now + RATE_LIMIT_WINDOW_MS };
    naturalFillRateLimits.set(key, bucket);
  }

  bucket.count += 1;
  if (naturalFillRateLimits.size > 500) {
    pruneRateLimitBuckets(now);
  }

  if (bucket.count > RATE_LIMIT_MAX) {
    return { ok: false, retryAfterMs: Math.max(0, bucket.resetAt - now) };
  }

  return { ok: true };
}

function pruneRateLimitBuckets(now) {
  naturalFillRateLimits.forEach((bucket, key) => {
    if (now >= bucket.resetAt) {
      naturalFillRateLimits.delete(key);
    }
  });
}

function getClientIp(request) {
  const forwardedFor = String(request.headers["x-forwarded-for"] || "").split(",")[0].trim();
  return forwardedFor || request.socket.remoteAddress || "unknown";
}

function loadEnvFile(envPath) {
  if (!fs.existsSync(envPath)) {
    return;
  }

  const lines = fs.readFileSync(envPath, "utf8").split(/\r?\n/);
  lines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) {
      return;
    }
    const separatorIndex = trimmed.indexOf("=");
    if (separatorIndex <= 0) {
      return;
    }
    const key = trimmed.slice(0, separatorIndex).trim();
    let value = trimmed.slice(separatorIndex + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (!process.env[key]) {
      process.env[key] = value;
    }
  });
}
