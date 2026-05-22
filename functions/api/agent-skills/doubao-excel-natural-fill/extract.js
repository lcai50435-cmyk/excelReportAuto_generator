// Cloudflare Pages Function: AI proxy for Excel smart fill
// Handles POST /api/agent-skills/doubao-excel-natural-fill/extract

const RATE_LIMIT_WINDOW_MS = 60000;
const RATE_LIMIT_MAX = 20;
const MAX_BODY_BYTES = 1024 * 1024;
const DEFAULT_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3";

// In-memory rate limiting (per-isolate, best-effort in Workers)
const rateLimitBuckets = new Map();

export async function onRequest(context) {
  const { request, env } = context;
  const corsHeaders = buildCorsHeaders();

  if (request.method === "OPTIONS") {
    return new Response(null, { status: 204, headers: corsHeaders });
  }

  if (request.method !== "POST") {
    return jsonResponse({ error: "不支持的请求方法" }, 405, corsHeaders);
  }

  // --- App password check ---
  const appPassword = env.APP_PASSWORD || "";
  if (appPassword) {
    const provided = (request.headers.get("X-App-Password") || "").trim();
    if (!provided || !timingSafeEqual(provided, appPassword)) {
      return jsonResponse({ error: "Invalid app password" }, 401, corsHeaders);
    }
  }

  // --- Rate limiting ---
  const rateLimit = checkRateLimit(request);
  if (!rateLimit.ok) {
    const headers = new Headers(corsHeaders);
    headers.set("Retry-After", String(Math.ceil(rateLimit.retryAfterMs / 1000)));
    return jsonResponse({ error: "Too many AI requests. Please try again later." }, 429, headers);
  }

  // --- Parse & validate body ---
  let payload;
  try {
    const bodyText = await readRequestBody(request);
    payload = JSON.parse(bodyText);
  } catch {
    return jsonResponse({ error: "请求体不是有效 JSON" }, 400, corsHeaders);
  }

  const validation = validateExtractPayload(payload);
  if (!validation.ok) {
    return jsonResponse({ error: validation.error }, 400, corsHeaders);
  }

  const apiKey = env.DOUBAO_API_KEY;
  const model = env.DOUBAO_MODEL;
  const baseUrl = (env.DOUBAO_BASE_URL || DEFAULT_BASE_URL).replace(/\/+$/, "");

  if (!apiKey) return jsonResponse({ error: "请先设置 DOUBAO_API_KEY 环境变量" }, 500, corsHeaders);
  if (!model) return jsonResponse({ error: "请先设置 DOUBAO_MODEL 环境变量" }, 500, corsHeaders);

  const fields = normalizeRequestFields(payload.fields);
  if (!fields.length) {
    return jsonResponse({ error: "缺少可填充字段" }, 400, corsHeaders);
  }

  // --- Build AI prompt ---
  const prompt = buildExtractionPrompt(fields, payload.text);
  const doubaoPayload = {
    model,
    temperature: 0,
    response_format: { type: "json_object" },
    messages: [{ role: "user", content: prompt }],
  };
  if (shouldDisableThinking(baseUrl, model)) {
    doubaoPayload.thinking = { type: "disabled" };
  }

  // --- Call upstream AI ---
  let upstreamResponse;
  try {
    upstreamResponse = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(doubaoPayload),
    });
  } catch (error) {
    return jsonResponse(
      { error: `无法连接 AI 服务：${error.message || "网络错误"}` },
      502,
      corsHeaders
    );
  }

  const upstreamText = await upstreamResponse.text();
  if (!upstreamResponse.ok) {
    const status = upstreamResponse.status >= 500 ? 502 : upstreamResponse.status;
    return jsonResponse(
      { error: normalizeUpstreamError(upstreamText) || `AI 服务返回 ${upstreamResponse.status}` },
      status,
      corsHeaders
    );
  }

  // --- Parse AI response ---
  const rawContent = extractMessageContent(upstreamText);
  if (!rawContent) {
    return jsonResponse({ error: "AI 服务没有返回可解析内容" }, 502, corsHeaders);
  }

  let extracted;
  try {
    extracted = JSON.parse(extractJsonText(rawContent));
  } catch {
    return jsonResponse({ error: "AI 返回内容不是有效 JSON" }, 502, corsHeaders);
  }

  const result = normalizeExtractionResult(extracted, fields);
  return jsonResponse(result, 200, corsHeaders);
}

// ============ Helpers ============

function jsonResponse(payload, status, headers) {
  const body = JSON.stringify(payload);
  const h = new Headers(headers);
  h.set("Content-Type", "application/json; charset=utf-8");
  return new Response(body, { status, headers: h });
}

function buildCorsHeaders() {
  return {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, HEAD, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, X-App-Password",
    "Access-Control-Max-Age": "86400",
  };
}

function timingSafeEqual(a, b) {
  const encoder = new TextEncoder();
  const left = encoder.encode(a);
  const right = encoder.encode(b);
  if (left.length !== right.length) return false;
  let result = 0;
  for (let i = 0; i < left.length; i++) {
    result |= left[i] ^ right[i];
  }
  return result === 0;
}

function checkRateLimit(request) {
  if (!RATE_LIMIT_MAX || !RATE_LIMIT_WINDOW_MS) return { ok: true };

  const now = Date.now();
  const key = getClientIp(request);
  let bucket = rateLimitBuckets.get(key);

  if (!bucket || now >= bucket.resetAt) {
    bucket = { count: 0, resetAt: now + RATE_LIMIT_WINDOW_MS };
    rateLimitBuckets.set(key, bucket);
  }

  bucket.count += 1;

  // Prune old entries to prevent memory leak
  if (rateLimitBuckets.size > 500) {
    for (const [k, v] of rateLimitBuckets) {
      if (now >= v.resetAt) rateLimitBuckets.delete(k);
    }
  }

  if (bucket.count > RATE_LIMIT_MAX) {
    return { ok: false, retryAfterMs: Math.max(0, bucket.resetAt - now) };
  }
  return { ok: true };
}

function getClientIp(request) {
  const cf = request.headers.get("CF-Connecting-IP");
  if (cf) return cf;
  const forwarded = (request.headers.get("X-Forwarded-For") || "").split(",")[0].trim();
  return forwarded || "unknown";
}

async function readRequestBody(request) {
  const contentLength = Number(request.headers.get("Content-Length") || 0);
  if (contentLength > MAX_BODY_BYTES) {
    throw new Error("请求内容过大，请分段填写");
  }
  return request.text();
}

// ============ Validation & Normalization (from server.js) ============

function validateExtractPayload(payload) {
  if (!payload || typeof payload !== "object") {
    return { ok: false, error: "请求体必须是 JSON 对象" };
  }
  const text = String(payload.text || "").trim();
  if (!text) return { ok: false, error: "请输入要填入表格的自然语言内容" };
  if (text.length > 20000) return { ok: false, error: "自然语言内容过长，请分段填写" };
  if (!Array.isArray(payload.fields) || !payload.fields.length) {
    return { ok: false, error: "缺少可填充字段" };
  }
  return { ok: true };
}

function normalizeRequestFields(fields) {
  return fields
    .filter((f) => f && typeof f === "object" && f.key && f.type !== "image")
    .map((f) => ({
      key: String(f.key),
      label: String(f.label || ""),
      group: String(f.group || ""),
      type: ["text", "number", "date", "select"].includes(f.type) ? f.type : "text",
      options: Array.isArray(f.options) ? f.options.map(String).filter(Boolean) : [],
      required: Boolean(f.required),
    }));
}

function buildExtractionPrompt(fields, text) {
  return [
    "你是用于 Excel 自动填表的自然语言填行 AgentSkill。",
    "根据用户输入的自然语言描述，按给定 Excel 表头字段生成一行或多行数据。",
    "只返回 JSON，不要 Markdown、解释、代码块或额外文字。",
    '返回格式必须是：{"rows":[{"values":{"field_key":"value"}}],"warnings":[]}',
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

function normalizeExtractionResult(extracted, fields) {
  const allowedFields = new Map(fields.map((f) => [f.key, f]));
  const rows = Array.isArray(extracted.rows) ? extracted.rows : [];

  const normalizedRows = rows
    .map((row) => {
      const sourceValues = row && row.values && typeof row.values === "object" ? row.values : row;
      const values = {};
      if (!sourceValues || typeof sourceValues !== "object") return { values };

      for (const key of Object.keys(sourceValues)) {
        const field = allowedFields.get(key);
        if (!field) continue;
        const value = normalizeFieldValue(sourceValues[key], field);
        if (value !== "") values[key] = value;
      }
      return { values };
    })
    .filter((row) => Object.keys(row.values).length);

  const warnings = Array.isArray(extracted.warnings)
    ? extracted.warnings.map(String).filter(Boolean).slice(0, 5)
    : [];

  return { rows: normalizedRows, warnings };
}

function normalizeFieldValue(value, field) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";

  if (field.type === "number") {
    const cleaned = text
      .replace(/,/g, "")
      .replace(/[¥￥元米个件套平方㎡mM]/g, "")
      .replace(/[^\d.]/g, "");
    const num = Number(cleaned);
    if (!Number.isFinite(num)) return "";
    return num < 0 ? "0" : cleaned;
  }

  if (field.type === "date") {
    const match = text.match(/(\d{4})[./-](\d{1,2})[./-](\d{1,2})/);
    if (match) return `${match[1]}-${match[2].padStart(2, "0")}-${match[3].padStart(2, "0")}`;
  }

  if (field.type === "select" && field.options.length) {
    const exact = field.options.find((o) => o === text);
    if (exact) return exact;
    const normalized = normalizeText(text);
    const fuzzy = field.options.find((o) => normalizeText(o) === normalized);
    return fuzzy || text;
  }

  return text;
}

function extractMessageContent(text) {
  try {
    const parsed = JSON.parse(text);
    return parsed?.choices?.[0]?.message?.content
      ? String(parsed.choices[0].message.content)
      : "";
  } catch {
    return "";
  }
}

function extractJsonText(text) {
  const trimmed = String(text || "").trim();
  const fenced = trimmed.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  if (fenced) return fenced[1].trim();
  const first = trimmed.indexOf("{");
  const last = trimmed.lastIndexOf("}");
  if (first >= 0 && last > first) return trimmed.slice(first, last + 1);
  return trimmed;
}

function normalizeUpstreamError(text) {
  try {
    const parsed = JSON.parse(text);
    if (parsed?.error) return typeof parsed.error === "string" ? parsed.error : parsed.error.message;
    if (parsed?.message) return parsed.message;
  } catch {
    return String(text || "").slice(0, 180);
  }
  return "";
}

function normalizeText(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[()（）]/g, "")
    .toLowerCase();
}

function shouldDisableThinking(baseUrl, model) {
  const b = String(baseUrl || "").toLowerCase();
  const m = String(model || "").toLowerCase();
  return b.includes("deepseek") && m.includes("deepseek");
}
