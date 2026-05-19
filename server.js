"use strict";

const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const { URL } = require("node:url");
const os = require("node:os");

const ROOT_DIR = __dirname;
const HOST = process.env.HOST || "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const MAX_BODY_BYTES = 1024 * 1024;
const DEFAULT_DOUBAO_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3";
const SKILL_PATH = "/api/agent-skills/doubao-excel-natural-fill/extract";

loadEnvFile(path.join(ROOT_DIR, ".env"));

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
    setCorsHeaders(response);

    if (request.method === "OPTIONS") {
      response.writeHead(204);
      response.end();
      return;
    }

    if (request.method === "POST" && url.pathname === SKILL_PATH) {
      await handleNaturalFillExtract(request, response);
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

  const prompt = buildExtractionPrompt(fields, payload.text);
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
  try {
    upstreamResponse = await fetch(`${baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(doubaoPayload),
    });
  } catch (error) {
    sendJson(response, 502, { error: `无法连接豆包服务：${error.message || "网络错误"}` });
    return;
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

  const result = normalizeExtractionResult(extracted, fields);
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
    .filter((field) => field && typeof field === "object" && field.key && field.type !== "image")
    .map((field) => ({
      key: String(field.key),
      label: String(field.label || ""),
      group: String(field.group || ""),
      type: ["text", "number", "date", "select"].includes(field.type) ? field.type : "text",
      options: Array.isArray(field.options) ? field.options.map(String).filter(Boolean) : [],
      required: Boolean(field.required),
    }));
}

function normalizeExtractionResult(extracted, fields) {
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
        if (!field) {
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

  return { rows: normalizedRows, warnings };
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

function setCorsHeaders(response) {
  response.setHeader("Access-Control-Allow-Origin", "*");
  response.setHeader("Access-Control-Allow-Methods", "GET, HEAD, POST, OPTIONS");
  response.setHeader("Access-Control-Allow-Headers", "Content-Type");
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
