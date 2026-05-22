import { readJsonBody } from "./body.js";
import {
  buildEnhancedExtractionPrompt,
  extractJsonText,
  extractLocalStructuralChanges,
  extractMessageContent,
  hasLocalStructuralResult,
  normalizeCalculationRules,
  normalizeExtractionResult,
  normalizeRequestFields,
  normalizeUpstreamError,
  shouldDisableThinking,
  validateExtractPayload,
} from "./natural-fill-core.js";
import { getEffectiveLlmConfig } from "./llm-config-store.js";
import { jsonResponse } from "./response.js";
import { validateAppPassword } from "./security.js";

async function handleNaturalFillExtract(request, context) {
  const { config, rateLimiter } = context;
  const auth = validateAppPassword(request, config);
  if (!auth.ok) {
    return jsonResponse(auth.statusCode, { error: auth.error });
  }

  const rateLimit = rateLimiter.check(request, config);
  if (!rateLimit.ok) {
    return jsonResponse(429, { error: "Too many AI requests. Please try again later." }, {
      "Retry-After": String(Math.ceil(rateLimit.retryAfterMs / 1000)),
    });
  }

  const payload = await readJsonBody(request, config.maxBodyBytes);
  const validation = validateExtractPayload(payload);
  if (!validation.ok) {
    return jsonResponse(400, { error: validation.error });
  }

  const llmConfig = await getEffectiveLlmConfig(config);
  if (!llmConfig.apiKey || !llmConfig.model || !llmConfig.baseUrl) {
    return jsonResponse(500, { error: "请先在大模型配置中填写 Base URL、模型名称和 API Key" });
  }

  const fields = normalizeRequestFields(payload.fields);
  if (!fields.length) {
    return jsonResponse(400, { error: "缺少可填充字段" });
  }

  const calculationRules = normalizeCalculationRules(payload.calculationRules, fields);
  const localStructuralResult = extractLocalStructuralChanges(payload.text, fields, calculationRules);
  if (hasLocalStructuralResult(localStructuralResult)) {
    return jsonResponse(200, localStructuralResult);
  }

  const prompt = buildEnhancedExtractionPrompt(fields, calculationRules, payload.text);
  const doubaoPayload = {
    model: llmConfig.model,
    temperature: 0,
    response_format: { type: "json_object" },
    messages: [
      {
        role: "user",
        content: prompt,
      },
    ],
  };
  if (shouldDisableThinking(llmConfig.baseUrl, llmConfig.model)) {
    doubaoPayload.thinking = { type: "disabled" };
  }

  let upstreamResponse;
  const controller = typeof AbortController === "function" ? new AbortController() : null;
  const timeout = controller ? setTimeout(() => controller.abort(), llmConfig.timeoutMs) : 0;
  try {
    upstreamResponse = await fetch(`${llmConfig.baseUrl}/chat/completions`, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${llmConfig.apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(doubaoPayload),
      signal: controller ? controller.signal : undefined,
    });
  } catch (error) {
    if (error && (error.name === "AbortError" || /abort|timeout/i.test(error.message || ""))) {
      return jsonResponse(504, { error: "豆包服务响应超时，请稍后重试或先使用更简单的表头/规则指令" });
    }
    return jsonResponse(502, { error: `无法连接豆包服务：${error.message || "网络错误"}` });
  } finally {
    if (timeout) {
      clearTimeout(timeout);
    }
  }

  const upstreamText = await upstreamResponse.text();
  if (!upstreamResponse.ok) {
    return jsonResponse(upstreamResponse.status >= 500 ? 502 : upstreamResponse.status, {
      error: normalizeUpstreamError(upstreamText) || `豆包服务返回 ${upstreamResponse.status}`,
    });
  }

  const rawContent = extractMessageContent(upstreamText);
  if (!rawContent) {
    return jsonResponse(502, { error: "豆包服务没有返回可解析内容" });
  }

  let extracted;
  try {
    extracted = JSON.parse(extractJsonText(rawContent));
  } catch (error) {
    return jsonResponse(502, { error: "豆包返回内容不是有效 JSON" });
  }

  const result = normalizeExtractionResult(extracted, fields, calculationRules);
  return jsonResponse(200, result);
}

export { handleNaturalFillExtract };
