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

export {
  shouldDisableThinking,
  buildExtractionPrompt,
  buildEnhancedExtractionPrompt,
  validateExtractPayload,
  normalizeRequestFields,
  normalizeExtractionResult,
  normalizeFieldChanges,
  normalizeCalculationRuleChanges,
  extractLocalStructuralChanges,
  hasLocalStructuralResult,
  normalizeCalculationRules,
  normalizeFieldValue,
  extractMessageContent,
  extractJsonText,
  normalizeUpstreamError,
  normalizeText,
};
