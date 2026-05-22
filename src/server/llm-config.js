import { readJsonBody } from "./body.js";
import { getPublicLlmConfig, saveLlmConfigUpdate } from "./llm-config-store.js";
import { jsonResponse } from "./response.js";
import { validateAppPassword } from "./security.js";

async function handleLlmConfigGet(request, context) {
  const { config } = context;
  const auth = validateAppPassword(request, config);
  if (!auth.ok) {
    return jsonResponse(auth.statusCode, { error: auth.error });
  }

  const llmConfig = await getPublicLlmConfig(config);
  return jsonResponse(200, llmConfig);
}

async function handleLlmConfigSave(request, context) {
  const { config } = context;
  const auth = validateAppPassword(request, config);
  if (!auth.ok) {
    return jsonResponse(auth.statusCode, { error: auth.error });
  }

  const payload = await readJsonBody(request, config.maxBodyBytes);
  try {
    const llmConfig = await saveLlmConfigUpdate(payload, config);
    return jsonResponse(200, llmConfig);
  } catch (error) {
    return jsonResponse(400, { error: error.message || "大模型配置保存失败" });
  }
}

export { handleLlmConfigGet, handleLlmConfigSave };
