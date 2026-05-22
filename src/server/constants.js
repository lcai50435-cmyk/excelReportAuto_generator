const DEFAULT_DOUBAO_BASE_URL = "https://ark.cn-beijing.volces.com/api/v3";
const SKILL_PATH = "/api/agent-skills/doubao-excel-natural-fill/extract";
const EXPORT_PATH_PREFIX = "/api/exports/xlsx";
const LLM_CONFIG_API_PATH = "/api/config/llm";
const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".xlsx": MIME_XLSX,
  ".log": "text/plain; charset=utf-8",
};

export {
  DEFAULT_DOUBAO_BASE_URL,
  EXPORT_PATH_PREFIX,
  LLM_CONFIG_API_PATH,
  MIME_TYPES,
  MIME_XLSX,
  SKILL_PATH,
};
