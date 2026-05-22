const MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const NATURAL_FILL_ENDPOINT = "/api/agent-skills/doubao-excel-natural-fill/extract";
const EXPORT_ENDPOINT = "/api/exports/xlsx";
const LLM_CONFIG_ENDPOINT = "/api/config/llm";
const NATURAL_FILL_PASSWORD_STORAGE_KEY = "excelNaturalFillAppPassword";

export {
  EXPORT_ENDPOINT,
  LLM_CONFIG_ENDPOINT,
  MIME_XLSX,
  NATURAL_FILL_ENDPOINT,
  NATURAL_FILL_PASSWORD_STORAGE_KEY,
};
