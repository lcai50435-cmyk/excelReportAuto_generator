import { readBinaryBody } from "./body.js";
import { EXPORT_PATH_PREFIX, MIME_XLSX } from "./constants.js";
import { encodeRFC5987ValueChars, sanitizeDownloadFilename } from "./export-store.js";
import { jsonResponse, textResponse } from "./response.js";

async function handleExportCreate(request, url, context) {
  const { config, exportStore } = context;
  const contentType = String(request.headers.get("content-type") || "").split(";")[0].trim().toLowerCase();
  if (contentType && contentType !== MIME_XLSX) {
    return jsonResponse(415, { error: "只支持上传 xlsx 文件" });
  }

  let buffer;
  try {
    buffer = await readBinaryBody(request, config.exportMaxBytes);
  } catch (error) {
    return jsonResponse(413, { error: error.message || "导出文件过大，请减少图片或行数后重试" });
  }

  if (!buffer.byteLength) {
    return jsonResponse(400, { error: "导出文件为空" });
  }

  const requestedName = url.searchParams.get("filename") || "自动填表数据.xlsx";
  const filename = sanitizeDownloadFilename(requestedName);
  const item = exportStore.create(buffer, filename, config);

  return jsonResponse(200, {
    url: item.url,
    expiresInSeconds: item.expiresInSeconds,
  });
}

function handleExportDownload(request, url, context) {
  const { config, exportStore } = context;
  const parts = url.pathname.split("/").filter(Boolean);
  const id = parts[3] || "";
  const item = exportStore.get(id, config);
  if (!item) {
    return textResponse(404, "Download expired", {
      "Content-Type": "text/plain; charset=utf-8",
    });
  }

  const filename = item.filename || "自动填表数据.xlsx";
  const headers = new Headers({
    "Content-Type": MIME_XLSX,
    "Content-Length": String(item.buffer.byteLength),
    "Content-Disposition": `attachment; filename*=UTF-8''${encodeRFC5987ValueChars(filename)}`,
    "Cache-Control": "private, no-store",
  });
  if (request.method === "HEAD") {
    return new Response(null, {
      status: 200,
      headers,
    });
  }
  return new Response(item.buffer, {
    status: 200,
    headers,
  });
}

function isExportDownloadPath(pathname) {
  return pathname.startsWith(`${EXPORT_PATH_PREFIX}/`);
}

export { handleExportCreate, handleExportDownload, isExportDownloadPath };
