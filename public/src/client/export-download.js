import { EXPORT_ENDPOINT, MIME_XLSX } from "./constants.js";

const root = typeof window !== "undefined" ? window : globalThis;

function shouldUseServerExportDownload() {
  return Boolean(root.location && /^https?:$/.test(root.location.protocol) && root.fetch);
}

async function createServerExportDownload(blob, filename) {
  const response = await fetch(`${EXPORT_ENDPOINT}?filename=${encodeURIComponent(filename)}`, {
    method: "POST",
    headers: {
      "Content-Type": MIME_XLSX,
    },
    body: blob,
  });
  let payload = null;
  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }
  if (!response.ok || !payload || !payload.url) {
    throw new Error((payload && payload.error) || "无法创建下载链接");
  }
  return payload.url;
}

function downloadBlobLocally(blob, filename) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  document.body.appendChild(anchor);
  anchor.click();
  if (anchor.parentNode) {
    anchor.parentNode.removeChild(anchor);
  }
  setTimeout(() => URL.revokeObjectURL(url), 1200);
}

export {
  createServerExportDownload,
  downloadBlobLocally,
  shouldUseServerExportDownload,
};
