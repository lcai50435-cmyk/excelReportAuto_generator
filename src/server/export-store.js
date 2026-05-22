function createExportStore() {
  const files = new Map();

  function create(buffer, filename, config) {
    prune(config);
    const id = createId();
    files.set(id, {
      buffer,
      filename,
      expiresAt: Date.now() + config.exportTtlMs,
    });
    prune(config);
    return {
      id,
      url: `/api/exports/xlsx/${id}/${encodeURIComponent(filename)}`,
      expiresInSeconds: Math.max(1, Math.floor(config.exportTtlMs / 1000)),
    };
  }

  function get(id, config) {
    prune(config);
    const item = files.get(id);
    if (!item || item.expiresAt <= Date.now()) {
      files.delete(id);
      return null;
    }
    return item;
  }

  function prune(config) {
    const now = Date.now();
    files.forEach((item, id) => {
      if (!item || item.expiresAt <= now) {
        files.delete(id);
      }
    });

    if (files.size <= config.exportMaxItems) {
      return;
    }

    const oldest = Array.from(files.entries()).sort((left, right) => left[1].expiresAt - right[1].expiresAt);
    oldest.slice(0, Math.max(0, files.size - config.exportMaxItems)).forEach(([id]) => {
      files.delete(id);
    });
  }

  return { create, get, prune };
}

function sanitizeDownloadFilename(value) {
  const text = String(value || "")
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_")
    .trim()
    .slice(0, 120);
  const filename = text || "自动填表数据.xlsx";
  return /\.xlsx$/i.test(filename) ? filename : `${filename}.xlsx`;
}

function createId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) {
    return crypto.randomUUID().replace(/-/g, "");
  }
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    const bytes = new Uint8Array(16);
    crypto.getRandomValues(bytes);
    return Array.from(bytes, (byte) => byte.toString(16).padStart(2, "0")).join("");
  }
  return `${Date.now().toString(16)}${Math.random().toString(16).slice(2).padEnd(16, "0")}`.slice(0, 32);
}

function encodeRFC5987ValueChars(value) {
  return encodeURIComponent(value)
    .replace(/['()]/g, escape)
    .replace(/\*/g, "%2A");
}

export {
  createExportStore,
  encodeRFC5987ValueChars,
  sanitizeDownloadFilename,
};
