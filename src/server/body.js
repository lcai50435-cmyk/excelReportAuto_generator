function byteLength(value) {
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(value)) {
    return value.length;
  }
  if (value instanceof Uint8Array) {
    return value.byteLength;
  }
  if (typeof Buffer !== "undefined") {
    return Buffer.byteLength(String(value));
  }
  return new TextEncoder().encode(String(value)).byteLength;
}

async function readJsonBody(request, maxBytes) {
  const text = await request.text();
  if (byteLength(text) > maxBytes) {
    throw new Error("请求内容过大，请分段填写");
  }

  try {
    return text ? JSON.parse(text) : {};
  } catch (error) {
    throw new Error("请求体不是有效 JSON");
  }
}

async function readBinaryBody(request, maxBytes) {
  const body = new Uint8Array(await request.arrayBuffer());
  if (body.byteLength > maxBytes) {
    throw new Error("导出文件过大，请减少图片或行数后重试");
  }
  return body;
}

export { byteLength, readBinaryBody, readJsonBody };
