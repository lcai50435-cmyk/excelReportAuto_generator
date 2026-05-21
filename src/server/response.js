function jsonResponse(statusCode, payload, headers = {}) {
  const body = JSON.stringify(payload);
  const responseHeaders = new Headers(headers);
  responseHeaders.set("Content-Type", "application/json; charset=utf-8");
  responseHeaders.set("Content-Length", byteLength(body));
  return new Response(body, {
    status: statusCode,
    headers: responseHeaders,
  });
}

function textResponse(statusCode, body, headers = {}) {
  const text = String(body || "");
  const responseHeaders = new Headers(headers);
  if (!responseHeaders.has("Content-Type")) {
    responseHeaders.set("Content-Type", "text/plain; charset=utf-8");
  }
  responseHeaders.set("Content-Length", byteLength(text));
  return new Response(text, {
    status: statusCode,
    headers: responseHeaders,
  });
}

function emptyResponse(statusCode, headers = {}) {
  return new Response(null, {
    status: statusCode,
    headers,
  });
}

function byteLength(value) {
  if (typeof Buffer !== "undefined") {
    return String(Buffer.byteLength(value));
  }
  return String(new TextEncoder().encode(String(value)).length);
}

export { emptyResponse, jsonResponse, textResponse };
