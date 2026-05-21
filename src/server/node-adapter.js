import { Readable } from "node:stream";

function toWebRequest(request, host, port) {
  const url = new URL(request.url || "/", `http://${request.headers.host || `${host}:${port}`}`);
  const headers = new Headers();
  Object.entries(request.headers).forEach(([key, value]) => {
    if (Array.isArray(value)) {
      value.forEach((item) => headers.append(key, item));
    } else if (value != null) {
      headers.set(key, String(value));
    }
  });
  if (!headers.has("x-forwarded-for") && request.socket.remoteAddress) {
    headers.set("x-forwarded-for", request.socket.remoteAddress);
  }

  const init = {
    method: request.method,
    headers,
  };
  if (request.method !== "GET" && request.method !== "HEAD") {
    init.body = request;
    init.duplex = "half";
  }
  return new Request(url, init);
}

async function writeWebResponse(response, nodeResponse) {
  nodeResponse.statusCode = response.status;
  response.headers.forEach((value, key) => {
    nodeResponse.setHeader(key, value);
  });

  if (!response.body) {
    nodeResponse.end();
    return;
  }

  Readable.fromWeb(response.body).pipe(nodeResponse);
}

export { toWebRequest, writeWebResponse };
