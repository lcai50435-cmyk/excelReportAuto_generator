# Edgefunc / Node Hosting Deployment

This project now separates the app into a generic Fetch-style edge handler and a
small Node.js local adapter. The browser talks to the same origin for the page,
the AI endpoint, and export downloads, while the AI provider key stays on the
server.

## Edgefunc Entry

The edge-compatible entry is `src/server/edge-handler.js`.

```js
import { fetch } from "./src/server/edge-handler.js";
```

The exported `fetch(request, env)` function accepts standard Web `Request`
objects and returns standard Web `Response` objects. The `env` object uses the
same variable names as `.env` and hosted Node deployments.

The local `server.js` file only adapts Node `http` requests into that Fetch
handler and serves static files for local/Node hosting.

## Recommended Platform

Use a Node.js hosting platform such as Zeabur, Render, Railway, or Fly.io. For
small private use with an `.eu.cc` domain, Zeabur is the simplest fit. For edge
platforms, wire incoming requests to `fetch(request, env)` and serve the static
files from the repository root.

## Required Environment Variables

Set these in the hosting platform dashboard:

```env
NODE_ENV=production
PUBLIC_ORIGIN=https://your-name.eu.cc
APP_PASSWORD=change_this_app_password
DOUBAO_API_KEY=your_volcengine_ark_or_proxy_api_key
DOUBAO_MODEL=your_text_model_name
DOUBAO_BASE_URL=https://ark.cn-beijing.volces.com/api/v3
RATE_LIMIT_WINDOW_MS=60000
RATE_LIMIT_MAX=20
```

The platform normally injects `PORT` automatically. Do not set `HOST` on most
hosted platforms unless the provider asks for it.

## Start Command

```sh
npm start
```

## Custom Domain

1. Add `your-name.eu.cc` as a custom domain in the hosting platform.
2. Follow the platform DNS instructions, usually a `CNAME` or `A` record.
3. Wait until HTTPS is active.
4. Open `https://your-name.eu.cc/` on mobile data and test smart fill.

## Security Checks

- The browser must never contain `DOUBAO_API_KEY`.
- The AI endpoint requires `X-App-Password`.
- Production CORS allows only same-origin requests and configured origins.
- The `.env` file is blocked from static access.
