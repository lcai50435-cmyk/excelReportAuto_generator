# Next.js / Node Server Deployment

This project has been converted to a Next.js app for deployment on your own Node.js server.

## Architecture

- `app/page.jsx` renders the original single-page Excel UI.
- The original browser logic is served from `public/app.js` and `public/src/client/*`.
- Next.js route handlers proxy API requests to the existing Fetch-style backend in `src/server/edge-handler.js`.
- The app password is entered by the user and saved in localStorage as a lightweight login token.
- The LLM API key is configured from the web UI, saved on the server, and is never sent back to the browser.
- Excel files are still generated in the browser. For HTTP/HTTPS pages, the browser uploads the generated xlsx to `/api/exports/xlsx`; the server stores it temporarily in memory and returns a download URL.

## Required Environment Variables

Set only the app-level variables on your server:

```env
NODE_ENV=production
PUBLIC_ORIGIN=https://your-domain.com
APP_PASSWORD=change_this_app_password
```

`APP_PASSWORD` is required in production. The browser will prompt for it when a protected API returns 401, then saves it in localStorage as a lightweight login token.

The LLM settings are configured from the web UI after login:

- Base URL
- Model name
- API Key
- Timeout, default 60000ms

They are saved on the server to `data/llm-config.json` by default. Use `LLM_CONFIG_PATH=/secure/path/llm-config.json` if you want to store it elsewhere.

Optional operational variables:

```env
RATE_LIMIT_WINDOW_MS=60000
RATE_LIMIT_MAX=20
MAX_EXPORT_BYTES=31457280
EXPORT_MAX_ITEMS=50
EXPORT_TTL_MS=600000
LLM_CONFIG_PATH=/absolute/path/to/llm-config.json
```

## Install and Run

```sh
npm install
npm run build
npm start
```

By default, `next start` listens on port `3000`. To use another port:

```sh
PORT=4173 npm start
```

For development:

```sh
npm run dev
```

## Reverse Proxy Example

Use Nginx/Caddy/Traefik to proxy your domain to the Next.js process, for example `127.0.0.1:3000`.

Make sure your public URL matches `PUBLIC_ORIGIN` exactly, including protocol.

## Notes

- The legacy Node server is still available as `npm run start:legacy`, but the recommended deployment path is now Next.js.
- The export download store is in memory. It is suitable for a single Node process. If you run multiple replicas, either use sticky sessions or replace `src/server/export-store.js` with Redis/S3/R2/object storage.
- Do not expose `.env` or `data/llm-config.json`, and do not commit API keys.
