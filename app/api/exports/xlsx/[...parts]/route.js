import { fetch as edgeFetch } from "../../../../../src/server/edge-handler.js";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

async function handle(request) {
  return edgeFetch(request, process.env);
}

export const GET = handle;
export const HEAD = handle;
export const OPTIONS = handle;
export const POST = handle;
