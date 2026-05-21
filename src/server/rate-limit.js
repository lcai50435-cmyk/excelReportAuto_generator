function createRateLimiter() {
  const buckets = new Map();

  function check(request, config) {
    if (!config.rateLimitMax || !config.rateLimitWindowMs) {
      return { ok: true };
    }

    const now = Date.now();
    const key = getClientIp(request);
    let bucket = buckets.get(key);
    if (!bucket || now >= bucket.resetAt) {
      bucket = { count: 0, resetAt: now + config.rateLimitWindowMs };
      buckets.set(key, bucket);
    }

    bucket.count += 1;
    if (buckets.size > 500) {
      prune(now);
    }

    if (bucket.count > config.rateLimitMax) {
      return { ok: false, retryAfterMs: Math.max(0, bucket.resetAt - now) };
    }

    return { ok: true };
  }

  function prune(now = Date.now()) {
    buckets.forEach((bucket, key) => {
      if (now >= bucket.resetAt) {
        buckets.delete(key);
      }
    });
  }

  return { check, prune };
}

function getClientIp(request) {
  const forwardedFor = String(request.headers.get("x-forwarded-for") || "").split(",")[0].trim();
  return forwardedFor || "unknown";
}

export { createRateLimiter, getClientIp };
