function validateAppPassword(request, config) {
  if (!config.appPassword) {
    if (config.isProduction) {
      return { ok: false, statusCode: 500, error: "APP_PASSWORD is not configured" };
    }
    return { ok: true };
  }

  const provided = String(request.headers.get("x-app-password") || "");
  if (!provided || !timingSafeEqualText(provided, config.appPassword)) {
    return { ok: false, statusCode: 401, error: "Invalid app password" };
  }

  return { ok: true };
}

function timingSafeEqualText(leftValue, rightValue) {
  const left = String(leftValue);
  const right = String(rightValue);
  let mismatch = left.length ^ right.length;
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index += 1) {
    mismatch |= left.charCodeAt(index % left.length || 0) ^ right.charCodeAt(index % right.length || 0);
  }
  return mismatch === 0;
}

export { timingSafeEqualText, validateAppPassword };
