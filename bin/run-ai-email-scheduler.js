#!/usr/bin/env node

function cleanString(value, maxLength = 2000) {
  return String(value || "").trim().slice(0, maxLength);
}

function schedulerUrl(baseUrl) {
  let url;
  try {
    url = new URL("/api/ai-email-intake/scheduled-check", cleanString(baseUrl, 2000));
  } catch (error) {
    throw new Error("AI_EMAIL_APP_BASE_URL must be a valid URL.");
  }
  const localHttp = url.protocol === "http:" && ["localhost", "127.0.0.1", "::1"].includes(url.hostname);
  if (url.protocol !== "https:" && !localHttp) {
    throw new Error("AI_EMAIL_APP_BASE_URL must use HTTPS.");
  }
  url.username = "";
  url.password = "";
  url.search = "";
  url.hash = "";
  return url.toString();
}

async function runSchedulerTrigger({
  baseUrl = process.env.AI_EMAIL_APP_BASE_URL,
  token = process.env.AI_EMAIL_SCHEDULER_TOKEN,
  fetchImpl = global.fetch,
  timeoutMs = 10 * 60 * 1000
} = {}) {
  const secret = cleanString(token, 4096);
  if (secret.length < 32) {
    throw new Error("AI_EMAIL_SCHEDULER_TOKEN must contain at least 32 characters.");
  }
  if (typeof fetchImpl !== "function") {
    throw new Error("A fetch implementation is required.");
  }
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetchImpl(schedulerUrl(baseUrl), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${secret}`,
        Accept: "application/json"
      },
      signal: controller.signal
    });
    const text = await response.text();
    let payload = {};
    try {
      payload = text ? JSON.parse(text) : {};
    } catch (error) {
      payload = {};
    }
    if ([200, 202, 409, 429].includes(response.status)) {
      return {
        ok: response.status === 200,
        delivered: true,
        processingComplete: response.status === 200,
        requiresFollowUp: response.status !== 200,
        httpStatus: response.status,
        status: cleanString(payload.status, 40) || "accepted",
        backlogRemaining: Boolean(payload.backlogRemaining),
        counts: payload.counts && typeof payload.counts === "object" ? payload.counts : {}
      };
    }
    throw new Error(`Scheduled intake trigger failed with HTTP ${response.status}.`);
  } finally {
    clearTimeout(timeout);
  }
}

if (require.main === module) {
  runSchedulerTrigger()
    .then((result) => {
      process.stdout.write(JSON.stringify(result) + "\n");
    })
    .catch((error) => {
      process.stderr.write(`${cleanString(error && error.message, 500) || "Scheduled intake trigger failed."}\n`);
      process.exitCode = 1;
    });
}

module.exports = {
  runSchedulerTrigger,
  schedulerUrl
};
