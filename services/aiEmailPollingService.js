const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const STATE_VERSION = 1;
const MAX_ALERTS = 20;

function cleanString(value, maxLength = 1000) {
  return String(value || "").trim().slice(0, maxLength);
}

function boundedInteger(value, fallback, minimum, maximum) {
  if (value === undefined || value === null || String(value).trim() === "") return fallback;
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.max(minimum, Math.min(maximum, Math.trunc(number)));
}

function safeErrorMessage(error) {
  const message = cleanString(error && error.message ? error.message : error, 500)
    .replace(/Bearer\s+[^\s]+/gi, "Bearer [redacted]")
    .replace(/(client_secret|access_token|refresh_token)=([^&\s]+)/gi, "$1=[redacted]")
    .replace(/\b(sk-[a-z0-9_-]{8})[a-z0-9_-]+\b/gi, "$1[redacted]");
  return message || "Scheduled intake failed.";
}

function verifySchedulerBearerToken(headerValue, expectedToken) {
  const expected = cleanString(expectedToken, 4096);
  if (expected.length < 32) return false;
  const match = String(headerValue || "").match(/^Bearer\s+([^\s]+)$/i);
  if (!match) return false;
  const suppliedDigest = crypto.createHash("sha256").update(match[1]).digest();
  const expectedDigest = crypto.createHash("sha256").update(expected).digest();
  return crypto.timingSafeEqual(suppliedDigest, expectedDigest);
}

function normalizeCounts(value) {
  return {
    checked: Math.max(0, Number(value && value.checked) || 0),
    processed: Math.max(0, Number(value && value.processed) || 0),
    skipped: Math.max(0, Number(value && value.skipped) || 0),
    failed: Math.max(0, Number(value && value.failed) || 0),
    proposalsCreated: Math.max(0, Number(value && value.proposalsCreated) || 0)
  };
}

function addCounts(left, right) {
  const result = normalizeCounts(left);
  const incoming = normalizeCounts(right);
  for (const key of Object.keys(result)) result[key] += incoming[key];
  return result;
}

function defaultState() {
  return {
    version: STATE_VERSION,
    status: "idle",
    checkpointAt: "",
    windowStartedAt: "",
    continuationCursor: "",
    lastAttemptAt: "",
    lastSuccessAt: "",
    lastCompletedAt: "",
    lastRunId: "",
    consecutiveFailures: 0,
    nextRetryAt: "",
    lastError: "",
    lastResult: normalizeCounts(),
    backlogRemaining: false,
    alerts: []
  };
}

function normalizeState(value) {
  const source = value && typeof value === "object" ? value : {};
  return {
    ...defaultState(),
    version: STATE_VERSION,
    status: cleanString(source.status, 40) || "idle",
    checkpointAt: cleanString(source.checkpointAt, 80),
    windowStartedAt: cleanString(source.windowStartedAt, 80),
    continuationCursor: cleanString(source.continuationCursor, 8000),
    lastAttemptAt: cleanString(source.lastAttemptAt, 80),
    lastSuccessAt: cleanString(source.lastSuccessAt, 80),
    lastCompletedAt: cleanString(source.lastCompletedAt, 80),
    lastRunId: cleanString(source.lastRunId, 120),
    consecutiveFailures: Math.max(0, Number(source.consecutiveFailures) || 0),
    nextRetryAt: cleanString(source.nextRetryAt, 80),
    lastError: cleanString(source.lastError, 500),
    lastResult: normalizeCounts(source.lastResult),
    backlogRemaining: Boolean(source.backlogRemaining),
    alerts: (Array.isArray(source.alerts) ? source.alerts : []).slice(0, MAX_ALERTS).map((alert) => ({
      id: cleanString(alert && alert.id, 120),
      type: cleanString(alert && alert.type, 80),
      message: cleanString(alert && alert.message, 500),
      createdAt: cleanString(alert && alert.createdAt, 80)
    })).filter((alert) => alert.id && alert.message)
  };
}

function createAiEmailPollingService({
  stateFile,
  lockFile,
  checkForNewEmails,
  readJsonFile,
  writeJsonFile,
  enabled = false,
  initialLookbackMinutes = 1440,
  overlapMinutes = 5,
  maxPagesPerRun = 10,
  maxRetries = 2,
  retryBaseMs = 1000,
  lockTtlMinutes = 30,
  now = () => new Date(),
  sleep = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)),
  fsImpl = fs
} = {}) {
  if (!stateFile || !lockFile || typeof checkForNewEmails !== "function") {
    throw new Error("Email polling service requires state, lock, and intake dependencies.");
  }
  const config = {
    enabled: Boolean(enabled),
    initialLookbackMinutes: boundedInteger(initialLookbackMinutes, 1440, 1, 60 * 24 * 30),
    overlapMinutes: boundedInteger(overlapMinutes, 5, 0, 60),
    maxPagesPerRun: boundedInteger(maxPagesPerRun, 10, 1, 100),
    maxRetries: boundedInteger(maxRetries, 2, 0, 5),
    retryBaseMs: boundedInteger(retryBaseMs, 1000, 0, 60000),
    lockTtlMinutes: boundedInteger(lockTtlMinutes, 30, 1, 120)
  };

  function readState() {
    return normalizeState(readJsonFile(stateFile, defaultState()));
  }

  function writeState(state) {
    const normalized = normalizeState(state);
    fsImpl.mkdirSync(path.dirname(stateFile), { recursive: true });
    const tempFile = `${stateFile}.${process.pid}.${crypto.randomUUID()}.tmp`;
    fsImpl.writeFileSync(tempFile, JSON.stringify(normalized, null, 2) + "\n", {
      encoding: "utf8",
      mode: 0o600
    });
    fsImpl.renameSync(tempFile, stateFile);
    return normalized;
  }

  function appendAlert(state, type, message, createdAt = now()) {
    const alert = {
      id: crypto.randomUUID(),
      type: cleanString(type, 80) || "error",
      message: cleanString(message, 500),
      createdAt: createdAt.toISOString()
    };
    return { ...state, alerts: [alert, ...(state.alerts || [])].slice(0, MAX_ALERTS) };
  }

  function lockAgeMs(lock) {
    const started = new Date(lock && (lock.heartbeatAt || lock.startedAt) || 0).getTime();
    return Number.isFinite(started) ? now().getTime() - started : Number.POSITIVE_INFINITY;
  }

  function readLock() {
    try {
      return JSON.parse(fsImpl.readFileSync(lockFile, "utf8"));
    } catch (error) {
      return null;
    }
  }

  function tryAcquireLock(runId) {
    fsImpl.mkdirSync(path.dirname(lockFile), { recursive: true });
    const startedAt = now().toISOString();
    const lock = { runId, startedAt, heartbeatAt: startedAt };
    try {
      const descriptor = fsImpl.openSync(lockFile, "wx", 0o600);
      fsImpl.writeFileSync(descriptor, JSON.stringify(lock) + "\n", "utf8");
      fsImpl.closeSync(descriptor);
      return { acquired: true, recovered: false, lock };
    } catch (error) {
      if (!error || error.code !== "EEXIST") throw error;
      const existing = readLock();
      if (lockAgeMs(existing) <= config.lockTtlMinutes * 60 * 1000) {
        return { acquired: false, recovered: false, lock: existing };
      }
      try {
        fsImpl.unlinkSync(lockFile);
      } catch (unlinkError) {
        if (!unlinkError || unlinkError.code !== "ENOENT") throw unlinkError;
      }
      const acquired = tryAcquireLock(runId);
      return { ...acquired, recovered: acquired.acquired };
    }
  }

  function releaseLock(runId) {
    const current = readLock();
    if (!current || current.runId !== runId) return false;
    try {
      fsImpl.unlinkSync(lockFile);
      return true;
    } catch (error) {
      if (error && error.code === "ENOENT") return false;
      throw error;
    }
  }

  function refreshLock(runId) {
    const current = readLock();
    if (!current || current.runId !== runId) {
      throw new Error("Scheduled intake lost its persistent lease.");
    }
    // Refresh the existing inode rather than replacing the lock pathname.
    // A different owner must never be overwritten by an old worker.
    const descriptor = fsImpl.openSync(lockFile, "r+");
    try {
      const onDisk = JSON.parse(fsImpl.readFileSync(descriptor, "utf8"));
      if (onDisk.runId !== runId) throw new Error("Scheduled intake lost its persistent lease.");
      fsImpl.ftruncateSync(descriptor, 0);
      fsImpl.writeSync(descriptor, JSON.stringify({ ...onDisk, heartbeatAt: now().toISOString() }) + "\n", 0, "utf8");
      fsImpl.fsyncSync(descriptor);
    } finally { fsImpl.closeSync(descriptor); }
  }

  async function runPageWithRetries({ user, since, continuationCursor }) {
    let lastError = null;
    for (let attempt = 0; attempt <= config.maxRetries; attempt += 1) {
      try {
        const result = await checkForNewEmails({ user, since, continuationCursor, scheduled: true });
        if (!result || result.configured === false || result.error || Number(result.failed || 0) > 0) {
          throw new Error(result && result.error || "One or more messages failed scheduled intake.");
        }
        return result;
      } catch (error) {
        lastError = error;
        if (attempt >= config.maxRetries) break;
        await sleep(config.retryBaseMs * (2 ** attempt));
      }
    }
    throw lastError || new Error("Scheduled intake failed.");
  }

  function windowStartFor(state, startedAt) {
    if (state.continuationCursor && state.windowStartedAt) return state.windowStartedAt;
    const checkpoint = new Date(state.checkpointAt || 0).getTime();
    const fallback = startedAt.getTime() - config.initialLookbackMinutes * 60 * 1000;
    const base = Number.isFinite(checkpoint) && checkpoint > 0 ? checkpoint : fallback;
    return new Date(base - config.overlapMinutes * 60 * 1000).toISOString();
  }

  async function runScheduled({ user } = {}) {
    if (!config.enabled) {
      return { status: "disabled", httpStatus: 503, error: "Automatic Microsoft 365 intake is disabled." };
    }
    if (!user || !user.email) {
      return { status: "misconfigured", httpStatus: 503, error: "The configured monitor user is unavailable." };
    }
    let state = readState();
    const currentTime = now();
    const retryAt = new Date(state.nextRetryAt || 0).getTime();
    if (Number.isFinite(retryAt) && retryAt > currentTime.getTime()) {
      return { status: "deferred", httpStatus: 429, nextRetryAt: state.nextRetryAt };
    }
    const runId = crypto.randomUUID();
    const lockResult = tryAcquireLock(runId);
    if (!lockResult.acquired) {
      return { status: "locked", httpStatus: 409, activeRunId: cleanString(lockResult.lock && lockResult.lock.runId, 120) };
    }
    if (lockResult.recovered) {
      state = appendAlert(state, "stale-lock-recovered", "Recovered a stale scheduled-intake lease before starting a new run.", currentTime);
    }
    const windowStartedAt = windowStartFor(state, currentTime);
    let cursor = state.continuationCursor;
    let totals = normalizeCounts();
    let pageCount = 0;
    state = writeState({
      ...state,
      status: "running",
      windowStartedAt,
      lastAttemptAt: currentTime.toISOString(),
      lastRunId: runId,
      lastError: ""
    });
    let leaseLost = false;
    const heartbeat = setInterval(() => {
      try { refreshLock(runId); } catch (error) { leaseLost = true; }
    }, Math.max(1000, Math.floor(config.lockTtlMinutes * 60000 / 3)));
    if (typeof heartbeat.unref === "function") heartbeat.unref();
    try {
      while (pageCount < config.maxPagesPerRun) {
        if (leaseLost) throw new Error("Scheduled intake lost its persistent lease.");
        refreshLock(runId);
        const page = await runPageWithRetries({ user, since: windowStartedAt, continuationCursor: cursor });
        refreshLock(runId);
        if (leaseLost) throw new Error("Scheduled intake lost its persistent lease.");
        totals = addCounts(totals, page);
        pageCount += 1;
        cursor = cleanString(page.continuationCursor, 8000);
        state = writeState({
          ...state,
          status: cursor ? "catching-up" : "running",
          continuationCursor: cursor,
          windowStartedAt,
          lastResult: totals,
          backlogRemaining: Boolean(cursor)
        });
        if (!cursor) break;
      }
      const completedAt = now();
      const backlogRemaining = Boolean(cursor);
      state = writeState({
        ...state,
        status: backlogRemaining ? "catching-up" : "idle",
        checkpointAt: backlogRemaining ? state.checkpointAt : currentTime.toISOString(),
        windowStartedAt: backlogRemaining ? windowStartedAt : "",
        continuationCursor: cursor,
        lastSuccessAt: completedAt.toISOString(),
        lastCompletedAt: completedAt.toISOString(),
        consecutiveFailures: 0,
        nextRetryAt: "",
        lastError: "",
        lastResult: totals,
        backlogRemaining
      });
      return {
        status: backlogRemaining ? "catching-up" : "completed",
        httpStatus: backlogRemaining ? 202 : 200,
        runId,
        pagesProcessed: pageCount,
        backlogRemaining,
        checkpointAt: state.checkpointAt,
        counts: totals
      };
    } catch (error) {
      const failedAt = now();
      const consecutiveFailures = state.consecutiveFailures + 1;
      const retryDelay = Math.min(60 * 60 * 1000, config.retryBaseMs * (2 ** Math.min(consecutiveFailures, 10)));
      const message = safeErrorMessage(error);
      state = appendAlert({
        ...state,
        status: "error",
        lastCompletedAt: failedAt.toISOString(),
        consecutiveFailures,
        nextRetryAt: new Date(failedAt.getTime() + retryDelay).toISOString(),
        lastError: message,
        lastResult: totals,
        backlogRemaining: Boolean(state.continuationCursor)
      }, "scheduled-intake-failed", message, failedAt);
      writeState(state);
      return {
        status: "error",
        httpStatus: 503,
        runId,
        error: message,
        nextRetryAt: state.nextRetryAt,
        counts: totals
      };
    } finally {
      clearInterval(heartbeat);
      releaseLock(runId);
    }
  }

  function getStatus() {
    const state = readState();
    return {
      config: { ...config },
      state: {
        ...state,
        continuationCursor: "",
        hasContinuationCursor: Boolean(state.continuationCursor)
      },
      lock: readLock() ? { active: true } : { active: false }
    };
  }

  return { getStatus, readState, runScheduled, tryAcquireLock, releaseLock };
}

module.exports = {
  addCounts,
  createAiEmailPollingService,
  normalizeState,
  verifySchedulerBearerToken
};
