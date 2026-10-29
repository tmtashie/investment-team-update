const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const {
  createAiEmailPollingService,
  verifySchedulerBearerToken
} = require("../services/aiEmailPollingService");

function createHarness(t, overrides = {}) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-polling-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  const readJsonFile = (file, fallback) => {
    try {
      return JSON.parse(fs.readFileSync(file, "utf8"));
    } catch (error) {
      return fallback;
    }
  };
  const writeJsonFile = (file, value) => {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(value, null, 2) + "\n", "utf8");
  };
  const current = overrides.current || { value: new Date("2026-10-29T12:00:00.000Z") };
  const service = createAiEmailPollingService({
    stateFile,
    lockFile,
    readJsonFile,
    writeJsonFile,
    enabled: true,
    initialLookbackMinutes: 60,
    overlapMinutes: 5,
    maxPagesPerRun: 10,
    maxRetries: 0,
    retryBaseMs: 10,
    lockTtlMinutes: 15,
    now: () => new Date(current.value),
    sleep: async () => {},
    ...overrides,
    current: undefined
  });
  return { service, stateFile, lockFile, readJsonFile, writeJsonFile, current, directory };
}

const monitorUser = { email: "monitor@example.test", role: "master-editor" };

test("scheduler bearer authentication requires a long exact token", () => {
  const token = "0123456789abcdef0123456789abcdef";
  assert.equal(verifySchedulerBearerToken(`Bearer ${token}`, token), true);
  assert.equal(verifySchedulerBearerToken(`bearer ${token}`, token), true);
  assert.equal(verifySchedulerBearerToken("Bearer wrong", token), false);
  assert.equal(verifySchedulerBearerToken(`Basic ${token}`, token), false);
  assert.equal(verifySchedulerBearerToken("Bearer short", "short"), false);
});

test("scheduled polling checkpoints only after all chronological Graph pages complete", async (t) => {
  const calls = [];
  const harness = createHarness(t, {
    checkForNewEmails: async (options) => {
      calls.push(options);
      return calls.length === 1
        ? { configured: true, checked: 2, processed: 2, skipped: 0, failed: 0, proposalsCreated: 2, continuationCursor: "trusted-next-page" }
        : { configured: true, checked: 1, processed: 1, skipped: 0, failed: 0, proposalsCreated: 1, continuationCursor: "" };
    }
  });

  const result = await harness.service.runScheduled({ user: monitorUser });

  assert.equal(result.status, "completed");
  assert.equal(result.pagesProcessed, 2);
  assert.deepEqual(result.counts, { checked: 3, processed: 3, skipped: 0, failed: 0, proposalsCreated: 3 });
  assert.equal(calls[0].since, "2026-10-29T10:55:00.000Z");
  assert.equal(calls[0].continuationCursor, "");
  assert.equal(calls[1].continuationCursor, "trusted-next-page");
  const state = harness.service.readState();
  assert.equal(state.checkpointAt, "2026-10-29T12:00:00.000Z");
  assert.equal(state.continuationCursor, "");
});

test("scheduled polling retries transient failures with bounded backoff", async (t) => {
  let attempts = 0;
  const sleeps = [];
  const harness = createHarness(t, {
    maxRetries: 1,
    sleep: async (milliseconds) => sleeps.push(milliseconds),
    checkForNewEmails: async () => {
      attempts += 1;
      if (attempts === 1) throw new Error("temporary Graph failure");
      return { configured: true, checked: 1, processed: 0, skipped: 1, failed: 0, proposalsCreated: 0 };
    }
  });

  const result = await harness.service.runScheduled({ user: monitorUser });

  assert.equal(result.status, "completed");
  assert.equal(attempts, 2);
  assert.deepEqual(sleeps, [10]);
  assert.equal(harness.service.readState().consecutiveFailures, 0);
});

test("a failed run preserves its checkpoint and records a sanitized operational alert", async (t) => {
  const harness = createHarness(t, {
    checkForNewEmails: async () => ({
      configured: true,
      checked: 1,
      processed: 0,
      skipped: 0,
      failed: 1,
      proposalsCreated: 0,
      error: "attachment failed Bearer super-secret-token"
    })
  });
  harness.writeJsonFile(harness.stateFile, {
    checkpointAt: "2026-10-29T10:00:00.000Z",
    status: "idle"
  });

  const result = await harness.service.runScheduled({ user: monitorUser });
  const state = harness.service.readState();

  assert.equal(result.status, "error");
  assert.equal(state.checkpointAt, "2026-10-29T10:00:00.000Z");
  assert.equal(state.status, "error");
  assert.equal(state.alerts[0].type, "scheduled-intake-failed");
  assert.equal(state.alerts[0].message, "attachment failed Bearer [redacted]");
  assert.equal(JSON.stringify(state).includes("super-secret-token"), false);
  assert.ok(state.nextRetryAt);
});

test("a restarted service resumes a persisted continuation cursor without replaying completed work", async (t) => {
  let firstPageProcessed = 0;
  const first = createHarness(t, {
    maxPagesPerRun: 1,
    checkForNewEmails: async () => {
      firstPageProcessed += 1;
      return { configured: true, checked: 1, processed: 1, skipped: 0, failed: 0, proposalsCreated: 1, continuationCursor: "trusted-next-page" };
    }
  });

  const firstResult = await first.service.runScheduled({ user: monitorUser });
  assert.equal(firstResult.status, "catching-up");
  assert.equal(first.service.readState().checkpointAt, "");

  const continuationCalls = [];
  const restarted = createAiEmailPollingService({
    stateFile: first.stateFile,
    lockFile: first.lockFile,
    readJsonFile: first.readJsonFile,
    writeJsonFile: first.writeJsonFile,
    enabled: true,
    initialLookbackMinutes: 60,
    overlapMinutes: 5,
    maxPagesPerRun: 10,
    maxRetries: 0,
    now: () => new Date("2026-10-29T12:05:00.000Z"),
    sleep: async () => {},
    checkForNewEmails: async (options) => {
      continuationCalls.push(options);
      return { configured: true, checked: 1, processed: 1, skipped: 0, failed: 0, proposalsCreated: 1, continuationCursor: "" };
    }
  });

  const restartedResult = await restarted.runScheduled({ user: monitorUser });

  assert.equal(firstPageProcessed, 1);
  assert.equal(continuationCalls[0].continuationCursor, "trusted-next-page");
  assert.equal(restartedResult.status, "completed");
  assert.equal(restarted.readState().checkpointAt, "2026-10-29T12:05:00.000Z");
});

test("restart recovery safely replays an interrupted page and relies on intake deduplication", async (t) => {
  let proposalWrites = 0;
  let sourceWasProcessed = false;
  const first = createHarness(t, {
    checkForNewEmails: async () => {
      if (!sourceWasProcessed) {
        sourceWasProcessed = true;
        proposalWrites += 1;
      }
      throw new Error("process interrupted after durable message state");
    }
  });

  const failed = await first.service.runScheduled({ user: monitorUser });
  assert.equal(failed.status, "error");
  const failedState = first.service.readState();
  first.writeJsonFile(first.stateFile, { ...failedState, nextRetryAt: "" });

  const restarted = createAiEmailPollingService({
    stateFile: first.stateFile,
    lockFile: first.lockFile,
    readJsonFile: first.readJsonFile,
    writeJsonFile: first.writeJsonFile,
    enabled: true,
    maxRetries: 0,
    now: () => new Date("2026-10-29T12:10:00.000Z"),
    sleep: async () => {},
    checkForNewEmails: async () => ({
      configured: true,
      checked: 1,
      processed: 0,
      skipped: sourceWasProcessed ? 1 : 0,
      failed: 0,
      proposalsCreated: 0,
      continuationCursor: ""
    })
  });

  const recovered = await restarted.runScheduled({ user: monitorUser });

  assert.equal(recovered.status, "completed");
  assert.equal(proposalWrites, 1);
  assert.equal(recovered.counts.proposalsCreated, 0);
});

test("a stale persistent lease is recovered and surfaced as an operational alert", async (t) => {
  const harness = createHarness(t, {
    checkForNewEmails: async () => ({
      configured: true,
      checked: 0,
      processed: 0,
      skipped: 0,
      failed: 0,
      proposalsCreated: 0,
      continuationCursor: ""
    })
  });
  fs.writeFileSync(harness.lockFile, JSON.stringify({
    runId: "interrupted-run",
    startedAt: "2026-10-29T11:00:00.000Z"
  }));

  const result = await harness.service.runScheduled({ user: monitorUser });
  const state = harness.service.readState();

  assert.equal(result.status, "completed");
  assert.equal(state.alerts.some((alert) => alert.type === "stale-lock-recovered"), true);
  assert.equal(fs.existsSync(harness.lockFile), false);
});

test("overlapping scheduler deliveries are rejected by the persistent lease", async (t) => {
  let releasePage;
  const pageGate = new Promise((resolve) => {
    releasePage = resolve;
  });
  const harness = createHarness(t, {
    checkForNewEmails: async () => {
      await pageGate;
      return {
        configured: true,
        checked: 0,
        processed: 0,
        skipped: 0,
        failed: 0,
        proposalsCreated: 0,
        continuationCursor: ""
      };
    }
  });

  const activeRun = harness.service.runScheduled({ user: monitorUser });
  await new Promise((resolve) => setImmediate(resolve));
  const overlappingRun = await harness.service.runScheduled({ user: monitorUser });
  releasePage();
  const completedRun = await activeRun;

  assert.equal(overlappingRun.status, "locked");
  assert.equal(overlappingRun.httpStatus, 409);
  assert.equal(completedRun.status, "completed");
});

test("monitor status never exposes the persisted Graph continuation cursor", async (t) => {
  const harness = createHarness(t, {
    maxPagesPerRun: 1,
    checkForNewEmails: async () => ({
      configured: true,
      checked: 0,
      processed: 0,
      skipped: 0,
      failed: 0,
      proposalsCreated: 0,
      continuationCursor: "secret-graph-cursor"
    })
  });
  await harness.service.runScheduled({ user: monitorUser });

  const status = harness.service.getStatus();

  assert.equal(status.state.continuationCursor, "");
  assert.equal(status.state.hasContinuationCursor, true);
  assert.equal(JSON.stringify(status).includes("secret-graph-cursor"), false);
});

test("scheduled route is independently authenticated and cannot approve or apply proposals", () => {
  const serverSource = fs.readFileSync(path.join(__dirname, "..", "server.js"), "utf8");
  const routeStart = serverSource.indexOf('url.pathname === "/api/ai-email-intake/scheduled-check"');
  const routeEnd = serverSource.indexOf('url.pathname === "/api/ai-email-intake/monitor-status"', routeStart);
  const routeSource = serverSource.slice(routeStart, routeEnd);
  const envSource = fs.readFileSync(path.join(__dirname, "..", ".env.example"), "utf8");
  const renderSource = fs.readFileSync(path.join(__dirname, "..", "render.yaml"), "utf8");

  assert.ok(routeStart > -1);
  assert.match(routeSource, /verifySchedulerBearerToken/);
  assert.match(routeSource, /AI_EMAIL_AUTO_MONITOR_ENABLED/);
  assert.match(routeSource, /getConfiguredAiEmailMonitorUser/);
  assert.match(routeSource, /aiEmailPollingService\.runScheduled/);
  assert.doesNotMatch(routeSource, /approve|applyApproved|createInvestment|sendEmail/);
  assert.match(envSource, /AI_EMAIL_AUTO_MONITOR_ENABLED=false/);
  assert.match(renderSource, /AI_EMAIL_AUTO_MONITOR_ENABLED\n\s+value: "false"/);
});
