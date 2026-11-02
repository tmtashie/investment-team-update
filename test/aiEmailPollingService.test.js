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


test("long-running Graph request refreshes its lease and blocks stale recovery by a concurrent worker", async (t) => {
  let resolvePage;
  let heartbeat;
  let heartbeatCleared = false;
  const first = createHarness(t, {
    lockTtlMinutes: 1,
    setIntervalImpl: (callback) => {
      heartbeat = callback;
      return { unref() {} };
    },
    clearIntervalImpl: () => {
      heartbeatCleared = true;
    },
    checkForNewEmails: () => new Promise((resolve) => { resolvePage = resolve; })
  });
  const pending = first.service.runScheduled({ user: monitorUser });

  assert.equal(typeof heartbeat, "function");
  first.current.value = new Date("2026-10-29T12:00:20.000Z");
  heartbeat();
  first.current.value = new Date("2026-10-29T12:01:10.000Z");

  const second = createAiEmailPollingService({
    stateFile: first.stateFile, lockFile: first.lockFile,
    readJsonFile: first.readJsonFile, writeJsonFile: first.writeJsonFile,
    enabled: true, lockTtlMinutes: 1,
    now: () => new Date(first.current.value),
    checkForNewEmails: async () => { throw new Error("must not run"); }
  });
  assert.equal((await second.runScheduled({ user: monitorUser })).status, "locked");

  resolvePage({ configured: true, checked: 0, processed: 0, skipped: 0, failed: 0, proposalsCreated: 0 });
  assert.equal((await pending).status, "completed");
  assert.equal(heartbeatCleared, true);
});

function createLeaseService({
  stateFile,
  lockFile,
  now,
  checkForNewEmails,
  setIntervalImpl = setInterval,
  clearIntervalImpl = clearInterval
}) {
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
  return createAiEmailPollingService({
    stateFile,
    lockFile,
    readJsonFile,
    writeJsonFile,
    enabled: true,
    lockTtlMinutes: 1,
    maxRetries: 0,
    retryBaseMs: 0,
    now,
    sleep: async () => {},
    checkForNewEmails,
    setIntervalImpl,
    clearIntervalImpl
  });
}

test("an expired lease during a long Graph request fails the old worker and recovers without duplicate processing", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-lease-expiry-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  const current = { value: new Date("2026-10-30T12:00:00.000Z") };
  let finishOldRequest;
  let finishNewRequest;
  let proposalWrites = 0;
  let sourceWasProcessed = false;

  const oldWorker = createLeaseService({
    stateFile,
    lockFile,
    now: () => new Date(current.value),
    setIntervalImpl: () => ({ unref() {} }),
    clearIntervalImpl: () => {},
    checkForNewEmails: () => new Promise((resolve) => {
      finishOldRequest = () => {
        if (!sourceWasProcessed) {
          sourceWasProcessed = true;
          proposalWrites += 1;
        }
        resolve({
          configured: true,
          checked: 1,
          processed: 1,
          skipped: 0,
          failed: 0,
          proposalsCreated: 1,
          continuationCursor: ""
        });
      };
    })
  });

  const oldRun = oldWorker.runScheduled({
    user: { email: "old-worker@example.test", role: "master-editor" }
  });
  await new Promise((resolve) => setImmediate(resolve));
  const oldLock = JSON.parse(fs.readFileSync(lockFile, "utf8"));

  current.value = new Date("2026-10-30T12:01:01.000Z");
  const newWorker = createLeaseService({
    stateFile,
    lockFile,
    now: () => new Date(current.value),
    checkForNewEmails: () => new Promise((resolve) => {
      finishNewRequest = () => resolve({
        configured: true,
        checked: 1,
        processed: 0,
        skipped: sourceWasProcessed ? 1 : 0,
        failed: 0,
        proposalsCreated: 0,
        continuationCursor: ""
      });
    })
  });
  const newRun = newWorker.runScheduled({
    user: { email: "new-worker@example.test", role: "master-editor" }
  });
  await new Promise((resolve) => setImmediate(resolve));

  const newLock = JSON.parse(fs.readFileSync(lockFile, "utf8"));
  assert.notEqual(newLock.runId, oldLock.runId);

  finishOldRequest();
  const oldResult = await oldRun;
  assert.equal(oldResult.status, "error");
  assert.match(oldResult.error, /lost its persistent lease/);
  assert.equal(JSON.parse(fs.readFileSync(lockFile, "utf8")).runId, newLock.runId);

  finishNewRequest();
  const newResult = await newRun;
  assert.equal(newResult.status, "completed");
  assert.equal(newResult.counts.skipped, 1);
  assert.equal(newResult.counts.proposalsCreated, 0);
  assert.equal(proposalWrites, 1);
  assert.equal(fs.existsSync(lockFile), false);
  assert.equal(
    newWorker.readState().alerts.some((alert) => alert.type === "stale-lock-recovered"),
    true
  );
});

test("replacement completion before the stale Graph request still prevents duplicate processing", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-lease-replacement-first-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  const current = { value: new Date("2026-10-30T12:00:00.000Z") };
  let finishOldRequest;
  let finishNewRequest;
  let proposalWrites = 0;
  let sourceWasProcessed = false;

  const processSource = () => {
    const duplicate = sourceWasProcessed;
    if (!duplicate) {
      sourceWasProcessed = true;
      proposalWrites += 1;
    }
    return {
      configured: true,
      checked: 1,
      processed: duplicate ? 0 : 1,
      skipped: duplicate ? 1 : 0,
      failed: 0,
      proposalsCreated: duplicate ? 0 : 1,
      continuationCursor: ""
    };
  };

  const oldWorker = createLeaseService({
    stateFile,
    lockFile,
    now: () => new Date(current.value),
    setIntervalImpl: () => ({ unref() {} }),
    clearIntervalImpl: () => {},
    checkForNewEmails: () => new Promise((resolve) => {
      finishOldRequest = () => resolve(processSource());
    })
  });
  const oldRun = oldWorker.runScheduled({
    user: { email: "old-worker@example.test", role: "master-editor" }
  });
  await new Promise((resolve) => setImmediate(resolve));
  const oldLock = JSON.parse(fs.readFileSync(lockFile, "utf8"));

  current.value = new Date("2026-10-30T12:01:01.000Z");
  const newWorker = createLeaseService({
    stateFile,
    lockFile,
    now: () => new Date(current.value),
    checkForNewEmails: () => new Promise((resolve) => {
      finishNewRequest = () => resolve(processSource());
    })
  });
  const newRun = newWorker.runScheduled({
    user: { email: "new-worker@example.test", role: "master-editor" }
  });
  await new Promise((resolve) => setImmediate(resolve));

  const newLock = JSON.parse(fs.readFileSync(lockFile, "utf8"));
  assert.notEqual(newLock.runId, oldLock.runId);

  finishNewRequest();
  const newResult = await newRun;
  assert.equal(newResult.status, "completed");
  assert.equal(newResult.counts.processed, 1);
  assert.equal(newResult.counts.proposalsCreated, 1);
  const completedState = newWorker.readState();
  assert.equal(completedState.status, "idle");
  assert.equal(completedState.lastRunId, newResult.runId);
  assert.equal(fs.existsSync(lockFile), false);

  finishOldRequest();
  const oldResult = await oldRun;
  assert.equal(oldResult.status, "error");
  assert.match(oldResult.error, /lost its persistent lease/);
  assert.equal(oldResult.counts.processed, 0);
  assert.equal(oldResult.counts.proposalsCreated, 0);
  assert.equal(proposalWrites, 1);
  assert.equal(fs.existsSync(lockFile), false);
  assert.deepEqual(newWorker.readState(), completedState);
});

test("a competing worker fails closed while the active lease heartbeat is temporarily unreadable", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-lease-heartbeat-write-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  const current = { value: new Date("2026-10-30T12:00:00.000Z") };
  let finishActiveRequest;

  const activeWorker = createLeaseService({
    stateFile,
    lockFile,
    now: () => new Date(current.value),
    checkForNewEmails: () => new Promise((resolve) => {
      finishActiveRequest = () => resolve({
        configured: true,
        checked: 0,
        processed: 0,
        skipped: 0,
        failed: 0,
        proposalsCreated: 0,
        continuationCursor: ""
      });
    })
  });
  const activeRun = activeWorker.runScheduled({
    user: { email: "active-worker@example.test", role: "master-editor" }
  });
  await new Promise((resolve) => setImmediate(resolve));

  current.value = new Date("2026-10-30T12:00:20.000Z");
  let simulateHeartbeatWriteWindow = true;
  const fsImpl = {
    ...fs,
    readFileSync(file, ...args) {
      if (file === lockFile && simulateHeartbeatWriteWindow) {
        simulateHeartbeatWriteWindow = false;
        return "{";
      }
      return fs.readFileSync(file, ...args);
    }
  };
  const competingWorker = createAiEmailPollingService({
    stateFile,
    lockFile,
    readJsonFile: (file, fallback) => {
      try {
        return JSON.parse(fs.readFileSync(file, "utf8"));
      } catch (error) {
        return fallback;
      }
    },
    writeJsonFile: (file, value) => fs.writeFileSync(file, JSON.stringify(value), "utf8"),
    enabled: true,
    lockTtlMinutes: 1,
    maxRetries: 0,
    retryBaseMs: 0,
    now: () => new Date(current.value),
    fsImpl,
    checkForNewEmails: async () => {
      throw new Error("a competing worker must not run while the active lease is fresh");
    }
  });

  const competingResult = await competingWorker.runScheduled({
    user: { email: "competing-worker@example.test", role: "master-editor" }
  });
  finishActiveRequest();
  const activeResult = await activeRun;

  assert.equal(competingResult.status, "locked");
  assert.equal(competingResult.httpStatus, 409);
  assert.equal(activeResult.status, "completed");
  assert.equal(fs.existsSync(lockFile), false);
});

test("an unreadable lease is recoverable after its filesystem age exceeds the TTL", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-corrupt-stale-lease-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  const current = new Date("2026-10-30T12:02:00.000Z");
  fs.writeFileSync(lockFile, "{", { encoding: "utf8", mode: 0o600 });

  const staleMtime = new Date("2026-10-30T12:00:00.000Z").getTime();
  const fsImpl = {
    ...fs,
    statSync(file, ...args) {
      const stats = fs.statSync(file, ...args);
      return file === lockFile ? { ...stats, mtimeMs: staleMtime } : stats;
    }
  };
  const service = createAiEmailPollingService({
    stateFile,
    lockFile,
    readJsonFile: (file, fallback) => {
      try {
        return JSON.parse(fs.readFileSync(file, "utf8"));
      } catch (error) {
        return fallback;
      }
    },
    writeJsonFile: (file, value) => fs.writeFileSync(file, JSON.stringify(value), "utf8"),
    enabled: true,
    lockTtlMinutes: 1,
    maxRetries: 0,
    retryBaseMs: 0,
    now: () => new Date(current),
    fsImpl,
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

  const result = await service.runScheduled({
    user: { email: "recovery-worker@example.test", role: "master-editor" }
  });

  assert.equal(result.status, "completed");
  assert.equal(
    service.readState().alerts.some((alert) => alert.type === "stale-lock-recovered"),
    true
  );
  assert.equal(fs.existsSync(lockFile), false);
});

test("an unreadable lease fails closed when its filesystem age cannot be inspected", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-uninspectable-lease-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  fs.writeFileSync(lockFile, "{", { encoding: "utf8", mode: 0o600 });
  let unlinkAttempted = false;

  const permissionError = Object.assign(new Error("permission denied"), { code: "EACCES" });
  const fsImpl = {
    ...fs,
    statSync(file, ...args) {
      if (file === lockFile) throw permissionError;
      return fs.statSync(file, ...args);
    },
    unlinkSync(file) {
      if (file === lockFile) unlinkAttempted = true;
      return fs.unlinkSync(file);
    }
  };
  const service = createAiEmailPollingService({
    stateFile,
    lockFile,
    readJsonFile: (file, fallback) => {
      try {
        return JSON.parse(fs.readFileSync(file, "utf8"));
      } catch (error) {
        return fallback;
      }
    },
    writeJsonFile: (file, value) => fs.writeFileSync(file, JSON.stringify(value), "utf8"),
    enabled: true,
    lockTtlMinutes: 1,
    now: () => new Date("2026-10-30T12:02:00.000Z"),
    fsImpl,
    checkForNewEmails: async () => {
      throw new Error("intake must not run while lease age is unknown");
    }
  });

  const result = await service.runScheduled({
    user: { email: "blocked-worker@example.test", role: "master-editor" }
  });

  assert.equal(result.status, "locked");
  assert.equal(result.httpStatus, 409);
  assert.equal(unlinkAttempted, false);
  assert.equal(fs.readFileSync(lockFile, "utf8"), "{");
});

test("stale recovery rechecks filesystem age when a heartbeat becomes unreadable", (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "ai-email-recovery-heartbeat-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const stateFile = path.join(directory, "polling-state.json");
  const lockFile = path.join(directory, "polling.lock");
  const recoveryFile = `${lockFile}.recovery`;
  const current = new Date("2026-11-01T12:00:00.000Z");
  const staleAt = new Date("2026-11-01T11:00:00.000Z");
  fs.writeFileSync(lockFile, JSON.stringify({
    runId: "active-worker",
    startedAt: staleAt.toISOString(),
    heartbeatAt: staleAt.toISOString()
  }) + "\n");
  fs.utimesSync(lockFile, staleAt, staleAt);
  let recoveryGuardOpened = false;
  let lockUnlinked = false;

  const fsImpl = {
    ...fs,
    openSync(file, flags, ...args) {
      const descriptor = fs.openSync(file, flags, ...args);
      if (file === recoveryFile && flags === "wx") {
        recoveryGuardOpened = true;
        fs.writeFileSync(lockFile, "", "utf8");
        fs.utimesSync(lockFile, current, current);
      }
      return descriptor;
    },
    unlinkSync(file) {
      if (file === lockFile) lockUnlinked = true;
      return fs.unlinkSync(file);
    }
  };
  const service = createAiEmailPollingService({
    stateFile,
    lockFile,
    readJsonFile: (file, fallback) => {
      try {
        return JSON.parse(fs.readFileSync(file, "utf8"));
      } catch (error) {
        return fallback;
      }
    },
    writeJsonFile: (file, value) => fs.writeFileSync(file, JSON.stringify(value), "utf8"),
    enabled: true,
    lockTtlMinutes: 1,
    now: () => new Date(current),
    fsImpl,
    checkForNewEmails: async () => ({ configured: true })
  });

  const result = service.tryAcquireLock("competing-worker");

  assert.equal(recoveryGuardOpened, true);
  assert.equal(result.acquired, false);
  assert.equal(result.recovered, false);
  assert.equal(lockUnlinked, false);
  assert.equal(fs.existsSync(lockFile), true);
  assert.equal(fs.existsSync(recoveryFile), false);
});

test("a crashed stale-recovery guard cannot block lease recovery forever", (t) => {
  const current = { value: new Date("2026-11-01T12:00:00.000Z") };
  const { service, lockFile } = createHarness(t, {
    current,
    lockTtlMinutes: 1,
    checkForNewEmails: async () => ({ configured: true })
  });
  const staleAt = new Date("2026-11-01T11:00:00.000Z");
  fs.writeFileSync(lockFile, JSON.stringify({
    runId: "stale-worker",
    startedAt: staleAt.toISOString(),
    heartbeatAt: staleAt.toISOString()
  }) + "\n");
  const recoveryFile = `${lockFile}.recovery`;
  fs.writeFileSync(recoveryFile, JSON.stringify({ runId: "crashed-recovery" }) + "\n");
  fs.utimesSync(recoveryFile, staleAt, staleAt);

  const result = service.tryAcquireLock("replacement-worker");

  assert.equal(result.acquired, true);
  assert.equal(result.recovered, true);
  assert.equal(JSON.parse(fs.readFileSync(lockFile, "utf8")).runId, "replacement-worker");
  assert.equal(fs.existsSync(recoveryFile), false);
  assert.deepEqual(
    fs.readdirSync(path.dirname(lockFile)).filter((name) => name.includes(".claim")),
    []
  );
});



