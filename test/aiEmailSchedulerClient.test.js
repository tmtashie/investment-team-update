const test = require("node:test");
const assert = require("node:assert/strict");
const {
  runSchedulerTrigger,
  schedulerUrl
} = require("../bin/run-ai-email-scheduler");

function response(status, payload) {
  return {
    status,
    async text() {
      return JSON.stringify(payload);
    }
  };
}

test("scheduler client builds only the fixed trigger path on an HTTPS origin", () => {
  assert.equal(
    schedulerUrl("https://investment-team-updates.onrender.com/ignored?token=bad"),
    "https://investment-team-updates.onrender.com/api/ai-email-intake/scheduled-check"
  );
  assert.throws(() => schedulerUrl("http://investment-team-updates.onrender.com"), /HTTPS/);
});

test("scheduler client sends the bearer token only in the authorization header", async () => {
  const calls = [];
  const token = "0123456789abcdef0123456789abcdef";
  const result = await runSchedulerTrigger({
    baseUrl: "https://investment-team-updates.onrender.com",
    token,
    fetchImpl: async (url, options) => {
      calls.push({ url, options });
      return response(202, {
        status: "catching-up",
        backlogRemaining: true,
        counts: { checked: 50, proposalsCreated: 3 }
      });
    }
  });

  assert.equal(result.ok, true);
  assert.equal(result.status, "catching-up");
  assert.equal(calls[0].url.includes(token), false);
  assert.equal(calls[0].options.headers.Authorization, `Bearer ${token}`);
  assert.equal(JSON.stringify(result).includes(token), false);
});

test("scheduler client treats lock and retry deferrals as safe no-op deliveries", async () => {
  const token = "0123456789abcdef0123456789abcdef";
  for (const status of [409, 429]) {
    const result = await runSchedulerTrigger({
      baseUrl: "https://investment-team-updates.onrender.com",
      token,
      fetchImpl: async () => response(status, { status: status === 409 ? "locked" : "deferred" })
    });
    assert.equal(result.ok, true);
    assert.equal(result.httpStatus, status);
  }
});

test("scheduler client fails closed on short credentials and server errors", async () => {
  await assert.rejects(
    runSchedulerTrigger({
      baseUrl: "https://investment-team-updates.onrender.com",
      token: "short",
      fetchImpl: async () => response(200, {})
    }),
    /at least 32/
  );
  await assert.rejects(
    runSchedulerTrigger({
      baseUrl: "https://investment-team-updates.onrender.com",
      token: "0123456789abcdef0123456789abcdef",
      fetchImpl: async () => response(503, { error: "disabled" })
    }),
    /HTTP 503/
  );
});
