const test = require("node:test");
const assert = require("node:assert/strict");
const { createExecutiveBriefDeliveryService, expectedSubject } = require("../services/executiveBriefDeliveryService");
const { createExecutiveBriefDeliveryStateService } = require("../services/executiveBriefDeliveryStateService");
const { createResendExecutiveBriefTransport, FIXED_EXECUTIVE_BRIEF_RECIPIENT } = require("../services/resendExecutiveBriefTransport");

const FIXED_NOW = new Date("2026-09-03T14:00:00.000Z");

function inputFor(briefType, overrides = {}) {
  return {
    briefType,
    subject: expectedSubject(briefType, "September 3, 2026"),
    html: "<h2>Executive Priorities</h2><p>Complete the operating review.</p>",
    text: "Executive Priorities\nComplete the operating review.",
    ...overrides
  };
}

function memoryState() {
  let stored = [];
  return {
    service: createExecutiveBriefDeliveryStateService({
      STATE_FILE: "executive-brief-deliveries.json",
      readJsonFile: () => stored,
      writeJsonFile: (file, value) => { stored = structuredClone(value); }
    }),
    read: () => stored
  };
}

function deliveryHarness(overrides = {}) {
  const state = memoryState();
  const calls = [];
  const audits = [];
  const service = createExecutiveBriefDeliveryService({
    enabled: true,
    stateService: state.service,
    transport: { send: async (payload) => { calls.push(payload); return { providerMessageId: "provider-1" }; } },
    now: () => FIXED_NOW,
    makeId: () => `delivery-${calls.length + 1}`,
    audit: (event) => audits.push(event),
    ...overrides
  });
  return { service, state, calls, audits };
}

test("morning brief is accepted for the fixed server-side recipient", async () => {
  const { service, calls, state } = deliveryHarness();
  const result = await service.sendExecutiveBrief(inputFor("morning"));
  assert.equal(result.status, "sent");
  assert.equal(result.briefType, "morning");
  assert.equal(calls.length, 1);
  assert.equal(state.read()[0].status, "sent");
});

test("evening brief is accepted for the fixed server-side recipient", async () => {
  const { service, calls } = deliveryHarness();
  const result = await service.sendExecutiveBrief(inputFor("evening"));
  assert.equal(result.briefType, "evening");
  assert.equal(calls.length, 1);
});

test("recipient, CC, BCC, reply, forward, attachment, and delivery fields are rejected", async () => {
  const forbidden = ["recipient", "to", "cc", "bcc", "replyTo", "forward", "attachments", "delivery"];
  for (const field of forbidden) {
    const { service, calls } = deliveryHarness();
    await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning", { [field]: "not-allowed" })), {
      code: "INVALID_INPUT"
    });
    assert.equal(calls.length, 0);
  }
});

test("invalid brief type, subject/type/date mismatch, and malformed content are rejected", async () => {
  const { service, calls } = deliveryHarness();
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("midday")), { code: "INVALID_BRIEF_TYPE" });
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning", { subject: inputFor("evening").subject })), { code: "INVALID_SUBJECT" });
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning", { subject: "BEAMAN VENTURES | MORNING EXECUTIVE BRIEF | September 2, 2026" })), { code: "INVALID_SUBJECT" });
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning", { html: "<script>bad()</script>" })), { code: "INVALID_HTML" });
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning", { text: "" })), { code: "INVALID_TEXT" });
  assert.equal(calls.length, 0);
});

test("duplicate and concurrent sends fail closed before a second provider call", async () => {
  let release;
  const pending = new Promise((resolve) => { release = resolve; });
  const state = memoryState();
  let calls = 0;
  const service = createExecutiveBriefDeliveryService({
    enabled: true,
    stateService: state.service,
    transport: { send: async () => { calls += 1; await pending; return { providerMessageId: "provider-1" }; } },
    now: () => FIXED_NOW,
    makeId: () => "delivery-1"
  });
  const first = service.sendExecutiveBrief(inputFor("morning"));
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning")), { code: "DUPLICATE_DELIVERY" });
  release();
  await first;
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning")), { code: "DUPLICATE_DELIVERY" });
  assert.equal(calls, 1);
});

test("provider failure is recorded without secrets or briefing contents and remains blocked", async () => {
  const secret = "re_super_secret_value";
  const state = memoryState();
  const audits = [];
  const service = createExecutiveBriefDeliveryService({
    enabled: true,
    stateService: state.service,
    transport: { send: async () => { throw Object.assign(new Error(`provider leaked ${secret}`), { code: "PROVIDER_FAILED" }); } },
    now: () => FIXED_NOW,
    makeId: () => "delivery-1",
    audit: (event) => audits.push(event)
  });
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning")), {
    code: "DELIVERY_FAILED",
    message: "Executive brief delivery failed closed."
  });
  const serialized = JSON.stringify({ audits, state: state.read() });
  assert.equal(serialized.includes(secret), false);
  assert.equal(serialized.includes("Complete the operating review"), false);
  assert.equal(state.read()[0].status, "failed-closed");
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning")), { code: "DUPLICATE_DELIVERY" });
});

test("delivery is disabled by default", async () => {
  const state = memoryState();
  const service = createExecutiveBriefDeliveryService({
    stateService: state.service,
    transport: { send: async () => assert.fail("transport must not be called") },
    now: () => FIXED_NOW
  });
  await assert.rejects(() => service.sendExecutiveBrief(inputFor("morning")), { code: "DELIVERY_DISABLED" });
});

test("Resend transport fixes the recipient and omits all unsupported delivery fields", async () => {
  let request;
  const transport = createResendExecutiveBriefTransport({
    apiKey: "re_secret_test_value",
    fromEmail: "briefs@beamanventures.com",
    fetchImpl: async (url, options) => {
      request = { url, options };
      return { ok: true, status: 200, json: async () => ({ id: "resend-1" }) };
    }
  });
  await transport.send({ subject: "subject", html: "<p>html</p>", text: "text", idempotencyKey: "key" });
  const payload = JSON.parse(request.options.body);
  assert.deepEqual(payload.to, [FIXED_EXECUTIVE_BRIEF_RECIPIENT]);
  assert.equal(FIXED_EXECUTIVE_BRIEF_RECIPIENT, "tyler@beamanventures.com");
  for (const field of ["cc", "bcc", "reply_to", "attachments", "recipient"]) {
    assert.equal(Object.prototype.hasOwnProperty.call(payload, field), false);
  }
  assert.equal(request.options.headers["Idempotency-Key"], "key");
});
