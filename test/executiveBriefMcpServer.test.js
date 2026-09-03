const test = require("node:test");
const assert = require("node:assert/strict");
const { createExecutiveBriefAuthService, createUnconfiguredExecutiveBriefAuthService } = require("../services/executiveBriefAuthService");
const { EXECUTIVE_BRIEF_TOOL, TOOL_NAME, createExecutiveBriefMcpServer } = require("../mcp/executiveBriefMcpServer");

test("MCP exposes exactly one narrowly scoped write tool", async () => {
  const server = createExecutiveBriefMcpServer({
    authService: createExecutiveBriefAuthService({ tokenVerifier: async () => ({ active: true, scope: "executive_brief:send" }) }),
    deliveryService: { sendExecutiveBrief: async () => assert.fail("delivery should not run") }
  });
  const response = await server.handle({
    headers: { authorization: "Bearer test-token" },
    body: { jsonrpc: "2.0", id: 1, method: "tools/list", params: {} }
  });
  assert.equal(response.body.result.tools.length, 1);
  assert.equal(response.body.result.tools[0].name, TOOL_NAME);
  assert.equal(TOOL_NAME, "send_executive_brief");
  assert.equal(EXECUTIVE_BRIEF_TOOL.inputSchema.additionalProperties, false);
  assert.deepEqual(Object.keys(EXECUTIVE_BRIEF_TOOL.inputSchema.properties).sort(), ["briefType", "html", "subject", "text"]);
  assert.deepEqual(EXECUTIVE_BRIEF_TOOL.inputSchema.properties.briefType.enum, ["morning", "evening"]);
  assert.deepEqual(EXECUTIVE_BRIEF_TOOL.annotations, { readOnlyHint: false, destructiveHint: true, openWorldHint: true });
});

test("provider-neutral auth requires a valid token with the dedicated send scope", async () => {
  const seen = [];
  const auth = createExecutiveBriefAuthService({ tokenVerifier: async (token) => { seen.push(token); return { active: true, scopes: ["executive_brief:send"], subject: "agent-owned" }; } });
  assert.deepEqual(await auth.authorize({ authorization: "Bearer opaque-token" }), { authorized: true, principalId: "agent-owned" });
  assert.deepEqual(seen, ["opaque-token"]);
  const insufficient = createExecutiveBriefAuthService({ tokenVerifier: async () => ({ active: true, scope: "other:scope" }) });
  await assert.rejects(() => insufficient.authorize({ authorization: "Bearer opaque-token" }), { code: "AUTH_SCOPE_REQUIRED" });
  await assert.rejects(() => auth.authorize({}), { code: "AUTH_REQUIRED" });
});

test("runtime auth fails closed until a managed provider is configured", async () => {
  const server = createExecutiveBriefMcpServer({
    authService: createUnconfiguredExecutiveBriefAuthService(),
    deliveryService: { sendExecutiveBrief: async () => assert.fail("delivery should not run") }
  });
  const response = await server.handle({
    headers: { authorization: "Bearer should-never-be-accepted" },
    body: { jsonrpc: "2.0", id: 2, method: "tools/list", params: {} }
  });
  assert.equal(response.statusCode, 503);
  assert.equal(JSON.stringify(response).includes("should-never-be-accepted"), false);
});

test("MCP call forwards only tool arguments and returns no secret or recipient", async () => {
  let received;
  const server = createExecutiveBriefMcpServer({
    authService: createExecutiveBriefAuthService({ tokenVerifier: async () => ({ active: true, scope: "executive_brief:send" }) }),
    deliveryService: { sendExecutiveBrief: async (args) => { received = args; return { status: "sent", deliveryId: "delivery-1", briefType: "morning", briefingDate: "2026-09-03" }; } }
  });
  const args = { briefType: "morning", subject: "subject", html: "<p>brief</p>", text: "brief" };
  const response = await server.handle({
    headers: { authorization: "Bearer secret-token" },
    body: { jsonrpc: "2.0", id: 3, method: "tools/call", params: { name: TOOL_NAME, arguments: args } }
  });
  assert.deepEqual(received, args);
  assert.equal(response.statusCode, 200);
  assert.equal(JSON.stringify(response).includes("secret-token"), false);
  assert.equal(JSON.stringify(response).includes("tyler@beamanventures.com"), false);
});
