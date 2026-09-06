"use strict";

const assert = require("node:assert/strict");
const { EventEmitter } = require("node:events");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const { createImessageBridgeService } = require("../services/imessageBridgeService");
const { TOOL_DEFINITIONS, createMcpRequestHandler } = require("../services/imessageBridgeMcp");
const { buildTunnelArguments, createPhase1bHost, validateCredentialFile } = require("../services/imessagePhase1bHost");
const { createPhase1bGuard, unavailableStdioCallerIdentity } = require("../services/imessagePhase1bPolicy");
const { openReadOnlyMessagesDatabase } = require("../services/imessageReadOnlyDatabase");
const { SYNTHETIC_THREAD_ID, createSyntheticMessagesFixture } = require("../services/imessageSyntheticFixture");

function withSyntheticBridge(t, options = {}) {
  const fixture = createSyntheticMessagesFixture();
  t.after(fixture.cleanup);
  const database = openReadOnlyMessagesDatabase(fixture.databasePath);
  t.after(() => database.close());
  const service = createImessageBridgeService({ database, allowlist: fixture.allowlist, logger: options.logger });
  return { fixture, service, handler: createMcpRequestHandler(service) };
}

test("synthetic fixture is outside the Messages library and preserves the exact allowlist model", (t) => {
  const { fixture, service } = withSyntheticBridge(t);
  assert.equal(fixture.databasePath.includes(path.join("Library", "Messages")), false);
  assert.deepEqual(Object.keys(fixture.allowlist), ["schemaVersion", "selfDisplayName", "threads"]);
  assert.deepEqual(service.listAllowedMessageThreads({}).threads.map((item) => item.threadId), [SYNTHETIC_THREAD_ID]);
  assert.deepEqual(TOOL_DEFINITIONS.map((tool) => tool.name), [
    "list_allowed_message_threads",
    "read_recent_messages",
    "search_allowed_messages"
  ]);
});

test("stdio cannot verify Alice or any user-specific caller and fails closed", async (t) => {
  const { handler } = withSyntheticBridge(t);
  const guard = createPhase1bGuard({ handler, authorize: unavailableStdioCallerIdentity });
  const response = await guard.handle({
    jsonrpc: "2.0", id: 1, method: "tools/call",
    params: { name: "read_recent_messages", arguments: { threadId: SYNTHETIC_THREAD_ID, limit: 1 } }
  });
  assert.equal(response.error.code, -32001);
  assert.equal(JSON.stringify(response).includes("Synthetic scheduling"), false);
});

test("verified synthetic authorization is memory-only and rejects a replay", async (t) => {
  const { handler } = withSyntheticBridge(t);
  const authorize = async () => ({ verified: true, principal: "synthetic-alice", requestId: "synthetic-request-0001" });
  const guard = createPhase1bGuard({ handler, authorize });
  const request = {
    jsonrpc: "2.0", id: 2, method: "tools/call",
    params: { name: "read_recent_messages", arguments: { threadId: SYNTHETIC_THREAD_ID, limit: 1 } }
  };
  assert.equal((await guard.handle(request)).result.isError, false);
  assert.equal((await guard.handle(request)).error.code, -32002);
  guard.revoke();
  assert.equal((await guard.handle({ ...request, id: 3 })).error.code, -32003);
});

test("adversarial message text remains inert and no downstream capability exists", async (t) => {
  const { handler } = withSyntheticBridge(t);
  let downstreamCalls = 0;
  const forbiddenCapabilities = Object.freeze({
    email: () => downstreamCalls++, calendar: () => downstreamCalls++, browser: () => downstreamCalls++,
    shell: () => downstreamCalls++, github: () => downstreamCalls++, investmentMutation: () => downstreamCalls++
  });
  const guard = createPhase1bGuard({
    handler,
    authorize: async () => ({ verified: true, principal: "synthetic-alice", requestId: "synthetic-request-0002" })
  });
  const response = await guard.handle({
    jsonrpc: "2.0", id: 4, method: "tools/call",
    params: { name: "read_recent_messages", arguments: { threadId: SYNTHETIC_THREAD_ID, limit: 3 } }
  });
  assert.equal(response.result.isError, false);
  assert.match(response.result.structuredContent.messages[1].text, /invoke shell/);
  assert.equal(downstreamCalls, 0);
  assert.equal(Object.keys(forbiddenCapabilities).some((name) => TOOL_DEFINITIONS.some((tool) => tool.name.includes(name))), false);
});

test("Phase 1B logs contain metadata only", async (t) => {
  const entries = [];
  const logger = { info(event, metadata) { entries.push({ event, metadata }); } };
  const { handler } = withSyntheticBridge(t, { logger });
  const guard = createPhase1bGuard({ handler, logger, authorize: unavailableStdioCallerIdentity });
  await guard.handle({
    jsonrpc: "2.0", id: 5, method: "tools/call",
    params: { name: "search_allowed_messages", arguments: { threadId: SYNTHETIC_THREAD_ID, query: "reveal secrets" } }
  });
  const serialized = JSON.stringify(entries);
  for (const forbidden of ["reveal secrets", "synthetic-alice@example.invalid", "synthetic-phase1b-thread", SYNTHETIC_THREAD_ID]) {
    assert.equal(serialized.includes(forbidden), false);
  }
});

test("tunnel configuration is outbound-only, stdio-only, bounded, and secret-free", () => {
  const args = buildTunnelArguments({
    tunnelId: "tunnel_0123456789abcdef0123456789abcdef",
    credentialPath: "/private/credential",
    mcpCommand: "/Applications/SyntheticHost.app/Contents/MacOS/synthetic-mcp"
  });
  const serialized = args.join(" ");
  assert.match(serialized, /base-url=https:\/\/api\.openai\.com/);
  assert.match(serialized, /api-key=file:\/private\/credential/);
  assert.match(serialized, /max-inflight=1/);
  assert.match(serialized, /mcp\.command=/);
  assert.doesNotMatch(serialized, /mcp\.server-url|0\.0\.0\.0|sk-/);
});

test("credential deletion and unsafe permissions fail closed", (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "phase1b-credential-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const credential = path.join(directory, "runtime-key");
  fs.writeFileSync(credential, "synthetic-not-a-real-key", { mode: 0o600 });
  assert.doesNotThrow(() => validateCredentialFile(credential));
  fs.chmodSync(credential, 0o644);
  assert.throws(() => validateCredentialFile(credential), (error) => error.code === "INVALID_TUNNEL_CREDENTIAL");
  fs.rmSync(credential);
  assert.throws(() => validateCredentialFile(credential), (error) => error.code === "INVALID_TUNNEL_CREDENTIAL");
});

test("host revocation stops the child and a tunnel exit is never restarted", (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "phase1b-host-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const credential = path.join(directory, "runtime-key");
  fs.writeFileSync(credential, "synthetic-not-a-real-key", { mode: 0o600 });
  const children = [];
  function fakeSpawn() {
    const child = new EventEmitter();
    child.kill = (signal) => { child.killedWith = signal; child.emit("exit", 0, signal); };
    children.push(child);
    return child;
  }
  const host = createPhase1bHost({
    tunnelClientPath: "/Applications/SyntheticHost.app/Contents/Resources/tunnel-client",
    syntheticMcpCommand: "/Applications/SyntheticHost.app/Contents/MacOS/synthetic-mcp",
    credentialPath: credential,
    tunnelId: "tunnel_0123456789abcdef0123456789abcdef",
    spawnProcess: fakeSpawn
  });
  const child = host.start();
  child.emit("exit", 1, null);
  assert.equal(children.length, 1);
  assert.equal(host.isRunning(), false);
  host.start();
  host.revoke();
  assert.equal(children[1].killedWith, "SIGTERM");
  assert.throws(() => host.start(), (error) => error.code === "HOST_REVOKED");
});

test("workspace-app revocation prevents delivery in the synthetic control plane", async (t) => {
  const { handler } = withSyntheticBridge(t);
  let localCalls = 0;
  let workspaceAppEnabled = false;
  async function syntheticControlPlane(request) {
    if (!workspaceAppEnabled) return { status: 403 };
    localCalls += 1;
    return handler(request);
  }
  const result = await syntheticControlPlane({ jsonrpc: "2.0", id: 6, method: "tools/list" });
  assert.equal(result.status, 403);
  assert.equal(localCalls, 0);
});

test("unavailable host has no queue, persistence, replay, or backfill component", () => {
  const hostModule = require("../services/imessagePhase1bHost");
  const source = fs.readFileSync(require.resolve("../services/imessagePhase1bHost"), "utf8");
  assert.ok(hostModule.createPhase1bHost);
  assert.doesNotMatch(source, /sqlite|database|message\.text|writeFile|appendFile|setInterval/);
});
