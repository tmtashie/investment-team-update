"use strict";

const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const { createImessageBridgeService } = require("../services/imessageBridgeService");
const { TOOL_DEFINITIONS, createMcpRequestHandler } = require("../services/imessageBridgeMcp");
const { buildHttpTunnelArguments, createPhase1bHttpApplication } = require("../services/imessagePhase1bHost");
const { createLocalMcpHttpServer, createReplayGuard } = require("../services/imessagePhase1bHttp");
const { createJwtAccessTokenVerifier } = require("../services/imessagePhase1bJwt");
const { openReadOnlyMessagesDatabase } = require("../services/imessageReadOnlyDatabase");
const { SYNTHETIC_THREAD_ID, createSyntheticMessagesFixture } = require("../services/imessageSyntheticFixture");

const ISSUER = "https://synthetic-idp.example.invalid";
const AUDIENCE = "https://synthetic-imessage-mcp.example.invalid";
const SUBJECT = "synthetic-tyler-subject";
const KID = "synthetic-key-1";

function encode(value) {
  return Buffer.from(JSON.stringify(value)).toString("base64url");
}

function signToken(privateKey, claims, header = { alg: "EdDSA", kid: KID, typ: "JWT" }) {
  const input = `${encode(header)}.${encode(claims)}`;
  return `${input}.${crypto.sign(null, Buffer.from(input), privateKey).toString("base64url")}`;
}

function request(socketPath, token, rpcRequest, extraHeaders = {}, serialize = JSON.stringify) {
  const body = serialize(rpcRequest);
  return new Promise((resolve, reject) => {
    const req = http.request({
      socketPath,
      path: "/mcp",
      method: "POST",
      headers: {
        authorization: token ? `Bearer ${token}` : "",
        "content-type": "application/json",
        "content-length": Buffer.byteLength(body),
        ...extraHeaders
      }
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => resolve({
        status: response.statusCode,
        headers: response.headers,
        body: JSON.parse(Buffer.concat(chunks).toString("utf8"))
      }));
    });
    req.on("error", reject);
    req.end(body);
  });
}

async function setup(t, { current = 2000, minimumIssuedAt = 1990, allowedSubject = SUBJECT, logger = null } = {}) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "phase1b-http-"));
  const socketPath = path.join(directory, "mcp.sock");
  const fixture = createSyntheticMessagesFixture();
  const database = openReadOnlyMessagesDatabase(fixture.databasePath);
  const keys = crypto.generateKeyPairSync("ed25519");
  const verifier = createJwtAccessTokenVerifier({
    issuer: ISSUER,
    audience: AUDIENCE,
    allowedSubjects: new Set([allowedSubject]),
    publicKeys: new Map([[KID, keys.publicKey]]),
    now: () => current,
    minimumIssuedAt
  });
  const service = createImessageBridgeService({ database, allowlist: fixture.allowlist });
  const server = createLocalMcpHttpServer({
    handler: createMcpRequestHandler(service),
    tokenVerifier: verifier,
    issuer: ISSUER,
    resource: AUDIENCE,
    socketPath,
    logger,
    replayGuard: createReplayGuard({ now: () => current })
  });
  await server.start();
  t.after(async () => {
    await server.stop();
    database.close();
    fixture.cleanup();
    fs.rmSync(directory, { recursive: true, force: true });
  });
  function claims(overrides = {}) {
    return {
      iss: ISSUER,
      aud: AUDIENCE,
      sub: SUBJECT,
      scope: "messages.read",
      iat: current,
      exp: current + 30,
      jti: crypto.randomUUID(),
      ...overrides
    };
  }
  return { socketPath, server, keys, claims };
}

function readRequest(id = 1) {
  return {
    jsonrpc: "2.0",
    id,
    method: "tools/call",
    params: { name: "read_recent_messages", arguments: { threadId: SYNTHETIC_THREAD_ID, limit: 1 } }
  };
}

test("Unix socket HTTP accepts a correctly signed token and preserves the three tools", async (t) => {
  const context = await setup(t);
  const token = signToken(context.keys.privateKey, context.claims());
  const response = await request(context.socketPath, token, {
    jsonrpc: "2.0", id: 1, method: "tools/list"
  });
  assert.equal(response.status, 200);
  assert.deepEqual(response.body.result.tools.map((tool) => tool.name), TOOL_DEFINITIONS.map((tool) => tool.name));
});

test("spoofed plain identity headers and a missing bearer token are rejected", async (t) => {
  const context = await setup(t);
  const response = await request(context.socketPath, null, readRequest(), {
    "x-user": SUBJECT,
    "x-caller-identity": "alice"
  });
  assert.equal(response.status, 401);
  assert.match(response.headers["www-authenticate"], /messages\.read/);
});

test("expired and not-yet-valid access tokens are rejected", async (t) => {
  const context = await setup(t);
  for (const claims of [
    context.claims({ iat: 1900, exp: 1999 }),
    context.claims({ iat: 2010, nbf: 2010, exp: 2030 })
  ]) {
    assert.equal((await request(context.socketPath, signToken(context.keys.privateKey, claims), readRequest())).status, 401);
  }
});

test("altered signatures and unknown signing keys are rejected", async (t) => {
  const context = await setup(t);
  const valid = signToken(context.keys.privateKey, context.claims());
  const parts = valid.split(".");
  const signature = Buffer.from(parts[2], "base64url");
  signature[0] ^= 0x01;
  const altered = `${parts[0]}.${parts[1]}.${signature.toString("base64url")}`;
  assert.equal((await request(context.socketPath, altered, readRequest())).status, 401);
  const unknownKid = signToken(context.keys.privateKey, context.claims(), { alg: "EdDSA", kid: "unknown" });
  assert.equal((await request(context.socketPath, unknownKid, readRequest())).status, 401);
});

test("wrong issuer, audience, user, and scope are rejected", async (t) => {
  const context = await setup(t);
  for (const override of [
    { iss: "https://wrong-issuer.invalid" },
    { aud: "https://wrong-audience.invalid" },
    { sub: "another-workspace-user" },
    { scope: "messages.write" }
  ]) {
    const token = signToken(context.keys.privateKey, context.claims(override));
    assert.equal((await request(context.socketPath, token, readRequest())).status, 401);
  }
});

test("an exact replay is rejected while a distinct authorized request can proceed", async (t) => {
  const context = await setup(t);
  const token = signToken(context.keys.privateKey, context.claims());
  const rpc = readRequest(7);
  assert.equal((await request(context.socketPath, token, rpc)).status, 200);
  assert.equal((await request(context.socketPath, token, rpc, {}, (value) => JSON.stringify({
    params: value.params,
    method: value.method,
    id: value.id,
    jsonrpc: value.jsonrpc
  }))).status, 409);
  assert.equal((await request(context.socketPath, token, readRequest(8))).status, 200);
});

test("queued delivery from before a host epoch is rejected after reconnect", async (t) => {
  const context = await setup(t, { current: 2000, minimumIssuedAt: 1995 });
  const queued = signToken(context.keys.privateKey, context.claims({ iat: 1990, exp: 2020 }));
  assert.equal((await request(context.socketPath, queued, readRequest())).status, 401);
});

test("revocation is immediate and clears replay state", async (t) => {
  const context = await setup(t);
  const token = signToken(context.keys.privateKey, context.claims());
  context.server.revoke();
  assert.equal((await request(context.socketPath, token, readRequest())).status, 403);
  assert.equal(context.server.isRevoked(), true);
});

test("HTTP authorization logs contain only fixed metadata", async (t) => {
  const entries = [];
  const context = await setup(t, { logger: { info(event, metadata) { entries.push({ event, metadata }); } } });
  const token = signToken(context.keys.privateKey, context.claims({ sub: "wrong-user", jti: "private-token-id" }));
  await request(context.socketPath, token, readRequest());
  const serialized = JSON.stringify(entries);
  for (const forbidden of [token, "wrong-user", "private-token-id", SYNTHETIC_THREAD_ID, "read_recent_messages"]) {
    assert.equal(serialized.includes(forbidden), false);
  }
});

test("tunnel uses a Unix socket HTTP binding with no public MCP listener", () => {
  const args = buildHttpTunnelArguments({
    tunnelId: "tunnel_0123456789abcdef0123456789abcdef",
    credentialPath: "/private/runtime-key",
    socketPath: "/private/imessage-bridge/mcp.sock"
  });
  const serialized = args.join(" ");
  assert.match(serialized, /mcp\.server-url=channel=main,url=http:\/\/localhost\/mcp,unix-socket=/);
  assert.doesNotMatch(serialized, /mcp\.command|0\.0\.0\.0|sk-/);
  assert.match(serialized, /max-inflight=1/);
});

test("host starts the local socket before the tunnel and revokes both immediately", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "phase1b-http-host-"));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const credentialPath = path.join(directory, "runtime-key");
  fs.writeFileSync(credentialPath, "synthetic-not-a-real-key", { mode: 0o600 });
  const events = [];
  const localServer = {
    async start() { events.push("server-start"); },
    async stop() { events.push("server-stop"); },
    revoke() { events.push("server-revoke"); }
  };
  const child = new (require("node:events").EventEmitter)();
  child.kill = (signal) => { events.push(`child-${signal}`); };
  const application = createPhase1bHttpApplication({
    localServer,
    tunnelClientPath: "/Applications/SyntheticHost.app/Contents/Resources/tunnel-client",
    tunnelId: "tunnel_0123456789abcdef0123456789abcdef",
    credentialPath,
    socketPath: path.join(directory, "mcp.sock"),
    spawnProcess(executable, args, options) {
      events.push("tunnel-start");
      assert.equal(options.shell, false);
      assert.deepEqual(options.stdio, ["ignore", "ignore", "ignore"]);
      assert.match(args.join(" "), /unix-socket=/);
      return child;
    }
  });
  await application.start();
  assert.deepEqual(events, ["server-start", "tunnel-start"]);
  await application.revoke();
  assert.deepEqual(events, ["server-start", "tunnel-start", "server-revoke", "child-SIGTERM", "server-stop"]);
  assert.equal(application.isRevoked(), true);
});

test("public TCP binds and ambiguous local transports are rejected", () => {
  const noop = async () => ({ jsonrpc: "2.0", id: 1, result: {} });
  const verifier = { verify() { return { expiresAt: 9999999999 }; } };
  assert.throws(() => createLocalMcpHttpServer({
    handler: noop,
    tokenVerifier: verifier,
    issuer: ISSUER,
    resource: AUDIENCE,
    tcp: { host: "0.0.0.0", port: 8080 }
  }), /loopback/);
  assert.throws(() => createLocalMcpHttpServer({
    handler: noop,
    tokenVerifier: verifier,
    issuer: ISSUER,
    resource: AUDIENCE,
    socketPath: "/tmp/mcp.sock",
    tcp: { host: "127.0.0.1", port: 0 }
  }), /exactly one/);
});

test("an existing Unix socket path is never unlinked or replaced", async (t) => {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "phase1b-existing-"));
  const socketPath = path.join(directory, "mcp.sock");
  fs.writeFileSync(socketPath, "do-not-replace");
  t.after(() => fs.rmSync(directory, { recursive: true, force: true }));
  const server = createLocalMcpHttpServer({
    handler: async () => ({ jsonrpc: "2.0", id: 1, result: {} }),
    tokenVerifier: { verify() { return { expiresAt: 9999999999 }; } },
    issuer: ISSUER,
    resource: AUDIENCE,
    socketPath
  });
  await assert.rejects(() => server.start(), /already exists/);
  assert.equal(fs.readFileSync(socketPath, "utf8"), "do-not-replace");
});
