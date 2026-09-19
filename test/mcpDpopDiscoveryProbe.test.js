"use strict";

const assert = require("node:assert/strict");
const crypto = require("node:crypto");
const fs = require("node:fs");
const http = require("node:http");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const { createMcpDpopDiscoveryProbe, createProbeHttpServer, publicJwkThumbprint } = require("../services/mcpDpopDiscoveryProbe");

const ISSUER = "https://synthetic-idp.example.invalid";
const RESOURCE = "https://synthetic-probe.example.invalid";
const CLIENT_ID = "https://synthetic-client.example.invalid/metadata.json";
const REDIRECT_URI = "https://chatgpt.com/connector/oauth/callback";

function encode(value) { return Buffer.from(JSON.stringify(value)).toString("base64url"); }
function decodeClaims(token) { return JSON.parse(Buffer.from(token.split(".")[1], "base64url").toString("utf8")); }

function signDpop(keys, claims, alg = "EdDSA") {
  const jwk = keys.publicKey.export({ format: "jwk" });
  const header = encode({ typ: "dpop+jwt", alg, jwk });
  const payload = encode(claims);
  const input = `${header}.${payload}`;
  const signature = crypto.sign(null, Buffer.from(input, "ascii"), keys.privateKey).toString("base64url");
  return `${input}.${signature}`;
}

function setup(current = 2000, logger = null) {
  const now = { value: current };
  const probe = createMcpDpopDiscoveryProbe({
    issuer: ISSUER,
    resource: RESOURCE,
    clientId: CLIENT_ID,
    redirectUris: new Set([REDIRECT_URI]),
    now: () => now.value,
    logger
  });
  return { probe, now };
}

function issueCode(probe, slot, verifier) {
  return probe.issueAuthorizationCode({
    principal_slot: slot,
    client_id: CLIENT_ID,
    redirect_uri: REDIRECT_URI,
    response_type: "code",
    code_challenge_method: "S256",
    code_challenge: crypto.createHash("sha256").update(verifier).digest("base64url"),
    resource: RESOURCE,
    scope: "probe.discovery",
    state: "synthetic-state"
  });
}

function exchange(probe, slot, keys, tokenJti = `token-${slot}`) {
  const verifier = `verifier-${slot}-${"x".repeat(43)}`;
  const code = issueCode(probe, slot, verifier);
  const proof = signDpop(keys, {
    htm: "POST", htu: `${ISSUER}/token`, iat: 2000, jti: tokenJti, nonce: probe.tokenNonce()
  });
  return probe.exchangeAuthorizationCode({
    grant_type: "authorization_code", code, client_id: CLIENT_ID,
    redirect_uri: REDIRECT_URI, code_verifier: verifier
  }, proof);
}

function resourceProof(probe, keys, token, overrides = {}) {
  return signDpop(keys, {
    htm: "POST",
    htu: `${RESOURCE}/mcp`,
    iat: 2000,
    jti: crypto.randomUUID(),
    ath: crypto.createHash("sha256").update(token, "ascii").digest("base64url"),
    nonce: probe.resourceNonce(),
    ...overrides
  });
}

function httpRequest(socketPath, { method = "GET", url = "/", headers = {}, body = "" } = {}) {
  return new Promise((resolve, reject) => {
    const request = http.request({ socketPath, method, path: url, headers: {
      ...(body ? { "content-length": Buffer.byteLength(body) } : {}), ...headers
    } }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => resolve({
        status: response.statusCode, headers: response.headers,
        body: Buffer.concat(chunks).toString("utf8")
      }));
    });
    request.on("error", reject);
    request.end(body);
  });
}

test("two delegated synthetic principals receive distinct signed subjects", () => {
  const { probe } = setup();
  const aliceKeys = crypto.generateKeyPairSync("ed25519");
  const tylerKeys = crypto.generateKeyPairSync("ed25519");
  const alice = exchange(probe, "alice", aliceKeys);
  const tyler = exchange(probe, "tyler", tylerKeys);
  const aliceClaims = decodeClaims(alice.access_token);
  const tylerClaims = decodeClaims(tyler.access_token);
  assert.equal(alice.token_type, "DPoP");
  assert.equal(tyler.token_type, "DPoP");
  assert.equal(aliceClaims.sub, "synthetic-alice");
  assert.equal(tylerClaims.sub, "synthetic-tyler");
  assert.notEqual(aliceClaims.sub, tylerClaims.sub);
  assert.equal(aliceClaims.cnf.jkt, publicJwkThumbprint(aliceKeys.publicKey.export({ format: "jwk" })));
  assert.equal(tylerClaims.cnf.jkt, publicJwkThumbprint(tylerKeys.publicKey.export({ format: "jwk" })));
});

test("resource proof verifies method URI time ID token binding and nonce", () => {
  const events = [];
  const { probe } = setup(2000, { info(event, metadata) { events.push({ event, metadata }); } });
  const keys = crypto.generateKeyPairSync("ed25519");
  const token = exchange(probe, "alice", keys).access_token;
  const claims = probe.authenticateResource(token, resourceProof(probe, keys, token));
  assert.equal(claims.sub, "synthetic-alice");
  assert.deepEqual(events.at(-1).metadata.proof_fields_valid,
    { htm: true, htu: true, iat: true, jti: true, ath: true, nonce: true });
});

test("duplicate proof IDs and stale proofs are rejected", () => {
  const { probe } = setup();
  const keys = crypto.generateKeyPairSync("ed25519");
  const token = exchange(probe, "alice", keys).access_token;
  const first = resourceProof(probe, keys, token, { jti: "duplicate-proof-id" });
  probe.authenticateResource(token, first);
  assert.throws(() => probe.authenticateResource(token, first), /use_dpop_nonce/);
  assert.throws(() => probe.authenticateResource(token,
    resourceProof(probe, keys, token, { iat: 1900 })), /invalid_dpop_proof/);
});

test("wrong nonce and wrong token binding are rejected", () => {
  const { probe } = setup();
  const keys = crypto.generateKeyPairSync("ed25519");
  const token = exchange(probe, "alice", keys).access_token;
  assert.throws(() => probe.authenticateResource(token,
    resourceProof(probe, keys, token, { nonce: "stale-nonce" })), /invalid_dpop_proof/);
  assert.throws(() => probe.authenticateResource(token,
    resourceProof(probe, keys, token, { ath: "wrong-token-hash" })), /invalid_dpop_proof/);
});

test("freshness epoch rotation rejects a previously valid queued proof", () => {
  const { probe } = setup();
  const keys = crypto.generateKeyPairSync("ed25519");
  const token = exchange(probe, "alice", keys).access_token;
  const queued = resourceProof(probe, keys, token);
  probe.rotateFreshnessEpoch();
  assert.throws(() => probe.authenticateResource(token, queued), /invalid_dpop_proof/);
});

test("revoking Alice leaves Tyler valid", () => {
  const { probe } = setup();
  const aliceKeys = crypto.generateKeyPairSync("ed25519");
  const tylerKeys = crypto.generateKeyPairSync("ed25519");
  const alice = exchange(probe, "alice", aliceKeys).access_token;
  const tyler = exchange(probe, "tyler", tylerKeys).access_token;
  probe.revokePrincipal("alice");
  assert.throws(() => probe.authenticateResource(alice,
    resourceProof(probe, aliceKeys, alice)), /invalid_token/);
  assert.equal(probe.authenticateResource(tyler,
    resourceProof(probe, tylerKeys, tyler)).sub, "synthetic-tyler");
});

test("identity-only MCP discovery exposes no tools and rejects every call", () => {
  const { probe } = setup();
  assert.deepEqual(probe.handleMcp({ jsonrpc: "2.0", id: 1, method: "tools/list" }).result.tools, []);
  assert.equal(probe.handleMcp({ jsonrpc: "2.0", id: 2, method: "tools/call", params: {} }).error.code, -32601);
});

test("probe implementation cannot reach Messages or persistence APIs", () => {
  const source = fs.readFileSync(require.resolve("../services/mcpDpopDiscoveryProbe"), "utf8");
  assert.doesNotMatch(source, /chat\.db|Messages|sqlite|attachment|writeFile|appendFile|createWriteStream/);
});

test("categorical logs omit tokens subjects codes and proof IDs", () => {
  const events = [];
  const { probe } = setup(2000, { info(event, metadata) { events.push({ event, metadata }); } });
  const keys = crypto.generateKeyPairSync("ed25519");
  const token = exchange(probe, "alice", keys, "private-token-proof-id").access_token;
  const proof = resourceProof(probe, keys, token, { jti: "private-resource-proof-id" });
  probe.authenticateResource(token, proof);
  const serialized = JSON.stringify(events);
  for (const value of [token, proof, "synthetic-alice", "private-token-proof-id", "private-resource-proof-id"]) {
    assert.equal(serialized.includes(value), false);
  }
});

test("HTTP probe publishes OAuth metadata and challenges for DPoP without exposing tools", async (t) => {
  const { probe } = setup();
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "dpop-probe-"));
  const socketPath = path.join(directory, "probe.sock");
  const server = createProbeHttpServer({ probe, issuer: ISSUER, resource: RESOURCE, socketPath });
  await server.start();
  t.after(async () => { await server.stop(); fs.rmSync(directory, { recursive: true, force: true }); });

  const metadata = await httpRequest(socketPath, { url: "/.well-known/oauth-protected-resource" });
  assert.equal(metadata.status, 200);
  assert.equal(JSON.parse(metadata.body).dpop_bound_access_tokens_required, true);

  const denied = await httpRequest(socketPath, {
    method: "POST", url: "/mcp", headers: { "content-type": "application/json" },
    body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list" })
  });
  assert.equal(denied.status, 401);
  assert.equal(typeof denied.headers["dpop-nonce"], "string");
  assert.match(denied.headers["www-authenticate"], /^DPoP /);
});

test("HTTP probe authenticates DPoP discovery and rejects an exact proof replay", async (t) => {
  const { probe } = setup();
  const keys = crypto.generateKeyPairSync("ed25519");
  const token = exchange(probe, "alice", keys).access_token;
  const proof = resourceProof(probe, keys, token, { jti: "http-discovery-proof" });
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "dpop-probe-auth-"));
  const socketPath = path.join(directory, "probe.sock");
  const server = createProbeHttpServer({ probe, issuer: ISSUER, resource: RESOURCE, socketPath });
  await server.start();
  t.after(async () => { await server.stop(); fs.rmSync(directory, { recursive: true, force: true }); });
  const body = JSON.stringify({ jsonrpc: "2.0", id: 1, method: "tools/list" });
  const options = {
    method: "POST", url: "/mcp", body,
    headers: { authorization: `DPoP ${token}`, dpop: proof, "content-type": "application/json" }
  };
  const accepted = await httpRequest(socketPath, options);
  assert.equal(accepted.status, 200);
  assert.deepEqual(JSON.parse(accepted.body).result.tools, []);
  const replayed = await httpRequest(socketPath, options);
  assert.equal(replayed.status, 401);
  assert.match(replayed.headers["www-authenticate"], /use_dpop_nonce/);
});
