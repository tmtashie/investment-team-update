"use strict";

const crypto = require("node:crypto");
const http = require("node:http");
const { URL, URLSearchParams } = require("node:url");

const MAX_BODY_BYTES = 32 * 1024;
const PROOF_MAX_AGE_SECONDS = 30;
const CODE_LIFETIME_SECONDS = 60;
const TOKEN_LIFETIME_SECONDS = 120;

function oauthError(code, status = 400, headers = {}) {
  const error = new Error(code);
  error.oauthCode = code;
  error.status = status;
  error.headers = headers;
  return error;
}

function encodeJson(value) {
  return Buffer.from(JSON.stringify(value)).toString("base64url");
}

function decodeJson(value) {
  try { return JSON.parse(Buffer.from(value, "base64url").toString("utf8")); }
  catch { throw oauthError("invalid_dpop_proof", 401); }
}

function publicJwkThumbprint(jwk) {
  let value;
  if (jwk && jwk.kty === "EC" && jwk.crv === "P-256") {
    value = { crv: jwk.crv, kty: jwk.kty, x: jwk.x, y: jwk.y };
  } else if (jwk && jwk.kty === "OKP" && jwk.crv === "Ed25519") {
    value = { crv: jwk.crv, kty: jwk.kty, x: jwk.x };
  } else {
    throw oauthError("invalid_dpop_proof", 401);
  }
  if (Object.values(value).some((item) => typeof item !== "string") || jwk.d !== undefined) {
    throw oauthError("invalid_dpop_proof", 401);
  }
  return crypto.createHash("sha256").update(JSON.stringify(value)).digest("base64url");
}

function verifyJoseSignature(alg, jwk, input, signature) {
  let key;
  try { key = crypto.createPublicKey({ key: jwk, format: "jwk" }); }
  catch { throw oauthError("invalid_dpop_proof", 401); }
  const options = alg === "ES256" ? { key, dsaEncoding: "ieee-p1363" } : key;
  const algorithm = alg === "ES256" ? "sha256" : null;
  if (!crypto.verify(algorithm, Buffer.from(input, "ascii"), options, signature)) {
    throw oauthError("invalid_dpop_proof", 401);
  }
}

function parseJwt(value) {
  if (typeof value !== "string" || value.length > 8192) throw oauthError("invalid_dpop_proof", 401);
  const parts = value.split(".");
  if (parts.length !== 3 || parts.some((part) => !/^[A-Za-z0-9_-]+$/.test(part))) {
    throw oauthError("invalid_dpop_proof", 401);
  }
  return {
    header: decodeJson(parts[0]),
    claims: decodeJson(parts[1]),
    input: `${parts[0]}.${parts[1]}`,
    signature: Buffer.from(parts[2], "base64url")
  };
}

function createSingleUseCache({ now }) {
  const entries = new Map();
  return Object.freeze({
    consume(key, expiresAt) {
      const current = now();
      for (const [candidate, expiry] of entries) if (expiry < current) entries.delete(candidate);
      if (entries.has(key) || entries.size >= 1024) return false;
      entries.set(key, expiresAt);
      return true;
    },
    clear() { entries.clear(); }
  });
}

function verifyDpopProof({ proof, method, targetUrl, accessToken = null, nonce, now, cache }) {
  const parsed = parseJwt(proof);
  const { header, claims } = parsed;
  if (header.typ !== "dpop+jwt" || !["ES256", "EdDSA"].includes(header.alg) || !header.jwk) {
    throw oauthError("invalid_dpop_proof", 401);
  }
  verifyJoseSignature(header.alg, header.jwk, parsed.input, parsed.signature);
  const current = now();
  if (
    claims.htm !== method || claims.htu !== targetUrl ||
    !Number.isInteger(claims.iat) || Math.abs(current - claims.iat) > PROOF_MAX_AGE_SECONDS ||
    typeof claims.jti !== "string" || claims.jti.length < 8 || claims.jti.length > 200 ||
    claims.nonce !== nonce
  ) throw oauthError("invalid_dpop_proof", 401);
  if (accessToken !== null) {
    const expected = crypto.createHash("sha256").update(accessToken, "ascii").digest("base64url");
    if (claims.ath !== expected) throw oauthError("invalid_dpop_proof", 401);
  } else if (claims.ath !== undefined) {
    throw oauthError("invalid_dpop_proof", 401);
  }
  const thumbprint = publicJwkThumbprint(header.jwk);
  const replayKey = `${thumbprint}\0${claims.jti}`;
  if (!cache.consume(replayKey, current + PROOF_MAX_AGE_SECONDS)) throw oauthError("use_dpop_nonce", 401);
  return Object.freeze({
    thumbprint,
    fields: Object.freeze({ htm: true, htu: true, iat: true, jti: true, ath: accessToken !== null, nonce: true })
  });
}

function createMcpDpopDiscoveryProbe({
  issuer,
  resource,
  clientId,
  redirectUris,
  principals = new Map([["alice", "synthetic-alice"], ["tyler", "synthetic-tyler"]]),
  now = () => Math.floor(Date.now() / 1000),
  randomBytes = crypto.randomBytes,
  logger = null
}) {
  if (![issuer, resource, clientId].every((value) => typeof value === "string" && value.startsWith("https://"))) {
    throw new TypeError("Probe URLs and client ID must be fixed HTTPS identifiers.");
  }
  if (!(redirectUris instanceof Set) || redirectUris.size === 0 || !(principals instanceof Map) || principals.size !== 2) {
    throw new TypeError("Probe requires fixed redirect URIs and exactly two synthetic principals.");
  }
  const signingKeys = crypto.generateKeyPairSync("ed25519");
  const publicJwk = signingKeys.publicKey.export({ format: "jwk" });
  const codes = new Map();
  const revokedSubjects = new Set();
  const proofCache = createSingleUseCache({ now });
  let tokenNonce = randomBytes(24).toString("base64url");
  let resourceNonce = randomBytes(24).toString("base64url");
  const tokenUrl = `${issuer}/token`;
  const mcpUrl = `${resource}/mcp`;

  function log(event, metadata = {}) {
    if (logger && typeof logger.info === "function") logger.info(event, metadata);
  }

  function rotateNonce(kind) {
    const next = randomBytes(24).toString("base64url");
    if (kind === "token") tokenNonce = next;
    else resourceNonce = next;
    return next;
  }

  function signAccessToken(claims) {
    const header = encodeJson({ alg: "EdDSA", kid: "probe-key", typ: "at+jwt" });
    const payload = encodeJson(claims);
    const input = `${header}.${payload}`;
    const signature = crypto.sign(null, Buffer.from(input, "ascii"), signingKeys.privateKey).toString("base64url");
    return `${input}.${signature}`;
  }

  function verifyAccessToken(token) {
    const parsed = parseJwt(token);
    if (parsed.header.alg !== "EdDSA" || parsed.header.typ !== "at+jwt" || parsed.header.kid !== "probe-key") {
      throw oauthError("invalid_token", 401);
    }
    if (!crypto.verify(null, Buffer.from(parsed.input, "ascii"), signingKeys.publicKey, parsed.signature)) {
      throw oauthError("invalid_token", 401);
    }
    const c = parsed.claims;
    if (
      c.iss !== issuer || c.aud !== resource || c.client_id !== clientId || c.scope !== "probe.discovery" ||
      !principals.has(c.principal_slot) || principals.get(c.principal_slot) !== c.sub || revokedSubjects.has(c.sub) ||
      !Number.isInteger(c.iat) || !Number.isInteger(c.exp) || c.iat > now() + 2 || c.exp <= now() ||
      !c.cnf || typeof c.cnf.jkt !== "string"
    ) throw oauthError("invalid_token", 401);
    return c;
  }

  function issueAuthorizationCode(params) {
    if (
      !principals.has(params.principal_slot) || params.client_id !== clientId ||
      !redirectUris.has(params.redirect_uri) || params.response_type !== "code" ||
      params.code_challenge_method !== "S256" || !/^[A-Za-z0-9_-]{43,128}$/.test(params.code_challenge || "") ||
      params.resource !== resource || params.scope !== "probe.discovery" || typeof params.state !== "string"
    ) throw oauthError("invalid_request");
    const code = randomBytes(32).toString("base64url");
    codes.set(code, Object.freeze({
      subject: principals.get(params.principal_slot),
      principalSlot: params.principal_slot,
      clientId,
      redirectUri: params.redirect_uri,
      codeChallenge: params.code_challenge,
      expiresAt: now() + CODE_LIFETIME_SECONDS
    }));
    log("probe_authorization_code_issued", { principal_slot: params.principal_slot });
    return code;
  }

  function exchangeAuthorizationCode(params, dpopProof) {
    const record = codes.get(params.code);
    if (
      !record || record.expiresAt < now() || params.grant_type !== "authorization_code" ||
      params.client_id !== record.clientId || params.redirect_uri !== record.redirectUri ||
      crypto.createHash("sha256").update(params.code_verifier || "").digest("base64url") !== record.codeChallenge
    ) throw oauthError("invalid_grant");
    const dpop = verifyDpopProof({
      proof: dpopProof, method: "POST", targetUrl: tokenUrl, nonce: tokenNonce, now, cache: proofCache
    });
    codes.delete(params.code);
    const current = now();
    const token = signAccessToken({
      iss: issuer, sub: record.subject, aud: resource, client_id: clientId, scope: "probe.discovery",
      iat: current, exp: current + TOKEN_LIFETIME_SECONDS, jti: randomBytes(16).toString("base64url"),
      principal_slot: record.principalSlot, cnf: { jkt: dpop.thumbprint }
    });
    rotateNonce("token");
    log("probe_token_issued", { principal_slot: record.principalSlot, dpop_bound: true });
    return Object.freeze({ access_token: token, token_type: "DPoP", expires_in: TOKEN_LIFETIME_SECONDS, scope: "probe.discovery" });
  }

  function authenticateResource(token, proof) {
    const claims = verifyAccessToken(token);
    const dpop = verifyDpopProof({
      proof, method: "POST", targetUrl: mcpUrl, accessToken: token, nonce: resourceNonce, now, cache: proofCache
    });
    if (dpop.thumbprint !== claims.cnf.jkt) throw oauthError("invalid_token", 401);
    log("probe_mcp_authenticated", {
      principal_slot: claims.principal_slot,
      dpop_bound: true,
      proof_fields_valid: dpop.fields
    });
    return claims;
  }

  function handleMcp(request) {
    if (request && request.method === "initialize") {
      return { jsonrpc: "2.0", id: request.id, result: {
        protocolVersion: "2025-11-25", capabilities: { tools: {} },
        serverInfo: { name: "phase1b-dpop-discovery-probe", version: "1.0.0" }
      } };
    }
    if (request && request.method === "tools/list") {
      return { jsonrpc: "2.0", id: request.id, result: { tools: [] } };
    }
    if (request && request.method === "notifications/initialized") return null;
    return { jsonrpc: "2.0", id: request && request.id !== undefined ? request.id : null,
      error: { code: -32601, message: "Method not found" } };
  }

  return Object.freeze({
    metadata: Object.freeze({
      authorizationServer: {
        issuer, authorization_endpoint: `${issuer}/authorize`, token_endpoint: tokenUrl,
        jwks_uri: `${issuer}/.well-known/jwks.json`,
        response_types_supported: ["code"], grant_types_supported: ["authorization_code"],
        code_challenge_methods_supported: ["S256"], dpop_signing_alg_values_supported: ["ES256", "EdDSA"],
        token_endpoint_auth_methods_supported: ["none"], scopes_supported: ["probe.discovery"]
      },
      protectedResource: {
        resource, authorization_servers: [issuer], scopes_supported: ["probe.discovery"],
        dpop_bound_access_tokens_required: true
      },
      jwks: { keys: [{ ...publicJwk, kid: "probe-key", use: "sig", alg: "EdDSA" }] }
    }),
    issueAuthorizationCode,
    exchangeAuthorizationCode,
    authenticateResource,
    handleMcp,
    resourceNonce() { return resourceNonce; },
    tokenNonce() { return tokenNonce; },
    rotateFreshnessEpoch() { proofCache.clear(); rotateNonce("token"); rotateNonce("resource"); log("probe_freshness_epoch_rotated"); },
    revokePrincipal(slot) {
      if (!principals.has(slot)) throw new TypeError("Unknown synthetic principal.");
      revokedSubjects.add(principals.get(slot));
      log("probe_principal_revoked", { principal_slot: slot });
    },
    isRevoked(slot) { return principals.has(slot) && revokedSubjects.has(principals.get(slot)); }
  });
}

function readBody(request) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let length = 0;
    request.on("data", (chunk) => {
      length += chunk.length;
      if (length > MAX_BODY_BYTES) {
        reject(oauthError("request_too_large", 413));
        request.destroy();
      } else chunks.push(chunk);
    });
    request.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    request.on("error", reject);
  });
}

function createProbeHttpServer({ probe, issuer, resource, socketPath, logger = null }) {
  if (!probe || typeof probe.authenticateResource !== "function" || typeof socketPath !== "string") {
    throw new TypeError("Invalid probe HTTP configuration.");
  }
  const issuerUrl = new URL(issuer);
  const resourceUrl = new URL(resource);

  function send(response, status, value, headers = {}) {
    const body = typeof value === "string" ? value : JSON.stringify(value);
    response.writeHead(status, {
      "cache-control": "no-store",
      "content-type": typeof value === "string" ? "text/html; charset=utf-8" : "application/json",
      ...headers
    });
    response.end(body);
  }

  function dpopChallenge(response, status, error, nonce) {
    send(response, status, { error }, {
      "dpop-nonce": nonce,
      "www-authenticate": `DPoP error="${error}", resource_metadata="${resource}/.well-known/oauth-protected-resource"`
    });
  }

  function authorizationForm(query) {
    const hidden = [...query.entries()].map(([key, value]) =>
      `<input type="hidden" name="${escapeHtml(key)}" value="${escapeHtml(value)}">`).join("");
    return `<!doctype html><html><head><meta charset="utf-8"><meta name="referrer" content="no-referrer"><title>Synthetic principal</title></head><body><main><h1>Synthetic discovery-only principal</h1><form method="post" action="${escapeHtml(issuerUrl.pathname.replace(/\/$/, "") + "/authorize")}">${hidden}<button name="principal_slot" value="alice" type="submit">Synthetic Alice</button><button name="principal_slot" value="tyler" type="submit">Synthetic Tyler</button></form></main></body></html>`;
  }

  const server = http.createServer(async (request, response) => {
    try {
      const url = new URL(request.url, "http://localhost");
      if (request.method === "GET" && [
        "/.well-known/oauth-authorization-server",
        `${issuerUrl.pathname.replace(/\/$/, "")}/.well-known/oauth-authorization-server`
      ].includes(url.pathname)) return send(response, 200, probe.metadata.authorizationServer);
      if (request.method === "GET" && [
        "/.well-known/oauth-protected-resource",
        "/.well-known/oauth-protected-resource/mcp",
        `${resourceUrl.pathname.replace(/\/$/, "")}/.well-known/oauth-protected-resource`
      ].includes(url.pathname)) return send(response, 200, probe.metadata.protectedResource);
      if (request.method === "GET" && url.pathname === "/.well-known/jwks.json") {
        return send(response, 200, probe.metadata.jwks);
      }
      if (request.method === "GET" && url.pathname === issuerUrl.pathname.replace(/\/$/, "") + "/authorize") {
        return send(response, 200, authorizationForm(url.searchParams), {
          "content-security-policy": "default-src 'none'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'",
          "x-content-type-options": "nosniff"
        });
      }
      if (request.method === "POST" && url.pathname === issuerUrl.pathname.replace(/\/$/, "") + "/authorize") {
        const params = new URLSearchParams(await readBody(request));
        const code = probe.issueAuthorizationCode(Object.fromEntries(params));
        const redirect = new URL(params.get("redirect_uri"));
        redirect.searchParams.set("code", code);
        redirect.searchParams.set("state", params.get("state") || "");
        response.writeHead(303, { location: redirect.toString(), "cache-control": "no-store" });
        return response.end();
      }
      if (request.method === "POST" && url.pathname === issuerUrl.pathname.replace(/\/$/, "") + "/token") {
        try {
          const params = Object.fromEntries(new URLSearchParams(await readBody(request)));
          const result = probe.exchangeAuthorizationCode(params, request.headers.dpop);
          return send(response, 200, result, { "dpop-nonce": probe.tokenNonce() });
        } catch (error) {
          if (error.oauthCode === "invalid_dpop_proof" || error.oauthCode === "use_dpop_nonce") {
            return dpopChallenge(response, 400, "use_dpop_nonce", probe.tokenNonce());
          }
          throw error;
        }
      }
      if (request.method === "POST" && url.pathname === resourceUrl.pathname.replace(/\/$/, "") + "/mcp") {
        const authorization = /^DPoP ([A-Za-z0-9_.-]+)$/.exec(request.headers.authorization || "");
        if (!authorization || typeof request.headers.dpop !== "string") {
          return dpopChallenge(response, 401, "use_dpop_nonce", probe.resourceNonce());
        }
        try {
          probe.authenticateResource(authorization[1], request.headers.dpop);
        } catch (error) {
          return dpopChallenge(response, error.status || 401, error.oauthCode || "invalid_token", probe.resourceNonce());
        }
        let rpc;
        try { rpc = JSON.parse(await readBody(request)); }
        catch { return send(response, 400, { error: "invalid_json" }); }
        const result = probe.handleMcp(rpc);
        if (result === null) {
          response.writeHead(204, { "cache-control": "no-store" });
          return response.end();
        }
        return send(response, 200, result, { "dpop-nonce": probe.resourceNonce() });
      }
      return send(response, 404, { error: "not_found" });
    } catch (error) {
      if (logger && typeof logger.info === "function") logger.info("probe_http_denied", { reason: error.oauthCode || "invalid_request" });
      return send(response, error.status || 400, { error: error.oauthCode || "invalid_request" });
    }
  });

  return Object.freeze({
    async start() { await new Promise((resolve, reject) => { server.once("error", reject); server.listen(socketPath, resolve); }); },
    async stop() { if (server.listening) await new Promise((resolve) => server.close(resolve)); },
    address() { return server.address(); }
  });
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;"
  })[character]);
}

module.exports = {
  PROOF_MAX_AGE_SECONDS,
  createProbeHttpServer,
  createMcpDpopDiscoveryProbe,
  publicJwkThumbprint,
  verifyDpopProof
};
