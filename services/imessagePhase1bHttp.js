"use strict";

const crypto = require("node:crypto");
const fs = require("node:fs");
const http = require("node:http");
const path = require("node:path");

const MAX_REQUEST_BYTES = 64 * 1024;
const MAX_REPLAY_ENTRIES = 1024;

function fixedChallenge(resourceMetadataUrl) {
  return `Bearer resource_metadata="${resourceMetadataUrl}", scope="messages.read"`;
}

function extractBearer(header) {
  if (typeof header !== "string") return null;
  const match = /^Bearer ([A-Za-z0-9_.-]+)$/.exec(header);
  return match ? match[1] : null;
}

function canonicalJson(value) {
  if (Array.isArray(value)) return `[${value.map(canonicalJson).join(",")}]`;
  if (value && typeof value === "object") {
    const properties = Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${canonicalJson(value[key])}`);
    return `{${properties.join(",")}}`;
  }
  return JSON.stringify(value);
}

function createReplayGuard({ now = () => Math.floor(Date.now() / 1000) } = {}) {
  const fingerprints = new Map();

  return Object.freeze({
    consume(token, body, expiresAt) {
      const current = now();
      for (const [key, expiry] of fingerprints) {
        if (expiry < current) fingerprints.delete(key);
      }
      const fingerprint = crypto.createHash("sha256").update(token).update("\0").update(body).digest("hex");
      if (fingerprints.has(fingerprint)) return false;
      if (fingerprints.size >= MAX_REPLAY_ENTRIES) return false;
      fingerprints.set(fingerprint, expiresAt);
      return true;
    },
    clear() { fingerprints.clear(); },
    size() { return fingerprints.size; }
  });
}

function isLoopback(address) {
  return address === "127.0.0.1" || address === "::1" || address === "::ffff:127.0.0.1";
}

function createLocalMcpHttpServer({
  handler,
  tokenVerifier,
  issuer,
  resource,
  socketPath = null,
  tcp = null,
  logger = null,
  replayGuard = createReplayGuard()
}) {
  if (typeof handler !== "function" || !tokenVerifier || typeof tokenVerifier.verify !== "function") {
    throw new TypeError("Invalid local MCP HTTP configuration.");
  }
  if ((socketPath === null) === (tcp === null)) throw new TypeError("Configure exactly one local transport.");
  if (socketPath !== null && !path.isAbsolute(socketPath)) throw new TypeError("Unix socket path must be absolute.");
  if (tcp !== null && (!tcp || tcp.host !== "127.0.0.1" || !Number.isInteger(tcp.port))) {
    throw new TypeError("TCP transport must bind IPv4 loopback.");
  }
  const metadataPath = "/.well-known/oauth-protected-resource";
  const challenge = fixedChallenge(`${resource}${metadataPath}`);
  let revoked = false;

  function log(event, reason) {
    if (logger && typeof logger.info === "function") logger.info(event, reason ? { reason } : {});
  }

  function json(response, status, value, headers = {}) {
    const body = JSON.stringify(value);
    response.writeHead(status, {
      "cache-control": "no-store",
      "content-type": "application/json",
      ...headers
    });
    response.end(body);
  }

  const server = http.createServer((request, response) => {
    const remoteAddress = request.socket.remoteAddress;
    if (remoteAddress && !isLoopback(remoteAddress)) {
      log("imessage_phase1b_http_denied", "non_loopback");
      return json(response, 403, { error: "forbidden" });
    }
    if (request.method === "GET" && request.url === metadataPath) {
      return json(response, 200, {
        resource,
        authorization_servers: [issuer],
        scopes_supported: ["messages.read"]
      });
    }
    if (request.method !== "POST" || request.url !== "/mcp") {
      return json(response, 404, { error: "not_found" });
    }
    if (revoked) {
      log("imessage_phase1b_http_denied", "revoked");
      return json(response, 403, { error: "revoked" });
    }
    if (!/^application\/json(?:\s*;|$)/i.test(request.headers["content-type"] || "")) {
      return json(response, 415, { error: "unsupported_media_type" });
    }

    const chunks = [];
    let length = 0;
    request.on("data", (chunk) => {
      length += chunk.length;
      if (length > MAX_REQUEST_BYTES) request.destroy();
      else chunks.push(chunk);
    });
    request.on("end", async () => {
      const raw = Buffer.concat(chunks);
      const token = extractBearer(request.headers.authorization);
      let identity;
      try {
        if (!token) throw new Error("missing");
        identity = tokenVerifier.verify(token);
      } catch (error) {
        log("imessage_phase1b_http_denied", error && error.reason ? error.reason : "missing_token");
        return json(response, 401, { error: "unauthorized" }, { "www-authenticate": challenge });
      }
      let rpcRequest;
      try {
        rpcRequest = JSON.parse(raw.toString("utf8"));
      } catch {
        return json(response, 400, { error: "invalid_json" });
      }
      if (!replayGuard.consume(token, canonicalJson(rpcRequest), identity.expiresAt)) {
        log("imessage_phase1b_http_denied", "replay");
        return json(response, 409, { error: "replayed_or_capacity_exceeded" });
      }
      const rpcResponse = await handler(rpcRequest);
      log("imessage_phase1b_http_completed");
      return json(response, 200, rpcResponse);
    });
    request.on("error", () => {
      if (!response.headersSent) json(response, 413, { error: "request_too_large" });
    });
  });

  return Object.freeze({
    async start() {
      if (socketPath !== null) {
        fs.mkdirSync(path.dirname(socketPath), { recursive: true, mode: 0o700 });
        try {
          fs.lstatSync(socketPath);
          throw new Error("Unix socket path already exists.");
        } catch (error) {
          if (error && error.code !== "ENOENT") throw error;
        }
      }
      await new Promise((resolve, reject) => {
        server.once("error", reject);
        server.listen(socketPath !== null ? socketPath : tcp, resolve);
      });
      if (socketPath !== null) fs.chmodSync(socketPath, 0o600);
    },
    async stop() {
      replayGuard.clear();
      if (server.listening) await new Promise((resolve) => server.close(resolve));
      if (socketPath !== null) {
        try { fs.unlinkSync(socketPath); } catch {}
      }
    },
    revoke() {
      revoked = true;
      replayGuard.clear();
      log("imessage_phase1b_http_revoked");
    },
    address() { return server.address(); },
    isRevoked() { return revoked; }
  });
}

module.exports = {
  MAX_REPLAY_ENTRIES,
  canonicalJson,
  createLocalMcpHttpServer,
  createReplayGuard,
  extractBearer,
  isLoopback
};
