#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const { createMcpDpopDiscoveryProbe, createProbeHttpServer } = require("../services/mcpDpopDiscoveryProbe");

if (process.env.PHASE1B_SYNTHETIC_ONLY !== "1") {
  process.stderr.write('{"event":"probe_start_failed","reason":"synthetic_mode_required"}\n');
  process.exit(1);
}

const required = ["PHASE1B_PROBE_ISSUER", "PHASE1B_PROBE_RESOURCE", "PHASE1B_PROBE_CLIENT_ID",
  "PHASE1B_PROBE_REDIRECT_URI", "PHASE1B_PROBE_SOCKET"];
if (required.some((name) => typeof process.env[name] !== "string" || process.env[name].length === 0)) {
  process.stderr.write('{"event":"probe_start_failed","reason":"invalid_configuration"}\n');
  process.exit(1);
}

const socketPath = path.resolve(process.env.PHASE1B_PROBE_SOCKET);
try {
  fs.lstatSync(socketPath);
  process.stderr.write('{"event":"probe_start_failed","reason":"socket_exists"}\n');
  process.exit(1);
} catch (error) {
  if (error.code !== "ENOENT") {
    process.stderr.write('{"event":"probe_start_failed","reason":"socket_unavailable"}\n');
    process.exit(1);
  }
}
fs.mkdirSync(path.dirname(socketPath), { recursive: true, mode: 0o700 });

const logger = Object.freeze({
  info(event, metadata = {}) {
    process.stderr.write(`${JSON.stringify({ event, ...metadata })}\n`);
  }
});
const probe = createMcpDpopDiscoveryProbe({
  issuer: process.env.PHASE1B_PROBE_ISSUER,
  resource: process.env.PHASE1B_PROBE_RESOURCE,
  clientId: process.env.PHASE1B_PROBE_CLIENT_ID,
  redirectUris: new Set([process.env.PHASE1B_PROBE_REDIRECT_URI]),
  logger
});
const server = createProbeHttpServer({
  probe,
  issuer: process.env.PHASE1B_PROBE_ISSUER,
  resource: process.env.PHASE1B_PROBE_RESOURCE,
  socketPath,
  logger
});

let stopping = false;
async function stop() {
  if (stopping) return;
  stopping = true;
  await server.stop();
  try { fs.unlinkSync(socketPath); } catch {}
  logger.info("probe_stopped");
}

server.start().then(() => {
  fs.chmodSync(socketPath, 0o600);
  logger.info("probe_started", { transport: "unix_socket", tools: 0 });
}).catch(() => {
  logger.info("probe_start_failed", { reason: "listen_failed" });
  process.exitCode = 1;
});

process.on("SIGUSR1", () => probe.revokePrincipal("alice"));
process.on("SIGUSR2", () => probe.rotateFreshnessEpoch());
process.on("SIGTERM", () => { void stop().then(() => process.exit(0)); });
process.on("SIGINT", () => { void stop().then(() => process.exit(0)); });
