"use strict";

const fs = require("node:fs");
const path = require("node:path");
const { spawn } = require("node:child_process");

const TUNNEL_ID_PATTERN = /^tunnel_[a-f0-9]{32}$/;

function hostError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function validateCredentialFile(credentialPath, fsApi = fs) {
  let descriptor;
  try {
    descriptor = fsApi.openSync(credentialPath, fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW);
    const metadata = fsApi.fstatSync(descriptor);
    const ownerMatches = typeof process.getuid !== "function" || metadata.uid === process.getuid();
    if (!metadata.isFile() || !ownerMatches || (metadata.mode & 0o077) !== 0 || metadata.size < 1 || metadata.size > 4096) {
      throw hostError("INVALID_TUNNEL_CREDENTIAL", "Tunnel credential is unavailable.");
    }
  } catch (error) {
    if (error && error.code === "INVALID_TUNNEL_CREDENTIAL") throw error;
    throw hostError("INVALID_TUNNEL_CREDENTIAL", "Tunnel credential is unavailable.");
  } finally {
    if (descriptor !== undefined) {
      try { fsApi.closeSync(descriptor); } catch {}
    }
  }
}

function buildTunnelArguments({ tunnelId, credentialPath, mcpCommand }) {
  if (!TUNNEL_ID_PATTERN.test(tunnelId)) throw hostError("INVALID_TUNNEL_ID", "Tunnel configuration is invalid.");
  return Object.freeze([
    "run",
    `--control-plane.tunnel-id=${tunnelId}`,
    `--control-plane.api-key=file:${credentialPath}`,
    "--control-plane.base-url=https://api.openai.com",
    "--control-plane.max-inflight=1",
    "--mcp.max-concurrent-requests=1",
    `--mcp.command=${mcpCommand}`,
    "--health.listen-addr=127.0.0.1:0",
    "--admin-ui.open-browser=false",
    "--log.level=warn",
    "--log.format=json"
  ]);
}

function buildHttpTunnelArguments({ tunnelId, credentialPath, socketPath }) {
  if (!TUNNEL_ID_PATTERN.test(tunnelId) || typeof socketPath !== "string" || !path.isAbsolute(socketPath)) {
    throw hostError("INVALID_TUNNEL_CONFIGURATION", "Tunnel configuration is invalid.");
  }
  return Object.freeze([
    "run",
    `--control-plane.tunnel-id=${tunnelId}`,
    `--control-plane.api-key=file:${credentialPath}`,
    "--control-plane.base-url=https://api.openai.com",
    "--control-plane.max-inflight=1",
    "--mcp.max-concurrent-requests=1",
    `--mcp.server-url=channel=main,url=http://localhost/mcp,unix-socket=${socketPath}`,
    "--health.listen-addr=127.0.0.1:0",
    "--admin-ui.open-browser=false",
    "--log.level=warn",
    "--log.format=json"
  ]);
}

function createPhase1bHost({
  tunnelClientPath,
  tunnelId,
  credentialPath,
  syntheticMcpCommand,
  spawnProcess = spawn,
  fsApi = fs,
  logger = null
}) {
  const fixedPaths = [tunnelClientPath, credentialPath, syntheticMcpCommand];
  if (fixedPaths.some((value) => typeof value !== "string" || !path.isAbsolute(value))) {
    throw hostError("INVALID_HOST_CONFIGURATION", "Host configuration is invalid.");
  }
  let child = null;
  let revoked = false;

  function log(event, metadata = {}) {
    if (logger && typeof logger.info === "function") logger.info(event, metadata);
  }

  function start() {
    if (revoked) throw hostError("HOST_REVOKED", "The local bridge is revoked.");
    if (child) throw hostError("HOST_ALREADY_RUNNING", "The local bridge is already running.");
    validateCredentialFile(credentialPath, fsApi);
    const args = buildTunnelArguments({ tunnelId, credentialPath, mcpCommand: syntheticMcpCommand });
    child = spawnProcess(tunnelClientPath, args, {
      shell: false,
      stdio: ["ignore", "ignore", "ignore"],
      env: Object.freeze({
        HOME: process.env.HOME || "",
        PATH: "/usr/bin:/bin:/usr/sbin:/sbin",
        TMPDIR: process.env.TMPDIR || "/tmp",
        PHASE1B_SYNTHETIC_ONLY: "1"
      })
    });
    child.once("exit", (code, signal) => {
      child = null;
      log("imessage_phase1b_tunnel_stopped", { outcome: code === 0 ? "clean" : "failed", signal: signal ? "present" : "absent" });
    });
    log("imessage_phase1b_tunnel_started");
    return child;
  }

  function stop(reason = "operator") {
    if (child) child.kill("SIGTERM");
    log("imessage_phase1b_tunnel_stop_requested", { reason: reason === "revoked" ? "revoked" : "operator" });
  }

  return Object.freeze({
    start,
    stop,
    revoke() {
      revoked = true;
      stop("revoked");
    },
    isRunning() { return child !== null; },
    isRevoked() { return revoked; }
  });
}

function createPhase1bHttpApplication({
  localServer,
  tunnelClientPath,
  tunnelId,
  credentialPath,
  socketPath,
  spawnProcess = spawn,
  fsApi = fs,
  logger = null
}) {
  if (!localServer || typeof localServer.start !== "function" || typeof localServer.revoke !== "function") {
    throw hostError("INVALID_HOST_CONFIGURATION", "Host configuration is invalid.");
  }
  for (const value of [tunnelClientPath, credentialPath, socketPath]) {
    if (typeof value !== "string" || !path.isAbsolute(value)) {
      throw hostError("INVALID_HOST_CONFIGURATION", "Host configuration is invalid.");
    }
  }
  let child = null;
  let revoked = false;

  function log(event, metadata = {}) {
    if (logger && typeof logger.info === "function") logger.info(event, metadata);
  }

  async function start() {
    if (revoked) throw hostError("HOST_REVOKED", "The local bridge is revoked.");
    if (child) throw hostError("HOST_ALREADY_RUNNING", "The local bridge is already running.");
    validateCredentialFile(credentialPath, fsApi);
    await localServer.start();
    const args = buildHttpTunnelArguments({ tunnelId, credentialPath, socketPath });
    try {
      child = spawnProcess(tunnelClientPath, args, {
        shell: false,
        stdio: ["ignore", "ignore", "ignore"],
        env: Object.freeze({
          HOME: process.env.HOME || "",
          PATH: "/usr/bin:/bin:/usr/sbin:/sbin",
          TMPDIR: process.env.TMPDIR || "/tmp",
          PHASE1B_SYNTHETIC_ONLY: "1"
        })
      });
    } catch (error) {
      await localServer.stop();
      throw error;
    }
    child.once("exit", () => {
      child = null;
      void localServer.stop();
      log("imessage_phase1b_http_tunnel_stopped");
    });
    log("imessage_phase1b_http_tunnel_started");
  }

  async function revoke() {
    revoked = true;
    localServer.revoke();
    if (child) child.kill("SIGTERM");
    await localServer.stop();
    log("imessage_phase1b_http_host_revoked");
  }

  return Object.freeze({
    start,
    revoke,
    isRunning() { return child !== null; },
    isRevoked() { return revoked; }
  });
}

module.exports = {
  TUNNEL_ID_PATTERN,
  buildHttpTunnelArguments,
  buildTunnelArguments,
  createPhase1bHttpApplication,
  createPhase1bHost,
  validateCredentialFile
};
