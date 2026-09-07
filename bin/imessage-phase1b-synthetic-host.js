#!/usr/bin/env node
"use strict";

const os = require("node:os");
const path = require("node:path");
const { createPhase1bHost } = require("../services/imessagePhase1bHost");

const appRoot = path.resolve(__dirname, "..");
const supportRoot = path.join(os.homedir(), "Library", "Application Support", "Beaman Ventures", "iMessage Bridge");
const host = createPhase1bHost({
  tunnelClientPath: path.join(appRoot, "vendor", "tunnel-client"),
  syntheticMcpCommand: path.join(appRoot, "bin", "imessage-phase1b-synthetic-mcp.js"),
  credentialPath: path.join(supportRoot, "tunnel-runtime-key"),
  tunnelId: process.env.PHASE1B_TUNNEL_ID || ""
});

try {
  host.start();
} catch {
  process.stderr.write("The synthetic Phase 1B host could not start.\n");
  process.exit(1);
}

process.once("SIGTERM", () => host.revoke());
process.once("SIGINT", () => host.revoke());
