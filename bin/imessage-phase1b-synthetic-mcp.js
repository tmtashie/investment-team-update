#!/usr/bin/env node
"use strict";

const readline = require("node:readline");
const { createImessageBridgeService } = require("../services/imessageBridgeService");
const { createMcpRequestHandler } = require("../services/imessageBridgeMcp");
const { createPhase1bGuard, unavailableStdioCallerIdentity } = require("../services/imessagePhase1bPolicy");
const { openReadOnlyMessagesDatabase } = require("../services/imessageReadOnlyDatabase");
const { createSyntheticMessagesFixture } = require("../services/imessageSyntheticFixture");

if (process.env.PHASE1B_SYNTHETIC_ONLY !== "1") {
  process.stderr.write("Synthetic-only mode is required.\n");
  process.exit(1);
}

const fixture = createSyntheticMessagesFixture();
const database = openReadOnlyMessagesDatabase(fixture.databasePath);
const service = createImessageBridgeService({ database, allowlist: fixture.allowlist });
const guard = createPhase1bGuard({
  handler: createMcpRequestHandler(service),
  authorize: unavailableStdioCallerIdentity
});
const lines = readline.createInterface({ input: process.stdin, crlfDelay: Infinity, terminal: false });

lines.on("line", async (line) => {
  let request;
  try {
    request = JSON.parse(line);
  } catch {
    process.stdout.write(`${JSON.stringify({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Parse error" } })}\n`);
    return;
  }
  const response = await guard.handle(request);
  if (response) process.stdout.write(`${JSON.stringify(response)}\n`);
});

function close() {
  guard.revoke();
  try { database.close(); } catch {}
  fixture.cleanup();
}

lines.on("close", close);
process.once("SIGTERM", () => { close(); process.exit(0); });
process.once("SIGINT", () => { close(); process.exit(0); });
