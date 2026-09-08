#!/usr/bin/env node
"use strict";

const readline = require("node:readline");
const { loadAllowlist, MESSAGES_DATABASE_PATH } = require("../services/imessageBridgeConfig");
const { createImessageBridgeService } = require("../services/imessageBridgeService");
const { createMcpRequestHandler } = require("../services/imessageBridgeMcp");
const { openReadOnlyMessagesDatabase } = require("../services/imessageReadOnlyDatabase");

let database;
try {
  const allowlist = loadAllowlist();
  database = openReadOnlyMessagesDatabase(MESSAGES_DATABASE_PATH);
  const service = createImessageBridgeService({ database, allowlist });
  const handle = createMcpRequestHandler(service);
  const lines = readline.createInterface({ input: process.stdin, crlfDelay: Infinity, terminal: false });
  lines.on("line", async (line) => {
    let request;
    try {
      request = JSON.parse(line);
    } catch {
      process.stdout.write(`${JSON.stringify({ jsonrpc: "2.0", id: null, error: { code: -32700, message: "Parse error" } })}\n`);
      return;
    }
    const response = await handle(request);
    if (response) process.stdout.write(`${JSON.stringify(response)}\n`);
  });
  lines.on("close", () => database.close());
} catch {
  process.stderr.write("The read-only iMessage bridge could not start.\n");
  if (database) {
    try { database.close(); } catch {}
  }
  process.exitCode = 1;
}
