"use strict";

const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const CONFIG_DIRECTORY = path.join(os.homedir(), ".config", "beaman-ventures", "imessage-bridge");
const CONFIG_PATH = path.join(CONFIG_DIRECTORY, "allowlist.json");
const MESSAGES_DATABASE_PATH = path.join(os.homedir(), "Library", "Messages", "chat.db");

function configurationError() {
  const error = new Error("The iMessage bridge allowlist is invalid.");
  error.code = "INVALID_ALLOWLIST";
  return error;
}

function requireCleanString(value, maxLength) {
  if (typeof value !== "string" || value.length === 0 || value.length > maxLength || value.trim() !== value) {
    throw configurationError();
  }
  return value;
}

function assertOnlyKeys(value, allowedKeys) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw configurationError();
  if (Object.keys(value).some((key) => !allowedKeys.has(key))) throw configurationError();
}

function normalizeAllowlist(input) {
  assertOnlyKeys(input, new Set(["schemaVersion", "selfDisplayName", "threads"]));
  if (input.schemaVersion !== 1 || !Array.isArray(input.threads) || input.threads.length > 100) {
    throw configurationError();
  }

  const selfDisplayName = requireCleanString(input.selfDisplayName, 100);
  const conversationIds = new Set();
  const threads = input.threads.map((thread) => {
    assertOnlyKeys(thread, new Set(["conversationId", "displayName", "participants"]));
    const conversationId = requireCleanString(thread.conversationId, 500);
    const displayName = requireCleanString(thread.displayName, 200);
    if (conversationIds.has(conversationId) || !Array.isArray(thread.participants) || thread.participants.length === 0 || thread.participants.length > 32) {
      throw configurationError();
    }
    conversationIds.add(conversationId);

    const handles = new Set();
    const participants = thread.participants.map((participant) => {
      assertOnlyKeys(participant, new Set(["handle", "displayName"]));
      const handle = requireCleanString(participant.handle, 320);
      const participantDisplayName = requireCleanString(participant.displayName, 100);
      if (handles.has(handle)) throw configurationError();
      handles.add(handle);
      return Object.freeze({ handle, displayName: participantDisplayName });
    });

    return Object.freeze({ conversationId, displayName, participants: Object.freeze(participants) });
  });

  return Object.freeze({ schemaVersion: 1, selfDisplayName, threads: Object.freeze(threads) });
}

function loadAllowlist() {
  let descriptor;
  let raw;
  try {
    descriptor = fs.openSync(CONFIG_PATH, fs.constants.O_RDONLY | fs.constants.O_NOFOLLOW);
    const metadata = fs.fstatSync(descriptor);
    const ownerMatches = typeof process.getuid !== "function" || metadata.uid === process.getuid();
    if (!metadata.isFile() || !ownerMatches || (metadata.mode & 0o077) !== 0 || metadata.size > 1024 * 1024) {
      throw configurationError();
    }
    raw = fs.readFileSync(descriptor, { encoding: "utf8" });
  } catch {
    throw configurationError();
  } finally {
    if (descriptor !== undefined) {
      try { fs.closeSync(descriptor); } catch {}
    }
  }
  try {
    return normalizeAllowlist(JSON.parse(raw));
  } catch {
    throw configurationError();
  }
}

module.exports = {
  CONFIG_PATH,
  MESSAGES_DATABASE_PATH,
  loadAllowlist,
  normalizeAllowlist
};
