"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");
const { DatabaseSync } = require("node:sqlite");

const { normalizeAllowlist } = require("../services/imessageBridgeConfig");
const { appleTimestampToCentral, createImessageBridgeService } = require("../services/imessageBridgeService");
const { TOOL_DEFINITIONS, createMcpRequestHandler } = require("../services/imessageBridgeMcp");
const { openReadOnlyMessagesDatabase } = require("../services/imessageReadOnlyDatabase");

const THREAD_ONE = "approved_one";
const THREAD_GROUP = "approved_group";
const THREAD_BLOCKED = "blocked_thread";
const THREAD_ONE_GUID = "iMessage;-;synthetic-one";
const THREAD_GROUP_GUID = "iMessage;+;synthetic-group";
const THREAD_BLOCKED_GUID = "iMessage;-;synthetic-blocked";

function allowlist(overrides = {}) {
  return normalizeAllowlist({
    schemaVersion: 1,
    selfDisplayName: "Local User",
    threads: [
      {
        threadId: THREAD_ONE,
        chatGuid: THREAD_ONE_GUID,
        displayName: "Approved One-to-One",
        participants: [{ handle: "person-one@example.invalid", displayName: "Person One" }]
      },
      {
        threadId: THREAD_GROUP,
        chatGuid: THREAD_GROUP_GUID,
        displayName: "Approved Group",
        participants: [
          { handle: "person-one@example.invalid", displayName: "Person One" },
          { handle: "+15555550102", displayName: "Person Two" }
        ]
      }
    ],
    ...overrides
  });
}

function buildFixture() {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), "imessage-bridge-test-"));
  const databasePath = path.join(directory, "chat.db");
  const database = new DatabaseSync(databasePath);
  database.exec(`
    CREATE TABLE chat (guid TEXT NOT NULL);
    CREATE TABLE handle (id TEXT NOT NULL);
    CREATE TABLE chat_handle_join (chat_id INTEGER NOT NULL, handle_id INTEGER NOT NULL);
    CREATE TABLE message (
      guid TEXT,
      handle_id INTEGER,
      is_from_me INTEGER,
      date INTEGER,
      text TEXT,
      attributedBody BLOB,
      associated_message_type INTEGER DEFAULT 0,
      item_type INTEGER DEFAULT 0,
      is_empty INTEGER DEFAULT 0,
      date_edited INTEGER DEFAULT 0,
      date_retracted INTEGER DEFAULT 0,
      thread_originator_guid TEXT
    );
    CREATE TABLE chat_message_join (chat_id INTEGER NOT NULL, message_id INTEGER NOT NULL);
  `);
  const insertChat = database.prepare("INSERT INTO chat (guid) VALUES (?)");
  const insertHandle = database.prepare("INSERT INTO handle (id) VALUES (?)");
  const joinHandle = database.prepare("INSERT INTO chat_handle_join (chat_id, handle_id) VALUES (?, ?)");
  const insertMessage = database.prepare("INSERT INTO message (guid, handle_id, is_from_me, date, text) VALUES (?, ?, ?, ?, ?)");
  const joinMessage = database.prepare("INSERT INTO chat_message_join (chat_id, message_id) VALUES (?, ?)");

  const chats = [THREAD_ONE_GUID, THREAD_GROUP_GUID, THREAD_BLOCKED_GUID].map((guid) => Number(insertChat.run(guid).lastInsertRowid));
  const handles = ["person-one@example.invalid", "+15555550102", "blocked@example.invalid"].map((id) => Number(insertHandle.run(id).lastInsertRowid));
  joinHandle.run(chats[0], handles[0]);
  joinHandle.run(chats[1], handles[0]);
  joinHandle.run(chats[1], handles[1]);
  joinHandle.run(chats[2], handles[2]);

  function message(chatId, guid, handleId, fromMe, date, text) {
    const result = insertMessage.run(guid, handleId, fromMe, date, text);
    joinMessage.run(chatId, result.lastInsertRowid);
  }
  message(chats[0], "message-later", handles[0], 1, 704764800, "Synthetic outbound body");
  message(chats[0], "message-earlier", handles[0], 0, 704761200000000000, "Synthetic inbound body");
  message(chats[1], "group-one", handles[1], 0, 704768400, "Synthetic group body");
  message(chats[2], "blocked-one", handles[2], 0, 704772000, "Never return this body");
  database.close();

  return {
    directory,
    databasePath,
    cleanup() { fs.rmSync(directory, { recursive: true, force: true }); }
  };
}

function withService(t, options = {}) {
  const fixture = buildFixture();
  t.after(fixture.cleanup);
  const database = openReadOnlyMessagesDatabase(fixture.databasePath);
  t.after(() => database.close());
  return {
    database,
    fixture,
    service: createImessageBridgeService({ database, allowlist: options.allowlist || allowlist(), logger: options.logger })
  };
}

test("an allowlisted conversation can be read with normalized output", (t) => {
  const { service } = withService(t);
  const result = service.readRecentMessages({ threadId: THREAD_ONE, limit: 10 });
  assert.equal(result.messages.length, 2);
  assert.deepEqual(Object.keys(result.messages[0]), [
    "threadId", "conversationDisplayName", "participants", "messageId", "replyToMessageId", "sender", "direction", "timestamp", "text"
  ]);
  assert.equal(result.messages[0].sender, "Person One");
  assert.equal(result.messages[0].direction, "received");
  assert.equal(result.messages[1].sender, "Local User");
  assert.equal(result.messages[1].direction, "sent");
  assert.ok(result.messages[0].timestamp.endsWith("-05:00"));
  assert.ok(result.messages[0].timestamp < result.messages[1].timestamp);
  assert.equal(JSON.stringify(result).includes(THREAD_ONE_GUID), false);
});

test("a non-allowlisted conversation is rejected before message access", (t) => {
  const { service } = withService(t);
  assert.throws(
    () => service.readRecentMessages({ threadId: THREAD_BLOCKED }),
    (error) => error.code === "CONVERSATION_NOT_ALLOWED"
  );
});

test("thread listing returns only configured aliases and no raw handles", (t) => {
  const { service } = withService(t);
  const result = service.listAllowedMessageThreads({});
  assert.equal(result.threads.length, 2);
  const serialized = JSON.stringify(result);
  assert.equal(serialized.includes("person-one@example.invalid"), false);
  assert.equal(serialized.includes("+15555550102"), false);
  assert.equal(serialized.includes(THREAD_ONE_GUID), false);
  assert.equal(serialized.includes(THREAD_GROUP_GUID), false);
});

test("arbitrary SQL and filesystem paths are rejected by every tool", async (t) => {
  const { service } = withService(t);
  const handle = createMcpRequestHandler(service);
  for (const [name, args] of [
    ["list_allowed_message_threads", { sql: "SELECT * FROM message" }],
    ["read_recent_messages", { threadId: THREAD_ONE, databasePath: "/tmp/other.db" }],
    ["search_allowed_messages", { threadId: THREAD_ONE, query: "Synthetic", path: "/etc/passwd" }]
  ]) {
    const response = await handle({ jsonrpc: "2.0", id: name, method: "tools/call", params: { name, arguments: args } });
    assert.equal(response.result.isError, true);
    assert.equal(response.result.structuredContent.code, "INVALID_REQUEST");
  }
});

test("search is parameterized and confined to the selected allowed conversation", (t) => {
  const { service } = withService(t);
  assert.equal(service.searchAllowedMessages({ threadId: THREAD_ONE, query: "Synthetic", limit: 10 }).messages.length, 2);
  assert.equal(service.searchAllowedMessages({ threadId: THREAD_ONE, query: "Never return", limit: 10 }).messages.length, 0);
  assert.equal(service.searchAllowedMessages({ threadId: THREAD_ONE, query: "%' OR 1=1 --", limit: 10 }).messages.length, 0);
});

test("reactions, group events, and retracted messages are excluded", (t) => {
  const fixture = buildFixture();
  t.after(fixture.cleanup);
  const writer = new DatabaseSync(fixture.databasePath);
  const chatId = writer.prepare("SELECT ROWID FROM chat WHERE guid = ?").get(THREAD_ONE_GUID).rowid;
  const insert = writer.prepare(`
    INSERT INTO message (guid, handle_id, is_from_me, date, text, associated_message_type, item_type, date_retracted, is_empty, date_edited)
    VALUES (?, 1, 0, ?, ?, ?, ?, ?, ?, ?)
  `);
  for (const values of [
    ["reaction", 704780000, "Reaction payload", 2000, 0, 0, 0, 0],
    ["group-event", 704780001, "Group event payload", 0, 1, 0, 0, 0],
    ["retracted", 704780002, "Retracted payload", 0, 0, 704780003, 0, 0],
    ["retracted-sequoia", 704780004, "Alternate retracted payload", 0, 0, 0, 1, 704780005]
  ]) {
    const result = insert.run(...values);
    writer.prepare("INSERT INTO chat_message_join (chat_id, message_id) VALUES (?, ?)").run(chatId, result.lastInsertRowid);
  }
  writer.close();
  const database = openReadOnlyMessagesDatabase(fixture.databasePath);
  t.after(() => database.close());
  const service = createImessageBridgeService({ database, allowlist: allowlist() });
  const serialized = JSON.stringify(service.readRecentMessages({ threadId: THREAD_ONE, limit: 50 }));
  assert.equal(serialized.includes("payload"), false);
});

test("group conversations require an exact current participant match", (t) => {
  const { service, databasePath } = (() => {
    const fixture = buildFixture();
    t.after(fixture.cleanup);
    const writer = new DatabaseSync(fixture.databasePath);
    const chatId = writer.prepare("SELECT ROWID FROM chat WHERE guid = ?").get(THREAD_GROUP_GUID).rowid;
    const newHandle = writer.prepare("INSERT INTO handle (id) VALUES (?)").run("unapproved@example.invalid").lastInsertRowid;
    writer.prepare("INSERT INTO chat_handle_join (chat_id, handle_id) VALUES (?, ?)").run(chatId, newHandle);
    writer.close();
    const database = openReadOnlyMessagesDatabase(fixture.databasePath);
    t.after(() => database.close());
    return { service: createImessageBridgeService({ database, allowlist: allowlist() }), databasePath: fixture.databasePath };
  })();
  assert.ok(databasePath);
  assert.throws(
    () => service.readRecentMessages({ threadId: THREAD_GROUP }),
    (error) => error.code === "PARTICIPANTS_CHANGED"
  );
});

test("malformed records fail safely without returning partial data", (t) => {
  const fixture = buildFixture();
  t.after(fixture.cleanup);
  const writer = new DatabaseSync(fixture.databasePath);
  const chatId = writer.prepare("SELECT ROWID FROM chat WHERE guid = ?").get(THREAD_ONE_GUID).rowid;
  const result = writer.prepare("INSERT INTO message (guid, handle_id, is_from_me, date, text) VALUES (NULL, 1, 0, 704779200, 'Sensitive malformed body')").run();
  writer.prepare("INSERT INTO chat_message_join (chat_id, message_id) VALUES (?, ?)").run(chatId, result.lastInsertRowid);
  writer.close();
  const database = openReadOnlyMessagesDatabase(fixture.databasePath);
  t.after(() => database.close());
  const service = createImessageBridgeService({ database, allowlist: allowlist() });
  assert.throws(
    () => service.readRecentMessages({ threadId: THREAD_ONE }),
    (error) => error.code === "MALFORMED_RECORD" && !error.message.includes("Sensitive")
  );
});

test("message bodies and search text are never logged", (t) => {
  const entries = [];
  const logger = { info(...args) { entries.push(args); } };
  const { service } = withService(t, { logger });
  service.readRecentMessages({ threadId: THREAD_ONE });
  service.searchAllowedMessages({ threadId: THREAD_ONE, query: "Synthetic inbound" });
  const serialized = JSON.stringify(entries);
  assert.equal(serialized.includes("Synthetic inbound"), false);
  assert.equal(serialized.includes("Synthetic outbound"), false);
  assert.equal(serialized.includes(THREAD_ONE), false);
});

test("the MCP surface contains only three annotated read-only tools", () => {
  assert.deepEqual(TOOL_DEFINITIONS.map((tool) => tool.name), [
    "list_allowed_message_threads",
    "read_recent_messages",
    "search_allowed_messages"
  ]);
  for (const tool of TOOL_DEFINITIONS) {
    assert.deepEqual(tool.annotations, {
      readOnlyHint: true,
      destructiveHint: false,
      idempotentHint: true,
      openWorldHint: false
    });
    assert.equal(tool.inputSchema.additionalProperties, false);
  }
  assert.equal(TOOL_DEFINITIONS.some((tool) => /send|reply|react|edit|delete|attachment|mark/i.test(tool.name)), false);
});

test("the MCP handler completes the read-only handshake and ignores notifications", async (t) => {
  const { service } = withService(t);
  const handle = createMcpRequestHandler(service);
  const initialized = await handle({ jsonrpc: "2.0", method: "notifications/initialized" });
  assert.equal(initialized, null);
  const discovered = await handle({ jsonrpc: "2.0", id: 0, method: "server/discover" });
  assert.deepEqual(discovered.result.supportedVersions, ["2025-11-25"]);
  const listed = await handle({ jsonrpc: "2.0", id: 1, method: "tools/list" });
  assert.deepEqual(listed.result.tools, TOOL_DEFINITIONS);
  const ping = await handle({ jsonrpc: "2.0", id: 2, method: "ping" });
  assert.deepEqual(ping.result, {});
});

test("database access is enforced as read-only and query-only", (t) => {
  const { database } = withService(t);
  assert.equal(database.prepare("PRAGMA query_only").get().query_only, 1);
  assert.throws(() => database.exec("INSERT INTO chat (guid) VALUES ('forbidden')"), /not authorized|read-?only/i);
});

test("Apple epoch seconds and nanoseconds normalize across Central daylight time", () => {
  assert.equal(appleTimestampToCentral(704764800), "2023-05-02T19:00:00-05:00");
  assert.equal(appleTimestampToCentral(704764800000000000n), "2023-05-02T19:00:00-05:00");
  assert.equal(appleTimestampToCentral(694310400), "2023-01-01T18:00:00-06:00");
});

test("allowlist parsing rejects extra fields and empty participant sets", () => {
  assert.throws(() => normalizeAllowlist({ ...allowlist(), databasePath: "/tmp/chat.db" }), /allowlist is invalid/i);
  assert.throws(() => normalizeAllowlist({ schemaVersion: 1, selfDisplayName: "Local User", threads: [{ threadId: "x", chatGuid: "guid", displayName: "x", participants: [] }] }), /allowlist is invalid/i);
});
