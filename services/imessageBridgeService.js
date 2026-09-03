"use strict";

const MAX_MESSAGES = 50;
const MAX_QUERY_LENGTH = 200;
const MAX_MESSAGE_TEXT_LENGTH = 50000;
const TIME_ZONE = "America/Chicago";

function bridgeError(code, message) {
  const error = new Error(message);
  error.code = code;
  return error;
}

function assertExactKeys(value, allowedKeys) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw bridgeError("INVALID_REQUEST", "Invalid request.");
  }
  if (Object.keys(value).some((key) => !allowedKeys.has(key))) {
    throw bridgeError("INVALID_REQUEST", "Invalid request.");
  }
}

function requireConversationId(value) {
  if (typeof value !== "string" || value.length === 0 || value.length > 500 || value.trim() !== value) {
    throw bridgeError("INVALID_REQUEST", "Invalid request.");
  }
  return value;
}

function normalizeLimit(value) {
  if (value === undefined) return 20;
  if (!Number.isInteger(value) || value < 1 || value > MAX_MESSAGES) {
    throw bridgeError("INVALID_REQUEST", "Invalid request.");
  }
  return value;
}

function normalizeSearchQuery(value) {
  if (typeof value !== "string" || value.trim().length < 2 || value.length > MAX_QUERY_LENGTH) {
    throw bridgeError("INVALID_REQUEST", "Invalid request.");
  }
  return value.trim();
}

function appleTimestampToCentral(value) {
  if (typeof value !== "number" && typeof value !== "bigint") {
    throw bridgeError("MALFORMED_RECORD", "An allowed message record is malformed.");
  }
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) throw bridgeError("MALFORMED_RECORD", "An allowed message record is malformed.");
  const appleSeconds = Math.abs(numeric) > 100000000000 ? numeric / 1000000000 : numeric;
  const date = new Date((appleSeconds + 978307200) * 1000);
  if (Number.isNaN(date.getTime())) throw bridgeError("MALFORMED_RECORD", "An allowed message record is malformed.");

  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: TIME_ZONE,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23",
    timeZoneName: "longOffset"
  }).formatToParts(date);
  const byType = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  const offset = byType.timeZoneName.replace("GMT", "") || "+00:00";
  return `${byType.year}-${byType.month}-${byType.day}T${byType.hour}:${byType.minute}:${byType.second}${offset}`;
}

function createImessageBridgeService({ database, allowlist, logger = null }) {
  if (!database || !allowlist || !Array.isArray(allowlist.threads)) {
    throw bridgeError("INVALID_CONFIGURATION", "The iMessage bridge is not configured.");
  }
  const allowedThreads = new Map(allowlist.threads.map((thread) => [thread.conversationId, thread]));

  function log(event, metadata) {
    if (logger && typeof logger.info === "function") logger.info(event, metadata);
  }

  function requireAllowedThread(conversationId) {
    const thread = allowedThreads.get(requireConversationId(conversationId));
    if (!thread) throw bridgeError("CONVERSATION_NOT_ALLOWED", "Conversation is not allowed.");

    let rows;
    try {
      rows = database.prepare(`
        SELECT DISTINCT h.id AS handle
        FROM chat AS c
        JOIN chat_handle_join AS chj ON chj.chat_id = c.ROWID
        JOIN handle AS h ON h.ROWID = chj.handle_id
        WHERE c.guid = ?
      `).all(thread.conversationId);
    } catch {
      throw bridgeError("MESSAGES_DATABASE_UNAVAILABLE", "The Messages database is unavailable or incompatible.");
    }
    const actualHandles = new Set(rows.map((row) => row.handle));
    const configuredHandles = new Set(thread.participants.map((participant) => participant.handle));
    if (actualHandles.size !== configuredHandles.size || [...actualHandles].some((handle) => !configuredHandles.has(handle))) {
      throw bridgeError("PARTICIPANTS_CHANGED", "Conversation participants do not match the allowlist.");
    }
    return thread;
  }

  function normalizeMessage(row, thread) {
    if (
      typeof row.message_id !== "string" || row.message_id.length === 0 || row.message_id.length > 500 ||
      typeof row.text !== "string" || row.text.length > MAX_MESSAGE_TEXT_LENGTH ||
      (row.is_from_me !== 0 && row.is_from_me !== 1)
    ) {
      throw bridgeError("MALFORMED_RECORD", "An allowed message record is malformed.");
    }
    const direction = row.is_from_me === 1 ? "sent" : "received";
    let sender = allowlist.selfDisplayName;
    if (direction === "received") {
      const participant = thread.participants.find((item) => item.handle === row.sender_handle);
      if (!participant) throw bridgeError("MALFORMED_RECORD", "An allowed message record is malformed.");
      sender = participant.displayName;
    }
    return {
      conversationId: thread.conversationId,
      conversationDisplayName: thread.displayName,
      participants: thread.participants.map((participant) => participant.displayName),
      messageId: row.message_id,
      replyToMessageId: typeof row.reply_to_message_id === "string" && row.reply_to_message_id.length > 0
        ? row.reply_to_message_id.slice(0, 500)
        : null,
      sender,
      direction,
      timestamp: appleTimestampToCentral(row.apple_seconds),
      text: row.text
    };
  }

  function selectMessages(thread, { limit, query = null }) {
    const searchClause = query === null ? "" : "AND instr(lower(m.text), lower(?)) > 0";
    const parameters = query === null
      ? [thread.conversationId, limit]
      : [thread.conversationId, query, limit];
    let rows;
    try {
      rows = database.prepare(`
        SELECT * FROM (
          SELECT
            m.guid AS message_id,
            m.thread_originator_guid AS reply_to_message_id,
            m.text AS text,
            CASE
              WHEN abs(m.date) > 100000000000 THEN m.date / 1000000000.0
              ELSE m.date * 1.0
            END AS apple_seconds,
            m.is_from_me AS is_from_me,
            h.id AS sender_handle
          FROM chat AS c
          JOIN chat_message_join AS cmj ON cmj.chat_id = c.ROWID
          JOIN message AS m ON m.ROWID = cmj.message_id
          LEFT JOIN handle AS h ON h.ROWID = m.handle_id
          WHERE c.guid = ?
            AND m.text IS NOT NULL
            AND length(m.text) > 0
            AND coalesce(m.associated_message_type, 0) = 0
            AND coalesce(m.item_type, 0) = 0
            AND coalesce(m.date_retracted, 0) = 0
            AND NOT (coalesce(m.is_empty, 0) = 1 AND coalesce(m.date_edited, 0) != 0)
            ${searchClause}
          ORDER BY m.date DESC, m.ROWID DESC
          LIMIT ?
        ) AS recent
        ORDER BY apple_seconds ASC, message_id ASC
      `).all(...parameters);
    } catch {
      throw bridgeError("MESSAGES_DATABASE_UNAVAILABLE", "The Messages database is unavailable or incompatible.");
    }
    return rows.map((row) => normalizeMessage(row, thread));
  }

  function listAllowedMessageThreads(input = {}) {
    assertExactKeys(input, new Set());
    const threads = [];
    for (const configured of allowlist.threads) {
      const thread = requireAllowedThread(configured.conversationId);
      threads.push({
        conversationId: thread.conversationId,
        displayName: thread.displayName,
        participants: thread.participants.map((participant) => participant.displayName)
      });
    }
    log("imessage_bridge_list_allowed_threads", { count: threads.length });
    return { threads };
  }

  function readRecentMessages(input) {
    assertExactKeys(input, new Set(["conversationId", "limit"]));
    const thread = requireAllowedThread(input.conversationId);
    const messages = selectMessages(thread, { limit: normalizeLimit(input.limit) });
    log("imessage_bridge_read_recent", { count: messages.length });
    return { messages };
  }

  function searchAllowedMessages(input) {
    assertExactKeys(input, new Set(["conversationId", "query", "limit"]));
    const thread = requireAllowedThread(input.conversationId);
    const messages = selectMessages(thread, {
      limit: normalizeLimit(input.limit),
      query: normalizeSearchQuery(input.query)
    });
    log("imessage_bridge_search", { count: messages.length });
    return { messages };
  }

  return Object.freeze({
    listAllowedMessageThreads,
    readRecentMessages,
    searchAllowedMessages
  });
}

module.exports = {
  MAX_MESSAGES,
  appleTimestampToCentral,
  createImessageBridgeService
};
