"use strict";

const TOOL_DEFINITIONS = Object.freeze([
  {
    name: "list_allowed_message_threads",
    description: "List only Apple Messages conversations explicitly approved in the local allowlist.",
    inputSchema: { type: "object", properties: {}, additionalProperties: false },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }
  },
  {
    name: "read_recent_messages",
    description: "Read a bounded number of recent text messages from one explicitly allowlisted conversation.",
    inputSchema: {
      type: "object",
      properties: {
        threadId: { type: "string", pattern: "^[a-z0-9][a-z0-9_-]{0,63}$" },
        limit: { type: "integer", minimum: 1, maximum: 50, default: 20 }
      },
      required: ["threadId"],
      additionalProperties: false
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }
  },
  {
    name: "search_allowed_messages",
    description: "Search text messages within one explicitly allowlisted conversation.",
    inputSchema: {
      type: "object",
      properties: {
        threadId: { type: "string", pattern: "^[a-z0-9][a-z0-9_-]{0,63}$" },
        query: { type: "string", minLength: 2, maxLength: 200 },
        limit: { type: "integer", minimum: 1, maximum: 50, default: 20 }
      },
      required: ["threadId", "query"],
      additionalProperties: false
    },
    annotations: { readOnlyHint: true, destructiveHint: false, idempotentHint: true, openWorldHint: false }
  }
]);

function safeError(error) {
  const permittedCodes = new Set([
    "CONVERSATION_NOT_ALLOWED",
    "INVALID_REQUEST",
    "MALFORMED_RECORD",
    "MESSAGES_DATABASE_UNAVAILABLE",
    "PARTICIPANTS_CHANGED"
  ]);
  const code = permittedCodes.has(error && error.code) ? error.code : "BRIDGE_ERROR";
  const messages = {
    BRIDGE_ERROR: "The iMessage bridge request failed.",
    CONVERSATION_NOT_ALLOWED: "Conversation is not allowed.",
    INVALID_REQUEST: "Invalid request.",
    MALFORMED_RECORD: "An allowed message record is malformed.",
    MESSAGES_DATABASE_UNAVAILABLE: "The Messages database is unavailable or incompatible.",
    PARTICIPANTS_CHANGED: "Conversation participants do not match the allowlist."
  };
  return { code, message: messages[code] };
}

function createMcpRequestHandler(service) {
  const operations = Object.freeze({
    list_allowed_message_threads: service.listAllowedMessageThreads,
    read_recent_messages: service.readRecentMessages,
    search_allowed_messages: service.searchAllowedMessages
  });

  return async function handle(request) {
    if (!request || request.jsonrpc !== "2.0" || typeof request.method !== "string") {
      return { jsonrpc: "2.0", id: request && request.id !== undefined ? request.id : null, error: { code: -32600, message: "Invalid Request" } };
    }
    if (!("id" in request)) return null;
    if (request.method === "server/discover") {
      return {
        jsonrpc: "2.0",
        id: request.id,
        result: {
          resultType: "complete",
          supportedVersions: ["2026-07-28", "2025-11-25"],
          capabilities: { tools: { listChanged: false } },
          serverInfo: { name: "beaman-imessage-readonly", version: "0.1.0" }
        }
      };
    }
    if (request.method === "initialize") {
      return {
        jsonrpc: "2.0",
        id: request.id,
        result: {
          protocolVersion: "2025-11-25",
          capabilities: { tools: { listChanged: false } },
          serverInfo: { name: "beaman-imessage-readonly", version: "0.1.0" }
        }
      };
    }
    if (request.method === "tools/list") {
      return { jsonrpc: "2.0", id: request.id, result: { tools: TOOL_DEFINITIONS } };
    }
    if (request.method === "ping") {
      return { jsonrpc: "2.0", id: request.id, result: {} };
    }
    if (request.method !== "tools/call") {
      return { jsonrpc: "2.0", id: request.id, error: { code: -32601, message: "Method not found" } };
    }
    const name = request.params && request.params.name;
    const operation = operations[name];
    if (!operation) return { jsonrpc: "2.0", id: request.id, error: { code: -32601, message: "Tool not found" } };
    try {
      const output = await operation(request.params.arguments || {});
      return {
        jsonrpc: "2.0",
        id: request.id,
        result: {
          content: [{ type: "text", text: JSON.stringify(output) }],
          structuredContent: output,
          isError: false
        }
      };
    } catch (error) {
      const output = safeError(error);
      return {
        jsonrpc: "2.0",
        id: request.id,
        result: {
          content: [{ type: "text", text: JSON.stringify(output) }],
          structuredContent: output,
          isError: true
        }
      };
    }
  };
}

module.exports = {
  TOOL_DEFINITIONS,
  createMcpRequestHandler,
  safeError
};
