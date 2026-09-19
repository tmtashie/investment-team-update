"use strict";

const READ_TOOLS = new Set([
  "list_allowed_message_threads",
  "read_recent_messages",
  "search_allowed_messages"
]);
const MAX_REPLAY_IDS = 1024;

function policyError(id, code, message) {
  return { jsonrpc: "2.0", id, error: { code, message } };
}

function unavailableStdioCallerIdentity() {
  return Object.freeze({ verified: false, reason: "stdio_has_no_caller_identity" });
}

function createPhase1bGuard({ handler, authorize = unavailableStdioCallerIdentity, logger = null, now = Date.now }) {
  if (typeof handler !== "function" || typeof authorize !== "function") throw new TypeError("Invalid guard configuration.");
  const completedRequestIds = new Set();
  let revoked = false;

  function log(event, metadata = {}) {
    if (logger && typeof logger.info === "function") logger.info(event, metadata);
  }

  async function handle(request) {
    if (!request || request.method !== "tools/call") return handler(request);
    const toolName = request.params && request.params.name;
    if (!READ_TOOLS.has(toolName)) return handler(request);
    if (revoked) {
      log("imessage_phase1b_denied", { reason: "host_revoked" });
      return policyError(request.id ?? null, -32003, "The local bridge is revoked.");
    }

    const decision = await authorize(Object.freeze({ toolName, receivedAt: now() }));
    if (!decision || decision.verified !== true || typeof decision.principal !== "string") {
      log("imessage_phase1b_denied", { reason: "caller_unverified" });
      return policyError(request.id ?? null, -32001, "Caller identity cannot be verified by the local stdio host.");
    }
    if (typeof decision.requestId !== "string" || decision.requestId.length < 8 || decision.requestId.length > 200) {
      log("imessage_phase1b_denied", { reason: "freshness_unverified" });
      return policyError(request.id ?? null, -32002, "Request freshness cannot be verified.");
    }
    if (completedRequestIds.has(decision.requestId)) {
      log("imessage_phase1b_denied", { reason: "replay" });
      return policyError(request.id ?? null, -32002, "Replayed request rejected.");
    }
    if (completedRequestIds.size >= MAX_REPLAY_IDS) {
      log("imessage_phase1b_denied", { reason: "replay_guard_full" });
      return policyError(request.id ?? null, -32002, "Request freshness guard is unavailable.");
    }
    completedRequestIds.add(decision.requestId);
    const response = await handler(request);
    log("imessage_phase1b_completed", { outcome: response && response.error ? "error" : "success" });
    return response;
  }

  return Object.freeze({
    handle,
    revoke() {
      revoked = true;
      completedRequestIds.clear();
      log("imessage_phase1b_revoked");
    },
    isRevoked() { return revoked; }
  });
}

module.exports = {
  MAX_REPLAY_IDS,
  READ_TOOLS,
  createPhase1bGuard,
  unavailableStdioCallerIdentity
};
