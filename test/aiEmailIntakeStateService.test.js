const test = require("node:test");
const assert = require("node:assert/strict");
const {
  createAiEmailIntakeStateService,
  messageDedupeKey
} = require("../services/aiEmailIntakeStateService");

function createMemoryStateService(initial = []) {
  let stored = initial;
  return {
    service: createAiEmailIntakeStateService({
      STATE_FILE: "state.json",
      readJsonFile: () => stored,
      writeJsonFile: (file, value) => {
        stored = value;
      }
    }),
    getStored: () => stored
  };
}

test("message dedupe key prefers internetMessageId before Graph id", () => {
  assert.equal(
    messageDedupeKey({ internetMessageId: "<mail@example.test>", id: "graph-id" }),
    "<mail@example.test>"
  );
  assert.equal(messageDedupeKey({ id: "graph-id" }), "graph-id");
});

test("intake state upserts by message and merges proposal ids and attachment hashes", () => {
  const { service, getStored } = createMemoryStateService();

  service.upsertEntry({
    graphMessageId: "graph-1",
    internetMessageId: "<mail-1@example.test>",
    attachmentHashes: ["hash-1"],
    proposalIds: ["proposal-1"],
    status: "processed"
  });
  service.upsertEntry({
    graphMessageId: "graph-1",
    internetMessageId: "<mail-1@example.test>",
    attachmentHashes: ["hash-2"],
    proposalIds: ["proposal-2"],
    status: "processed"
  });

  assert.equal(getStored().length, 1);
  assert.deepEqual(getStored()[0].attachmentHashes, ["hash-1", "hash-2"]);
  assert.deepEqual(getStored()[0].proposalIds, ["proposal-1", "proposal-2"]);
  assert.equal(service.hasAttachmentHash("hash-2"), true);
  assert.equal(
    service.findByMessage({ id: "graph-1", internetMessageId: "<mail-1@example.test>" }).status,
    "processed"
  );
});
