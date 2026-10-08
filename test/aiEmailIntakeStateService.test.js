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

test("message identity matching does not cross Graph and internet ID namespaces", () => {
  const { service, getStored } = createMemoryStateService([
    {
      graphMessageId: "<older@example.test>",
      internetMessageId: "<newer@example.test>",
      subject: "Newer message",
      status: "processed"
    },
    {
      graphMessageId: "graph-older",
      internetMessageId: "<older@example.test>",
      subject: "Older message",
      status: "processed"
    }
  ]);

  const older = service.findByMessage({
    id: "graph-older",
    internetMessageId: "<older@example.test>"
  });
  assert.equal(older.subject, "Older message");

  service.upsertEntry({
    graphMessageId: "graph-older",
    internetMessageId: "<older@example.test>",
    status: "skipped"
  });
  assert.equal(getStored()[0].status, "processed");
  assert.equal(getStored()[1].status, "skipped");
});

test("conflicting message identifiers fail closed without mutating state", () => {
  const initial = [
    {
      graphMessageId: "graph-a",
      internetMessageId: "<mail-a@example.test>",
      subject: "Message A",
      status: "processed"
    },
    {
      graphMessageId: "graph-b",
      internetMessageId: "<mail-b@example.test>",
      subject: "Message B",
      status: "processed"
    }
  ];
  const { service, getStored } = createMemoryStateService(initial);
  const conflicting = {
    id: "graph-b",
    internetMessageId: "<mail-a@example.test>"
  };

  assert.throws(
    () => service.findByMessage(conflicting),
    /identifiers resolve ambiguously across intake state entries/i
  );
  const reservation = service.claimMessage(conflicting);
  assert.equal(reservation.claimed, false);
  assert.match(reservation.reason, /identifiers resolve ambiguously across intake state entries/i);
  assert.throws(
    () => service.upsertEntry({
      graphMessageId: conflicting.id,
      internetMessageId: conflicting.internetMessageId,
      status: "skipped"
    }),
    /identifiers resolve ambiguously across intake state entries/i
  );
  assert.deepEqual(getStored(), initial);
});

test("duplicate identifiers within one namespace fail closed without mutating state", () => {
  for (const identity of [
    {
      field: "internetMessageId",
      value: "<duplicate@example.test>",
      message: { internetMessageId: "<duplicate@example.test>" }
    },
    {
      field: "graphMessageId",
      value: "graph-duplicate",
      message: { id: "graph-duplicate" }
    }
  ]) {
    const initial = [
      {
        graphMessageId: identity.field === "graphMessageId" ? identity.value : "graph-a",
        internetMessageId: identity.field === "internetMessageId" ? identity.value : "<mail-a@example.test>",
        subject: "Message A",
        status: "processed"
      },
      {
        graphMessageId: identity.field === "graphMessageId" ? identity.value : "graph-b",
        internetMessageId: identity.field === "internetMessageId" ? identity.value : "<mail-b@example.test>",
        subject: "Message B",
        status: "processed"
      }
    ];
    const { service, getStored } = createMemoryStateService(initial);

    assert.throws(
      () => service.findByMessage(identity.message),
      /identifiers resolve ambiguously across intake state entries/i
    );
    const reservation = service.claimMessage(identity.message);
    assert.equal(reservation.claimed, false);
    assert.match(reservation.reason, /identifiers resolve ambiguously across intake state entries/i);
    assert.throws(
      () => service.upsertEntry({ [identity.field]: identity.value, status: "skipped" }),
      /identifiers resolve ambiguously across intake state entries/i
    );
    assert.deepEqual(getStored(), initial);
  }
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

test("message reservation prevents concurrent and completed reprocessing", () => {
  const { service } = createMemoryStateService();
  const message = { id: "graph-1", internetMessageId: "<mail-1@example.test>" };
  assert.equal(service.claimMessage(message).claimed, true);
  assert.equal(service.claimMessage(message).claimed, false);
  service.upsertEntry({ ...message, graphMessageId: message.id, status: "skipped", processedAt: new Date().toISOString() });
  assert.equal(service.claimMessage(message).claimed, false);
});

test("preserved attachment metadata is reusable by content hash", () => {
  const { service } = createMemoryStateService();
  service.upsertEntry({
    graphMessageId: "graph-1",
    attachments: [{ hash: "hash-1", name: "Deck.pdf", storedName: "deck.pdf", url: "/uploads/deck.pdf" }],
    status: "processed"
  });
  assert.equal(service.findAttachmentByHash("hash-1").storedName, "deck.pdf");
});
