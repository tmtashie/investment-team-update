const test = require("node:test");
const assert = require("node:assert/strict");
const { createAiUpdateProposalService } = require("../services/aiUpdateProposalService");

function createHarness() {
  let stored = [];
  const service = createAiUpdateProposalService({
    AI_UPDATE_PROPOSALS_FILE: "proposals.json",
    readJsonFile: () => stored,
    writeJsonFile: (_file, value) => { stored = value; },
    writeMetadata: () => {},
    normalizeAiUpdateProposal: (value) => ({
      id: value.id || `proposal-${stored.length + 1}`,
      proposalType: value.proposalType || "investment-update",
      status: value.status || "pending",
      sourceMessageKey: value.sourceMessageKey || "",
      sourceMessageKeys: value.sourceMessageKeys || [],
      opportunityFingerprint: value.opportunityFingerprint || "",
      documents: value.documents || [],
      createdAt: value.createdAt || new Date().toISOString(),
      ...value
    }),
    createBackupSnapshot: () => {},
    applyApprovedAiUpdateProposal: () => ({ applied: false })
  });
  return { service, getStored: () => stored };
}

test("same source message cannot create two new-deal proposals", () => {
  const harness = createHarness();
  const first = harness.service.saveAiUpdateProposal({ proposalType: "new-deal", sourceMessageKey: "message-1", opportunityFingerprint: "fp-1" });
  const second = harness.service.saveAiUpdateProposal({ proposalType: "new-deal", sourceMessageKey: "message-1", opportunityFingerprint: "fp-1" });
  assert.equal(first.id, second.id);
  assert.equal(harness.getStored().length, 1);
});

test("one source message may create distinct proposals but not duplicate a source-opportunity pair", () => {
  const harness = createHarness();
  const first = harness.service.saveAiUpdateProposal({ proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-1", opportunityFingerprint: "fp-1" });
  const second = harness.service.saveAiUpdateProposal({ proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-2", opportunityFingerprint: "fp-2" });
  const duplicate = harness.service.saveAiUpdateProposal({ proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-1", opportunityFingerprint: "fp-1" });
  assert.notEqual(first.id, second.id);
  assert.equal(duplicate.id, first.id);
  assert.equal(harness.getStored().length, 2);
});

test("source-opportunity idempotency also protects an existing-investment proposal route", () => {
  const harness = createHarness();
  const first = harness.service.saveAiUpdateProposal({ proposalType: "investment-update", sourceMessageKey: "message-1", opportunityId: "opp-1", investmentId: "existing-1" });
  const duplicate = harness.service.saveAiUpdateProposal({ proposalType: "investment-update", sourceMessageKey: "message-1", opportunityId: "opp-1", investmentId: "existing-1" });
  assert.equal(duplicate.id, first.id);
  assert.equal(harness.getStored().length, 1);
});

test("multiple emails for one opportunity coalesce attachments by hash", () => {
  const harness = createHarness();
  const first = harness.service.saveAiUpdateProposal({
    proposalType: "new-deal", sourceMessageKey: "message-1", sourceMessageKeys: ["message-1"], opportunityFingerprint: "fp-1",
    documents: [{ name: "Deck.pdf", hash: "hash-1" }]
  });
  const second = harness.service.saveAiUpdateProposal({
    proposalType: "new-deal", sourceMessageKey: "message-2", sourceMessageKeys: ["message-2"], opportunityFingerprint: "fp-1",
    documents: [{ name: "Deck copy.pdf", hash: "hash-1" }, { name: "Terms.pdf", hash: "hash-2" }]
  });
  assert.equal(first.id, second.id);
  assert.equal(harness.getStored().length, 1);
  assert.deepEqual(harness.getStored()[0].sourceMessageKeys.sort(), ["message-1", "message-2"]);
  assert.equal(harness.getStored()[0].documents.length, 2);
});
