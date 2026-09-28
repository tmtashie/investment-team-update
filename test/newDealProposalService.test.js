const test = require("node:test");
const assert = require("node:assert/strict");
const { createAiUpdateProposalService } = require("../services/aiUpdateProposalService");

function createHarness(initial = []) {
  let stored = structuredClone(initial);
  let generatedId = 0;
  const service = createAiUpdateProposalService({
    AI_UPDATE_PROPOSALS_FILE: "proposals.json",
    readJsonFile: () => stored,
    writeJsonFile: (_file, value) => { stored = value; },
    writeMetadata: () => {},
    normalizeAiUpdateProposal: (value) => ({
      id: value.id || `proposal-${++generatedId}`,
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

test("explicit reanalysis may refresh a pending source-opportunity proposal in place", () => {
  const harness = createHarness();
  const first = harness.service.saveAiUpdateProposal({
    proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-1",
    opportunityFingerprint: "fp-1", summary: "Stale analysis"
  });
  const refreshed = harness.service.saveAiUpdateProposal({
    proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-1",
    opportunityFingerprint: "fp-1", summary: "Corrected analysis"
  }, { replacePendingSourceOpportunity: true });
  assert.equal(refreshed.id, first.id);
  assert.equal(refreshed.summary, "Corrected analysis");
  assert.equal(refreshed.status, "pending");
  assert.equal(harness.getStored().length, 1);
});

test("explicit reanalysis matches a pending source opportunity through canonical identity aliases", () => {
  const harness = createHarness();
  const first = harness.service.saveAiUpdateProposal({
    proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "canonical-pure",
    opportunityIdentityKeys: ["canonical-pure"], opportunityFingerprint: "fp-1", summary: "Project Pure"
  });
  const refreshed = harness.service.saveAiUpdateProposal({
    proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "new-model-name",
    opportunityIdentityKeys: ["canonical-pure", "new-model-name"], opportunityFingerprint: "fp-1", summary: "Project Pure Co-Investment"
  }, { replacePendingSourceOpportunity: true });
  assert.equal(refreshed.id, first.id);
  assert.equal(harness.getStored().length, 1);
  assert.equal(refreshed.summary, "Project Pure Co-Investment");
});

test("explicit reanalysis never overwrites an approved or rejected source opportunity", () => {
  for (const status of ["approved", "rejected"]) {
    const harness = createHarness();
    const terminal = harness.service.saveAiUpdateProposal({
      proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-1",
      opportunityFingerprint: "fp-1", summary: "Reviewed", status
    });
    const result = harness.service.saveAiUpdateProposal({
      proposalType: "new-deal", sourceMessageKey: "message-1", opportunityId: "opp-1",
      opportunityFingerprint: "fp-1", summary: "Replacement", status: "pending"
    }, { replacePendingSourceOpportunity: true });
    assert.equal(result.id, terminal.id);
    assert.equal(result.status, status);
    assert.equal(result.summary, "Reviewed");
    assert.equal(harness.getStored().length, 1);
  }
});

test("source reconciliation supersedes pending aliases but preserves reviewed aliases", () => {
  const initial = [
    { id: "pure", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "Project Pure", opportunityId: "pure-id", status: "pending", updatedAt: "v1", documents: [] },
    { id: "pure-alias", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "Project Pure Co-Investment", opportunityId: "legacy-pure-id", status: "pending", updatedAt: "v1", documents: [] },
    { id: "pure-approved", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "Project Pure Co-Investment", opportunityId: "legacy-approved-id", status: "approved", updatedAt: "v1", documents: [] },
    { id: "care", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "Project Care", opportunityId: "care-id", status: "pending", updatedAt: "v1", documents: [] },
    { id: "care-rejected", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "Project Care Co-Investment", opportunityId: "legacy-rejected-id", status: "rejected", updatedAt: "v1", documents: [] }
  ];
  const harness = createHarness(initial);
  const snapshot = harness.service.sourceProposalSnapshot("message-1");
  const result = harness.service.reconcilePendingSourceOpportunities({
    sourceMessageKey: "message-1",
    expectedSnapshot: snapshot,
    reviewer: "master@example.test",
    reconciledAt: "2026-09-22T22:00:00.000Z",
    canonicalEntries: [
      { proposalType: "new-deal", opportunityName: "Project Pure", opportunityId: require("../services/newDealAnalysisService").opportunityIdentity("Project Pure"), status: "pending" },
      { proposalType: "new-deal", opportunityName: "Project Care", opportunityId: require("../services/newDealAnalysisService").opportunityIdentity("Project Care"), status: "pending" }
    ]
  });
  assert.deepEqual(result.supersededProposals.map((proposal) => proposal.id), ["pure-alias"]);
  assert.equal(harness.getStored().find((proposal) => proposal.id === "pure-approved").status, "approved");
  assert.equal(harness.getStored().find((proposal) => proposal.id === "care-rejected").status, "rejected");
});

test("source reconciliation fails closed when a proposal changes after the snapshot", () => {
  const harness = createHarness([{
    id: "pure", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "Project Pure",
    opportunityId: "pure-id", status: "pending", updatedAt: "v1", documents: []
  }]);
  const snapshot = harness.service.sourceProposalSnapshot("message-1");
  harness.service.updateAiUpdateProposal("pure", { status: "approved" });
  assert.throws(() => harness.service.reconcilePendingSourceOpportunities({
    sourceMessageKey: "message-1",
    expectedSnapshot: snapshot,
    reviewer: "master@example.test",
    canonicalEntries: [{ proposalType: "new-deal", opportunityName: "Project Pure", opportunityId: "pure-id" }]
  }), /changed during reanalysis/);
  assert.equal(harness.getStored()[0].status, "approved");
});

test("source reconciliation leaves an unpartitioned legacy proposal for the existing legacy supersession path", () => {
  const harness = createHarness([{
    id: "legacy", proposalType: "new-deal", sourceMessageKey: "message-1", opportunityName: "",
    opportunityId: "", status: "pending", updatedAt: "v1",
    documents: [{ graphAttachmentId: "core" }, { graphAttachmentId: "pure" }, { graphAttachmentId: "care" }]
  }]);
  const result = harness.service.reconcilePendingSourceOpportunities({
    sourceMessageKey: "message-1",
    expectedSnapshot: harness.service.sourceProposalSnapshot("message-1"),
    reviewer: "master@example.test",
    canonicalEntries: [
      { proposalType: "new-deal", opportunityName: "Core Fund", opportunityId: "core", documents: [{ graphAttachmentId: "core" }] },
      { proposalType: "new-deal", opportunityName: "Project Pure", opportunityId: "pure", documents: [{ graphAttachmentId: "pure" }] },
      { proposalType: "new-deal", opportunityName: "Project Care", opportunityId: "care", documents: [{ graphAttachmentId: "care" }] }
    ]
  });
  assert.equal(result.proposals.length, 3);
  assert.equal(result.supersededProposals.length, 0);
  assert.equal(harness.getStored().find((proposal) => proposal.id === "legacy").status, "pending");
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
