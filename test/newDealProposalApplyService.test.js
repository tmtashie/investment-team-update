const test = require("node:test");
const assert = require("node:assert/strict");
const { createNewDealProposalApplyService } = require("../services/newDealProposalApplyService");

function claim(value, evidenceStatus = "verified") {
  return { value, authoritativeValue: evidenceStatus === "verified" ? value : "", evidenceStatus };
}

function createHarness(overrides = {}) {
  let proposals = [{
    id: "proposal-1",
    proposalType: "new-deal",
    status: "pending",
    proposedEntity: "Beaman Ventures",
    entityConfirmed: true,
    noExistingMatchConfirmed: true,
    amountConfirmed: true,
    opportunityFingerprint: "fingerprint-1",
    sourceMessageKey: "message-1",
    matchResult: { status: "no-match" },
    dealData: {
      companyName: claim("NewCo"),
      roundType: claim("Seed preferred"),
      proposedCheckSize: claim("250000"),
      amountBeingRaised: claim("5000000"),
      dealSummary: claim("NewCo builds financial software."),
      contactName: claim("Founder"),
      contactEmail: claim("founder@newco.example"),
      keyInvestmentPoints: [claim("Strong early traction")],
      keyRisks: [claim("Early-stage execution risk")],
      nextSteps: [claim("Schedule diligence call")],
      deadlines: [],
      relevantUrls: []
    },
    documents: [{
      id: "doc-1", name: "Deck.pdf", storedName: "deck-1.pdf", url: "/uploads/deck-1.pdf",
      hash: "hash-1", preservationStatus: "preserved", extractionStatus: "parsed"
    }]
  }];
  let investments = overrides.investments || [];
  let documents = [];
  let nextInvestment = 1;
  const service = createNewDealProposalApplyService({
    readInvestments: () => investments,
    saveInvestment: (entry) => {
      const saved = { ...entry, id: `investment-${nextInvestment++}` };
      investments = [saved, ...investments];
      return saved;
    },
    readAiUpdateProposals: () => proposals,
    updateAiUpdateProposal: (id, updates) => {
      proposals = proposals.map((proposal) => proposal.id === id ? { ...proposal, ...updates } : proposal);
      return proposals.find((proposal) => proposal.id === id);
    },
    readCompanyDocuments: () => documents,
    saveCompanyDocument: (entry) => {
      const saved = { ...entry, id: `company-doc-${documents.length + 1}` };
      documents = [saved, ...documents];
      return saved;
    },
    normalizeCompanyKey: (value) => String(value || "").trim().toLowerCase()
  });
  return { service, getInvestments: () => investments, getDocuments: () => documents, getProposals: () => proposals };
}

test("approval creates one financially empty New Lead and associates the same binary", async () => {
  const harness = createHarness();
  const result = await harness.service.approveNewDealProposal("proposal-1", "tyler@example.test");
  assert.equal(result.investment.assetType, "Private Investment");
  assert.equal(result.investment.status, "New Lead");
  assert.equal(result.investment.amount, "250000");
  assert.deepEqual(result.investment.capitalActivity, []);
  assert.deepEqual(result.investment.valuationHistory, []);
  assert.deepEqual(result.investment.ownershipHistory, []);
  assert.equal(harness.getDocuments()[0].storedName, "deck-1.pdf");
  assert.equal(harness.getDocuments()[0].sourceProposalId, "proposal-1");
});

test("concurrent and repeated approvals cannot duplicate investments or documents", async () => {
  const harness = createHarness();
  const [first, second] = await Promise.all([
    harness.service.approveNewDealProposal("proposal-1", "tyler@example.test"),
    harness.service.approveNewDealProposal("proposal-1", "tyler@example.test")
  ]);
  assert.equal(harness.getInvestments().length, 1);
  assert.equal(harness.getDocuments().length, 1);
  assert.equal(first.investment.id, second.investment.id);
  assert.equal(second.idempotent, true);
});

test("approval-time duplicate check blocks an existing company", async () => {
  const harness = createHarness({ investments: [{ id: "existing", company: "NewCo", investmentAliases: [] }] });
  await assert.rejects(
    () => harness.service.approveNewDealProposal("proposal-1", "tyler@example.test"),
    /matching investment already exists/i
  );
  assert.equal(harness.getInvestments().length, 1);
});

test("ambiguous proposal cannot be approved without explicit no-match confirmation", async () => {
  const harness = createHarness();
  harness.getProposals()[0].matchResult.status = "ambiguous";
  harness.getProposals()[0].noExistingMatchConfirmed = false;
  await assert.rejects(
    () => harness.service.approveNewDealProposal("proposal-1", "tyler@example.test"),
    /Resolve the ambiguous/i
  );
});

test("round size never becomes pipeline amount when check size is unresolved", async () => {
  const harness = createHarness();
  harness.getProposals()[0].dealData.proposedCheckSize = claim("250000", "unresolved");
  const result = await harness.service.approveNewDealProposal("proposal-1", "tyler@example.test");
  assert.equal(result.investment.amount, "");
  assert.equal(result.investment.dealData, undefined);
});

test("server approval authority is master-editor only", () => {
  const { canApprovePotentialNewDeal } = require("../server")._test;
  assert.equal(canApprovePotentialNewDeal({ role: "master-editor" }), true);
  assert.equal(canApprovePotentialNewDeal({ role: "editor" }), false);
  assert.equal(canApprovePotentialNewDeal({ role: "dashboard-viewer" }), false);
});

test("manual edits cannot upgrade changed financial terms to verified", () => {
  const { applyNewDealEdits } = require("../server")._test;
  const proposal = { dealData: { proposedCheckSize: claim("250000", "verified") } };
  const unchanged = applyNewDealEdits(proposal, { dealData: { proposedCheckSize: "250000" } });
  const changed = applyNewDealEdits(proposal, { dealData: { proposedCheckSize: "500000" } });
  assert.equal(unchanged.proposedCheckSize.evidenceStatus, "verified");
  assert.equal(changed.proposedCheckSize.evidenceStatus, "unresolved");
  assert.equal(changed.proposedCheckSize.authoritativeValue, "");
});
