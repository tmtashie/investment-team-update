const test = require("node:test");
const assert = require("node:assert/strict");
const {
  applyApprovedAiUpdateProposalToInvestment,
  buildApprovedMaterialDevelopmentsReportUpdate
} = require("../services/aiUpdateProposalApplyService");

function createProposal() {
  return {
    id: "proposal-123",
    investmentId: "healing-id",
    sourceType: "PDF",
    sourceIdentifier: "PDF | Healing Board Deck.pdf",
    sourceDate: "2026-08-28",
    reviewedBy: "approver@example.test",
    reviewedAt: "2026-08-28T20:30:00.000Z",
    summary: "• 5 Units Sold (3 RaaS + 2 DP).",
    documents: [
      {
        name: "Healing Board Deck.pdf",
        url: "/uploads/healing-board-deck.pdf",
        storedName: "healing-board-deck.pdf"
      }
    ],
    extractedData: {
      source: {
        filename: "Healing Board Deck.pdf"
      },
      materialDevelopments: [
        {
          category: "sales performance",
          summary: "5 Units Sold (3 RaaS + 2 DP).",
          sourceEvidence: "5 Units Sold (3 RaaS + 2 DP)",
          sourcePage: 3,
          evidenceStatus: "verified",
          confidence: 95,
          riskLevel: "low",
          importance: "medium",
          verification: {
            numericMatchFound: true,
            contextMatchFound: true
          }
        },
        {
          category: "financial",
          summary: "Revenue increased to 21.7M.",
          sourceEvidence: "",
          sourcePage: 9,
          evidenceStatus: "unresolved",
          confidence: 97
        }
      ],
      unverifiedClaims: [
        {
          field: "revenue",
          value: "21.7M",
          evidenceStatus: "unresolved"
        }
      ]
    },
    proposedChanges: [
      {
        field: "revenue",
        proposedValue: "21.7M",
        evidenceStatus: "unresolved"
      }
    ]
  };
}

test("approved verified material developments build a report update with source and evidence audit metadata", () => {
  const update = buildApprovedMaterialDevelopmentsReportUpdate(
    createProposal(),
    "approver@example.test",
    "2026-08-28T20:30:00.000Z"
  );

  assert.equal(update.type, "AI Material Developments");
  assert.equal(update.date, "2026-08-28");
  assert.equal(update.sourceType, "PDF");
  assert.equal(update.aiProposalId, "proposal-123");
  assert.equal(update.aiApprovedBy, "approver@example.test");
  assert.equal(update.aiApprovedAt, "2026-08-28T20:30:00.000Z");
  assert.equal(update.aiSourceFilename, "Healing Board Deck.pdf");
  assert.equal(update.attachmentLink, "/uploads/healing-board-deck.pdf");
  assert.match(update.summary, /5 Units Sold/);
  assert.match(update.originalNotes, /Page 3: 5 Units Sold/);
  assert.equal(update.aiMaterialDevelopments.length, 1);
  assert.equal(update.aiMaterialDevelopments[0].sourceEvidence, "5 Units Sold (3 RaaS + 2 DP)");
  assert.equal(update.aiMaterialDevelopments[0].sourcePage, "3");
  assert.equal(update.aiMaterialDevelopments[0].confidence, 95);
  assert.doesNotMatch(JSON.stringify(update), /21\.7M|unverifiedClaims|proposedChanges/);
});

test("approval apply appends verified material developments but ignores structured proposed changes", () => {
  const calls = [];
  const result = applyApprovedAiUpdateProposalToInvestment({
    proposal: createProposal(),
    investment: {
      id: "healing-id",
      company: "Healing Innovations",
      reportUpdates: [{ date: "2026-01-01", type: "Manual Update", summary: "Existing row" }]
    },
    approver: "approver@example.test",
    approvedAt: "2026-08-28T20:30:00.000Z",
    normalizeStructuredRows: (rows) => rows || [],
    updateInvestment: (id, updates) => {
      calls.push({ id, updates });
      return { id, ...updates };
    }
  });

  assert.equal(result.applied, true);
  assert.equal(result.ignoredProposedChanges, 1);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].id, "healing-id");
  assert.equal(calls[0].updates.reportUpdates.length, 2);
  assert.equal(calls[0].updates.reportUpdates[0].aiProposalId, "proposal-123");
  assert.match(calls[0].updates.reportUpdates[0].summary, /5 Units Sold/);
  assert.doesNotMatch(JSON.stringify(calls[0].updates), /21\.7M/);
  assert.equal(calls[0].updates.reportUpdates[1].summary, "Existing row");
});

test("approval apply does not mutate investment when no verified material developments exist", () => {
  const proposal = createProposal();
  proposal.extractedData.materialDevelopments = proposal.extractedData.materialDevelopments.map((item) => ({
    ...item,
    evidenceStatus: "unresolved"
  }));

  const result = applyApprovedAiUpdateProposalToInvestment({
    proposal,
    investment: { id: "healing-id", reportUpdates: [] },
    approver: "approver@example.test",
    approvedAt: "2026-08-28T20:30:00.000Z",
    updateInvestment: () => {
      throw new Error("updateInvestment should not be called");
    },
    normalizeStructuredRows: (rows) => rows || []
  });

  assert.equal(result.applied, false);
  assert.equal(result.appendedReportUpdate, null);
});
