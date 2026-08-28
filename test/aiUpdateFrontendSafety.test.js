const test = require("node:test");
const assert = require("node:assert/strict");
const {
  buildUserFacingWarnings,
  getReportUpdatesEmptyMessage,
  isHighRiskNumeric,
  sanitizeForActionableView,
  shouldRefreshInvestmentsAfterAiProposalAction,
  warningMessage
} = require("../public/aiUpdateSafety");

test("frontend safety helper blocks unresolved financial items from actionable sections", () => {
  const sanitized = sanitizeForActionableView({
    extractedFacts: [
      {
        field: "revenue",
        value: "1.18M",
        evidenceStatus: "unresolved"
      }
    ],
    proposedChanges: [
      {
        field: "revenue",
        currentValue: "18.2M",
        proposedValue: "1.18M",
        evidenceStatus: "unresolved"
      }
    ],
    whatChanged: ["Q2 2026 Revenue decreased from 18.2M to 1.18M."],
    unverifiedClaims: []
  });

  assert.equal(isHighRiskNumeric({ field: "revenue", value: "1.18M" }), true);
  assert.equal(sanitized.extractedFacts.length, 0);
  assert.equal(sanitized.proposedChanges.length, 0);
  assert.deepEqual(sanitized.whatChanged, [
    "No verified portfolio changes identified from this document."
  ]);
  assert.equal(
    sanitized.unverifiedClaims.some(
      (claim) => claim.field === "revenue" && claim.value === "1.18M"
    ),
    true
  );
});

test("frontend warning helper never renders structured warning objects as object serialization", () => {
  assert.equal(warningMessage({ message: "Readable warning." }), "Readable warning.");
  assert.equal(warningMessage({ reason: "Readable reason." }), "Readable reason.");
  assert.equal(warningMessage({ nested: { value: true } }), "");
  assert.notEqual(warningMessage({ message: "Readable warning." }), "[object Object]");
});

test("frontend user-facing warnings hide internal sanitizer messages", () => {
  const warnings = buildUserFacingWarnings({
    unverifiedClaims: [
      { field: "revenue", value: "21.7M", sourcePage: 3, reason: "Removed unsupported proposed change." },
      { field: "revenue", value: "21.7M", sourcePage: 9, reason: "Could not verify numeric value on source page 9." }
    ]
  });

  assert.deepEqual(warnings, [
    "Revenue: 21.7M — Could not verify against source document."
  ]);
  assert.equal(warnings.some((warning) => /Removed unsupported|source page/i.test(warning)), false);
});

test("frontend rejected material development is visible only as unverified and not actionable", () => {
  const sanitized = sanitizeForActionableView({
    materialDevelopments: [
      {
        category: "recurring revenue",
        summary: "Monthly Recurring Revenue was $172,200.",
        evidenceStatus: "unresolved",
        sourcePage: 6
      }
    ],
    whatChanged: ["Monthly Recurring Revenue was $172,200."],
    unverifiedClaims: []
  });

  assert.equal(sanitized.materialDevelopments.length, 0);
  assert.equal(sanitized.unverifiedClaims.length, 1);
  assert.equal(sanitized.unverifiedClaims[0].source, "materialDevelopments");
  assert.deepEqual(sanitized.userFacingWarnings, [
    "Monthly Recurring Revenue: $172,200 — Could not verify against source document."
  ]);
  assert.doesNotMatch(sanitized.whatChanged.join(" "), /172,200/);
});

test("successful approval action triggers investment data refresh decision", () => {
  assert.equal(
    shouldRefreshInvestmentsAfterAiProposalAction("approve", { proposal: { id: "proposal-1" } }),
    true
  );
  assert.equal(
    shouldRefreshInvestmentsAfterAiProposalAction("reject", { proposal: { id: "proposal-1" } }),
    false
  );
  assert.equal(shouldRefreshInvestmentsAfterAiProposalAction("approve", {}), false);
});

test("report update empty state distinguishes no rows from filtered-out rows", () => {
  assert.equal(
    getReportUpdatesEmptyMessage(0, 0),
    "No saved updates or reports yet. Add your first monthly report, quarterly letter, capital call, or call note above."
  );
  assert.equal(
    getReportUpdatesEmptyMessage(1, 0),
    "No reports match the current filters."
  );
  assert.equal(
    getReportUpdatesEmptyMessage(2, 1),
    "No saved updates or reports yet. Add your first monthly report, quarterly letter, capital call, or call note above."
  );
});
