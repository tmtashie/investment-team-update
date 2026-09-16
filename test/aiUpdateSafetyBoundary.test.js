const test = require("node:test");
const assert = require("node:assert/strict");
const {
  buildUserFacingWarnings,
  enforceProposalSafetyInvariant,
  finalizeAnalysisForResponse
} = require("../services/aiUpdateSafetyBoundary");

function createUnsafeRevenueAnalysis() {
  return {
    investmentMatch: {
      investmentId: "healing-innovations",
      investmentName: "Healing Innovations",
      confidence: 97
    },
    extractedFacts: [
      {
        category: "Financial",
        field: "revenue",
        value: "1.18M",
        evidenceStatus: "unresolved",
        sourceEvidence: "Not verified",
        confidence: 97
      }
    ],
    proposedChanges: [
      {
        field: "revenue",
        currentValue: "18.2M",
        proposedValue: "1.18M",
        evidenceStatus: "unresolved",
        sourceEvidence: "Not verified",
        confidence: 97
      }
    ],
    whatChanged: ["Q2 2026 Revenue decreased from 18.2M to 1.18M."],
    unverifiedClaims: [],
    warnings: ["Could not verify source evidence for revenue."]
  };
}

test("final analysis response removes unverified high-risk numeric revenue claims", () => {
  const sanitized = finalizeAnalysisForResponse(createUnsafeRevenueAnalysis());

  assert.equal(
    sanitized.extractedFacts.some((fact) => String(fact.field).toLowerCase() === "revenue"),
    false
  );
  assert.equal(
    sanitized.proposedChanges.some((change) => String(change.field).toLowerCase() === "revenue"),
    false
  );
  assert.equal(sanitized.whatChanged.some((line) => /1\.18M|18\.2M/i.test(line)), false);
  assert.deepEqual(sanitized.whatChanged, [
    "No verified portfolio changes identified from this document."
  ]);
  assert.equal(
    sanitized.unverifiedClaims.some(
      (claim) => claim.field === "revenue" && claim.value === "1.18M" && claim.source === "extractedFacts"
    ),
    true
  );
  assert.equal(
    sanitized.unverifiedClaims.some(
      (claim) => claim.field === "revenue" && claim.value === "1.18M" && claim.source === "proposedChanges"
    ),
    true
  );
});

test("proposal creation boundary removes browser-submitted unverified revenue updates", () => {
  const unsafe = createUnsafeRevenueAnalysis();
  const proposal = enforceProposalSafetyInvariant(
    {
      investmentId: "healing-innovations",
      summary: unsafe.whatChanged.join("\n"),
      extractedData: {
        facts: unsafe.extractedFacts,
        warnings: unsafe.warnings,
        unverifiedClaims: unsafe.unverifiedClaims
      },
      proposedChanges: unsafe.proposedChanges
    },
    {
      id: "healing-innovations",
      company: "Healing Innovations",
      officialValue: "50,000,000"
    }
  );

  assert.equal(proposal.proposedChanges.length, 0);
  assert.equal(
    proposal.extractedData.facts.some((fact) => String(fact.field).toLowerCase() === "revenue"),
    false
  );
  assert.equal(/1\.18M|18\.2M/i.test(proposal.summary), false);
  assert.equal(
    proposal.extractedData.unverifiedClaims.some(
      (claim) => claim.field === "revenue" && claim.source === "proposedChanges"
    ),
    true
  );
});

test("verified proposed changes and material developments survive the safety boundary", () => {
  const sanitized = finalizeAnalysisForResponse({
    extractedFacts: [
      {
        category: "Financial",
        field: "MRR",
        value: "$172,200",
        evidenceStatus: "verified",
        confidence: 92
      }
    ],
    proposedChanges: [
      {
        field: "MRR",
        currentValue: "Not currently recorded",
        proposedValue: "$172,200",
        evidenceStatus: "verified",
        confidence: 92
      }
    ],
    materialDevelopments: [
      {
        category: "Customer",
        summary: "Nobis Rehabilitation Partners added as a corporate account.",
        evidenceStatus: "verified",
        confidence: 88
      }
    ],
    whatChanged: ["Model text that should not be preserved."]
  });

  assert.equal(sanitized.extractedFacts.length, 1);
  assert.equal(sanitized.proposedChanges.length, 1);
  assert.equal(sanitized.materialDevelopments.length, 1);
  assert.deepEqual(sanitized.whatChanged, [
    "MRR changed from Not currently recorded to $172,200.",
    "Nobis Rehabilitation Partners added as a corporate account."
  ]);
});

test("final safety boundary normalizes structured warning objects", () => {
  const sanitized = finalizeAnalysisForResponse({
    extractedFacts: [],
    proposedChanges: [],
    materialDevelopments: [],
    warnings: [{ message: "Readable warning." }, { unused: "raw object" }],
    unresolved: [{ reason: "Readable reason." }]
  });

  assert.deepEqual(sanitized.warnings, ["Readable warning."]);
  assert.deepEqual(sanitized.unresolved, ["Readable reason."]);
  assert.equal(sanitized.warnings.includes("[object Object]"), false);
});

test("multiple verification failures for same field and value collapse to one user-facing warning", () => {
  const sanitized = finalizeAnalysisForResponse({
    extractedFacts: [
      { field: "revenue", value: "21.7M", evidenceStatus: "unresolved", sourcePage: 3 },
      { field: "revenue", value: "21.7M", evidenceStatus: "unresolved", sourcePage: 9 }
    ],
    proposedChanges: [
      { field: "revenue", proposedValue: "21.7M", evidenceStatus: "unresolved", sourcePage: 3 }
    ],
    materialDevelopments: [],
    warnings: [
      "Removed unsupported proposed change for revenue.",
      "Detected unsupported revenue claim.",
      "Could not verify numeric value 21.7M for revenue on source page 3.",
      "Could not verify numeric value 21.7M for revenue on source page 9."
    ],
    unresolved: ["Unsupported proposed change removed for revenue."]
  });

  assert.deepEqual(sanitized.userFacingWarnings, [
    "Revenue: 21.7M — Could not verify against source document."
  ]);
  assert.equal(sanitized.warnings.some((warning) => /Removed unsupported/i.test(warning)), true);
  assert.equal(sanitized.userFacingWarnings.some((warning) => /Removed unsupported|source page 3|source page 9/i.test(warning)), false);
});

test("user-facing warnings dedupe rejected claims across multiple pages", () => {
  const warnings = buildUserFacingWarnings({
    unverifiedClaims: [
      { field: "revenue", value: "21.7M", sourcePage: 3 },
      { field: "revenue", value: "21.7M", sourcePage: 9 }
    ]
  });

  assert.deepEqual(warnings, [
    "Revenue: 21.7M — Could not verify against source document."
  ]);
});

test("rejected material development remains visible but non-actionable", () => {
  const sanitized = finalizeAnalysisForResponse({
    extractedFacts: [],
    proposedChanges: [],
    materialDevelopments: [
      {
        category: "recurring revenue",
        summary: "Monthly Recurring Revenue was $172,200.",
        evidenceStatus: "unresolved",
        sourcePage: 6
      }
    ],
    whatChanged: ["Monthly Recurring Revenue was $172,200."],
    warnings: [],
    unresolved: []
  });

  assert.equal(sanitized.materialDevelopments.length, 0);
  assert.equal(sanitized.unverifiedClaims.length, 1);
  assert.equal(sanitized.unverifiedClaims[0].source, "materialDevelopments");
  assert.equal(sanitized.userFacingWarnings[0], "Monthly Recurring Revenue: $172,200 — Could not verify against source document.");
  assert.doesNotMatch(sanitized.whatChanged.join(" "), /172,200/);
});

test("verified developments remain unchanged in user-facing presentation", () => {
  const sanitized = finalizeAnalysisForResponse({
    extractedFacts: [],
    proposedChanges: [],
    materialDevelopments: [
      {
        category: "customer win",
        summary: "NEW Corporate account added: Nobis Rehabilitation Partners.",
        sourceEvidence: "NEW Corporate account: Nobis Rehabilitation Partners",
        evidenceStatus: "verified",
        sourcePage: 3
      }
    ],
    warnings: [],
    unresolved: []
  });

  assert.equal(sanitized.materialDevelopments.length, 1);
  assert.equal(sanitized.materialDevelopments[0].sourceEvidence, "NEW Corporate account: Nobis Rehabilitation Partners");
  assert.deepEqual(sanitized.whatChanged, ["NEW Corporate account added: Nobis Rehabilitation Partners."]);
  assert.deepEqual(sanitized.userFacingWarnings, []);
});
