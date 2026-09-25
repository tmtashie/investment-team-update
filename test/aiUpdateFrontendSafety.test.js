const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const {
  buildUserFacingWarnings,
  formatDealClaimValue,
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

test("structured deal claims render evidence-backed semantic labels with their values", () => {
  const proposal = {
    dealData: {
      financingTerms: [
        { value: "1.75%", semanticLabel: "Management fee", sourceEvidence: "Management fee is 1.75%." },
        { value: "17.5%", semanticLabel: "Performance fee/carry", sourceEvidence: "Performance fee/carry is 17.5%." },
        { value: "10 years", semanticLabel: "Fund term", sourceEvidence: "The fund term is 10 years." }
      ],
      customersContractsDeployments: [
        { value: "$15.0MM", semanticLabel: "Alpha Services", sourceEvidence: "$15.0MM investment in Alpha Services" },
        { value: "$7.5MM", semanticLabel: "Beta Industrial", sourceEvidence: "$7.5MM investment in Beta Industrial" }
      ],
      historicalTargetDifference: {
        value: "$255.5MM",
        currentAvailability: false,
        sourceEvidence: "The $255.5MM difference was historically unfunded."
      }
    }
  };
  assert.equal(formatDealClaimValue(proposal, "financingTerms"), [
    "Management fee: 1.75%", "Performance fee/carry: 17.5%", "Fund term: 10 years"
  ].join("\n"));
  assert.equal(formatDealClaimValue(proposal, "customersContractsDeployments"), [
    "Alpha Services: $15.0MM", "Beta Industrial: $7.5MM"
  ].join("\n"));
  assert.equal(formatDealClaimValue(proposal, "historicalTargetDifference"), "$255.5MM");
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

test("intake preview is master-editor-only and exposed in the AI Update Inbox", () => {
  const serverSource = fs.readFileSync(path.join(__dirname, "..", "server.js"), "utf8");
  const indexSource = fs.readFileSync(path.join(__dirname, "..", "public", "index.html"), "utf8");
  const appSource = fs.readFileSync(path.join(__dirname, "..", "public", "app.js"), "utf8");
  const routeStart = serverSource.indexOf('url.pathname === "/api/ai-email-intake/preview"');
  const routeSource = serverSource.slice(routeStart, routeStart + 500);

  assert.notEqual(routeStart, -1);
  assert.match(routeSource, /requireMasterEditor\(request, response\)/);
  assert.match(indexSource, /id="previewAiEmailIntakeButton"/);
  assert.match(indexSource, /id="aiEmailIntakePreviewResult"/);
  assert.match(appSource, /fetchJson\("\/api\/ai-email-intake\/preview"\)/);
  assert.match(appSource, /isMasterEditor\(\).*previewAiEmailIntakeButton/s);
});

test("source-message reanalysis is explicit, master-editor-only, and does not call approval", () => {
  const serverSource = fs.readFileSync(path.join(__dirname, "..", "server.js"), "utf8");
  const appSource = fs.readFileSync(path.join(__dirname, "..", "public", "app.js"), "utf8");
  const routeStart = serverSource.indexOf('url.pathname === "/api/ai-email-intake/reanalyze"');
  const routeSource = serverSource.slice(routeStart, routeStart + 3000);
  assert.notEqual(routeStart, -1);
  assert.match(routeSource, /requireMasterEditor\(request, response\)/);
  assert.match(routeSource, /status: "superseded"/);
  assert.doesNotMatch(routeSource, /approveNewDealProposal|saveInvestment/);
  assert.match(appSource, /window\.confirm\("Reanalyze this preserved source email\?/);
  assert.match(appSource, /fetchJson\("\/api\/ai-email-intake\/reanalyze"/);
  assert.match(appSource, /No investment was created/);
  assert.match(appSource, /item\.status === "pending" && item\.opportunityId/);
  assert.match(appSource, /superseded.*redundant pending aliases/);
});
