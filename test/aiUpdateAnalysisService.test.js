const test = require("node:test");
const assert = require("node:assert/strict");
const {
  createAiUpdateAnalysisService,
  normalizeAnalysisResult,
  _test
} = require("../services/aiUpdateAnalysisService");

function normalizeEntityName(value) {
  return String(value || "").trim();
}

const investments = [
  {
    id: "inv-1",
    company: "Northwind Ventures",
    entity: "Beaman Ventures",
    assetType: "Private Investment",
    stage: "Growth",
    status: "Active",
    amount: "1,000,000",
    currency: "USD",
    officialValue: "18,200,000",
    internalValue: "18,200,000",
    valuationDate: "2026-03-31",
    capitalCallAmount: "",
    distributionAmount: "",
    notes: "Q1 revenue was $18.2M."
  },
  {
    id: "inv-2",
    company: "Southridge Fund II",
    entity: "Lee Beaman",
    assetType: "Fund",
    status: "Active"
  }
];

const entities = ["Beaman Ventures", "Lee Beaman"];

const finsyncInvestments = [
  {
    id: "finsync-id",
    company: "FINSYNC",
    entity: "Beaman Ventures",
    assetType: "Private Investment",
    status: "Active",
    investmentAliases: ["FINSYNC, Inc."]
  },
  {
    id: "vanguard-id",
    company: "VANGUARD SHORT-TERM CORPORATE BOND ETF",
    entity: "Lee Beaman",
    assetType: "Public Stock",
    status: "Active"
  }
];

function createService(rawResponse) {
  let calls = 0;
  const service = createAiUpdateAnalysisService({
    normalizeEntityName,
    getNow: () => new Date("2026-08-27T12:00:00.000Z"),
    callModel: async () => {
      calls += 1;
      return rawResponse;
    }
  });
  return { service, getCalls: () => calls };
}

function createSequenceService(rawResponses) {
  let calls = 0;
  const prompts = [];
  const service = createAiUpdateAnalysisService({
    normalizeEntityName,
    getNow: () => new Date("2026-08-27T12:00:00.000Z"),
    callModel: async (prompt) => {
      prompts.push(prompt);
      const response = rawResponses[Math.min(calls, rawResponses.length - 1)];
      calls += 1;
      return response;
    }
  });
  return { service, getCalls: () => calls, prompts };
}

const healingInvestments = [
  {
    id: "healing-id",
    company: "Healing Innovations",
    entity: "Beaman Ventures",
    assetType: "Private Investment",
    status: "Active",
    investmentAliases: ["Healing Innovations, Inc.", "Healing Innovations Board"]
  },
  {
    id: "other-id",
    company: "Other Medical Device Company",
    entity: "Lee Beaman",
    assetType: "Private Investment",
    status: "Active"
  }
];

test("strict numeric equivalence preserves financial scale", () => {
  assert.equal(_test.numericValuesExactlyEquivalent("21.7M", "21,700,000"), true);
  assert.equal(_test.numericValuesExactlyEquivalent("21.7M", "$21.7 million"), true);
  assert.equal(_test.numericValuesExactlyEquivalent("21.7M", "21"), false);
  assert.equal(_test.numericValuesExactlyEquivalent("21.7M", "21.7"), false);
  assert.equal(_test.numericValuesExactlyEquivalent("21.7M", "21,700"), false);
  assert.equal(_test.numericValuesExactlyEquivalent("172.2K", "172,200"), true);
  assert.equal(_test.numericValuesExactlyEquivalent("172.2K", "172.2"), false);
  assert.equal(_test.numericValuesExactlyEquivalent("3M", "3,000,000"), true);
  assert.equal(_test.numericValuesExactlyEquivalent("3M", "3"), false);
});

test("strict numeric matcher rejects Page 9 KPI tokens for 21.7M", () => {
  const page9 =
    "Q2 KPIs Total Units Sold: 57 Sales Meetings: 89 Total Units Deployed: 53 Total Units Trained: 51 Total Steps Last Quarter: 3,920,989 Total Sessions: 23,908 Total Patients: 4,093 Service Calls: 45 Operations FTE: 21 (+1) Part-Time: 4 Key Contractors: 6 People";

  const match = _test.findNumericMatch(page9, "21.7M");

  assert.equal(match.numericMatchFound, false);
  assert.equal(match.matchedNumericText, "");
});

test("source explicitly says FINSYNC and portfolio contains FINSYNC", async () => {
  const { service } = createService({
    investmentMatch: {
      investmentId: "finsync-id",
      investmentName: "FINSYNC",
      confidence: 80,
      reason: "Model saw FINSYNC."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: ["July revenue increased."],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      subject: "FINSYNC July customer and revenue update",
      sender: "updates@finsync.com",
      sourceText: "FINSYNC July revenue was $3.0 million."
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.equal(result.analysis.investmentMatch.confidence, 98);
  assert.match(result.analysis.investmentMatch.reason, /Exact source body match for 'FINSYNC'/);
});

test("FINSYNC subject and sender domain deterministically match FINSYNC instead of Healing Innovations", async () => {
  const { service } = createService({
    investmentMatch: {
      investmentId: "healing-id",
      investmentName: "Healing Innovations",
      confidence: 84,
      reason: "Model guessed from strategic update language."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "Email",
      subject: "HIGHLY CONFIDENTIAL: FINSYNC August Update",
      sender: "tuckermathis@finsync.com",
      sourceText: "August investor update covering customer growth, monthly recurring revenue, and sales pipeline."
    },
    investments: [
      ...finsyncInvestments,
      {
        id: "healing-id",
        company: "Healing Innovations",
        entity: "Beaman Ventures",
        assetType: "Private Investment",
        status: "Active",
        investmentAliases: ["Healing Innovations, Inc."]
      }
    ],
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.notEqual(result.analysis.investmentMatch.investmentId, "healing-id");
  assert.equal(result.analysis.candidates.some((candidate) => candidate.investmentId === "finsync-id"), true);
  assert.match(result.analysis.investmentMatch.reason, /Exact subject match for 'FINSYNC'/);
});

test("explicit FINSYNC source must not match Vanguard", async () => {
  const { service } = createService({
    investmentMatch: {
      investmentId: "vanguard-id",
      investmentName: "VANGUARD SHORT-TERM CORPORATE BOND ETF",
      confidence: 85,
      reason: "Model guessed from financial context."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: ["FINSYNC reported July operating metrics."],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      subject: "FINSYNC July customer and revenue update",
      sender: "updates@finsync.com",
      sourceText: "FINSYNC total customers reached 62,824 in July."
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.notEqual(result.analysis.investmentMatch.investmentId, "vanguard-id");
  assert.match(result.analysis.warnings.join(" "), /conflicted with explicit source evidence/);
});

test("exact name match produces high confidence", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 72 },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC, Inc. issued its monthly investor update." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.ok(result.analysis.investmentMatch.confidence >= 95);
});

test("semantic-only match produces lower confidence and warning", async () => {
  const { service } = createService({
    investmentMatch: {
      investmentId: "vanguard-id",
      investmentName: "VANGUARD SHORT-TERM CORPORATE BOND ETF",
      confidence: 91,
      reason: "Financial update sounded similar."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      subject: "July customer and revenue update",
      sourceText: "Total customers reached 62,824 in July and revenue was $3.0 million."
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "vanguard-id");
  assert.equal(result.analysis.investmentMatch.confidence, 84);
  assert.match(result.analysis.warnings.join(" "), /lacks explicit/);
});

test("competing candidates lower confidence", async () => {
  const investmentsWithCompetition = finsyncInvestments.concat({
    id: "finsync-holdings-id",
    company: "FINSYNC Holdings",
    entity: "Lee Beaman",
    assetType: "Private Investment",
    status: "Active"
  });
  const { service } = createService({
    investmentMatch: {
      investmentId: "finsync-holdings-id",
      investmentName: "FINSYNC Holdings",
      confidence: 96,
      reason: "Exact name appeared."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC Holdings sent an investor update." },
    investments: investmentsWithCompetition,
    entities
  });

  assert.ok(result.analysis.investmentMatch.confidence <= 88);
  assert.match(result.analysis.warnings.join(" "), /Multiple plausible/);
});

test("$3,000,000 source text verifies normalized 3.0M revenue evidence", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "3.0M",
        unit: "USD",
        sourceEvidence: "revenue was 3.0M",
        confidence: 94
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC reported July revenue of $3,000,000." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts[0].sourceEvidence, "FINSYNC reported July revenue of $3,000,000.");
  assert.equal(result.analysis.extractedFacts[0].evidenceStatus, "verified");
});

test("$3.0 million source text verifies normalized 3.0M revenue evidence", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "3.0M",
        unit: "USD",
        sourceEvidence: "July revenue",
        confidence: 94
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC July revenue increased to $3.0 million from $2.83 million." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts[0].sourceEvidence, "FINSYNC July revenue increased to $3.0 million from $2.83 million.");
  assert.equal(result.analysis.extractedFacts[0].evidenceStatus, "verified");
});

test("62,824 source text verifies normalized 62824 customer evidence", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "customerCount",
        value: "62824",
        sourceEvidence: "customers reached 62824",
        confidence: 96
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "Customer count reached 62,824 in July for FINSYNC." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts[0].sourceEvidence, "Customer count reached 62,824 in July for FINSYNC.");
  assert.equal(result.analysis.extractedFacts[0].evidenceStatus, "verified");
});

test("context prevents selecting a cash sentence as evidence for revenue", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "3.0M",
        sourceEvidence: "Revenue was $3.0M",
        confidence: 90
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "Cash balance was $3.0M. Revenue was $2.0M." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts.length, 0);
  assert.equal(result.analysis.unverifiedClaims[0].field, "revenue");
  assert.equal(result.analysis.unverifiedClaims[0].verification.numericMatchFound, true);
  assert.equal(result.analysis.unverifiedClaims[0].verification.contextMatchFound, false);
  assert.match(result.analysis.unresolved.join(" "), /Could not verify source evidence for revenue/);
});

test("fabricated model evidence is not displayed as verified", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "9.0M",
        sourceEvidence: "Revenue was $9.0M",
        confidence: 90
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC July revenue was $3.0 million." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts.length, 0);
  assert.equal(result.analysis.unverifiedClaims[0].field, "revenue");
  assert.equal(result.analysis.unverifiedClaims[0].verification.numericMatchFound, false);
});

test("duplicate fact and change evidence warnings are deduplicated", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "9.0M",
        sourceEvidence: "Revenue was $9.0M",
        confidence: 90
      }
    ],
    whatChanged: [],
    proposedChanges: [
      {
        field: "revenue",
        currentValue: "2.83M",
        proposedValue: "9.0M",
        sourceEvidence: "Revenue was $9.0M",
        confidence: 90
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC July revenue was $3.0 million." },
    investments: finsyncInvestments,
    entities
  });
  const evidenceWarnings = result.analysis.unresolved.filter((item) =>
    item.includes("Could not verify source evidence for revenue")
  );

  assert.equal(evidenceWarnings.length, 1);
});

test("revenue comparison produces useful summary language", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: ["Updated revenue metrics."],
    proposedChanges: [
      {
        field: "revenue",
        currentValue: "2.83M",
        proposedValue: "3.0M",
        period: "July",
        sourceEvidence: "July revenue increased to $3.0 million from $2.83 million.",
        confidence: 95
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC July revenue increased to $3.0 million from $2.83 million." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.proposedChanges[0].currentValue, "Not currently recorded");
  assert.equal(result.analysis.whatChanged[0], "No verified portfolio changes identified from this document.");
});

test("customerCount comparison produces useful summary language", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: ["Updated customer count metrics."],
    proposedChanges: [
      {
        field: "customerCount",
        currentValue: "58,626",
        proposedValue: "62,824",
        sourceEvidence: "Customer count increased to 62,824 in July.",
        confidence: 95
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "Customer count increased to 62,824 in July." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.proposedChanges[0].currentValue, "Not currently recorded");
  assert.equal(result.analysis.whatChanged[0], "No verified portfolio changes identified from this document.");
});

test("summary does not claim increase or decrease when comparison is not valid", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 80 },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: ["Management changed the revenue reporting format."],
    proposedChanges: [
      {
        field: "revenue",
        currentValue: "",
        proposedValue: "3.0M",
        sourceEvidence: "Revenue was $3.0 million.",
        confidence: 95
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "FINSYNC revenue was $3.0 million." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.whatChanged[0], "No verified portfolio changes identified from this document.");
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /increased|decreased/);
});

test("FINSYNC regression returns grounded evidence and specific what changed", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", investmentName: "FINSYNC", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "3.0M",
        unit: "USD",
        sourceEvidence: "July revenue was 3.0M",
        confidence: 97
      },
      {
        category: "operating metrics",
        field: "customerCount",
        value: "62824",
        sourceEvidence: "customers reached 62824",
        confidence: 97
      }
    ],
    whatChanged: ["Updated revenue and customer count metrics."],
    proposedChanges: [
      {
        field: "revenue",
        currentValue: "2.83M",
        proposedValue: "3.0M",
        period: "July",
        sourceEvidence: "July revenue was 3.0M",
        confidence: 97,
        riskLevel: "medium"
      },
      {
        field: "customerCount",
        currentValue: "58,626",
        proposedValue: "62,824",
        period: "July",
        sourceEvidence: "customers reached 62824",
        confidence: 97,
        riskLevel: "medium"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      subject: "FINSYNC July customer and revenue update",
      sourceText: [
        "FINSYNC July revenue increased to $3.0 million from $2.83 million.",
        "Customer count increased to 62,824 in July."
      ].join(" ")
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.equal(result.analysis.proposedChanges[0].sourceEvidence, "FINSYNC July revenue increased to $3.0 million from $2.83 million.");
  assert.equal(result.analysis.proposedChanges[0].evidenceStatus, "verified");
  assert.equal(result.analysis.proposedChanges[0].currentValue, "Not currently recorded");
  assert.equal(result.analysis.proposedChanges[1].sourceEvidence, "Customer count increased to 62,824 in July.");
  assert.equal(result.analysis.proposedChanges[1].evidenceStatus, "verified");
  assert.equal(result.analysis.proposedChanges[1].currentValue, "Not currently recorded");
  assert.equal(result.analysis.whatChanged[0], "No verified portfolio changes identified from this document.");
});

test("email body source-backed operating metrics survive when model misses them and unsupported revenue is removed", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", investmentName: "FINSYNC", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 97
      }
    ],
    materialDevelopments: [
      {
        category: "customer growth",
        summary: "Customer base increased to 62,824.",
        sourceEvidence: "Customer base increased to 62,824.",
        confidence: 93
      }
    ],
    proposedChanges: [
      {
        field: "revenue",
        proposedValue: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 97,
        riskLevel: "medium"
      }
    ],
    whatChanged: [
      "Revenue increased to 21.7M.",
      "Customer base increased to 62,824."
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "Email",
      subject: "HIGHLY CONFIDENTIAL: FINSYNC August Update",
      sender: "tuckermathis@finsync.com",
      sourceText: "FINSYNC update. Based on results through July, we are now at approximately 62,824 customers and $3.0 million in monthly recurring revenue."
    },
    investments: finsyncInvestments,
    entities
  });

  const customerFact = result.analysis.extractedFacts.find((fact) => fact.field === "customerCount");
  const mrrFact = result.analysis.extractedFacts.find((fact) => fact.field === "monthly recurring revenue");

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.equal(customerFact.evidenceStatus, "verified");
  assert.equal(customerFact.value, "62,824");
  assert.equal(customerFact.sourceEvidence, "Based on results through July, we are now at approximately 62,824 customers and $3.0 million in monthly recurring revenue.");
  assert.equal(mrrFact.evidenceStatus, "verified");
  assert.equal(mrrFact.value, "$3.0 million");
  assert.equal(mrrFact.sourceEvidence, "Based on results through July, we are now at approximately 62,824 customers and $3.0 million in monthly recurring revenue.");
  assert.equal(result.analysis.extractedFacts.some((fact) => fact.field === "monthly recurring revenue" && fact.value === "62,824"), false);
  assert.equal(result.analysis.extractedFacts.some((fact) => fact.field === "revenue" && String(fact.value).includes("3.0")), false);
  assert.equal(result.analysis.extractedFacts.some((fact) => fact.field === "revenue" && String(fact.value).includes("21.7")), false);
  assert.equal(result.analysis.proposedChanges.length, 0);
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /21\.7M|revenue increased/i);
  assert.equal(result.analysis.unverifiedClaims.some((claim) => claim.field === "revenue" && claim.value === "21.7M"), true);
});

test("email operating-metric extraction keeps customer counts separate from monetary MRR", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", investmentName: "FINSYNC", confidence: 80 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "$3.0 million",
        sourceEvidence: "$3.0 million in monthly recurring revenue",
        confidence: 97
      },
      {
        category: "operating metrics",
        field: "revenue",
        value: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 97
      }
    ],
    materialDevelopments: [],
    proposedChanges: [
      {
        field: "revenue",
        proposedValue: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 97,
        riskLevel: "medium"
      }
    ],
    whatChanged: ["Revenue increased to 21.7M."],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "Email",
      subject: "HIGHLY CONFIDENTIAL: Operating Update",
      sourceText: [
        "By the end of April, we had grown to 58,626 customers and $2.83 million in monthly recurring revenue.",
        "Based on results through July, we are now at approximately 62,824 customers and $3.0 million in monthly recurring revenue.",
        "A different company claims to have five times our revenue and 200 times our revenue."
      ].join(" ")
    },
    investments: finsyncInvestments,
    entities
  });

  const factKeys = result.analysis.extractedFacts.map((fact) => `${fact.field}:${fact.value}`);

  assert.equal(result.analysis.extractedFacts.filter((fact) => fact.field === "customerCount" && fact.value === "58,626").length, 1);
  assert.equal(result.analysis.extractedFacts.filter((fact) => fact.field === "customerCount" && fact.value === "62,824").length, 1);
  assert.equal(result.analysis.extractedFacts.filter((fact) => fact.field === "monthly recurring revenue" && fact.value === "$2.83 million").length, 1);
  assert.equal(result.analysis.extractedFacts.filter((fact) => fact.field === "monthly recurring revenue" && fact.value === "$3.0 million").length, 1);
  assert.equal(factKeys.includes("monthly recurring revenue:62,824"), false);
  assert.equal(factKeys.includes("monthly recurring revenue:58,626"), false);
  assert.equal(result.analysis.extractedFacts.some((fact) => fact.field === "revenue" && /\$?3\.0/i.test(String(fact.value))), false);
  assert.equal(result.analysis.extractedFacts.some((fact) => fact.field === "revenue" && /21\.7/i.test(String(fact.value))), false);
  assert.equal(result.analysis.proposedChanges.length, 0);
  assert.doesNotMatch(JSON.stringify(result.analysis), /five times our revenue|200 times our revenue/);
  assert.equal(new Set(factKeys).size, factKeys.length);
});

test("successful manual text analysis returns normalized result", async () => {
  const { service, getCalls } = createService({
    investmentMatch: {
      investmentId: "inv-1",
      investmentName: "Northwind Ventures",
      confidence: 95,
      reason: "Subject and body mention Northwind."
    },
    entityMatch: {
      entityId: "Beaman Ventures",
      entityName: "Beaman Ventures",
      confidence: 90,
      reason: "Only matching position entity."
    },
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "$21.7M",
        period: "Q2 2026",
        factType: "actual historical result",
        sourceEvidence: "Q2 revenue was $21.7M",
        confidence: 97
      }
    ],
    whatChanged: ["Q2 revenue increased to $21.7M."],
    proposedChanges: [
      {
        action_type: "update operating metric",
        field: "revenue",
        current_value: "$18.2M",
        proposed_value: "$21.7M",
        confidence: 97,
        risk_level: "medium"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "Q2 revenue was $21.7M." },
    investments,
    entities
  });

  assert.equal(getCalls(), 1);
  assert.equal(result.analysis.investmentMatch.investmentId, "inv-1");
  assert.equal(result.analysis.extractedFacts[0].field, "revenue");
  assert.equal(result.analysis.proposedChanges[0].currentValue, "Not currently recorded");
  assert.equal(result.analyzedAt, "2026-08-27T12:00:00.000Z");
});

test("manual investment override is authoritative", async () => {
  const { service } = createService({
    investmentMatch: {
      investmentId: "inv-2",
      investmentName: "Southridge Fund II",
      confidence: 91,
      reason: "Model preferred Southridge."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "Northwind update" },
    investments,
    entities,
    investmentOverrideId: "inv-1"
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "inv-1");
  assert.equal(result.analysis.investmentMatch.confidence, 100);
  assert.match(result.analysis.warnings.join(" "), /Manual investment selection/);
});

test("low-confidence match is not treated as authoritative without override", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: {
        investmentId: "inv-1",
        investmentName: "Northwind Ventures",
        confidence: 62,
        reason: "Weak name overlap."
      },
      entityMatch: {},
      extractedFacts: [],
      whatChanged: [],
      proposedChanges: [],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.investmentMatch.investmentId, "");
  assert.match(result.warnings.join(" "), /below 70/);
});

test("invalid model-supplied investment ID is rejected", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: {
        investmentId: "not-real",
        investmentName: "Northwind Ventures",
        confidence: 98,
        reason: "Looks like Northwind."
      },
      entityMatch: {},
      extractedFacts: [],
      whatChanged: [],
      proposedChanges: [],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.investmentMatch.investmentId, "");
  assert.match(result.warnings.join(" "), /not in the current portfolio/);
});

test("entity matching uses the matched investment entity as strong evidence", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: {
        investmentId: "inv-1",
        investmentName: "Northwind Ventures",
        confidence: 94,
        reason: "Matched by name."
      },
      entityMatch: {
        entityId: "",
        entityName: "",
        confidence: 0,
        reason: ""
      },
      extractedFacts: [],
      whatChanged: [],
      proposedChanges: [],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.entityMatch.entityName, "Beaman Ventures");
  assert.equal(result.entityMatch.confidence, 90);
});

test("extraction output validation clamps confidence and normalizes risk", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: { investmentId: "inv-1", confidence: 99 },
      entityMatch: {},
      extractedFacts: [{ field: "cash balance", confidence: 140 }],
      whatChanged: [],
      proposedChanges: [{ field: "cash", confidence: -10, riskLevel: "surprising" }],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.extractedFacts[0].confidence, 100);
  assert.equal(result.proposedChanges[0].confidence, 0);
  assert.equal(result.proposedChanges[0].riskLevel, "medium");
});

test("comparison fills known current values from the matched investment", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: { investmentId: "inv-1", confidence: 99 },
      entityMatch: {},
      extractedFacts: [],
      whatChanged: ["Official value changed."],
      proposedChanges: [{ field: "officialValue", proposedValue: "21,700,000" }],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.proposedChanges[0].currentValue, "18,200,000");
});

test("malformed AI response fails safely", async () => {
  const { service } = createService("not json");

  await assert.rejects(
    service.analyzeInvestmentUpdate({
      source: { sourceText: "Update" },
      investments,
      entities
    }),
    /malformed JSON/
  );
});

test("analysis does not create a proposal or modify live investments", async () => {
  const before = JSON.stringify(investments);
  const { service } = createService({
    investmentMatch: { investmentId: "inv-1", confidence: 92 },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  await service.analyzeInvestmentUpdate({
    source: { sourceText: "Update" },
    investments,
    entities
  });

  assert.equal(JSON.stringify(investments), before);
});

test("capital call extraction remains a high-risk staged proposal", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: { investmentId: "inv-1", confidence: 95 },
      entityMatch: {},
      extractedFacts: [
        {
          category: "capital activity",
          field: "capital call amount",
          value: "$250,000",
          sourceEvidence: "Capital call due September 15",
          confidence: 93
        }
      ],
      whatChanged: ["A $250,000 capital call was detected."],
      proposedChanges: [
        {
          actionType: "capital call detected",
          field: "capitalCallAmount",
          currentValue: "",
          proposedValue: "$250,000",
          sourceEvidence: "Capital call due September 15",
          confidence: 93,
          riskLevel: "high"
        }
      ],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.proposedChanges[0].riskLevel, "high");
  assert.equal(result.proposedChanges[0].field, "capitalCallAmount");
});

test("distribution extraction remains a high-risk staged proposal", () => {
  const result = normalizeAnalysisResult({
    raw: {
      investmentMatch: { investmentId: "inv-1", confidence: 95 },
      entityMatch: {},
      extractedFacts: [
        {
          category: "capital activity",
          field: "distribution amount",
          value: "$75,000",
          sourceEvidence: "Distribution paid August 1",
          confidence: 93
        }
      ],
      whatChanged: ["A $75,000 distribution was detected."],
      proposedChanges: [
        {
          actionType: "distribution detected",
          field: "distributionAmount",
          currentValue: "",
          proposedValue: "$75,000",
          sourceEvidence: "Distribution paid August 1",
          confidence: 93,
          riskLevel: "high"
        }
      ],
      warnings: [],
      unresolved: []
    },
    investments,
    entities,
    normalizeEntityName
  });

  assert.equal(result.proposedChanges[0].riskLevel, "high");
  assert.equal(result.proposedChanges[0].field, "distributionAmount");
});

test("PDF page spans preserve source page on verified evidence", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 76 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "operating metrics",
        field: "revenue",
        value: "$3.0M",
        sourceEvidence: "July revenue was $3.0 million",
        confidence: 94
      }
    ],
    whatChanged: [],
    proposedChanges: [
      {
        actionType: "update operating metric",
        field: "revenue",
        currentValue: "",
        proposedValue: "$3.0M",
        sourceEvidence: "July revenue was $3.0 million",
        confidence: 94,
        riskLevel: "medium"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      filename: "FINSYNC July investor update.pdf",
      pageCount: 2,
      sourceText: "Page 1:\nFINSYNC investor update.\n\nPage 2:\nJuly revenue was $3.0 million.",
      pages: [
        { pageNumber: 1, text: "FINSYNC investor update." },
        { pageNumber: 2, text: "July revenue was $3.0 million." }
      ]
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts[0].evidenceStatus, "verified");
  assert.equal(result.analysis.extractedFacts[0].sourcePage, 2);
  assert.equal(result.analysis.proposedChanges[0].sourcePage, 2);
});

test("PDF filename can support matching but does not override explicit body evidence", async () => {
  const { service } = createService({
    investmentMatch: {
      investmentId: "vanguard-id",
      investmentName: "VANGUARD SHORT-TERM CORPORATE BOND ETF",
      confidence: 82,
      reason: "Filename mentioned Vanguard."
    },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      filename: "Vanguard update.pdf",
      sourceText: "FINSYNC July investor update reported revenue of $3.0 million.",
      pages: [{ pageNumber: 1, text: "FINSYNC July investor update reported revenue of $3.0 million." }]
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.notEqual(result.analysis.investmentMatch.investmentId, "vanguard-id");
});

test("PDF filename-only match stays below high-confidence body matches", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 82 },
    entityMatch: {},
    extractedFacts: [],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      filename: "FINSYNC board package.pdf",
      sourceText: "July revenue was $3.0 million.",
      pages: [{ pageNumber: 1, text: "July revenue was $3.0 million." }]
    },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(result.analysis.investmentMatch.investmentId, "finsync-id");
  assert.equal(result.analysis.investmentMatch.confidence, 78);
  assert.match(result.analysis.investmentMatch.reason, /filename/);
});

test("unsupported numeric facts cannot create proposed changes or What Changed claims", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "healing-id", confidence: 90 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "financial metric",
        field: "revenue",
        value: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 95
      }
    ],
    whatChanged: ["Healing Innovations revenue increased to 21.7M."],
    proposedChanges: [
      {
        actionType: "update financial metric",
        field: "revenue",
        proposedValue: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 95,
        riskLevel: "high"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceText: "Healing Innovations sold 5 units in Q2. No reliable revenue value was extracted."
    },
    investments: healingInvestments,
    entities
  });

  assert.equal(result.analysis.proposedChanges.length, 0);
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /21\.7M|revenue increased/i);
  assert.match(result.analysis.warnings.join(" "), /Removed unsupported proposed change for revenue/);
});

test("verified high-risk fact can support a proposed change and preserve page", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "healing-id", confidence: 90 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "recurring revenue",
        field: "subscription MRR",
        value: "$172,200",
        sourceEvidence: "subscription MRR of $172,200",
        confidence: 95
      }
    ],
    whatChanged: [],
    proposedChanges: [
      {
        actionType: "update recurring revenue",
        field: "subscription MRR",
        proposedValue: "$172,200",
        sourceEvidence: "",
        confidence: 92,
        riskLevel: "high"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      sourceText: "Page 6:\nRaaS subscription MRR of $172,200 across 41 units under contract.",
      pages: [{ pageNumber: 6, text: "RaaS subscription MRR of $172,200 across 41 units under contract." }]
    },
    investments: healingInvestments,
    entities
  });

  assert.equal(result.analysis.proposedChanges.length, 1);
  assert.equal(result.analysis.proposedChanges[0].evidenceStatus, "verified");
  assert.equal(result.analysis.proposedChanges[0].sourcePage, 6);
});

test("verified narrative material development is preserved without a structured field", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "healing-id", confidence: 90 },
    entityMatch: {},
    extractedFacts: [],
    materialDevelopments: [
      {
        category: "customer win",
        summary: "Healing Innovations added Nobis Rehabilitation Partners as a new corporate account.",
        sourceEvidence: "new corporate account: Nobis Rehabilitation Partners",
        sourcePage: 3,
        confidence: 95,
        riskLevel: "low"
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      sourceText: "Page 3:\nnew corporate account: Nobis Rehabilitation Partners",
      pages: [{ pageNumber: 3, text: "new corporate account: Nobis Rehabilitation Partners" }]
    },
    investments: healingInvestments,
    entities
  });

  assert.equal(result.analysis.proposedChanges.length, 0);
  assert.equal(result.analysis.materialDevelopments[0].sourcePage, 3);
  assert.match(result.analysis.whatChanged.join(" "), /Nobis Rehabilitation Partners/);
});

test("Healing Innovations deck regression rejects unsupported revenue and keeps verified developments", async () => {
  const pageResults = Array.from({ length: 5 }, () => ({
    extractedFacts: [],
    materialDevelopments: [],
    warnings: [],
    unresolved: []
  }));
  const finalResult = {
    investmentMatch: { investmentId: "healing-id", confidence: 90 },
    entityMatch: { entityName: "Beaman Ventures", confidence: 90 },
    extractedFacts: [
      {
        category: "sales",
        field: "units sold",
        value: "5",
        period: "Q2 2026",
        sourceEvidence: "Q2: 5 units sold",
        confidence: 94
      },
      {
        category: "recurring revenue",
        field: "subscription MRR",
        value: "$172,200",
        sourceEvidence: "subscription MRR of $172,200",
        confidence: 95
      },
      {
        category: "pipeline",
        field: "proposal pipeline",
        value: "$2.5M",
        period: "2026",
        sourceEvidence: "proposal $2.5M",
        confidence: 93
      },
      {
        category: "financial metric",
        field: "revenue",
        value: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 95
      }
    ],
    materialDevelopments: [
      {
        category: "operations",
        summary: "Healing Innovations sold 5 units in Q2 and installed 7 units.",
        sourceEvidence: "Q2: 5 units sold; 7 units installed",
        sourcePage: 3,
        confidence: 95,
        riskLevel: "medium"
      },
      {
        category: "customer win",
        summary: "Nobis Rehabilitation Partners was added as a new corporate account.",
        sourceEvidence: "new corporate account: Nobis Rehabilitation Partners",
        sourcePage: 3,
        confidence: 95,
        riskLevel: "low"
      },
      {
        category: "recurring revenue",
        summary: "RaaS subscription MRR reached $172,200 across 41 units under contract.",
        sourceEvidence: "subscription MRR of $172,200; 41 units under contract",
        sourcePage: 6,
        confidence: 95,
        riskLevel: "high"
      },
      {
        category: "pipeline",
        summary: "The 2026 sales pipeline includes $2.5M in proposal stage and $2.1M in contracting.",
        sourceEvidence: "proposal $2.5M; contracting $2.1M",
        sourcePage: 8,
        confidence: 94,
        riskLevel: "medium"
      },
      {
        category: "capital support commitment",
        summary: "Beaman Ventures committed to support an expansion of the line of credit to $5M.",
        sourceEvidence: "line of credit expansion to $5M; Beaman Ventures committed to support the expansion",
        sourcePage: 15,
        confidence: 96,
        riskLevel: "high"
      }
    ],
    whatChanged: ["Revenue increased to 21.7M."],
    proposedChanges: [
      {
        actionType: "update financial metric",
        field: "revenue",
        proposedValue: "21.7M",
        sourceEvidence: "revenue was 21.7M",
        confidence: 95,
        riskLevel: "high"
      }
    ],
    warnings: [],
    unresolved: []
  };
  const { service, getCalls, prompts } = createSequenceService(pageResults.concat(finalResult));
  const pages = [
    { pageNumber: 3, text: "Healing Innovations Q2: 5 units sold; 7 units installed; 6 trainings; new corporate account: Nobis Rehabilitation Partners." },
    { pageNumber: 6, text: "RaaS subscription MRR of $172,200; 41 units under contract." },
    { pageNumber: 8, text: "2026 pipeline: discovery $2.3M; demo $215K; proposal $2.5M; contracting $2.1M; closed won $1.2M." },
    { pageNumber: 9, text: "Q2 KPIs: 57 total units sold; 53 total units deployed; 51 total units trained; 21 FTE." },
    { pageNumber: 15, text: "Line of credit expansion to $5M. Beaman Ventures committed to support the expansion." }
  ];

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      filename: "Healing Innovations Board Deck July 29 2026.pdf",
      pageCount: 16,
      sourceText: pages.map((page) => `Page ${page.pageNumber}:\n${page.text}`).join("\n\n"),
      pages
    },
    investments: healingInvestments,
    entities
  });

  assert.equal(getCalls(), 6);
  assert.match(prompts[0], /Page 3:/);
  assert.match(prompts[5], /Page-level candidate extraction results/);
  assert.equal(result.analysis.investmentMatch.investmentId, "healing-id");
  assert.equal(result.analysis.proposedChanges.length, 0);
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /21\.7M|Revenue increased/i);
  assert.match(result.analysis.whatChanged.join(" "), /RaaS subscription MRR reached \$172,200/);
  assert.match(result.analysis.whatChanged.join(" "), /line of credit to \$5M/);
  assert.equal(result.analysis.materialDevelopments.find((item) => /line of credit/i.test(item.summary)).sourcePage, 15);
  assert.ok(result.analysis.materialDevelopments.every((item) => item.sourcePage));
});

test("numeric verification rejects 21.7M revenue when cited page lacks that number", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "healing-id", confidence: 97 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "financial metric",
        field: "revenue",
        value: "21.7M",
        sourceEvidence: "Q2 2026 Review 5 Units Sold 7 Units Installed 6 Trainings",
        sourcePage: 3,
        confidence: 97
      }
    ],
    whatChanged: ["Revenue increased to 21.7M."],
    proposedChanges: [
      {
        actionType: "update financial metric",
        field: "revenue",
        proposedValue: "21.7M",
        sourceEvidence: "Q2 2026 Review 5 Units Sold 7 Units Installed 6 Trainings",
        sourcePage: 3,
        confidence: 99,
        riskLevel: "high"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      sourceText: "Page 3:\nQ2 2026 Review\n5 Units Sold\n7 Units Installed\n6 Trainings\nNEW Corporate account: Nobis Rehabilitation Partners",
      pages: [
        {
          pageNumber: 3,
          text: "Q2 2026 Review\n5 Units Sold\n7 Units Installed\n6 Trainings\nNEW Corporate account: Nobis Rehabilitation Partners"
        }
      ]
    },
    investments: healingInvestments,
    entities
  });

  assert.equal(result.analysis.extractedFacts.length, 0);
  assert.equal(result.analysis.unverifiedClaims[0].field, "revenue");
  assert.equal(result.analysis.unverifiedClaims[0].verification.numericMatchFound, false);
  assert.equal(result.analysis.proposedChanges.length, 0);
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /21\.7M|Revenue increased/i);
  assert.match(result.analysis.unresolved.join(" "), /Could not verify numeric value 21\.7M for revenue on source page 3/);
});

test("revenue 3.0M verifies from revenue context but not cash context", async () => {
  const verified = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 90 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "financial metric",
        field: "revenue",
        value: "3.0M",
        sourceEvidence: "Revenue was $3.0 million",
        confidence: 95
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });
  const verifiedResult = await verified.service.analyzeInvestmentUpdate({
    source: { sourceText: "Revenue was $3.0 million." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(verifiedResult.analysis.extractedFacts[0].evidenceStatus, "verified");
  assert.equal(verifiedResult.analysis.extractedFacts[0].verification.numericMatchFound, true);
  assert.equal(verifiedResult.analysis.extractedFacts[0].verification.contextMatchFound, true);

  const rejected = createService({
    investmentMatch: { investmentId: "finsync-id", confidence: 90 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "financial metric",
        field: "revenue",
        value: "3.0M",
        sourceEvidence: "Cash balance was $3.0 million",
        confidence: 95
      }
    ],
    whatChanged: ["Revenue increased to 3.0M."],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });
  const rejectedResult = await rejected.service.analyzeInvestmentUpdate({
    source: { sourceText: "Cash balance was $3.0 million." },
    investments: finsyncInvestments,
    entities
  });

  assert.equal(rejectedResult.analysis.extractedFacts.length, 0);
  assert.equal(rejectedResult.analysis.unverifiedClaims[0].verification.numericMatchFound, true);
  assert.equal(rejectedResult.analysis.unverifiedClaims[0].verification.contextMatchFound, false);
  assert.doesNotMatch(rejectedResult.analysis.whatChanged.join(" "), /Revenue increased/);
});

test("Healing numeric positives verify MRR units installed sold and trainings", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "healing-id", confidence: 97 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "recurring revenue",
        field: "MRR",
        value: "172,200",
        sourceEvidence: "Subscription Program TOTAL $172,200",
        sourcePage: 6,
        confidence: 96
      },
      {
        category: "operations",
        field: "units sold",
        value: "5",
        sourceEvidence: "5 Units Sold",
        sourcePage: 3,
        confidence: 96
      },
      {
        category: "operations",
        field: "units installed",
        value: "7",
        sourceEvidence: "7 Units Installed",
        sourcePage: 3,
        confidence: 96
      },
      {
        category: "operations",
        field: "trainings",
        value: "6",
        sourceEvidence: "6 Trainings",
        sourcePage: 3,
        confidence: 96
      }
    ],
    whatChanged: [],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      sourceText: [
        "Page 3:\nQ2 2026 Review\n5 Units Sold\n7 Units Installed\n6 Trainings",
        "Page 6:\nSubscription Program\nTOTAL $172,200\n41 Units under contract for the RaaS Program"
      ].join("\n\n"),
      pages: [
        { pageNumber: 3, text: "Q2 2026 Review\n5 Units Sold\n7 Units Installed\n6 Trainings" },
        { pageNumber: 6, text: "Subscription Program\nTOTAL $172,200\n41 Units under contract for the RaaS Program" }
      ]
    },
    investments: healingInvestments,
    entities
  });

  assert.deepEqual(
    result.analysis.extractedFacts.map((fact) => fact.evidenceStatus),
    ["verified", "verified", "verified", "verified"]
  );
  assert.deepEqual(
    result.analysis.extractedFacts.map((fact) => fact.verification.numericMatchFound),
    [true, true, true, true]
  );
});

test("final response uses portfolio currentValue instead of model supplied currentValue", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "inv-1", confidence: 96 },
    entityMatch: {},
    extractedFacts: [
      {
        category: "valuation",
        field: "valuation",
        value: "$21.7M",
        sourceEvidence: "Valuation was $21.7M",
        confidence: 96
      }
    ],
    materialDevelopments: [
      {
        category: "valuation",
        summary: "Valuation was updated to $21.7M.",
        sourceEvidence: "Valuation was $21.7M",
        confidence: 96,
        riskLevel: "high"
      }
    ],
    whatChanged: ["Valuation increased from $1.18M to $21.7M."],
    proposedChanges: [
      {
        field: "valuation",
        currentValue: "$1.18M",
        proposedValue: "$21.7M",
        sourceEvidence: "Valuation was $21.7M",
        confidence: 96,
        riskLevel: "high"
      }
    ],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: { sourceText: "Northwind valuation was $21.7M." },
    investments,
    entities
  });

  assert.equal(result.analysis.proposedChanges[0].currentValue, "18,200,000");
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /1\.18M/);
  assert.match(result.analysis.whatChanged.join(" "), /18.2M.*21.7M/);
});

test("structured warning object normalizes to message instead of object serialization", () => {
  assert.equal(_test.warningMessage({ message: "Use this warning." }), "Use this warning.");
  assert.equal(_test.warningMessage({ reason: "Use this reason." }), "Use this reason.");
  assert.equal(_test.warningMessage({ unused: "raw object" }), "");
});

test("material evidence locator returns exact short evidence for units sold", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "operations",
      summary: "Healing Innovations sold 5 units in Q2, including 3 RaaS and 2 DP.",
      sourceEvidence: "Page 3 broad block",
      sourcePage: 3
    },
    sourceText: "",
    pages: [
      {
        pageNumber: 3,
        text: "Q2 2026 Review\nNEW Corporate account: Nobis Rehabilitation Partners\n5 Units Sold (3 RaaS + 2 DP)\n7 Units Installed + 6 Trainings"
      }
    ],
    sourcePage: 3
  });

  assert.equal(located.evidenceStatus, "verified");
  assert.equal(located.sourcePage, 3);
  assert.equal(located.evidence, "5 Units Sold (3 RaaS + 2 DP)");
});

test("material evidence locator returns exact short evidence for installed units and trainings", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "operations",
      summary: "Healing Innovations installed 7 units and completed 6 trainings.",
      sourcePage: 3
    },
    sourceText: "",
    pages: [
      {
        pageNumber: 3,
        text: "Q2 2026 Review\n5 Units Sold (3 RaaS + 2 DP)\n7 Units Installed + 6 Trainings\nNEW Corporate account: Nobis Rehabilitation Partners"
      }
    ],
    sourcePage: 3
  });

  assert.equal(located.evidenceStatus, "verified");
  assert.equal(located.sourcePage, 3);
  assert.equal(located.evidence, "7 Units Installed + 6 Trainings");
});

test("numeric material development cannot verify against unrelated same-page evidence", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "clinical",
      summary: "2 papers in the publishing process",
      sourcePage: 3
    },
    sourceText: "",
    pages: [
      {
        pageNumber: 3,
        text: "Q2 2026 Review\n5 Units Sold (3 RaaS + 2 DP)\n7 Units Installed + 6 Trainings"
      }
    ],
    sourcePage: 3
  });

  assert.equal(located.evidenceStatus, "unresolved");
});

test("numeric material development verifies when evidence supports the publishing claim", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "clinical",
      summary: "2 papers in the publishing process",
      sourcePage: 3
    },
    sourceText: "",
    pages: [
      {
        pageNumber: 3,
        text: "Clinical progress: 2 papers in the publishing process\n5 Units Sold (3 RaaS + 2 DP)"
      }
    ],
    sourcePage: 3
  });

  assert.equal(located.evidenceStatus, "verified");
  assert.equal(located.sourcePage, 3);
  assert.equal(located.evidence, "Clinical progress: 2 papers in the publishing process");
});

test("broad page text cannot independently verify a material development", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "customer win",
      summary: "Nobis Rehabilitation Partners became a new corporate account.",
      sourcePage: 3
    },
    sourceText: "",
    pages: [
      {
        pageNumber: 3,
        text: "Q2 2026 Review " +
          "Total Units Sold: 57 Sales Meetings: 89 Total Units Deployed: 53 Total Units Trained: 51 " +
          "Service Calls: 45 Operations FTE: 21 Part-Time: 4 Key Contractors: 6 People " +
          "The deck mentions Nobis Rehabilitation Partners in one section and corporate planning in another section with account operations elsewhere."
      }
    ],
    sourcePage: 3
  });

  assert.equal(located.evidenceStatus, "unresolved");
});

test("numeric material-development mismatch fails claim-level verification", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "operations",
      summary: "Healing Innovations sold 5 units in Q2.",
      sourcePage: 3
    },
    sourceText: "",
    pages: [{ pageNumber: 3, text: "Q2 2026 Review\n4 Units Sold (3 RaaS + 1 DP)" }],
    sourcePage: 3
  });

  assert.equal(located.evidenceStatus, "unresolved");
});

test("$5M LOC material development requires local LOC context", () => {
  const located = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "financing",
      summary: "The company expanded its line of credit to $5M.",
      sourcePage: 15
    },
    sourceText: "",
    pages: [{ pageNumber: 15, text: "The company expects $5M of future capacity. The operations team discussed credit discipline." }],
    sourcePage: 15
  });

  assert.equal(located.evidenceStatus, "unresolved");
});

test("commitment material development requires local commitment or support language", () => {
  const failed = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "capital support commitment",
      summary: "Beaman Ventures committed to support an expansion of the line of credit to $5M.",
      sourcePage: 15
    },
    sourceText: "",
    pages: [{ pageNumber: 15, text: "Line of credit expansion to $5M. Beaman Ventures was listed on the page." }],
    sourcePage: 15
  });
  const verified = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "capital support commitment",
      summary: "Beaman Ventures committed to support an expansion of the line of credit to $5M.",
      sourcePage: 15
    },
    sourceText: "",
    pages: [{ pageNumber: 15, text: "Line of credit expansion to $5M. Beaman Ventures committed to support the expansion." }],
    sourcePage: 15
  });

  assert.equal(failed.evidenceStatus, "unresolved");
  assert.equal(verified.evidenceStatus, "verified");
  assert.equal(verified.evidence, "Line of credit expansion to $5M. Beaman Ventures committed to support the expansion.");
});

test("personnel material developments verify against their respective personnel evidence", () => {
  const pages = [
    {
      pageNumber: 8,
      text: "HR Update\nWade Lawrence joined as VP of Operations.\nEmily Wilson hired as Clinical Lead."
    }
  ];
  const wade = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "staffing",
      summary: "Wade Lawrence joined as VP of Operations.",
      sourcePage: 8
    },
    sourceText: "",
    pages,
    sourcePage: 8
  });
  const emily = _test.findMaterialDevelopmentEvidence({
    item: {
      category: "staffing",
      summary: "Emily Wilson hired as Clinical Lead.",
      sourcePage: 8
    },
    sourceText: "",
    pages,
    sourcePage: 8
  });

  assert.equal(wade.evidenceStatus, "verified");
  assert.equal(wade.evidence, "Wade Lawrence joined as VP of Operations.");
  assert.equal(emily.evidenceStatus, "verified");
  assert.equal(emily.evidence, "Emily Wilson hired as Clinical Lead.");
});

test("failed material development is absent from What Changed and not actionable from page-wide overlap", async () => {
  const { service } = createService({
    investmentMatch: { investmentId: "healing-id", confidence: 96 },
    entityMatch: { entityName: "Beaman Ventures", confidence: 100 },
    extractedFacts: [],
    materialDevelopments: [
      {
        category: "capital support commitment",
        summary: "Beaman Ventures committed to support an expansion of the line of credit to $5M.",
        sourceEvidence: "Line of credit expansion to $5M. Beaman Ventures was listed on the page.",
        sourcePage: 15,
        confidence: 96,
        riskLevel: "high"
      }
    ],
    whatChanged: ["Beaman Ventures committed to support an expansion of the line of credit to $5M."],
    proposedChanges: [],
    warnings: [],
    unresolved: []
  });

  const result = await service.analyzeInvestmentUpdate({
    source: {
      sourceType: "PDF",
      sourceText: "Page 15:\nLine of credit expansion to $5M. Beaman Ventures was listed on the page.",
      pages: [{ pageNumber: 15, text: "Line of credit expansion to $5M. Beaman Ventures was listed on the page." }]
    },
    investments: healingInvestments,
    entities
  });

  assert.equal(result.analysis.materialDevelopments.length, 0);
  assert.equal(result.analysis.unverifiedClaims.some((claim) => claim.sourcePage === 15 && /line of credit/i.test(claim.value)), true);
  assert.doesNotMatch(result.analysis.whatChanged.join(" "), /\$5M|line of credit|committed/i);
  assert.match(result.analysis.warnings.join(" "), /Removed unsupported material development/);
});
