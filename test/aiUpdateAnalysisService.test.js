const test = require("node:test");
const assert = require("node:assert/strict");
const { createAiUpdateAnalysisService, normalizeAnalysisResult } = require("../services/aiUpdateAnalysisService");

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
  assert.equal(result.analysis.proposedChanges[0].currentValue, "$18.2M");
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
