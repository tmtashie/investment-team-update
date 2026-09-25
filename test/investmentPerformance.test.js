const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadPerformanceHelpers() {
  const noop = () => {};
  let element;
  const elementTarget = {
    addEventListener: noop,
    removeEventListener: noop,
    appendChild: noop,
    remove: noop,
    querySelector: () => element,
    querySelectorAll: () => [],
    closest: () => null,
    classList: {
      add: noop,
      remove: noop,
      toggle: noop,
      contains: () => false
    },
    dataset: {},
    style: {},
    files: [],
    options: [],
    value: "",
    checked: false,
    disabled: false,
    getAttribute: () => null,
    setAttribute: noop,
    scrollIntoView: noop,
    reset: noop
  };

  element = new Proxy(elementTarget, {
    get(target, property) {
      if (property === "elements") {
        return new Proxy({}, { get: () => element });
      }
      return target[property];
    }
  });

  const document = {
    getElementById: () => element,
    querySelector: () => element,
    querySelectorAll: () => [],
    createElement: () => element,
    body: element
  };
  const context = {
    console: { log: noop, warn: noop, error: noop },
    document,
    addEventListener: noop,
    removeEventListener: noop,
    localStorage: { getItem: () => null, setItem: noop, removeItem: noop },
    fetch: async () => ({
      ok: false,
      status: 401,
      json: async () => ({ message: "Not signed in." })
    }),
    setTimeout,
    clearTimeout,
    FormData: class {},
    FileReader: class {},
    Blob: class {},
    URL: { createObjectURL: () => "", revokeObjectURL: noop },
    navigator: {},
    crypto: require("node:crypto").webcrypto,
    Intl,
    Date,
    Math,
    JSON,
    Number,
    String,
    Array,
    Object,
    Map,
    Set,
    RegExp,
    Error,
    Promise,
    encodeURIComponent,
    decodeURIComponent
  };
  context.window = context;
  context.globalThis = context;

  vm.createContext(context);
  const appSource = fs.readFileSync(path.join(__dirname, "../public/app.js"), "utf8");
  vm.runInContext(appSource, context, { filename: "public/app.js" });

  return {
    buildPerformanceInputs: context.buildPerformanceInputs,
    dealClaimStatus: context.dealClaimStatus,
    dealClaimValue: context.dealClaimValue,
    getAiProposalTypeLabel: context.getAiProposalTypeLabel,
    renderDealFieldEvidence: context.renderDealFieldEvidence
  };
}

test("proposal labels distinguish existing updates, ambiguous review, and potential new deals", () => {
  const { getAiProposalTypeLabel } = loadPerformanceHelpers();
  assert.equal(getAiProposalTypeLabel({ proposalType: "investment-update" }), "Existing Investment Update");
  assert.equal(getAiProposalTypeLabel({ proposalType: "new-deal", matchResult: { status: "ambiguous" } }), "Ambiguous Review");
  assert.equal(getAiProposalTypeLabel({ proposalType: "new-deal", matchResult: { status: "no-match" } }), "Potential New Deal");
});

test("structured new-deal claims render readable values while retaining evidence", () => {
  const { dealClaimStatus, dealClaimValue, renderDealFieldEvidence } = loadPerformanceHelpers();
  const proposal = { dealData: {
    tractionRevenue: [
      { value: "$94.5MM historical commitments", sourceEvidence: "Initial close on $94.5MM.", evidenceStatus: "verified" }
    ],
    customersContractsDeployments: [
      { value: "Control investments", sourceEvidence: "The portfolio includes control investments.", evidenceStatus: "verified" }
    ],
    financingTerms: [
      { value: "Ten-year term", sourceEvidence: "The term is ten years.", evidenceStatus: "verified" },
      { value: "2% management fee", sourceEvidence: "Management fee is 2%.", evidenceStatus: "verified" }
    ]
  } };
  assert.equal(dealClaimValue(proposal, "tractionRevenue"), "$94.5MM historical commitments");
  assert.equal(dealClaimValue(proposal, "customersContractsDeployments"), "Control investments");
  assert.equal(dealClaimValue(proposal, "financingTerms"), "Ten-year term\n2% management fee");
  assert.equal(dealClaimStatus(proposal, "financingTerms"), "verified");
  const evidence = renderDealFieldEvidence(proposal, "financingTerms");
  assert.match(evidence, /The term is ten years/);
  assert.match(evidence, /Management fee is 2%/);
  assert.doesNotMatch(evidence, /\[object Object\]/);
  assert.doesNotMatch(dealClaimValue(proposal, "tractionRevenue"), /\[object Object\]/);
  assert.doesNotMatch(dealClaimValue(proposal, "customersContractsDeployments"), /\[object Object\]/);
});

test("pipeline contributions are excluded from performance inputs", () => {
  const { buildPerformanceInputs } = loadPerformanceHelpers();
  const result = buildPerformanceInputs([
    {
      id: "pipeline-update",
      status: "New Lead",
      createdAt: "2026-01-01",
      capitalActivity: [
        { type: "Investment Amount", amount: 100, date: "2026-01-01" },
        { type: "Capital Call", amount: 75, date: "2026-01-02" },
        { type: "Fee", amount: 5, date: "2026-01-03" }
      ]
    },
    {
      id: "pipeline-stage-update",
      status: "Approved",
      stage: "Under Review",
      createdAt: "2026-01-04",
      capitalActivity: [
        { type: "Capital Call", amount: 25, date: "2026-01-04" }
      ]
    }
  ]);

  assert.equal(result.investedCapital, 0);
  assert.deepEqual(Array.from(result.baseCashFlows), []);
});

test("funded contributions remain included in performance inputs", () => {
  const { buildPerformanceInputs } = loadPerformanceHelpers();
  const result = buildPerformanceInputs([
    {
      id: "funded-update",
      status: "Funded",
      createdAt: "2026-01-01",
      capitalActivity: [
        { type: "Capital Call", amount: 75, date: "2026-01-02" },
        { type: "Fee", amount: 5, date: "2026-01-03" }
      ]
    }
  ]);

  assert.equal(result.investedCapital, 80);
  assert.deepEqual(
    Array.from(result.baseCashFlows, (cashFlow) => cashFlow.amount),
    [-75, -5]
  );
});

test("server performance snapshot excludes New Lead and Under Review activity and marks", () => {
  const { calculateCompanyPerformanceSnapshot } = require("../server")._test;
  const result = calculateCompanyPerformanceSnapshot({
    capitalActivities: [
      { type: "Capital Call", amount: "500000", date: "2026-01-01", sourceStatus: "New Lead" },
      { type: "Fee", amount: "10000", date: "2026-01-02", sourceStage: "Under Review" }
    ],
    valuationHistory: [
      { officialValue: "1000000", internalValue: "1200000", date: "2026-02-01", sourceStatus: "New Lead" }
    ]
  });
  assert.equal(result.totalInvestedCapital, 0);
  assert.equal(result.totalDistributions, 0);
  assert.equal(result.officialValue, 0);
  assert.equal(result.internalValue, 0);
  assert.equal(result.official.moic, null);
  assert.equal(result.official.xirr, null);
});

test("server performance snapshot retains funded activity", () => {
  const { calculateCompanyPerformanceSnapshot } = require("../server")._test;
  const result = calculateCompanyPerformanceSnapshot({
    capitalActivities: [{ type: "Capital Call", amount: "100", date: "2026-01-01", sourceStatus: "Funded" }],
    valuationHistory: [{ officialValue: "150", date: "2026-09-01", sourceStatus: "Funded" }]
  });
  assert.equal(result.totalInvestedCapital, 100);
  assert.equal(result.officialValue, 150);
  assert.equal(result.official.moic, 1.5);
});
