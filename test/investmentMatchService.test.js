const test = require("node:test");
const assert = require("node:assert/strict");
const { generateInvestmentMatchCandidates } = require("../services/investmentMatchService");

test("shared matcher returns a confident deterministic existing investment", () => {
  const result = generateInvestmentMatchCandidates({
    source: { subject: "Acme Robotics seed round", sender: "founder@acmerobotics.com", sourceText: "Acme Robotics investor update" },
    investments: [{ id: "acme", company: "Acme Robotics", entity: "Beaman Ventures", investmentAliases: ["Acme"] }]
  });
  assert.equal(result.status, "existing-confident");
  assert.equal(result.best.investmentId, "acme");
});

test("competing deterministic candidates are ambiguous", () => {
  const result = generateInvestmentMatchCandidates({
    source: { subject: "Atlas update", sourceText: "Atlas Holdings and Atlas Health are referenced in this opportunity." },
    investments: [
      { id: "atlas-holdings", company: "Atlas Holdings", entity: "Beaman Ventures" },
      { id: "atlas-health", company: "Atlas Health", entity: "Beaman Ventures" }
    ]
  });
  assert.equal(result.status, "ambiguous");
  assert.equal(result.hasCompetingCandidate, true);
  assert.equal(result.candidates.length, 2);
});

test("Beaman Ventures house domain does not support Company Ventures by domain alone", () => {
  const result = generateInvestmentMatchCandidates({
    source: { sender: "tyler@beamanventures.com" },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }]
  });

  assert.equal(result.status, "no-match");
  assert.equal(result.candidates.length, 0);
});

test("legitimate portfolio-company domain still supports Company Ventures", () => {
  const result = generateInvestmentMatchCandidates({
    source: { sender: "founder@companyventures.com" },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }]
  });

  assert.equal(result.status, "existing-possible");
  assert.equal(result.best.investmentId, "company-ventures");
  assert.equal(result.best.hasDomainEvidence, true);
});

test("explicit Company Ventures text still matches from the house domain", () => {
  const result = generateInvestmentMatchCandidates({
    source: {
      sender: "tyler@updates.beamanventures.com",
      sourceText: "Company Ventures reported a quarterly revenue and customer pipeline update."
    },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }]
  });

  assert.equal(result.status, "existing-confident");
  assert.equal(result.best.investmentId, "company-ventures");
  assert.equal(result.best.hasExplicitNameEvidence, true);
  assert.equal(result.best.hasDomainEvidence, false);
});
