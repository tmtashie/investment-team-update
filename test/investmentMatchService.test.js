const test = require("node:test");
const assert = require("node:assert/strict");
const { generateInvestmentMatchCandidates } = require("../services/investmentMatchService");
const { hasAutomatedExplicitInvestmentMatch } = require("../services/aiEmailIntakeService");

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
    source: { subject: "Test CHRP", sender: "tyler@beamanventures.com", sourceText: "CHRP introduction and attached deck." },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }],
    houseDomains: "beamanventures.com"
  });

  assert.equal(result.status, "no-match");
  assert.equal(result.candidates.length, 0);
});

test("legitimate company domain still supports Company Ventures", () => {
  const result = generateInvestmentMatchCandidates({
    source: { subject: "Investor update", sender: "founder@companyventures.com", sourceText: "Quarterly investor update." },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }],
    houseDomains: "beamanventures.com"
  });

  assert.equal(result.status, "existing-possible");
  assert.equal(result.best.investmentId, "company-ventures");
  assert.equal(result.best.hasDomainEvidence, true);
});

test("explicit Company Ventures evidence still matches from a house-domain sender", () => {
  const result = generateInvestmentMatchCandidates({
    source: {
      subject: "Company Ventures quarterly update",
      sender: "tyler@subdomain.beamanventures.com",
      sourceText: "Company Ventures reported its quarterly operating results."
    },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }],
    houseDomains: ["beamanventures.com"]
  });

  assert.equal(result.status, "existing-confident");
  assert.equal(result.best.hasExplicitNameEvidence, true);
  assert.equal(result.best.hasDomainEvidence, false);
});

test("house-domain-only false match cannot pass automated proposal eligibility", () => {
  const result = generateInvestmentMatchCandidates({
    source: { subject: "Test CHRP", sender: "tyler@beamanventures.com", sourceText: "CHRP introduction." },
    investments: [{ id: "company-ventures", company: "Company Ventures", entity: "Beaman Ventures" }],
    houseDomains: "beamanventures.com"
  });
  const analysis = {
    investmentMatch: result.best
      ? { investmentId: result.best.investmentId, reason: result.best.reason }
      : { investmentId: "", reason: "" },
    candidates: result.candidates
  };

  assert.equal(hasAutomatedExplicitInvestmentMatch(analysis), false);
});
