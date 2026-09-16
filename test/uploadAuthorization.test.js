const test = require("node:test");
const assert = require("node:assert/strict");
const { canViewStoredUpload } = require("../server")._test;

test("unassociated uploads fail closed", () => {
  assert.equal(canViewStoredUpload({ role: "master-editor" }, "orphan.pdf"), false);
});

test("proposal attachment visibility follows proposal entity permissions", () => {
  const proposals = [{
    id: "proposal-1",
    proposalType: "new-deal",
    proposedEntity: "Lee Beaman IRA",
    documents: [{ storedName: "private.pdf" }]
  }];
  assert.equal(canViewStoredUpload({ role: "editor" }, "private.pdf", { proposals }), false);
  assert.equal(canViewStoredUpload({ role: "master-editor" }, "private.pdf", { proposals }), true);
});

test("company-document and proposal associations must both authorize", () => {
  const companyDocuments = [{ company: "PrivateCo", entity: "Lee Beaman IRA", storedName: "shared.pdf" }];
  const proposals = [{ proposedEntity: "Beaman Ventures", documents: [{ storedName: "shared.pdf" }] }];
  assert.equal(canViewStoredUpload({ role: "editor" }, "shared.pdf", { companyDocuments, proposals }), false);
});

test("legacy investment update documents retain investment visibility", () => {
  const investments = [{ company: "PortfolioCo", entity: "Beaman Ventures", documents: [{ storedName: "update.pdf" }] }];
  assert.equal(canViewStoredUpload({ role: "editor" }, "update.pdf", { investments }), true);
});
