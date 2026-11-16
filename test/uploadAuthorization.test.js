const test = require("node:test");
const assert = require("node:assert/strict");
const { canViewStoredUpload, decodeBase64FileData } = require("../server")._test;

test("document uploads reject malformed Base64 before writing files", () => {
  const content = Buffer.from("quarterly update", "utf8");
  const padded = content.toString("base64");

  assert.deepEqual(decodeBase64FileData(padded), content);
  assert.deepEqual(decodeBase64FileData(padded.replace(/=+$/, "")), content);
  assert.equal(decodeBase64FileData("!!!!"), null);
  assert.equal(decodeBase64FileData(`${padded}invalid`), null);
  assert.equal(decodeBase64FileData("AA=A"), null);
  assert.equal(decodeBase64FileData("="), null);
});

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
