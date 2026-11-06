const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "investment-workbook-test-"));
process.env.DATA_DIR = dataDir;

const {
  buildInvestmentsWorkbookBuffer,
  importWorkbookIntoInvestments
} = require("../server")._test;

test("SheetJS security upgrade preserves application workbook export/import fields", (t) => {
  t.after(() => fs.rmSync(dataDir, { recursive: true, force: true }));
  const existingInvestment = {
    id: "existing-investment",
    entity: "Beaman Ventures",
    company: "Existing Financial Record",
    amount: "4000000.00",
    officialValue: "5250000.00",
    internalValue: "5500000.00",
    ownershipPercent: "18.25",
    entityOwnershipPercent: "11.75",
    notes: "Must survive workbook import unchanged",
    documents: [],
    recipients: [],
    createdAt: "2026-10-01T12:00:00.000Z"
  };
  fs.writeFileSync(
    path.join(dataDir, "investments.json"),
    JSON.stringify([existingInvestment], null, 2) + "\n"
  );

  const workbook = buildInvestmentsWorkbookBuffer([{
    entity: "Beaman Ventures",
    company: "Workbook Security Regression",
    amount: "1250000.50",
    currency: "USD",
    status: "Funded",
    owner: "Workbook test",
    officialValue: "1750000.25",
    internalValue: "1900000.75",
    ownershipPercent: "12.5",
    entityOwnershipPercent: "7.25",
    ownershipNotes: "Representative round-trip data",
    documents: [],
    recipients: [],
    createdAt: "2026-10-30T12:00:00.000Z"
  }]);

  const result = importWorkbookIntoInvestments(
    workbook,
    { email: "workbook-test@example.test" },
    "security-regression.xlsx"
  );
  const investments = JSON.parse(
    fs.readFileSync(path.join(dataDir, "investments.json"), "utf8")
  );
  const imported = investments.find(
    (investment) => investment.company === "Workbook Security Regression"
  );
  const preserved = investments.find(
    (investment) => investment.id === existingInvestment.id
  );

  assert.equal(require("xlsx").version, "0.20.3");
  assert.equal(result.importedCount, 1);
  assert.deepEqual(result.sheets, ["Investment Updates"]);
  assert.equal(investments.length, 2);
  assert.equal(imported.company, "Workbook Security Regression");
  assert.equal(imported.entity, "Beaman Ventures");
  assert.equal(imported.amount, "1250000.50");
  assert.equal(imported.officialValue, "1750000.25");
  assert.equal(imported.internalValue, "1900000.75");
  assert.equal(imported.ownershipPercent, "12.5");
  assert.equal(imported.entityOwnershipPercent, "7.25");
  assert.equal(preserved.amount, existingInvestment.amount);
  assert.equal(preserved.officialValue, existingInvestment.officialValue);
  assert.equal(preserved.internalValue, existingInvestment.internalValue);
  assert.equal(preserved.ownershipPercent, existingInvestment.ownershipPercent);
  assert.equal(
    preserved.entityOwnershipPercent,
    existingInvestment.entityOwnershipPercent
  );

  const stateAfterImport = fs.readFileSync(
    path.join(dataDir, "investments.json"),
    "utf8"
  );
  assert.throws(
    () => importWorkbookIntoInvestments(
      workbook,
      { email: "workbook-test@example.test" },
      "security-regression.xlsx"
    ),
    /No importable rows were found/
  );
  assert.equal(
    fs.readFileSync(path.join(dataDir, "investments.json"), "utf8"),
    stateAfterImport,
    "replaying the same workbook must not mutate financial records"
  );
});
