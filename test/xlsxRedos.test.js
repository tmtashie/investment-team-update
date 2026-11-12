const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const { Worker } = require("node:worker_threads");

const PARSE_TIMEOUT_MS = 2000;
const REPEATED_COMMENT_COUNT = 150387;

function runHostileImport({ dataDir, serverPath }) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(
      `
        const { parentPort, workerData } = require("node:worker_threads");
        process.env.DATA_DIR = workerData.dataDir;

        try {
          const { importWorkbookIntoInvestments } =
            require(workerData.serverPath)._test;
          const hostileWorkbook = Buffer.from(
            '<?xml version="1.0"?><Workbook>' +
              '<!--'.repeat(workerData.repeatedCommentCount),
            "utf8"
          );

          try {
            importWorkbookIntoInvestments(
              hostileWorkbook,
              { email: "workbook-test@example.test" },
              "redos-regression.xlsx"
            );
            parentPort.postMessage({ completed: true, rejected: false });
          } catch (error) {
            parentPort.postMessage({
              completed: true,
              rejected: true,
              message: error && error.message
            });
          }
        } catch (error) {
          parentPort.postMessage({
            completed: false,
            rejected: true,
            message: error && error.stack
          });
        }
      `,
      {
        eval: true,
        workerData: {
          dataDir,
          repeatedCommentCount: REPEATED_COMMENT_COUNT,
          serverPath
        }
      }
    );

    let settled = false;
    const timeout = setTimeout(() => {
      settled = true;
      worker.terminate().finally(() => {
        reject(new Error(
          `SheetJS exceeded the ${PARSE_TIMEOUT_MS}ms hostile XML parse limit`
        ));
      });
    }, PARSE_TIMEOUT_MS);

    worker.once("message", (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      worker.terminate().finally(() => resolve(result));
    });
    worker.once("error", (error) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      reject(error);
    });
    worker.once("exit", (code) => {
      if (settled) return;
      settled = true;
      clearTimeout(timeout);
      reject(new Error(`hostile workbook worker exited with code ${code}`));
    });
  });
}

test("application import rejects hostile repeated XML comments within a bounded time", async (t) => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), "xlsx-redos-test-"));
  t.after(() => fs.rmSync(dataDir, { recursive: true, force: true }));

  const investmentsFile = path.join(dataDir, "investments.json");
  const existingInvestments = [{
    id: "existing-investment",
    entity: "Beaman Ventures",
    company: "Existing Financial Record",
    amount: "4000000.00",
    officialValue: "5250000.00",
    internalValue: "5500000.00",
    ownershipPercent: "18.25",
    entityOwnershipPercent: "11.75",
    documents: [],
    recipients: [],
    createdAt: "2026-10-01T12:00:00.000Z"
  }];
  fs.writeFileSync(
    investmentsFile,
    JSON.stringify(existingInvestments, null, 2) + "\n",
    "utf8"
  );
  const result = await runHostileImport({
    dataDir,
    serverPath: path.resolve(__dirname, "../server.js")
  });

  assert.equal(result.completed, true, result.message);
  assert.equal(result.rejected, true, "the malformed workbook must be rejected");
  const preserved = JSON.parse(fs.readFileSync(investmentsFile, "utf8"))
    .find((investment) => investment.id === "existing-investment");
  assert.equal(preserved.amount, existingInvestments[0].amount);
  assert.equal(preserved.officialValue, existingInvestments[0].officialValue);
  assert.equal(preserved.internalValue, existingInvestments[0].internalValue);
  assert.equal(preserved.ownershipPercent, existingInvestments[0].ownershipPercent);
  assert.equal(
    preserved.entityOwnershipPercent,
    existingInvestments[0].entityOwnershipPercent
  );
});
