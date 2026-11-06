const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const projectRoot = path.resolve(__dirname, "..");
const packageJson = JSON.parse(
  fs.readFileSync(path.join(projectRoot, "package.json"), "utf8")
);
const packageLock = JSON.parse(
  fs.readFileSync(path.join(projectRoot, "package-lock.json"), "utf8")
);

const expectedTarball = "https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz";
const expectedIntegrity =
  "sha512-oLDq3jw7AcLqKWH2AhCpVTZl8mf6X2YReP+Neh0SJUzV/BdZYjth94tG5toiMB1PPrYtxOCfaoUCkvtuH+3AJA==";

test("SheetJS remediation stays pinned to the reviewed 0.20.3 artifact", () => {
  assert.equal(packageJson.dependencies.xlsx, expectedTarball);
  assert.equal(packageLock.packages[""].dependencies.xlsx, expectedTarball);

  const locked = packageLock.packages["node_modules/xlsx"];
  assert.ok(locked, "expected a locked SheetJS package");
  assert.equal(locked.version, "0.20.3");
  assert.equal(locked.resolved, expectedTarball);
  assert.equal(locked.integrity, expectedIntegrity);
});

test("SheetJS remediation does not restore the vulnerable 0.18 dependency set", () => {
  const removedPackages = [
    "adler-32",
    "cfb",
    "codepage",
    "crc-32",
    "frac",
    "ssf",
    "wmf",
    "word"
  ];

  for (const dependency of removedPackages) {
    assert.equal(
      packageLock.packages[`node_modules/${dependency}`],
      undefined,
      `unexpected legacy SheetJS dependency: ${dependency}`
    );
  }
});
