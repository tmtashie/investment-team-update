const test = require("node:test");
const assert = require("node:assert/strict");
const { normalizeNextStepReminderDays } = require("../server")._test;

test("invalid next-step reminder configuration falls back to fourteen days", () => {
  for (const value of [
    undefined,
    null,
    "",
    "   ",
    "not-a-number",
    "Infinity",
    "-Infinity",
    "NaN",
    "1e308",
    "-1e308"
  ]) {
    assert.equal(normalizeNextStepReminderDays(value), 14);
  }
});

test("finite next-step reminder configuration preserves existing numeric behavior", () => {
  assert.equal(normalizeNextStepReminderDays("0"), 0);
  assert.equal(normalizeNextStepReminderDays("21"), 21);
  assert.equal(normalizeNextStepReminderDays("2.5"), 2.5);
  assert.equal(normalizeNextStepReminderDays("-3"), -3);
});
