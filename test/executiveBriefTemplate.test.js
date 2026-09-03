const test = require("node:test");
const assert = require("node:assert/strict");
const { buildExecutiveBriefTemplate, validateRestrictedHtml } = require("../services/executiveBriefTemplate");

test("email HTML and text contain required presidential letterhead and footer", () => {
  const rendered = buildExecutiveBriefTemplate({
    briefType: "morning",
    displayDate: "September 3, 2026",
    html: "<h2>Executive Priorities</h2><p>Priority one.</p>",
    text: "Executive Priorities\nPriority one."
  });
  for (const required of [
    "BEAMAN VENTURES",
    "OFFICE OF THE CHIEF OF STAFF",
    "MORNING EXECUTIVE BRIEF",
    "Prepared for Tyler Tashie",
    "September 3, 2026",
    "CONFIDENTIAL &mdash; BEAMAN VENTURES",
    "Prepared by the Office of the Chief of Staff"
  ]) {
    assert.match(rendered.html, new RegExp(required.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  }
  assert.match(rendered.text, /CONFIDENTIAL \u2014 BEAMAN VENTURES/);
  assert.match(rendered.html, /max-width:600px/);
});

test("restricted HTML accepts section formatting but rejects attributes, links, images, and unbalanced tags", () => {
  assert.equal(validateRestrictedHtml("<h2>Schedule</h2><ul><li><strong>9:00</strong> Review</li></ul>"), "<h2>Schedule</h2><ul><li><strong>9:00</strong> Review</li></ul>");
  for (const value of [
    "<p style=\"color:red\">bad</p>",
    "<a href=\"https://example.com\">bad</a>",
    "<img src=\"https://example.com/pixel\">",
    "<p>unbalanced"
  ]) {
    assert.throws(() => validateRestrictedHtml(value));
  }
});
