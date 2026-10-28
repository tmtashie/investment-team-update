const test = require("node:test");
const assert = require("node:assert/strict");
const { Readable } = require("node:stream");
const { parseRequestBody } = require("../server")._test;

test("request body parser preserves UTF-8 split across stream chunks", async () => {
  const expected = { note: "Quarterly update \u{1F4C8}" };
  const payload = Buffer.from(JSON.stringify(expected));
  const character = Buffer.from("\u{1F4C8}");
  const characterStart = payload.indexOf(character);
  assert.notEqual(characterStart, -1);

  const request = Readable.from([
    payload.subarray(0, characterStart + 2),
    payload.subarray(characterStart + 2)
  ]);

  assert.deepEqual(await parseRequestBody(request), expected);
});
