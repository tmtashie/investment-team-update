const test = require("node:test");
const assert = require("node:assert/strict");
const { parseCookies } = require("../server")._test;

test("malformed cookie values are ignored without hiding valid cookies", () => {
  const cookies = parseCookies({
    headers: {
      cookie: "bad=%; investment_team_session=signed.value%3D; preference=compact"
    }
  });

  assert.deepEqual(cookies, {
    investment_team_session: "signed.value=",
    preference: "compact"
  });
});

test("a malformed session cookie fails closed", () => {
  const cookies = parseCookies({
    headers: {
      cookie: "investment_team_session=%"
    }
  });

  assert.equal(cookies.investment_team_session, undefined);
});
