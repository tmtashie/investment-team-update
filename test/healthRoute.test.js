const test = require("node:test");
const assert = require("node:assert/strict");
const http = require("node:http");
const { server } = require("../server");

function request(port, path) {
  return new Promise((resolve, reject) => {
    const clientRequest = http.get({ host: "127.0.0.1", port, path }, (response) => {
      let body = "";
      response.setEncoding("utf8");
      response.on("data", (chunk) => {
        body += chunk;
      });
      response.on("end", () => {
        resolve({ statusCode: response.statusCode, body });
      });
    });
    clientRequest.on("error", reject);
  });
}

test("health shortcut matches only the exact route", async (context) => {
  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(0, "127.0.0.1", resolve);
  });
  context.after(() => new Promise((resolve) => server.close(resolve)));

  const { port } = server.address();
  assert.deepEqual(await request(port, "/api/health"), {
    statusCode: 200,
    body: '{"ok":true}'
  });
  assert.deepEqual(await request(port, "/api/health?probe=1"), {
    statusCode: 200,
    body: '{"ok":true}'
  });
  assert.equal((await request(port, "/api/healthcheck")).statusCode, 404);
  assert.equal((await request(port, "/api/healthy")).statusCode, 404);
});
