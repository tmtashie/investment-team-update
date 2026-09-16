const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function loadPerformanceHelpers() {
  const noop = () => {};
  let element;
  const elementTarget = {
    addEventListener: noop,
    removeEventListener: noop,
    appendChild: noop,
    remove: noop,
    querySelector: () => element,
    querySelectorAll: () => [],
    closest: () => null,
    classList: {
      add: noop,
      remove: noop,
      toggle: noop,
      contains: () => false
    },
    dataset: {},
    style: {},
    files: [],
    options: [],
    value: "",
    checked: false,
    disabled: false,
    getAttribute: () => null,
    setAttribute: noop,
    scrollIntoView: noop,
    reset: noop
  };

  element = new Proxy(elementTarget, {
    get(target, property) {
      if (property === "elements") {
        return new Proxy({}, { get: () => element });
      }
      return target[property];
    }
  });

  const document = {
    getElementById: () => element,
    querySelector: () => element,
    querySelectorAll: () => [],
    createElement: () => element,
    body: element
  };
  const context = {
    console: { log: noop, warn: noop, error: noop },
    document,
    addEventListener: noop,
    removeEventListener: noop,
    localStorage: { getItem: () => null, setItem: noop, removeItem: noop },
    fetch: async () => ({
      ok: false,
      status: 401,
      json: async () => ({ message: "Not signed in." })
    }),
    setTimeout,
    clearTimeout,
    FormData: class {},
    FileReader: class {},
    Blob: class {},
    URL: { createObjectURL: () => "", revokeObjectURL: noop },
    navigator: {},
    crypto: require("node:crypto").webcrypto,
    Intl,
    Date,
    Math,
    JSON,
    Number,
    String,
    Array,
    Object,
    Map,
    Set,
    RegExp,
    Error,
    Promise,
    encodeURIComponent,
    decodeURIComponent
  };
  context.window = context;
  context.globalThis = context;

  vm.createContext(context);
  const appSource = fs.readFileSync(path.join(__dirname, "../public/app.js"), "utf8");
  vm.runInContext(appSource, context, { filename: "public/app.js" });

  return {
    buildPerformanceInputs: context.buildPerformanceInputs
  };
}

test("pipeline contributions are excluded from performance inputs", () => {
  const { buildPerformanceInputs } = loadPerformanceHelpers();
  const result = buildPerformanceInputs([
    {
      id: "pipeline-update",
      status: "New Lead",
      createdAt: "2026-01-01",
      capitalActivity: [
        { type: "Investment Amount", amount: 100, date: "2026-01-01" },
        { type: "Capital Call", amount: 75, date: "2026-01-02" },
        { type: "Fee", amount: 5, date: "2026-01-03" }
      ]
    },
    {
      id: "pipeline-stage-update",
      status: "Approved",
      stage: "Under Review",
      createdAt: "2026-01-04",
      capitalActivity: [
        { type: "Capital Call", amount: 25, date: "2026-01-04" }
      ]
    }
  ]);

  assert.equal(result.investedCapital, 0);
  assert.deepEqual(Array.from(result.baseCashFlows), []);
});

test("funded contributions remain included in performance inputs", () => {
  const { buildPerformanceInputs } = loadPerformanceHelpers();
  const result = buildPerformanceInputs([
    {
      id: "funded-update",
      status: "Funded",
      createdAt: "2026-01-01",
      capitalActivity: [
        { type: "Capital Call", amount: 75, date: "2026-01-02" },
        { type: "Fee", amount: 5, date: "2026-01-03" }
      ]
    }
  ]);

  assert.equal(result.investedCapital, 80);
  assert.deepEqual(
    Array.from(result.baseCashFlows, (cashFlow) => cashFlow.amount),
    [-75, -5]
  );
});
