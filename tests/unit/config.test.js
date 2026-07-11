const test = require("node:test");
const assert = require("node:assert/strict");
const configApi = require("../../src/shared/config.js");

test("normalizes missing config to defaults", () => {
  const config = configApi.normalizeConfig(null);
  assert.equal(config.enabled, true);
  assert.equal(config.insertMode, "cursor");
  assert.ok(config.phrases.length >= 3);
});

test("preserves user phrase order and clamps columns", () => {
  const config = configApi.normalizeConfig({
    enabled: false,
    insertMode: "append",
    columns: 99,
    compactMode: true,
    phrases: [
      { id: "a", name: "A", text: "one" },
      { id: "b", name: "B", text: "two" }
    ]
  });

  assert.equal(config.enabled, false);
  assert.equal(config.insertMode, "append");
  assert.equal(config.compactMode, true);
  assert.equal(config.columns, 10);
  assert.deepEqual(config.phrases.map((phrase) => phrase.id), ["a", "b"]);
});

test("imports exported payload and raw config payload", () => {
  const source = configApi.normalizeConfig({
    phrases: [{ name: "测试", text: "hello" }],
    insertMode: "replace"
  });
  const exported = configApi.buildExportPayload(source);
  const imported = configApi.parseImportedConfig(JSON.stringify(exported));
  const rawImported = configApi.parseImportedConfig(JSON.stringify(source));

  assert.equal(imported.insertMode, "replace");
  assert.equal(imported.phrases[0].name, "测试");
  assert.equal(rawImported.phrases[0].text, "hello");
});

test("rejects invalid json import", () => {
  assert.throws(() => configApi.parseImportedConfig("{bad json"), /JSON/);
});
