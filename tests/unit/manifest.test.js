const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.resolve(__dirname, "../..");
const manifest = JSON.parse(fs.readFileSync(path.join(root, "manifest.json"), "utf8"));

test("manifest is MV3 with minimal permissions", () => {
  assert.equal(manifest.manifest_version, 3);
  assert.deepEqual(manifest.permissions, ["storage"]);
  assert.equal(manifest.host_permissions, undefined);
});

test("manifest uses bundled extension icons", () => {
  assert.deepEqual(manifest.icons, {
    16: "icons/icon-16.png",
    32: "icons/icon-32.png",
    48: "icons/icon-48.png",
    128: "icons/icon-128.png"
  });
  assert.deepEqual(manifest.action.default_icon, manifest.icons);
});

test("manifest matches supported Teams web origins", () => {
  const matches = manifest.content_scripts[0].matches;
  assert.ok(matches.includes("https://teams.live.com/v2/*"));
  assert.ok(matches.includes("https://teams.microsoft.com/*"));
  assert.ok(matches.includes("https://teams.cloud.microsoft/*"));
});

test("manifest separates MAIN bridge and ISOLATED toolbar", () => {
  assert.equal(manifest.content_scripts[0].world, "MAIN");
  assert.deepEqual(manifest.content_scripts[0].js, ["src/content/main-bridge.js"]);
  assert.equal(manifest.content_scripts[1].world, "ISOLATED");
  assert.ok(manifest.content_scripts[1].js.includes("src/content/content-script.js"));
});

test("all manifest referenced files exist", () => {
  assert.ok(fs.existsSync(path.join(root, manifest.action.default_popup)));
  for (const iconPath of Object.values(manifest.icons || {})) {
    assert.ok(fs.existsSync(path.join(root, iconPath)), `${iconPath} missing`);
  }
  for (const iconPath of Object.values(manifest.action.default_icon || {})) {
    assert.ok(fs.existsSync(path.join(root, iconPath)), `${iconPath} missing`);
  }
  for (const contentScript of manifest.content_scripts) {
    for (const script of contentScript.js || []) {
      assert.ok(fs.existsSync(path.join(root, script)), `${script} missing`);
    }
    for (const css of contentScript.css || []) {
      assert.ok(fs.existsSync(path.join(root, css)), `${css} missing`);
    }
  }
});
