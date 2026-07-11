const fs = require("node:fs");
const path = require("node:path");
const assert = require("node:assert/strict");

const root = path.resolve(__dirname, "..");
const manifestPath = path.join(root, "manifest.json");
const manifest = JSON.parse(fs.readFileSync(manifestPath, "utf8"));

function exists(relativePath) {
  assert.ok(fs.existsSync(path.join(root, relativePath)), `${relativePath} is missing`);
}

assert.equal(manifest.manifest_version, 3, "manifest_version must be 3");
assert.deepEqual(manifest.permissions, ["storage"], "extension should only request storage permission");
assert.ok(!manifest.host_permissions, "host_permissions are not needed because content script matches are explicit");
assert.equal(manifest.action.default_popup, "src/popup/popup.html");
assert.deepEqual(manifest.icons, {
  16: "icons/icon-16.png",
  32: "icons/icon-32.png",
  48: "icons/icon-48.png",
  128: "icons/icon-128.png"
});
assert.deepEqual(manifest.action.default_icon, manifest.icons);

const matches = manifest.content_scripts[0].matches;
assert.ok(matches.includes("https://teams.live.com/v2/*"), "teams.live.com/v2/* match is missing");
assert.ok(matches.includes("https://teams.microsoft.com/*"), "teams.microsoft.com/* match is missing");
assert.ok(matches.includes("https://teams.cloud.microsoft/*"), "teams.cloud.microsoft/* match is missing");

[
  "src/shared/config.js",
  "src/import-export/export-data.js",
  "src/import-export/migrate-import.js",
  "src/import-export/validate-import.js",
  "src/import-export/import-data.js",
  "src/import-export/import-backup.js",
  "src/content/main-bridge.js",
  "src/content/content-script.js",
  "src/content/toolbar.css",
  "src/popup/popup.html",
  "src/popup/popup.css",
  "src/popup/popup.js",
  "icons/icon-16.png",
  "icons/icon-32.png",
  "icons/icon-48.png",
  "icons/icon-128.png",
  "README.md"
].forEach(exists);

for (const contentScript of manifest.content_scripts) {
  for (const script of contentScript.js || []) {
    exists(script);
  }
  for (const css of contentScript.css || []) {
    exists(css);
  }
}

for (const iconPath of Object.values(manifest.icons || {})) {
  exists(iconPath);
}
for (const iconPath of Object.values(manifest.action.default_icon || {})) {
  exists(iconPath);
}

console.log("Extension file and manifest validation passed.");
