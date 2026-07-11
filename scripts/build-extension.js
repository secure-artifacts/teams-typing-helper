const fs = require("node:fs");
const path = require("node:path");
const { execFileSync } = require("node:child_process");

const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");

function copyRecursive(source, target) {
  const stat = fs.statSync(source);
  if (stat.isDirectory()) {
    fs.mkdirSync(target, { recursive: true });
    for (const item of fs.readdirSync(source)) {
      copyRecursive(path.join(source, item), path.join(target, item));
    }
    return;
  }
  fs.copyFileSync(source, target);
}

execFileSync(process.execPath, [path.join(root, "scripts", "validate-extension.js")], {
  cwd: root,
  stdio: "inherit"
});

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(dist, { recursive: true });

for (const entry of ["manifest.json", "src", "icons"]) {
  copyRecursive(path.join(root, entry), path.join(dist, entry));
}

console.log("Extension build output written to dist/.");
