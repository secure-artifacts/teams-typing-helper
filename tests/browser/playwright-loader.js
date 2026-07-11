const path = require("node:path");

function loadPlaywright() {
  const candidatePaths = [
    process.cwd(),
    path.resolve(__dirname, "../.."),
    path.resolve(__dirname, "../../../chat-notify-filter/node_modules")
  ];
  try {
    const resolved = require.resolve("playwright", { paths: candidatePaths });
    return require(resolved);
  } catch (error) {
    return null;
  }
}

module.exports = { loadPlaywright };
