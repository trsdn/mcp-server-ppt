// Patches vscode-jsonrpc for Node 24 ESM compatibility
// Run via: npm run postinstall
const fs = require("fs");
const path = require("path");

const pkgPath = path.join(__dirname, "node_modules", "vscode-jsonrpc", "package.json");
if (fs.existsSync(pkgPath)) {
  const pkg = JSON.parse(fs.readFileSync(pkgPath, "utf-8"));
  pkg.exports = {
    ".": { import: "./lib/node/main.js", require: "./lib/node/main.js" },
    "./node": { import: "./lib/node/main.js", require: "./lib/node/main.js" },
    "./node.js": { import: "./lib/node/main.js", require: "./lib/node/main.js" },
  };
  fs.writeFileSync(pkgPath, JSON.stringify(pkg, null, 2));
  console.log("Patched vscode-jsonrpc exports for Node 24 ESM");
}
