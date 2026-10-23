/* Exercise the real development-server request handler without opening a socket. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const root = path.resolve(__dirname, "..");
let handler;
const reads = [];
const fakeHttps = { createServer(options, callback) { handler = callback; return { listen() {} }; } };
const fakeFs = { readFile(file, callback) { reads.push(file); callback(null, Buffer.from("ok")); } };
const context = { __dirname: root, console: { log() {}, error() {} }, process: { exit(code) { throw new Error("Exit " + code); } },
  require(name) {
    if (name === "https") return fakeHttps;
    if (name === "fs") return fakeFs;
    if (name === "path") return path;
    if (name === "office-addin-dev-certs") return { getHttpsServerOptions: async () => ({}) };
    throw new Error("Unexpected dependency " + name);
  } };
function request(url) {
  let status;
  handler({ url }, { writeHead(code) { status = code; }, end() {} });
  return status;
}
(async () => {
  vm.runInNewContext(fs.readFileSync(path.join(root, "server.js"), "utf8"), context);
  await new Promise(setImmediate);
  assert.equal(typeof handler, "function");
  assert.equal(request("/%E0%A4%A"), 400, "invalid percent encoding cannot crash the server");
  assert.equal(request("/%00"), 400, "reject null bytes");
  assert.equal(request("/../" + path.basename(root) + "-sibling/secrets.txt"), 403, "reject sibling paths sharing the root prefix");
  assert.equal(request("/%2e%2e/secrets.txt"), 403, "reject encoded traversal");
  assert.equal(reads.length, 0, "invalid requests never read files");
  assert.equal(request("/src/taskpane/taskpane.html?build=1"), 200);
  assert.equal(reads[0], path.join(root, "src/taskpane/taskpane.html"));
  assert.equal(request("/"), 200);
  assert.equal(reads[1], path.join(root, "index.html"));
  console.log("All development-server safety tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
