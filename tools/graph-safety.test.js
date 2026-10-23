/* Offline regression tests against the real Graph data layer. No account/network needed. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const graphPath = path.join(__dirname, "../src/graph.js");
const source = fs.readFileSync(fs.existsSync(graphPath) ? graphPath :
  path.join(__dirname, "../src/commands/commands.js"), "utf8");
const names = ["fetchRetry", "graphJson", "graphAll", "listItems", "listReplies", "copyFileAttachment"]
  .filter(name => new RegExp("async function " + name + "\\(").test(source));
const exportsCode = "self.transport = {" + names.join(",") + "};\n";
const instrumented = source.includes("  root.GraphData = {") ?
  source.replace("  root.GraphData = {", exportsCode + "  root.GraphData = {") :
  source.includes("  var GraphData = {") ?
    source.replace("  var GraphData = {", exportsCode + "  var GraphData = {") : source + "\n" + exportsCode;
function response(status, body, retryAfter) {
  return { status, ok: status >= 200 && status < 300,
    headers: { get: () => retryAfter == null ? null : retryAfter },
    json: async () => body || {}, text: async () => JSON.stringify(body || {}) };
}
function load(fetcher) {
  const calls = [], waits = [];
  const context = { self: {}, Date, Uint8Array, TextEncoder,
    atob: text => Buffer.from(text, "base64").toString("binary"),
    Office: { onReady() {}, actions: { associate() {} },
      context: { roamingSettings: { get() { return false; } } } },
    fetch: async (url, opts) => { calls.push({ url, opts }); return fetcher(url, opts, calls.length); },
    setTimeout: (fn, ms) => { waits.push(ms); fn(); return 1; }, clearTimeout() {} };
  context.window = context.self;
  vm.runInNewContext(instrumented, context);
  return { api: context.self.GraphData, transport: context.self.transport, calls, waits };
}
(async () => {
  // Exercise retry behavior with fake timers, including date/invalid headers.
  for (const status of [429, 503]) {
    for (const header of ["2", null, "invalid", "0", "60", new Date(Date.now() + 60000).toUTCString()]) {
      const g = load((url, opts, n) => response(n === 1 ? status : 200, {}, header));
      const long = header === "60" || (header && header.includes("GMT"));
      if (g.transport.fetchRetry) {
        const res = await g.transport.fetchRetry("https://graph.microsoft.com/v1.0/me", { method: "GET" });
        assert.equal(res.status, long ? status : 200);
      } else if (long) {
        await assert.rejects(g.transport.graphJson("token", "GET", "/me"), /Retry in/);
      } else {
        await g.transport.graphJson("token", "GET", "/me");
      }
      assert.equal(g.calls.length, long ? 1 : 2, "never retry early for a long throttle");
      assert.deepEqual(g.waits, long ? [] : [header === "2" ? 2000 : header === "0" ? 0 : 3000]);
    }
  }
  const plain = load(() => response(400));
  if (plain.transport.fetchRetry) { await plain.transport.fetchRetry("url", {}); }
  else { await assert.rejects(plain.transport.graphJson("token", "GET", "/me")); }
  assert.equal(plain.calls.length, 1, "do not retry unrelated errors");
  const twice = load(() => response(429, {}, "0"));
  if (twice.transport.fetchRetry) { assert.equal((await twice.transport.fetchRetry("url", {})).status, 429); }
  else { await assert.rejects(twice.transport.graphJson("token", "GET", "/me")); }
  assert.equal(twice.calls.length, 2, "retry stays bounded");

  if (names.includes("graphAll")) {
    const g = load((url, opts, n) => response(200, n === 1 ?
      { value: [{ id: "first" }], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me?opaque=abc%2B123" } :
      { value: [{ id: "second" }] }));
    const items = await g.transport.graphAll("token", "/me");
    assert.equal(items.length, 2);
    assert.equal(g.calls[1].url, "https://graph.microsoft.com/v1.0/me?opaque=abc%2B123");
    const foreign = load(() => response(200, {value: [], "@odata.nextLink": "https://untrusted.example/next"}));
    await assert.rejects(foreign.transport.graphAll("token", "/me"), /paging URL/);
    assert.equal(foreign.calls.length, 1, "never send a token to an unrelated paging host");
    const capped = load(() => response(200, { value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/me?next=1" }));
    await assert.rejects(capped.transport.graphAll("token", "/me"), /scan limit/);
    assert.ok(capped.calls.length <= 100);
  }
  for (const name of ["listItems", "listReplies"].filter(name => names.includes(name))) {
    const g = load(() => response(200, { value: [], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next" }));
    await assert.rejects(g.transport[name]("token", "site", "list", "message"), /scan limit/);
  }
  if (names.includes("copyFileAttachment")) {
    const small = load(() => response(201, { id: "attachment" }));
    await small.transport.copyFileAttachment("token", "draft", { name: "small.txt" }, "YWJj");
    assert.equal(small.calls.length, 1);
    assert.equal(JSON.parse(small.calls[0].opts.body).contentBytes, "YWJj");
    const bytes = Buffer.alloc(4000000, 97);
    const g = load((url, opts, n) => n === 1 ? response(201, { uploadUrl: "https://outlook.office.com/upload?auth=opaque" }) :
      response(n === 3 ? 201 : 200));
    await g.transport.copyFileAttachment("token", "draft", { name: "large.pdf" }, bytes.toString("base64"));
    assert.equal(g.calls.length, 3);
    assert.equal(JSON.parse(g.calls[0].opts.body).AttachmentItem.size, bytes.length);
    assert.equal(g.calls[1].opts.headers["Content-Range"], "bytes 0-3276799/4000000");
    assert.equal(g.calls[2].opts.headers["Content-Range"], "bytes 3276800-3999999/4000000");
    assert.equal(g.calls[1].opts.headers.Authorization, undefined);
    assert.deepEqual(Buffer.concat(g.calls.slice(1).map(c => Buffer.from(c.opts.body))), bytes);
    const fail = load((url, opts, n) => n === 1 ? response(201, { uploadUrl: "https://outlook.office.com/upload" }) : response(500));
    await assert.rejects(fail.transport.copyFileAttachment("token", "draft", { name: "large.pdf" }, bytes.toString("base64")), /upload/);
  }
  if (source.includes("async function searchMessages(token, query, onProgress)")) {
    for (const hasMore of [false, true]) {
      const g = load((url, opts, n) => response(200, { value: [],
        "@odata.nextLink": n < 40 || hasMore ? "https://graph.microsoft.com/v1.0/me/messages?next=1" : null }));
      assert.equal((await g.api.searchMessages("token", "records")).capped, hasMore);
    }
    const g = load((url, opts, n) => response(200, n === 1 ? { value: [],
      "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages/m/attachments?next=1" } :
      { value: [{ "@odata.type": "#microsoft.graph.fileAttachment", name: "a.txt", contentBytes: "YWJj" }] }));
    assert.equal((await g.api.getAttachments("token", "m")).length, 1);
    const missing = load(() => response(200, { value: [{ "@odata.type": "#microsoft.graph.fileAttachment", id: "a", name: "missing.txt" }] }));
    await assert.rejects(missing.api.getAttachments("token", "m"), /content unavailable/);
    const text = load(() => response(200, { body: { contentType: "text", content: "<script> & hello\nworld" } }));
    assert.equal(await text.api.getBody("token", "m"), "<pre>&lt;script&gt; &amp; hello\nworld</pre>");
  }
  console.log("All Graph safety regression tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
