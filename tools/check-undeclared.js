/*
 * Find identifiers that are referenced but never declared — the class of bug
 * that made connectRules() throw "Can't find variable: quiet" on every
 * successful run. Real AST scope analysis, not regex guessing.
 *
 * Usage: node tools/check-undeclared.js [file.js ...]   (defaults to src/)
 *
 * Honours each file's /* global A, B *\/ header, which is how this suite
 * declares the modules it loads via separate <script> tags.
 */
"use strict";
const fs = require("fs");
const path = require("path");

// acorn ships with several of the suite's dev toolchains; find one, and skip
// quietly rather than failing the build if none is installed here.
let acorn = null;
for (const p of ["acorn",
                 path.join(__dirname, "..", "node_modules", "acorn"),
                 path.join(__dirname, "..", "..", "Reply all with attachments", "node_modules", "acorn"),
                 path.join(__dirname, "..", "..", "biweekly-activity-report", "node_modules", "acorn")]) {
  try { acorn = require(p); break; } catch (e) { /* try the next one */ }
}
if (!acorn) {
  console.log("check-undeclared: acorn not installed here — skipping (run npm i in a sibling add-in to enable).");
  process.exit(0);
}

const GLOBALS = new Set(`
Office OfficeRuntime document window console JSON Math Date String Number Boolean Array Object
RegExp Promise Error TypeError RangeError SyntaxError setTimeout clearTimeout setInterval
clearInterval fetch Headers Request Response btoa atob encodeURIComponent decodeURIComponent
encodeURI decodeURI unescape escape parseInt parseFloat isNaN isFinite navigator location
history localStorage sessionStorage alert confirm prompt URL URLSearchParams Blob File FileReader
FormData TextEncoder TextDecoder Uint8Array ArrayBuffer DataView Map Set WeakMap WeakSet Symbol
Proxy Reflect BigInt Intl AbortController Image ClipboardItem DOMParser XMLSerializer Node Element
HTMLElement Event CustomEvent MutationObserver requestAnimationFrame cancelAnimationFrame
performance crypto structuredClone queueMicrotask globalThis self undefined NaN Infinity
module exports require process __dirname __filename
msal JSZip
`.trim().split(/\s+/));

function analyze(file) {
  const src = fs.readFileSync(file, "utf8");
  // honour the file's own /* global A, B */ directive — the suite loads its
  // modules as separate <script> tags, so cross-file names are legitimate
  const declaredGlobals = new Set();
  const gm = src.match(/\/\*\s*globals?\s+([^*]+)\*\//g) || [];
  gm.forEach(function (block) {
    block.replace(/\/\*\s*globals?\s+|\*\//g, "").split(/[,\s]+/).forEach(function (n) {
      const name = n.split(":")[0].trim();
      if (name) { declaredGlobals.add(name); }
    });
  });
  let ast;
  try {
    ast = acorn.parse(src, { ecmaVersion: 2022, allowReturnOutsideFunction: true });
  } catch (e) {
    return [{ name: "(parse error)", line: e.loc ? e.loc.line : 0, detail: e.message }];
  }

  const scopes = [];
  function pushScope(node, kind) {
    const s = { node, kind, names: new Set(), parent: scopes[scopes.length - 1] || null };
    scopes.push(s);
    return s;
  }

  // --- pass 1: collect declarations, honouring function-scope hoisting -----
  const scopeOf = new Map();          // node -> scope
  const refs = [];                    // {name, line, scope}

  function declarePattern(pat, scope) {
    if (!pat) { return; }
    switch (pat.type) {
      case "Identifier": scope.names.add(pat.name); break;
      case "ObjectPattern": pat.properties.forEach(function (p) {
        declarePattern(p.value || p.argument, scope);
      }); break;
      case "ArrayPattern": pat.elements.forEach(function (el) { declarePattern(el, scope); }); break;
      case "AssignmentPattern": declarePattern(pat.left, scope); break;
      case "RestElement": declarePattern(pat.argument, scope); break;
      default: break;
    }
  }

  function walk(node, scope) {
    if (!node || typeof node.type !== "string") { return; }
    let current = scope;

    if (node.type === "Program") {
      current = pushScope(node, "program");
    } else if (/Function(Declaration|Expression)|ArrowFunctionExpression/.test(node.type)) {
      if (node.type === "FunctionDeclaration" && node.id) { scope.names.add(node.id.name); }
      current = pushScope(node, "function");
      if (node.id && node.type === "FunctionExpression") { current.names.add(node.id.name); }
      node.params.forEach(function (p) { declarePattern(p, current); });
      current.names.add("arguments");
    } else if (node.type === "CatchClause" && node.param) {
      current = pushScope(node, "catch");
      declarePattern(node.param, current);
    } else if (node.type === "VariableDeclaration") {
      node.declarations.forEach(function (d) { declarePattern(d.id, scope); });
    } else if (node.type === "ClassDeclaration" && node.id) {
      scope.names.add(node.id.name);
    }

    scopeOf.set(node, current);

    for (const key of Object.keys(node)) {
      if (key === "type" || key === "start" || key === "end" || key === "loc") { continue; }
      const child = node[key];
      // skip non-computed member property and object keys: they're not refs
      if (node.type === "MemberExpression" && key === "property" && !node.computed) { continue; }
      if (node.type === "Property" && key === "key" && !node.computed) { continue; }
      if (node.type === "MethodDefinition" && key === "key" && !node.computed) { continue; }
      if ((node.type === "LabeledStatement" || node.type === "BreakStatement" ||
           node.type === "ContinueStatement") && key === "label") { continue; }
      if (Array.isArray(child)) { child.forEach(function (c) { walk(c, current); }); }
      else if (child && typeof child.type === "string") { walk(child, current); }
    }
  }
  walk(ast, null);

  // --- pass 2: collect references against the built scopes ----------------
  function collect(node, scope) {
    if (!node || typeof node.type !== "string") { return; }
    const current = scopeOf.get(node) || scope;
    if (node.type === "Identifier") {
      refs.push({ name: node.name, line: lineOf(src, node.start), scope: scope });
    }
    for (const key of Object.keys(node)) {
      if (key === "type" || key === "start" || key === "end" || key === "loc") { continue; }
      const child = node[key];
      if (node.type === "MemberExpression" && key === "property" && !node.computed) { continue; }
      if (node.type === "Property" && key === "key" && !node.computed) { continue; }
      if (node.type === "MethodDefinition" && key === "key" && !node.computed) { continue; }
      if ((node.type === "LabeledStatement" || node.type === "BreakStatement" ||
           node.type === "ContinueStatement") && key === "label") { continue; }
      if (Array.isArray(child)) { child.forEach(function (c) { collect(c, current); }); }
      else if (child && typeof child.type === "string") { collect(child, current); }
    }
  }
  collect(ast, null);

  function resolves(name, scope) {
    for (let s = scope; s; s = s.parent) { if (s.names.has(name)) { return true; } }
    return false;
  }

  const seen = new Set();
  const out = [];
  refs.forEach(function (r) {
    if (GLOBALS.has(r.name) || declaredGlobals.has(r.name)) { return; }
    if (resolves(r.name, r.scope)) { return; }
    const key = r.name + ":" + r.line;
    if (seen.has(key)) { return; }
    seen.add(key);
    out.push({ name: r.name, line: r.line });
  });
  return out;
}

function lineOf(src, index) {
  let line = 1;
  for (let i = 0; i < index && i < src.length; i++) { if (src[i] === "\n") { line++; } }
  return line;
}

function defaultTargets() {
  const out = [];
  const roots = ["src", path.join("src", "taskpane"), path.join("src", "commands")];
  roots.forEach(function (dir) {
    let names = [];
    try { names = fs.readdirSync(dir); } catch (e) { return; }
    names.filter(function (n) { return n.endsWith(".js"); })
         .forEach(function (n) { out.push(path.join(dir, n)); });
  });
  return out;
}

let bad = 0;
const targets = process.argv.slice(2).length ? process.argv.slice(2) : defaultTargets();
targets.forEach(function (f) {
  const issues = analyze(f);
  if (issues.length) {
    bad += issues.length;
    console.log("\n" + f);
    issues.forEach(function (i) { console.log("  line " + i.line + ": " + i.name + (i.detail ? " — " + i.detail : "")); });
  }
});
console.log(bad ? "\n" + bad + " undeclared reference(s) — these throw at runtime under \"use strict\"."
                : "check-undeclared: no undeclared references in " + targets.length + " file(s).");
process.exit(bad ? 1 : 0);
