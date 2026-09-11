/*
 * Every page must load the scripts its code depends on.
 *
 * commands.js was refactored to use AddFlow, and commands.html was not updated
 * to load addflow.js. Nothing caught it: the unit tests inject AddFlow
 * directly, check-undeclared sees the `/* global AddFlow *\/` annotation and is
 * satisfied, and the manifest only names the page. The failure would have
 * appeared for the first time on a user's machine, as "Auto Contacts failed:
 * AddFlow is not defined" — after a sign-in.
 *
 * So: read what each module PROVIDES (its `root.X = ...` assignment), read what
 * each module NEEDS (its `/* global ... *\/` line), and check that every page
 * loading a module also loads the modules it needs.
 *
 * Run: node tools/check-pages.js
 */
"use strict";
const fs = require("fs");
const path = require("path");

// Supplied by the browser, Node, or the host - not by a file in this repo.
const AMBIENT = new Set([
  "window", "document", "console", "fetch", "self", "globalThis", "navigator",
  "location", "localStorage", "sessionStorage", "setTimeout", "clearTimeout",
  "setInterval", "clearInterval", "URL", "URLSearchParams", "Blob", "File",
  "FileReader", "TextDecoder", "TextEncoder", "DOMParser", "XMLSerializer",
  "atob", "btoa", "DecompressionStream", "CompressionStream", "crypto",
  "performance", "Intl", "AbortController", "Headers", "Request", "Response",
  // Node-only, and always reached through a typeof guard in this codebase.
  "module", "require", "process", "Buffer", "__dirname",
]);

/*
 * Globals that arrive from a CDN <script>, keyed by a fragment of the URL.
 *
 * These were being reported as missing because the scan skipped every
 * https: script tag outright - so a page loading JSZip from jsdelivr looked
 * like a page that had forgotten JSZip. Matching on the URL keeps the check
 * honest without pretending the name is ambient: a page that uses JSZip and
 * does NOT load it is still an error.
 */
const CDN_GLOBALS = [
  [/office\.js/i, ["Office", "OfficeRuntime", "OfficeExtension"]],
  [/msal-browser/i, ["msal"]],
  [/jszip/i, ["JSZip"]],
];

function walk(dir, out) {
  let entries = [];
  try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return out; }
  entries.forEach(function (e) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) { if (e.name !== "node_modules") { walk(full, out); } }
    else { out.push(full); }
  });
  return out;
}

const files = walk("src", []);
const scripts = files.filter(function (f) { return f.endsWith(".js"); });
const pages = files.filter(function (f) { return f.endsWith(".html"); });

// What each script provides, and what it needs.
const provides = new Map();   // basename -> [globals]
const needs = new Map();      // basename -> [globals]

scripts.forEach(function (f) {
  const src = fs.readFileSync(f, "utf8");
  const base = path.basename(f);

  // Modules in this suite publish themselves in several ways: root.X = (an
  // IIFE given the global), window.X =, self.X =. Matching only root.X made
  // window.X modules look like they provided nothing, which reported pages
  // that were in fact correct.
  const gives = [];
  const re = /\b(?:root|window|self|globalThis)\.([A-Za-z_$][\w$]*)\s*=(?!=)/g;
  let m;
  while ((m = re.exec(src))) { if (gives.indexOf(m[1]) < 0) { gives.push(m[1]); } }
  provides.set(base, gives);

  const g = src.match(/\/\*\s*global\s+([^*]+)\*\//);
  const wants = g
    ? g[1].split(",").map(function (s) { return s.trim(); }).filter(Boolean)
    : [];
  needs.set(base, wants);
});

let problems = 0;

pages.forEach(function (page) {
  const html = fs.readFileSync(page, "utf8");
  const loaded = [], cdn = [];
  const re = /<script[^>]*\ssrc=["']([^"']+)["']/g;
  let m;
  while ((m = re.exec(html))) {
    const src = m[1];
    if (/^https?:/i.test(src)) { cdn.push(src); continue; }
    loaded.push(path.basename(src.split("?")[0]));
  }
  if (!loaded.length) { return; }

  // Everything those scripts make available on this page.
  const available = new Set();
  loaded.forEach(function (b) { (provides.get(b) || []).forEach(function (g) { available.add(g); }); });
  cdn.forEach(function (url) {
    CDN_GLOBALS.forEach(function (pair) {
      if (pair[0].test(url)) { pair[1].forEach(function (g) { available.add(g); }); }
    });
  });

  loaded.forEach(function (b) {
    (needs.get(b) || []).forEach(function (want) {
      if (AMBIENT.has(want) || available.has(want)) { return; }
      problems++;
      const from = [...provides.entries()].find(function (e) { return e[1].indexOf(want) >= 0; });
      console.log("\n" + page);
      console.log("  " + b + " needs " + want + ", which this page never loads." +
        (from ? "\n  Add: <script src=\"" + from[0] + "\"></script>" : ""));
    });
  });

  // A script that is loaded but whose file is missing is equally fatal.
  loaded.forEach(function (b) {
    if (!provides.has(b) && !scripts.some(function (f) { return path.basename(f) === b; })) {
      problems++;
      console.log("\n" + page);
      console.log("  loads " + b + ", which does not exist in src/");
    }
  });
});

console.log(problems
  ? "\n" + problems + " page dependency problem(s)."
  : "check-pages: " + pages.length + " page(s) load everything their scripts need.");
process.exit(problems ? 1 : 0);
