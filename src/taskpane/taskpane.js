/*
 * Biweekly Activity Report — task pane UI wiring.
 * Pulls activity via GraphData, builds the report via Report (pure), renders the
 * preview, and lets the user copy it or save it as a Drafts email.
 */
/* global Office, GraphData, Report, document */
(function () {
  "use strict";

  var lastReport = null;
  var SETTINGS_KEY = "bar.settings";

  Office.onReady(function () {
    byId("generate").addEventListener("click", generate);
    try {
      var saved = JSON.parse(Office.context.roamingSettings.get(SETTINGS_KEY) || "{}");
      ["projectName", "projectKeywords", "sentEmailMin", "recvEmailMin"].forEach(function (k) {
        if (saved[k] != null && saved[k] !== "") { byId(k).value = saved[k]; }
      });
      if (saved.projectOnly) { byId("projectOnly").checked = true; }
    } catch (e) { /* defaults */ }
    ["projectName", "projectKeywords", "projectOnly", "sentEmailMin", "recvEmailMin"].forEach(function (id) {
      byId(id).addEventListener("change", function () {
        try {
          Office.context.roamingSettings.set(SETTINGS_KEY, JSON.stringify({
            projectName: byId("projectName").value,
            projectKeywords: byId("projectKeywords").value,
            projectOnly: byId("projectOnly").checked,
            sentEmailMin: byId("sentEmailMin").value,
            recvEmailMin: byId("recvEmailMin").value,
          }));
          Office.context.roamingSettings.saveAsync(function () {});
        } catch (e) { /* session-only */ }
      });
    });
    byId("copy").addEventListener("click", copyHtml);
    byId("copyText").addEventListener("click", copyText);
    byId("draft").addEventListener("click", saveDraft);
    // Phone-width panes: start with the options folded so the primary
    // action and the report get the space.
    if (window.innerWidth < 480) { byId("options").removeAttribute("open"); }
  });

  function byId(id) { return document.getElementById(id); }

  function setStatus(kind, text) {
    var el = byId("status");
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "status " + kind;
    el.textContent = text;
  }

  function cfgFromUI() {
    return {
      daysBack: clampInt(byId("daysBack").value, 1, 60, 14),
      mode: byId("mode").value === "span" ? "span" : "standard",
      standardDay: parseFloat(byId("standardDay").value) || 8,
      maxDay: parseFloat(byId("maxDay").value) || 11,
      workStart: byId("workStart").value || "08:00",
      workEnd: byId("workEnd").value || "16:30",
      countReceivedAsWork: byId("countReceived").checked,
      listReceived: byId("listReceived").checked,
      projectName: byId("projectName").value.trim(),
      projectKeywords: byId("projectKeywords").value.split(",").map(function (k) { return k.trim(); }).filter(Boolean),
      projectOnly: byId("projectOnly").checked,
      sentEmailMin: clampInt(byId("sentEmailMin").value, 0, 60, 5),
      recvEmailMin: clampInt(byId("recvEmailMin").value, 0, 60, 2),
    };
  }

  function clampInt(v, min, max, dflt) {
    var n = parseInt(v, 10);
    if (isNaN(n)) return dflt;
    return Math.max(min, Math.min(max, n));
  }

  async function generate() {
    var cfg = cfgFromUI();
    byId("generate").disabled = true;
    byId("output").hidden = true;
    setStatus("work", "Reading your calendar and sent mail…");
    try {
      var data = await GraphData.fetchActivity(cfg.daysBack);
      lastReport = Report.build(data, cfg);
      byId("preview").innerHTML = Report.renderHtml(lastReport);
      byId("output").hidden = false;
      // Fold the options away and bring the report into view — matters most
      // in the narrow mobile pane where the controls eat the viewport.
      byId("options").removeAttribute("open");
      if (byId("output").scrollIntoView) {
        byId("output").scrollIntoView({ behavior: "smooth", block: "start" });
      }

      // Diagnostics: show exactly what Graph returned so we can verify the data.
      var d = data.diagnostics || { counts: {}, errors: [] };
      var c = d.counts || {};
      var diag = "Fetched from Graph → meetings: " + (c.meetings != null ? c.meetings : "?") +
        " (raw " + (c.meetingsRaw != null ? c.meetingsRaw : "?") + "), sent: " +
        (c.sent != null ? c.sent : "?") + ", received: " + (c.received != null ? c.received : "?") +
        " · window " + (d.windowLocal || "") + " · tz " + (d.tz || "");
      if (d.calendars && d.calendars.length) {
        diag += " · calendars: " + d.calendars.map(function (x) {
          return x.name + "(" + (x.error ? "ERR" : (x.skipped ? "skip" : x.count)) + ")";
        }).join(", ");
      }
      if (d.errors && d.errors.length) {
        setStatus("error", "Graph errors: " + d.errors.join("  |  ") + "   —   " + diag);
      } else {
        setStatus("info", diag);
      }
    } catch (e) {
      var msg = (e && e.message) || String(e);
      if (/REPLACE_WITH_ENTRA_CLIENT_ID/.test(GraphData._config.clientId)) {
        msg = "Set your Entra client ID in src/graph.js before running. (" + msg + ")";
      }
      setStatus("error", "Could not build the report: " + msg);
    } finally {
      byId("generate").disabled = false;
    }
  }

  async function copyHtml() {
    if (!lastReport) return;
    var html = Report.renderHtml(lastReport);
    try {
      if (navigator.clipboard && window.ClipboardItem) {
        await navigator.clipboard.write([
          new ClipboardItem({ "text/html": new Blob([html], { type: "text/html" }) }),
        ]);
      } else {
        await navigator.clipboard.writeText(Report.renderText(lastReport));
      }
      setStatus("info", "Copied formatted report to the clipboard.");
    } catch (e) {
      setStatus("error", "Copy failed: " + ((e && e.message) || e));
    }
  }

  async function copyText() {
    if (!lastReport) return;
    try {
      await navigator.clipboard.writeText(Report.renderText(lastReport));
      setStatus("info", "Copied plain-text diary to the clipboard.");
    } catch (e) {
      setStatus("error", "Copy failed: " + ((e && e.message) || e));
    }
  }

  async function saveDraft() {
    if (!lastReport) return;
    byId("draft").disabled = true;
    setStatus("work", "Creating a draft in your mailbox…");
    try {
      var subject = "Timesheet diary " + lastReport.rangeLabel;
      var html = Report.renderHtml(lastReport);
      var draft = await GraphData.saveDraft(subject, html);
      setStatus("info", "Saved to Drafts. Open Outlook Drafts to review and send.");
      if (draft && draft.webLink) {
        Office.context.ui.openBrowserWindow
          ? Office.context.ui.openBrowserWindow(draft.webLink)
          : window.open(draft.webLink, "_blank");
      }
    } catch (e) {
      setStatus("error", "Could not save draft: " + ((e && e.message) || e));
    } finally {
      byId("draft").disabled = false;
    }
  }
})();
