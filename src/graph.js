/*
 * Biweekly Activity Report — Microsoft Graph data layer.
 *
 * This replaces the VBA macro's direct MAPI folder access:
 *   GetDefaultFolder(olFolderCalendar)  ->  GET /me/calendarView   (expands recurrences)
 *   GetDefaultFolder(olFolderSentMail)  ->  GET /me/mailFolders/sentitems/messages
 *   Restrict("[Start] >= ...")          ->  $filter / calendarView start-end window
 *
 * AUTH: Nested App Authentication (NAA) via MSAL — no backend, identical pattern
 * to the Reply-All add-in. The add-in only ever reads the SIGNED-IN user's own
 * mailbox (delegated Calendars.Read + Mail.ReadWrite).
 *
 * All times are requested in the user's mailbox time zone via the Prefer header,
 * and parsed into local Date objects so report.js can do plain getHours() math.
 *
 * Exposes a global `GraphData` object.
 */
/* global msal, window */
(function () {
  "use strict";

  // Reuse the PurposeBuiltSystems Entra app registration. If you keep this in
  // the same app as Reply-All, add the delegated scopes below to that app and
  // grant admin consent. Otherwise create a new registration and paste its id.
  // Reusing the PurposeBuiltSystems "Reply All with Attachments" Entra app.
  // That app must have the delegated scopes below added + admin-consented.
  var CLIENT_ID = "87764ff9-16e7-4e2f-8164-38eff9f3a895";
  var GRAPH = "https://graph.microsoft.com/v1.0";
  var SCOPES = ["Calendars.Read", "Mail.ReadWrite"]; // Mail.ReadWrite covers reading Sent + creating the draft

  var pcaPromise = null;

  function getPca() {
    if (!pcaPromise) {
      pcaPromise = msal.createNestablePublicClientApplication({
        auth: {
          clientId: CLIENT_ID,
          authority: "https://login.microsoftonline.com/common",
        },
      });
    }
    return pcaPromise;
  }

  // --- add-in sign-out state -------------------------------------------
  //
  // Certification rejected the naive version on a sibling add-in: "after
  // clicking sign-out there is no response or not signed out." The reason is
  // structural. Under nested app authentication Outlook owns the session and
  // getAllAccounts() reports the HOST's account, not a cache this add-in
  // controls - so clearing MSAL's cache changes nothing visible, the next
  // silent acquisition succeeds anyway, and the pane redraws as signed in.
  //
  // A sign-out this add-in cannot deliver should not be offered. What it CAN
  // deliver is refusing to act until the user authenticates again: while
  // signed out it reports itself signed out and will not use a silent token,
  // so the next action raises a real prompt. Outlook's own session is
  // untouched, and the pane says so.
  var SIGNED_OUT_KEY = "addinSignedOut";
  var signedOut = false;
  try { signedOut = Office.context.roamingSettings.get(SIGNED_OUT_KEY) === true; } catch (e) { signedOut = false; }

  function setSignedOut(v) {
    signedOut = !!v;
    try {
      Office.context.roamingSettings.set(SIGNED_OUT_KEY, signedOut);
      Office.context.roamingSettings.saveAsync(function () {});
    } catch (e) { /* in-memory is still correct for this session */ }
  }

  /** Signed-in account, or null. Reports null while signed out, by design. */

  /**
   * Microsoft Graph throttles, and this add-in makes bursts of calls - a
   * records bundle or a bulk post is dozens to hundreds. An unretried 429
   * aborts the whole run part-way, which is the worst possible failure for
   * work that is half-written. One respectful retry honouring Retry-After
   * absorbs the overwhelming majority of throttling without hammering the
   * service; anything past that is a real outage and should surface.
   */
  async function fetchRetry(url, opts) {
    var res = await fetch(url, opts);
    if (res.status === 429 || res.status === 503) {
      var wait = Number(res.headers.get("Retry-After") || 3) * 1000;
      await new Promise(function (r) { setTimeout(r, Math.min(wait, 15000)); });
      res = await fetch(url, opts);
    }
    return res;
  }

  async function currentAccount() {
    if (signedOut) { return null; }
    try {
      var pca = await getPca();
      var accts = (pca.getAllAccounts && pca.getAllAccounts()) || [];
      return accts.length ? (accts[0].username || accts[0].name || "signed in") : null;
    } catch (e) { return null; }
  }

  /**
   * Sign out of the add-in. The state flips SYNCHRONOUSLY before any awaiting
   * so the pane can respond instantly - awaiting a broker handshake first is
   * the "no response" half of the finding. Cache clearing is best-effort on
   * top; the enforced state is what makes this real.
   */
  function signOut() {
    setSignedOut(true);
    var pending = pcaPromise;      // only clear what exists; never start a handshake here
    pcaPromise = null;
    if (!pending) { return Promise.resolve(true); }
    return Promise.resolve(pending).then(function (pca) {
      var accts = (pca && pca.getAllAccounts && pca.getAllAccounts()) || [];
      var chain = Promise.resolve();
      accts.forEach(function (a) {
        chain = chain.then(function () {
          if (pca.clearCache) { return pca.clearCache({ account: a }); }
          if (pca.logoutPopup) { return pca.logoutPopup({ account: a }); }
        }).catch(function () { /* best effort; the enforced state stands */ });
      });
      return chain.then(function () { return true; });
    }).catch(function () { return true; });
  }

  /** True while the user has signed the add-in out. */
  function isSignedOut() { return signedOut; }




  /**
   * Sign-in must never hang the pane. An un-timed await on the popup flow
   * leaves a button disabled with nothing visible happening — which reads
   * to the user as "the button does nothing" and gives them nothing to act
   * on. Fail loudly instead, naming the two things that actually fix it.
   */
  function withTimeout(promise, ms, message) {
    var timer;
    return Promise.race([
      promise.then(function (v) { clearTimeout(timer); return v; },
                   function (e) { clearTimeout(timer); throw e; }),
      new Promise(function (_, reject) {
        timer = setTimeout(function () { reject(new Error(message)); }, ms);
      }),
    ]);
  }

  async function getToken() {
    var pca = await withTimeout(getPca(), 20000,
      "Sign-in didn't start. Fully quit Outlook (Cmd+Q) and reopen, then try again.");
    try {
      // Signed out means signed out: skip silent so the user must re-authenticate.
      if (signedOut) { throw new Error("signed out of the add-in"); }
      return (await withTimeout(pca.acquireTokenSilent({ scopes: SCOPES }), 20000, "silent timeout")).accessToken;
    } catch (e) {
      var interactive = await withTimeout(
        pca.acquireTokenPopup({ scopes: SCOPES }), 120000,
        "Sign-in didn't finish. A Microsoft sign-in window may have opened behind Outlook — " +
        "check for it (or Mission Control), finish signing in, and click again. If no window " +
        "appeared at all, fully quit Outlook (Cmd+Q), reopen, and retry.");
      setSignedOut(false);   // a real interactive sign-in ends the signed-out state
      return interactive.accessToken;
    }
  }

  function localTimeZone() {
    try { return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC"; }
    catch (e) { return "UTC"; }
  }

  async function graph(token, path, prefer) {
    var headers = { Authorization: "Bearer " + token };
    if (prefer) headers["Prefer"] = prefer;
    var res = await fetchRetry(GRAPH + path, { headers: headers });
    if (!res.ok) {
      var text = await res.text();
      throw new Error("Graph GET " + path + " -> " + res.status + " " + text);
    }
    return res.json();
  }

  /** Page through a Graph collection following @odata.nextLink. */
  async function graphAll(token, path, prefer) {
    var items = [];
    var url = path;
    var guard = 0;
    while (url && guard++ < 50) {
      var page = await graph(token, url, prefer);
      items = items.concat(page.value || []);
      var next = page["@odata.nextLink"];
      url = next ? next.substring(GRAPH.length) : null;
    }
    return items;
  }

  /**
   * Graph returns dateTime like "2026-06-10T09:30:00.0000000" already shifted to
   * the time zone we asked for (no offset). Parse the wall-clock components into
   * a local Date so report.js can reason in the user's day boundaries.
   */
  function parseLocal(graphDateTime) {
    var m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/.exec(graphDateTime || "");
    if (!m) return new Date(graphDateTime);
    return new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5], +m[6]);
  }

  function toIsoNoZ(date) {
    function p(n) { return (n < 10 ? "0" : "") + n; }
    return date.getFullYear() + "-" + p(date.getMonth() + 1) + "-" + p(date.getDate()) +
      "T" + p(date.getHours()) + ":" + p(date.getMinutes()) + ":" + p(date.getSeconds());
  }

  /**
   * Pull everything report.build() needs for the window.
   * @param {number} daysBack
   * @returns {Promise<{meetings:Array, sent:Array, received:Array}>}
   */
  async function fetchActivity(daysBack) {
    var token = await getToken();
    var tz = localTimeZone();
    var prefer = 'outlook.timezone="' + tz + '"';

    var now = new Date();
    var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    var rangeStart = new Date(startOfToday);
    rangeStart.setDate(rangeStart.getDate() - (daysBack - 1));
    var rangeEnd = new Date(startOfToday);
    rangeEnd.setDate(rangeEnd.getDate() + 1); // through end of today

    // $filter (mail) and calendarView are evaluated in UTC, so the window
    // boundaries must be TRUE UTC instants — toISOString() converts the local
    // midnight Date correctly. (Previously we pasted a local time and appended
    // "Z", which shifted the whole window by the timezone offset and mis-bucketed
    // email by ~5-6 hours.) We still send Prefer: outlook.timezone so Graph
    // RETURNS times in the user's zone; report.js then buckets into local days.
    var startUtc = rangeStart.toISOString();
    var endUtc = rangeEnd.toISOString();

    var diagnostics = {
      tz: tz,
      windowLocal: toIsoNoZ(rangeStart) + " to " + toIsoNoZ(rangeEnd),
      windowUtc: startUtc + " to " + endUtc,
      counts: { meetingsRaw: 0, meetings: 0, sent: 0, received: 0 },
      errors: [],
    };

    // --- calendar: read EVERY calendar (default + secondary), not just the
    //     default one. /me/calendarView covers ONLY the default calendar, which
    //     is why meetings on other calendars showed as zero. We enumerate
    //     /me/calendars and run calendarView per calendar (recurrences expanded),
    //     skipping obvious noise calendars (birthdays/holidays). ---
    var meetings = [];
    var seenEvents = {};
    diagnostics.calendars = [];

    function mapEvent(e) {
      var resp = (e.responseStatus && e.responseStatus.response) || "";
      return {
        subject: e.subject || "(no subject)",
        start: parseLocal(e.start && e.start.dateTime),
        end: parseLocal(e.end && e.end.dateTime),
        isAllDay: !!e.isAllDay,
        isRecurring: e.type === "occurrence" || e.type === "exception" || e.type === "seriesMaster",
        declined: resp === "declined",
        organizer: resp === "organizer",
      };
    }
    function ingest(raw) {
      raw.forEach(function (e) {
        if (e.isCancelled) return;
        var key = (e.subject || "") + "|" + (e.start && e.start.dateTime);
        if (seenEvents[key]) return;
        seenEvents[key] = 1;
        meetings.push(mapEvent(e));
      });
    }
    var viewQs = "calendarView?startDateTime=" + encodeURIComponent(startUtc) +
      "&endDateTime=" + encodeURIComponent(endUtc) +
      "&$select=subject,start,end,isAllDay,type,responseStatus,isCancelled" +
      "&$orderby=start/dateTime&$top=500";

    try {
      var calendars = await graphAll(token, "/me/calendars?$select=id,name&$top=100", null);
      if (!calendars.length) {
        var raw0 = await graphAll(token, "/me/" + viewQs, prefer);
        diagnostics.counts.meetingsRaw += raw0.length;
        diagnostics.calendars.push({ name: "(default)", count: raw0.length });
        ingest(raw0);
      } else {
        for (var ci = 0; ci < calendars.length; ci++) {
          var cal = calendars[ci];
          if (/birthday|holiday/i.test(cal.name || "")) {
            diagnostics.calendars.push({ name: cal.name, skipped: true });
            continue;
          }
          try {
            var raw = await graphAll(token, "/me/calendars/" + cal.id + "/" + viewQs, prefer);
            diagnostics.calendars.push({ name: cal.name, count: raw.length });
            diagnostics.counts.meetingsRaw += raw.length;
            ingest(raw);
          } catch (e2) {
            diagnostics.calendars.push({ name: cal.name, error: (e2 && e2.message) || String(e2) });
          }
        }
      }
    } catch (e) {
      diagnostics.errors.push("calendar: " + ((e && e.message) || e));
    }
    diagnostics.counts.meetings = meetings.length;

    // --- sent mail ---
    var sent = [];
    try {
      var sentFilter = "sentDateTime ge " + startUtc + " and sentDateTime lt " + endUtc;
      var sentPath = "/me/mailFolders/sentitems/messages?$filter=" +
        encodeURIComponent(sentFilter) +
        "&$select=subject,toRecipients,sentDateTime&$orderby=sentDateTime&$top=500";
      var sentRaw = await graphAll(token, sentPath, prefer);
      sent = sentRaw.map(function (mi) {
        var to = (mi.toRecipients || []).map(function (r) {
          return (r.emailAddress && (r.emailAddress.name || r.emailAddress.address)) || "";
        }).filter(Boolean).join("; ");
        return {
          subject: mi.subject || "(no subject)",
          to: to,
          sentOn: parseLocal(mi.sentDateTime),
        };
      });
    } catch (e) {
      diagnostics.errors.push("sent: " + ((e && e.message) || e));
    }
    diagnostics.counts.sent = sent.length;

    // --- received mail (inbox; volume + top-contact stats) ---
    var received = [];
    try {
      var recvFilter = "receivedDateTime ge " + startUtc + " and receivedDateTime lt " + endUtc;
      var recvPath = "/me/mailFolders/inbox/messages?$filter=" +
        encodeURIComponent(recvFilter) +
        "&$select=from,subject,receivedDateTime&$orderby=receivedDateTime&$top=500";
      var recvRaw = await graphAll(token, recvPath, prefer);
      received = recvRaw.map(function (mi) {
        var ea = (mi.from && mi.from.emailAddress) || {};
        return {
          from: ea.name || ea.address || "",
          fromAddress: (ea.address || "").toLowerCase(),
          subject: mi.subject || "(no subject)",
          receivedOn: parseLocal(mi.receivedDateTime),
        };
      });
    } catch (e) {
      diagnostics.errors.push("received: " + ((e && e.message) || e));
    }
    diagnostics.counts.received = received.length;

    return { meetings: meetings, sent: sent, received: received, diagnostics: diagnostics };
  }

  /** Create a draft email with the report (mirrors the macro's m.Display). */
  async function saveDraft(subject, htmlBody) {
    var token = await getToken();
    var res = await fetchRetry(GRAPH + "/me/messages", {
      method: "POST",
      headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
      body: JSON.stringify({
        subject: subject,
        body: { contentType: "HTML", content: htmlBody },
        // No toRecipients: it's a self-review draft, exactly like the VBA Display().
      }),
    });
    if (!res.ok) throw new Error("Create draft -> " + res.status + " " + (await res.text()));
    return res.json(); // includes webLink + id
  }

  var GraphData = {
    signOut: signOut,
    currentAccount: currentAccount,
    isSignedOut: isSignedOut,
    fetchActivity: fetchActivity,
    saveDraft: saveDraft,
    _config: { get clientId() { return CLIENT_ID; }, scopes: SCOPES },
  };
  if (typeof window !== "undefined") window.GraphData = GraphData;
})();
