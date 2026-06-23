/*
 * Biweekly Activity Report — report builder (pure logic).
 *
 * Descendant of the VBA macro BuildBiWeeklyActivityReport. Takes plain data
 * (meetings + sent/received email already pulled from Graph) and produces the
 * timesheet diary. NO AI, NO network, NO Office APIs — deterministic + testable.
 *
 * HOURS MODEL (v2 — email-aware, since most workdays here show up as email, not
 * meetings):
 *   - A weekday is "worked" if it has a meeting, a sent email, or (optionally)
 *     received email. Received-only days = inbox/triage days.
 *   - mode "standard" (default): a worked weekday = `standardDay` hours (8h),
 *     extended toward the actual activity span when work ran before workStart or
 *     after workEnd (overtime), capped at `maxDay`. Never less than the day's
 *     merged meeting time.
 *   - mode "span": hours = first→last activity span that day (with a small
 *     minimum), capped at `maxDay`. Conservative / honest "I worked X to Y".
 *   - Weekends only count when there's activity, and always use the span (no
 *     full-day assumption); flagged as weekend.
 *   - Empty weekdays = 0h with an honest "no activity logged" note (no
 *     fabricated narrative).
 *
 * Other improvements over the macro: overlapping meetings merged; all-day events
 * listed but not counted; declined meetings excluded; received volume + top
 * contacts; per-week subtotals; HTML + plain-text rendering.
 *
 * Exposes a global `Report` object (no bundler, matches the project style).
 */
/* global window, module */
(function () {
  "use strict";

  var DEFAULTS = {
    daysBack: 14,
    workStart: "08:00",
    workEnd: "16:30",
    standardDay: 8,          // hours credited to a worked weekday (standard mode)
    maxDay: 11,              // cap on any single day's hours
    minActive: 0.5,         // floor for a span-mode / weekend day that had any activity
    mode: "standard",       // "standard" | "span"
    countReceivedAsWork: true,
    fillerProse: false,     // off: honest notes instead of "I was continuing work…"
  };

  // --- helpers ---------------------------------------------------------------

  function pad(n) { return (n < 10 ? "0" : "") + n; }
  function hhmm(d) { return pad(d.getHours()) + ":" + pad(d.getMinutes()); }
  function parseClock(s) {
    var m = /^(\d{1,2}):(\d{2})$/.exec(s || "");
    return m ? parseInt(m[1], 10) * 60 + parseInt(m[2], 10) : 0;
  }
  function minutesOfDay(d) { return d.getHours() * 60 + d.getMinutes(); }
  function startOfDay(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
  function addDays(d, n) { var x = new Date(d); x.setDate(x.getDate() + n); return x; }
  function isWeekend(d) { var k = d.getDay(); return k === 0 || k === 6; }
  function round2(n) { return Math.round(n * 100) / 100; }
  function round25(n) { return Math.round(n * 4) / 4; } // nearest quarter hour
  function clamp(n, lo, hi) { return Math.max(lo, Math.min(hi, n)); }

  var WEEKDAY = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  var MONTH = ["January", "February", "March", "April", "May", "June", "July",
    "August", "September", "October", "November", "December"];
  function longDate(d) {
    return WEEKDAY[d.getDay()] + ", " + MONTH[d.getMonth()] + " " + d.getDate() + ", " + d.getFullYear();
  }
  function isoDate(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  /** Merge overlapping [start,end] intervals; return total hours covered. */
  function mergedHours(intervals) {
    if (!intervals.length) return 0;
    var s = intervals.slice().sort(function (a, b) { return a.start - b.start; });
    var total = 0, curS = s[0].start, curE = s[0].end;
    for (var i = 1; i < s.length; i++) {
      if (s[i].start <= curE) { if (s[i].end > curE) curE = s[i].end; }
      else { total += curE - curS; curS = s[i].start; curE = s[i].end; }
    }
    total += curE - curS;
    return total / 3600000;
  }

  // --- build -----------------------------------------------------------------

  function build(data, cfg) {
    var c = Object.assign({}, DEFAULTS, cfg || {});
    var meetings = (data.meetings || []).slice();
    var sent = (data.sent || []).slice();
    var received = (data.received || []).slice();

    var workStartMin = parseClock(c.workStart);
    var workEndMin = parseClock(c.workEnd);

    var today = startOfDay(new Date());
    var rangeStart = addDays(today, -(c.daysBack - 1));

    var days = [];
    var contacts = {};
    var weekTotals = [];
    var totals = {
      estimatedHours: 0, meetingHours: 0, meetingCount: 0, recurringCount: 0,
      emailsSent: 0, emailsReceived: 0, daysWorked: 0,
      busiestDay: null, busiestDayHours: 0,
    };

    for (var di = 0; di < c.daysBack; di++) {
      var dayStart = addDays(rangeStart, di);
      var dayEnd = addDays(dayStart, 1);
      var weekend = isWeekend(dayStart);

      // meetings (attended, not declined)
      var dm = meetings.filter(function (m) {
        return m.start >= dayStart && m.start < dayEnd && !m.declined;
      }).sort(function (a, b) { return a.start - b.start; });
      var timed = [], allDay = [];
      dm.forEach(function (m) {
        if (m.isAllDay || (m.end - m.start) >= 86400000) allDay.push(m);
        else if (m.end > m.start) timed.push(m);
      });
      var meetingHours = mergedHours(timed.map(function (m) { return { start: m.start, end: m.end }; }));

      // sent email
      var ds = sent.filter(function (e) { return e.sentOn >= dayStart && e.sentOn < dayEnd; })
        .sort(function (a, b) { return a.sentOn - b.sentOn; });
      ds.forEach(function (e) { var to = (e.to || "").trim(); if (to) contacts[to] = (contacts[to] || 0) + 1; });

      // received count
      var receivedCount = received.filter(function (e) {
        return e.receivedOn >= dayStart && e.receivedOn < dayEnd;
      }).length;

      // activity marks (your active output) for span / overtime
      var marks = [];
      timed.forEach(function (m) { marks.push(m.start); marks.push(m.end); });
      ds.forEach(function (e) { marks.push(e.sentOn); });
      var first = marks.length ? new Date(Math.min.apply(null, marks)) : null;
      var last = marks.length ? new Date(Math.max.apply(null, marks)) : null;
      var spanH = (first && last) ? (last - first) / 3600000 : 0;

      var hasActiveOutput = timed.length > 0 || ds.length > 0;
      var hasReceived = receivedCount > 0;
      var worked = hasActiveOutput || (c.countReceivedAsWork && hasReceived);

      // --- hours ---
      var hours = 0;
      if (worked) {
        if (c.mode === "span") {
          hours = clamp(spanH, c.minActive, c.maxDay);
        } else if (weekend) {
          // weekends: no full-day assumption — use actual span
          hours = clamp(Math.max(spanH, meetingHours), c.minActive, c.maxDay);
        } else {
          var base = c.standardDay;
          var ranLong = marks.length && (minutesOfDay(first) < workStartMin || minutesOfDay(last) > workEndMin);
          hours = ranLong ? clamp(Math.max(spanH, base), base, c.maxDay) : base;
        }
        hours = clamp(Math.max(hours, meetingHours), 0, c.maxDay);
        hours = round25(hours);
      }

      // --- note (honest) ---
      var note = "";
      if (!worked) {
        note = weekend ? "" : "No meetings, sent email, or received email logged.";
      } else if (!hasActiveOutput && hasReceived) {
        note = "Inbox/triage day — " + receivedCount + " email" +
          (receivedCount === 1 ? "" : "s") + " received; no meetings or sent email logged.";
      } else if (c.fillerProse && !timed.length && !ds.length) {
        note = "Follow-up and documentation.";
      }

      // accumulate
      if (worked) totals.daysWorked++;
      totals.estimatedHours += hours;
      totals.meetingHours += meetingHours;
      totals.meetingCount += timed.length;
      totals.recurringCount += timed.filter(function (m) { return m.isRecurring; }).length;
      totals.emailsSent += ds.length;
      totals.emailsReceived += receivedCount;
      if (hours > totals.busiestDayHours) { totals.busiestDayHours = hours; totals.busiestDay = longDate(dayStart); }
      var wk = Math.floor(di / 7);
      weekTotals[wk] = (weekTotals[wk] || 0) + hours;

      days.push({
        date: dayStart, longDate: longDate(dayStart), isoDate: isoDate(dayStart),
        weekend: weekend, worked: worked,
        meetings: timed.map(function (m) {
          return { time: hhmm(m.start) + "–" + hhmm(m.end), subject: m.subject, recurring: !!m.isRecurring };
        }),
        allDay: allDay.map(function (m) { return m.subject; }),
        sent: ds.map(function (e) { return { time: hhmm(e.sentOn), to: e.to || "(recipient not shown)", subject: e.subject }; }),
        receivedCount: receivedCount,
        emailWindow: (first && last && (ds.length || timed.length)) ? hhmm(first) + "–" + hhmm(last) : "",
        note: note,
        hours: hours,
      });
    }

    totals.estimatedHours = round2(totals.estimatedHours);
    totals.meetingHours = round2(totals.meetingHours);

    var topContacts = Object.keys(contacts)
      .map(function (k) { return { to: k, count: contacts[k] }; })
      .sort(function (a, b) { return b.count - a.count; }).slice(0, 5);

    return {
      rangeStart: rangeStart, rangeEnd: today,
      rangeLabel: isoDate(rangeStart) + " to " + isoDate(today),
      days: days, weekTotals: weekTotals.map(round2), totals: totals,
      topContacts: topContacts, config: c,
    };
  }

  // --- renderers -------------------------------------------------------------

  function renderText(r) {
    var out = "Activity diary for " + r.rangeLabel + "\n";
    out += "Estimated total: " + r.totals.estimatedHours.toFixed(2) + " hours across " +
      r.totals.daysWorked + " active days\n\n";
    r.days.forEach(function (day) {
      var quietWeekend = day.weekend && !day.worked;
      if (quietWeekend) return;
      out += day.longDate + (day.weekend ? " (weekend)" : "") + " — " + day.hours.toFixed(2) + " h\n";
      if (day.meetings.length) {
        out += "  Meetings:\n";
        day.meetings.forEach(function (m) {
          out += "    • " + m.time + " — " + m.subject + (m.recurring ? " (recurring)" : "") + "\n";
        });
      }
      if (day.allDay.length) {
        out += "  All-day (not counted): " + day.allDay.join(", ") + "\n";
      }
      if (day.sent.length) {
        out += "  Sent email:\n";
        day.sent.forEach(function (e) { out += "    • " + e.time + " — To: " + e.to + " — " + e.subject + "\n"; });
        if (day.emailWindow) out += "  Activity window: " + day.emailWindow + "\n";
      }
      if (day.receivedCount) out += "  Emails received: " + day.receivedCount + "\n";
      if (day.note) out += "  " + day.note + "\n";
      out += "\n";
    });
    out += "========================================\n";
    r.weekTotals.forEach(function (h, i) { out += "Week " + (i + 1) + " total: " + h.toFixed(2) + " hours\n"; });
    out += "Estimated biweekly total: " + r.totals.estimatedHours.toFixed(2) + " hours\n";
    return out;
  }

  function renderHtml(r) {
    var t = r.totals, h = [];
    h.push('<div class="bar-report">');
    h.push('<h2>Activity diary</h2>');
    h.push('<p class="range">' + escapeHtml(r.rangeLabel) + "</p>");

    h.push('<div class="dash">');
    h.push(stat(t.estimatedHours.toFixed(2), "est. hours"));
    h.push(stat(t.daysWorked, "active days"));
    h.push(stat(t.emailsSent, "emails sent"));
    h.push(stat(t.emailsReceived, "received"));
    if (t.meetingCount) h.push(stat(t.meetingCount, "meetings"));
    h.push("</div>");
    if (t.busiestDay) {
      h.push('<p class="muted">Busiest day: ' + escapeHtml(t.busiestDay) + " (" +
        t.busiestDayHours.toFixed(2) + " h).</p>");
    }
    if (r.topContacts.length) {
      h.push('<p class="muted">Top contacts: ' + r.topContacts.map(function (x) {
        return escapeHtml(x.to) + " (" + x.count + ")";
      }).join(", ") + "</p>");
    }

    r.days.forEach(function (day) {
      var quietWeekend = day.weekend && !day.worked;
      if (quietWeekend) return;
      h.push('<div class="day' + (day.weekend ? " weekend" : "") + (day.worked ? "" : " empty") + '">');
      h.push('<div class="day-head"><span>' + escapeHtml(day.longDate) +
        (day.weekend ? ' <span class="tag">wknd</span>' : "") +
        '</span><span class="hrs">' + day.hours.toFixed(2) + " h</span></div>");
      if (day.meetings.length) {
        h.push("<ul>");
        day.meetings.forEach(function (m) {
          h.push("<li><b>" + escapeHtml(m.time) + "</b> — " + escapeHtml(m.subject) +
            (m.recurring ? ' <span class="tag">recurring</span>' : "") + "</li>");
        });
        h.push("</ul>");
      }
      if (day.allDay.length) {
        h.push('<p class="muted">All-day (not counted): ' + day.allDay.map(escapeHtml).join(", ") + "</p>");
      }
      if (day.sent.length) {
        h.push("<details><summary>" + day.sent.length + " email" + (day.sent.length === 1 ? "" : "s") +
          " sent" + (day.emailWindow ? " · " + escapeHtml(day.emailWindow) : "") + "</summary><ul>");
        day.sent.forEach(function (e) {
          h.push("<li><b>" + escapeHtml(e.time) + "</b> → " + escapeHtml(e.to) + " — " + escapeHtml(e.subject) + "</li>");
        });
        h.push("</ul></details>");
      }
      if (day.receivedCount) h.push('<p class="muted">' + day.receivedCount + " emails received</p>");
      if (day.note) h.push('<p class="note">' + escapeHtml(day.note) + "</p>");
      h.push("</div>");
    });

    h.push('<div class="totals">');
    r.weekTotals.forEach(function (hrs, i) { h.push("<div>Week " + (i + 1) + ": <b>" + hrs.toFixed(2) + " h</b></div>"); });
    h.push('<div class="grand">Biweekly total: <b>' + t.estimatedHours.toFixed(2) + " h</b></div>");
    h.push("</div>");

    h.push("</div>");
    return h.join("");
  }

  function stat(value, label) {
    return '<div class="stat"><div class="num">' + escapeHtml(value) +
      '</div><div class="lbl">' + escapeHtml(label) + "</div></div>";
  }

  var Report = { build: build, renderText: renderText, renderHtml: renderHtml, DEFAULTS: DEFAULTS };
  if (typeof module !== "undefined" && module.exports) module.exports = Report;
  if (typeof window !== "undefined") window.Report = Report;
})();
