# Contributing

Thanks for your interest — user feedback and contributions genuinely shape
this suite. A few things about how this project works before you start:

## This repository IS the live product

The code on `main` is served directly (via GitHub Pages) to every user's
Outlook, and the add-in is certified through Microsoft's marketplace
review. That has two consequences:

1. **Every change is reviewed line-by-line by the maintainer** before merge.
2. **Merges and deployments are performed only by the maintainer.** This is
   the security posture the add-in's certification and its users' IT
   departments rely on — it isn't negotiable, and it isn't about you.

## The best ways to contribute

- **Open an issue first** — a clear description of a problem, a confusing
  behavior, or a feature idea is as valuable as code. Real-world usage
  reports are the thing a solo maintainer can't generate.
- **Small, focused pull requests** — one change per PR, matching the
  existing style (vanilla JS, no dependencies, no build step, comments
  explain *why*).
- **Run the tests** — `node test/*.test.js` runs offline with no setup.
  New behavior needs a test; bug fixes need a regression test.

## Architecture ground rules

These are load-bearing for the suite's privacy story — PRs that break them
can't be merged, however good the feature:

- No servers, no telemetry, no analytics, no external services.
- Delegated Microsoft Graph permissions only, and no new scopes without a
  documented reason the current ones can't work.
- Deterministic logic — no AI calls.
- No new runtime dependencies without discussion (currently: Office.js,
  MSAL, and in some repos JSZip — all loaded from Microsoft/jsdelivr CDNs).

## Security reports

Email Matthew@purposebuilt.systems — please don't open public issues for
vulnerabilities. See also the suite Trust Center:
https://purposebuiltsystems.github.io/trust.html
