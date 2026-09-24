# Fitness Journal

A self-hosted Google Apps Script workout journal backed by your own Google Sheet. Log exercises from a web app, restore unfinished sessions, reuse templates, and review training history through a dashboard and weekly email digest.

Features include per-set reps and loads, load PRs, best reps at the maximum recorded load, and longest isometric holds at the maximum recorded load.

## Files

- `Code.gs`: Apps Script backend, spreadsheet storage, dashboard, and digest.
- `Index.html`: Web app interface.
- `tests/pr-records.test.js`: Local regression tests for PR calculations and hints.

## Setup

1. Create or open a Google Sheet and add a tab named `StrengthWorkoutINPUT`.
2. Add these headers in columns A–G: `Date`, `Exercise`, `Sets`, `Reps`, `Load`, `Focus`, `Notes`.
3. Open **Extensions > Apps Script**, copy `Code.gs` into the script project, and add an HTML file named `Index` containing `Index.html`.
4. Review the configuration at the top of `Code.gs`, including historical import sheet names and the optional digest recipient. Set the script time zone to match the spreadsheet.
5. Deploy the script as a web app that executes as you, restrict access to yourself, and authorize the required Google services. This is a single-user journal: drafts, templates, and workout history are shared within one spreadsheet.

The app creates supporting tabs as needed. Reopen the spreadsheet to access the **Fitness Journal** menu for dashboard rebuilds, historical imports, exercise name merging, and optional weekly digest setup.

For isometric exercises, use a name containing `iso` and enter hold duration in seconds in the Reps/Hold field. PRs are derived from the existing workout rows; no separate record migration is required.

## Tests

With Node.js installed, run:

```sh
node --test tests/pr-records.test.js
```

Tests use mocked spreadsheet services. Live Apps Script deployment and Google service integration require separate verification.

## Deployment updates

This repository stores the source code. Pushing to GitHub does not automatically update the Apps Script deployment; copy changes to the script project and update its deployment separately.

## Configuration and data

No API keys are required. Apps Script uses Google authorization to access the bound spreadsheet and send the optional digest. Workout data stays in the spreadsheet; this repository contains source code and synthetic test fixtures only.

Historical imports are disabled by default (`IMPORT_SOURCE_SHEETS = []`). Add your own source tab names locally if needed. The optional digest recipient defaults to the user running the script; configure an override only in your own Apps Script project.

Keep spreadsheet exports, deployment identifiers, email addresses, and credentials out of public commits. Each installation should use its own spreadsheet and Apps Script project.
