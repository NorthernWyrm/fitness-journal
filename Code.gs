// Fitness Journal — Apps Script backend
// Bind this script to your existing fitness journal Google Sheet
// (Extensions > Apps Script from within the Sheet itself).

var SHEET_NAME = 'StrengthWorkoutINPUT';
var DASHBOARD_SHEET_NAME = 'Dashboard';
var TEMPLATES_SHEET_NAME = 'Templates';
var DRAFT_SHEET_NAME = '_Draft'; // hidden — holds the in-progress session so mobile browsers (iOS localStorage restrictions) don't lose data
var SESSIONS_SHEET_NAME = '_Sessions'; // hidden — one row per saved session: Date, Start, End (for Workout Density)

// Exercises need AT LEAST this many sessions AND AT LEAST this much total
// volume to appear in the exercise-level dashboard tables/chart. Keeps rarely
// -performed, low-load exercises from cluttering long-term stats.
var EXERCISE_MIN_SESSIONS = 2;
var EXERCISE_MIN_VOLUME_KG = 300;

// Even after the threshold above, the bar chart shows at most this many
// exercises — the top ones by total volume. Keeps the chart focused as
// volume accumulates over time instead of growing indefinitely.
var EXERCISE_CHART_MAX_COUNT = 15;

// One-time historical import: list any old tabs with the same
// Date/Exercise/Sets/Reps/Load/Focus/Notes structure here.
var IMPORT_SOURCE_SHEETS = []; // Optional: add your own historical sheet names.

// Weekly digest settings. Recipient defaults to whoever owns/runs the script;
// set an explicit address here to override.
var DIGEST_RECIPIENT_EMAIL = null;

// An exercise is flagged as plateaued if it qualifies for the dashboard chart
// (same EXERCISE_MIN_SESSIONS/EXERCISE_MIN_VOLUME_KG threshold above) and its
// all-time best load hasn't been beaten in this many sessions since.
var PLATEAU_LOOKBACK_SESSIONS = 4;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Fitness Journal")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Adds a manual "Rebuild Dashboard" / "Import Historical Data" menu when opening the Sheet
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Fitness Journal')
    .addItem('Rebuild Dashboard', 'updateDashboard_')
    .addItem('Import Historical Data (one-time)', 'importHistoricalData')
    .addItem('Merge Exercise Names...', 'mergeExerciseNames')
    .addSeparator()
    .addItem('Set Up Weekly Digest (Monday mornings)', 'setUpWeeklyDigestTrigger')
    .addItem('Send Weekly Digest Now (test)', 'sendWeeklyDigest')
    .addToUi();
}

/**
 * Called from the client. Expects:
 * {
 *   date: "2026-08-12",
 *   exercises: [
 *     { exercise, sets, reps, load, focus, notes },
 *     ...
 *   ]
 * }
 * Appends one row per exercise, all sharing the same date, then refreshes the dashboard.
 */
function submitSession(payload) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet tab "' + SHEET_NAME + '" not found.');
  }

  var date = payload.date;
  var exercises = payload.exercises || [];

  if (!date) throw new Error('Date is required.');
  if (exercises.length === 0) throw new Error('Add at least one exercise.');

  var knownExercises = getExistingExercises();
  var priorRecords = getPRRecords();

  var rows = exercises.map(function (ex) {
    return [
      date,
      normalizeExercise_(ex.exercise || '', knownExercises),
      ex.sets || '',
      ex.reps || '',
      ex.load || '',
      ex.focus || '',
      ex.notes || ''
    ];
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 7).setValues(rows);

  // Compare the best set across all cards for each exercise with saved history.
  var sessionRecords = {};
  rows.forEach(function (row) {
    mergePRRecord_(sessionRecords, row[1], getLoadRecord_(row[3], row[4]));
  });
  var prs = [];
  Object.keys(sessionRecords).forEach(function (name) {
    var current = sessionRecords[name], prior = priorRecords[name];
    var loadPR = !prior || current.maxLoad > prior.maxLoad;
    var repsPR = prior && current.maxLoad === prior.maxLoad && current.bestReps !== null &&
      (prior.bestReps === null || current.bestReps > prior.bestReps);
    if (loadPR || repsPR) {
      prs.push({ exercise: name, newMax: current.maxLoad, previousMax: prior ? prior.maxLoad : null,
        bestReps: current.bestReps, previousReps: prior ? prior.bestReps : null,
        type: loadPR ? 'load' : 'reps', isIsometric: isIsometricExercise_(name) });
    }
  });

  try {
    recordSessionTiming_(date, payload.sessionStart);
  } catch (e) {
    Logger.log('Session timing record failed: ' + e.message);
  }

  try {
    updateDashboard_();
  } catch (e) {
    // Don't let a dashboard hiccup block the actual save
    Logger.log('Dashboard update failed: ' + e.message);
  }

  return { success: true, rowsAdded: rows.length, prs: prs };
}

/**
 * Public wrapper: returns { exerciseName: heaviestLoadEverLogged } so the
 * client can show a live PR indicator while filling in the Load field,
 * before the session is even saved.
 */
function getMaxLoads() {
  return getMaxLoadByExercise_();
}

// Raw reps are seconds for isometrics here; do not apply the 1RM/volume conversion.
// Keep invalid/missing slots in place so a later rep count cannot shift to another load.
function getLoadRecord_(repsStr, loadStr) {
  function slots(value) {
    if (value === null || value === undefined || value === '') return [];
    return value.toString().split(',').map(function (part) {
      var n = parseFloat(part.trim());
      return isFinite(n) && n >= 0 ? n : null;
    });
  }
  var reps = slots(repsStr), loads = slots(loadStr);
  var record = null;
  for (var i = 0; i < Math.max(reps.length, loads.length); i++) {
    var load = loads.length ? loads[Math.min(i, loads.length - 1)] : null;
    var rep = reps.length === 1 ? reps[0] : (i < reps.length ? reps[i] : null);
    if (load === null) continue;
    if (!record || load > record.maxLoad) record = { maxLoad: load, bestReps: rep };
    else if (load === record.maxLoad && rep !== null &&
             (record.bestReps === null || rep > record.bestReps)) record.bestReps = rep;
  }
  return record;
}

function mergePRRecord_(records, name, record) {
  if (!name || !record) return;
  var prior = records[name];
  if (!prior || record.maxLoad > prior.maxLoad) {
    records[name] = { maxLoad: record.maxLoad, bestReps: record.bestReps };
  } else if (record.maxLoad === prior.maxLoad && record.bestReps !== null &&
             (prior.bestReps === null || record.bestReps > prior.bestReps)) {
    prior.bestReps = record.bestReps;
  }
}

// Derived from existing history, so no migration or extra sheet columns are needed.
function getPRRecords() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var records = {};
  if (!sheet || sheet.getLastRow() < 2) return records;
  sheet.getRange(2, 2, sheet.getLastRow() - 1, 4).getValues().forEach(function (row) {
    mergePRRecord_(records, (row[0] || '').toString().trim(), getLoadRecord_(row[2], row[3]));
  });
  return records;
}

/**
 * Returns { exerciseName: mostCommonFocusValue } from logged history, used
 * to auto-fill the Focus dropdown when you type an exercise you've logged
 * before — only applied client-side when the field is still empty, never
 * overriding a manual pick.
 */
function getFocusModeByExercise() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;

  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 5).getValues(); // B:F = Exercise, Sets, Reps, Load, Focus
  var counts = {}; // exercise -> { focusValue: count }
  data.forEach(function (row) {
    var name = (row[0] || '').toString().trim();
    var focus = (row[4] || '').toString().trim();
    if (!name || !focus) return;
    if (!counts[name]) counts[name] = {};
    counts[name][focus] = (counts[name][focus] || 0) + 1;
  });

  Object.keys(counts).forEach(function (name) {
    var best = null, bestCount = -1;
    Object.keys(counts[name]).forEach(function (f) {
      if (counts[name][f] > bestCount) { best = f; bestCount = counts[name][f]; }
    });
    result[name] = best;
  });

  return result;
}

/**
 * Public wrapper: returns [{ focus, lastDateKey, daysSince }] sorted
 * most-overdue-first, for the live "muscle group gaps" panel at the top
 * of the app. Same underlying logic as the weekly digest's list.
 */
function getFocusGapsForApp() {
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!source || source.getLastRow() < 2) return [];
  var tz = Session.getScriptTimeZone();
  var data = source.getRange(2, 1, source.getLastRow() - 1, 7).getValues();
  return computeFocusGaps_(data, tz, new Date());
}

/**
 * Returns { exerciseName: heaviestLoadEverLogged } from the existing sheet data,
 * used to detect PRs on new saves.
 */
function getMaxLoadByExercise_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;

  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 4).getValues(); // B:E = Exercise,Sets,Reps,Load
  data.forEach(function (row) {
    var name = (row[0] || '').toString().trim();
    if (!name) return;
    var loads = parseNumbers_(row[3]);
    if (loads.length === 0) return;
    var maxLoad = Math.max.apply(null, loads);
    if (!(name in result) || maxLoad > result[name]) {
      result[name] = maxLoad;
    }
  });
  return result;
}

/**
 * Returns a sorted list of unique exercise names already logged,
 * used to power autocomplete in the form.
 */
function getExistingExercises() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues(); // column B = Exercise
  var seen = {};
  values.forEach(function (row) {
    var name = (row[0] || '').toString().trim();
    if (name) seen[name] = true;
  });

  return Object.keys(seen).sort(function (a, b) {
    return a.localeCompare(b);
  });
}

/**
 * Returns a sorted list of unique Focus values already logged,
 * used to power autocomplete in the form.
 */
function getExistingFocuses() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var values = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues(); // column F = Focus
  var seen = {};
  values.forEach(function (row) {
    var name = (row[0] || '').toString().trim();
    if (name) seen[name] = true;
  });

  return Object.keys(seen).sort(function (a, b) {
    return a.localeCompare(b);
  });
}

/**
 * If the typed exercise name matches an existing one case-insensitively
 * (ignoring extra spaces), snap it to the existing casing so
 * "back squat" and "Back Squat" don't become two separate exercises.
 * Otherwise the typed name is kept as-is (it becomes a new canonical entry).
 */
function normalizeExercise_(typed, knownExercises) {
  var trimmed = typed.toString().trim().replace(/\s+/g, ' ');
  if (!trimmed) return trimmed;

  var lower = trimmed.toLowerCase();
  for (var i = 0; i < knownExercises.length; i++) {
    if (knownExercises[i].toLowerCase() === lower) {
      return knownExercises[i]; // use the existing canonical casing
    }
  }
  return trimmed;
}

// --- In-progress Session Draft ---------------------------------------------
// Stores the client's in-progress (unsaved) session so a reload — from a
// scroll/pull-to-refresh gesture, the browser tab sleeping, etc. — doesn't
// lose it. Lives in a hidden sheet tab rather than browser localStorage
// because iOS WebKit (Safari, and anything iOS-based like Brave iOS) can
// block or wipe localStorage inside the sandboxed frame Apps Script web
// apps render in; a server-side round-trip works identically everywhere.

function getDraftSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(DRAFT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(DRAFT_SHEET_NAME);
    sheet.getRange(1, 1).setValue('');
    try { sheet.hideSheet(); } catch (e) { /* fine if it can't hide, e.g. only sheet in file */ }
  }
  return sheet;
}

// draftJson: a JSON string, or '' / null to clear
function saveDraftServer(draftJson) {
  getDraftSheet_().getRange(1, 1).setValue(draftJson || '');
  return { success: true };
}

// Returns the stored draft JSON string, or '' if none
function getDraftServer() {
  var val = getDraftSheet_().getRange(1, 1).getValue();
  return val ? val.toString() : '';
}

function clearDraftServer() {
  getDraftSheet_().getRange(1, 1).setValue('');
  return { success: true };
}

// --- Workout Density (session start/end tracking) --------------------------
// One row per saved session, kept separate from the main log since duration
// is a session-level fact (same for every exercise in that session), not a
// per-exercise one — storing it here avoids touching the main sheet's column
// layout, which a lot of other code assumes is fixed at 7 columns.

function getSessionsSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SESSIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SESSIONS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['Date', 'Session Start', 'Session End']]).setFontWeight('bold');
    try { sheet.hideSheet(); } catch (e) { /* fine if it can't hide */ }
  }
  return sheet;
}

// Called from submitSession — sessionStartIso is the client's captured
// "session began" timestamp (survives a draft restore, so a reload mid-
// session doesn't reset the clock). Session end is just "now," server-side.
function recordSessionTiming_(dateStr, sessionStartIso) {
  if (!sessionStartIso) return; // client didn't send one — skip rather than guess
  var startDate = new Date(sessionStartIso);
  if (isNaN(startDate.getTime())) return;

  var sheet = getSessionsSheet_();
  var endDate = new Date();
  sheet.appendRow([dateStr, startDate, endDate]);
}

/**
 * Returns [{ dateKey, durationMin, density }] sorted chronologically.
 * density = that date's total volume ÷ session duration in minutes.
 * Sessions under 1 minute are skipped (clock anomaly, not real data).
 * Note: if you ever log two separate sessions on the same calendar date,
 * both share that date's total volume as the denominator — a known
 * limitation of matching by date rather than a session ID.
 */
function computeWorkoutDensity_(tz) {
  var sessionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET_NAME);
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sessionsSheet || sessionsSheet.getLastRow() < 2 || !source || source.getLastRow() < 2) return [];

  // Volume per date, from the main log
  var data = source.getRange(2, 1, source.getLastRow() - 1, 7).getValues();
  var volumeByDate = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    if (!dateObj) return;
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    volumeByDate[dateKey] = (volumeByDate[dateKey] || 0) + computeVolume_(row[2], row[3], row[4], row[1]);
  });

  var sessionRows = sessionsSheet.getRange(2, 1, sessionsSheet.getLastRow() - 1, 3).getValues();
  var results = [];
  sessionRows.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    var start = row[1], end = row[2];
    if (!dateObj || !(start instanceof Date) || !(end instanceof Date)) return;

    var durationMin = (end - start) / (1000 * 60);
    if (durationMin < 1) return; // anomaly, not real data

    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    var volume = volumeByDate[dateKey] || 0;
    results.push({
      dateKey: dateKey,
      durationMin: Math.round(durationMin * 10) / 10,
      density: Math.round((volume / durationMin) * 100) / 100
    });
  });

  return results.sort(function (a, b) { return a.dateKey.localeCompare(b.dateKey); });
}

// --- Session Templates ---------------------------------------------------

function getTemplatesSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TEMPLATES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TEMPLATES_SHEET_NAME);
    sheet.getRange(1, 1, 1, 7)
      .setValues([['TemplateName', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes']])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Returns a sorted list of saved template names
function getTemplates() {
  var sheet = getTemplatesSheet_();
  if (sheet.getLastRow() < 2) return [];

  var names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  var seen = {};
  names.forEach(function (row) {
    var name = (row[0] || '').toString().trim();
    if (name) seen[name] = true;
  });
  return Object.keys(seen).sort(function (a, b) { return a.localeCompare(b); });
}

// Returns the list of exercises saved under a given template name
function getTemplateExercises(templateName) {
  var sheet = getTemplatesSheet_();
  if (sheet.getLastRow() < 2) return [];

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  var result = [];
  data.forEach(function (row) {
    if ((row[0] || '').toString().trim() === templateName) {
      result.push({
        exercise: row[1] || '',
        sets: row[2] || '',
        reps: row[3] || '',
        load: row[4] || '',
        focus: row[5] || '',
        notes: row[6] || ''
      });
    }
  });
  return result;
}

/**
 * Saves (or overwrites) a template: deletes any existing rows under this
 * name, then writes the current exercise list under it.
 */
function saveTemplate(templateName, exercises) {
  templateName = (templateName || '').toString().trim();
  if (!templateName) throw new Error('Template name is required.');
  if (!exercises || exercises.length === 0) throw new Error('Add at least one exercise before saving a template.');

  var sheet = getTemplatesSheet_();

  // Remove any existing rows for this template name
  if (sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (var i = data.length - 1; i >= 0; i--) {
      if ((data[i][0] || '').toString().trim() === templateName) {
        sheet.deleteRow(i + 2); // +2: 1-indexed and header offset
      }
    }
  }

  var rows = exercises.map(function (ex) {
    return [
      templateName,
      ex.exercise || '',
      ex.sets || '',
      ex.reps || '',
      ex.load || '',
      ex.focus || '',
      ex.notes || ''
    ];
  });

  var startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 7).setValues(rows);

  return { success: true };
}

// --- Merge Exercise Names -------------------------------------------------

/**
 * Interactive merge, run from the Sheet's "Fitness Journal" menu.
 * Lets you fold a misspelled/duplicate exercise name into its canonical
 * form across both the main log and the Templates tab.
 */
function mergeExerciseNames() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    ui.alert('Sheet tab "' + SHEET_NAME + '" not found.');
    return;
  }

  var exercises = getExistingExercises();
  if (exercises.length === 0) {
    ui.alert('No exercises logged yet — nothing to merge.');
    return;
  }

  var listText = exercises.map(function (name, i) { return (i + 1) + '. ' + name; }).join('\n');

  var fromResp = ui.prompt(
    'Merge Exercise Names — Step 1 of 2',
    'Enter the NUMBER or exact name of the exercise to merge FROM (the variant to eliminate):\n\n' + listText,
    ui.ButtonSet.OK_CANCEL
  );
  if (fromResp.getSelectedButton() !== ui.Button.OK) return;
  var fromName = resolveExerciseSelection_(fromResp.getResponseText().trim(), exercises);
  if (!fromName) {
    ui.alert('Could not match that to an existing exercise. Merge cancelled.');
    return;
  }

  var toResp = ui.prompt(
    'Merge Exercise Names — Step 2 of 2',
    'Enter the NUMBER or exact name of the exercise to merge INTO (the canonical name to KEEP):\n\n' + listText,
    ui.ButtonSet.OK_CANCEL
  );
  if (toResp.getSelectedButton() !== ui.Button.OK) return;
  var toName = resolveExerciseSelection_(toResp.getResponseText().trim(), exercises);
  if (!toName) {
    ui.alert('Could not match that to an existing exercise. Merge cancelled.');
    return;
  }

  if (fromName === toName) {
    ui.alert('Source and target are the same exercise — nothing to merge.');
    return;
  }

  var confirm = ui.alert(
    'Confirm Merge',
    'All logged rows for "' + fromName + '" will be renamed to "' + toName + '" in ' + SHEET_NAME +
      ' and in Templates. This cannot be undone automatically. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  var logUpdated = renameExerciseInSheet_(sheet, fromName, toName, 2);

  var templatesUpdated = 0;
  var templatesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATES_SHEET_NAME);
  if (templatesSheet) {
    templatesUpdated = renameExerciseInSheet_(templatesSheet, fromName, toName, 2);
  }

  updateDashboard_();

  ui.alert(
    'Merge Complete',
    'Renamed ' + logUpdated + ' row(s) in ' + SHEET_NAME +
      (templatesUpdated ? ' and ' + templatesUpdated + ' row(s) in Templates' : '') + '.',
    ui.ButtonSet.OK
  );
}

// Resolves a user-typed number or name against the known exercises list
function resolveExerciseSelection_(input, exercises) {
  var idx = parseInt(input, 10);
  if (!isNaN(idx) && idx >= 1 && idx <= exercises.length) return exercises[idx - 1];

  var lower = input.toLowerCase();
  for (var i = 0; i < exercises.length; i++) {
    if (exercises[i].toLowerCase() === lower) return exercises[i];
  }
  return null;
}

// Renames every occurrence of fromName to toName in the given column of a sheet
function renameExerciseInSheet_(sheet, fromName, toName, col) {
  if (sheet.getLastRow() < 2) return 0;
  var range = sheet.getRange(2, col, sheet.getLastRow() - 1, 1);
  var values = range.getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if ((values[i][0] || '').toString().trim() === fromName) {
      values[i][0] = toName;
      count++;
    }
  }
  if (count > 0) range.setValues(values);
  return count;
}

// --- Historical Import ---------------------------------------------------

/**
 * One-time migration: pulls rows from each tab listed in IMPORT_SOURCE_SHEETS
 * (same Date/Exercise/Sets/Reps/Load/Focus/Notes structure) into the main
 * log, normalizing exercise names against what's already there so old
 * casing/typo variants merge instead of creating duplicate "exercises".
 * Safe to re-run — exact-duplicate rows are skipped.
 * Run from the Sheet's "Fitness Journal" menu, not from the phone app.
 */
function importHistoricalData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = ss.getSheetByName(SHEET_NAME);
  if (!target) {
    ui.alert('Sheet tab "' + SHEET_NAME + '" not found.');
    return;
  }

  // Build a set of existing rows (as composite keys) to avoid double-importing
  var existingKeys = {};
  var knownExercises = getExistingExercises();
  if (target.getLastRow() >= 2) {
    var existingData = target.getRange(2, 1, target.getLastRow() - 1, 7).getValues();
    existingData.forEach(function (row) {
      existingKeys[rowKey_(row)] = true;
    });
  }

  var allNewRows = [];
  var summary = [];

  IMPORT_SOURCE_SHEETS.forEach(function (tabName) {
    var source = ss.getSheetByName(tabName);
    if (!source) {
      summary.push(tabName + ': tab not found, skipped');
      return;
    }
    if (source.getLastRow() < 2) {
      summary.push(tabName + ': no data rows, skipped');
      return;
    }

    var data = source.getRange(2, 1, source.getLastRow() - 1, 7).getValues();
    var imported = 0;
    var skippedDupes = 0;
    var skippedBlank = 0;

    data.forEach(function (row) {
      var date = row[0];
      var exerciseRaw = (row[1] || '').toString().trim();

      // Skip stray header rows or fully blank rows
      if (!date && !exerciseRaw) { skippedBlank++; return; }
      if (exerciseRaw.toLowerCase() === 'exercise') { skippedBlank++; return; }
      if (!date || !exerciseRaw) { skippedBlank++; return; }

      var normalizedExercise = normalizeExercise_(exerciseRaw, knownExercises);
      if (knownExercises.indexOf(normalizedExercise) === -1) {
        knownExercises.push(normalizedExercise); // so later rows in this same import can match it too
      }

      var newRow = [
        date,
        normalizedExercise,
        row[2] || '',
        row[3] || '',
        row[4] || '',
        row[5] || '',
        row[6] || ''
      ];

      var key = rowKey_(newRow);
      if (existingKeys[key]) { skippedDupes++; return; }
      existingKeys[key] = true;

      allNewRows.push(newRow);
      imported++;
    });

    summary.push(tabName + ': ' + imported + ' imported, ' + skippedDupes + ' duplicates skipped, ' + skippedBlank + ' blank/header rows skipped');
  });

  if (allNewRows.length > 0) {
    var startRow = target.getLastRow() + 1;
    target.getRange(startRow, 1, allNewRows.length, 7).setValues(allNewRows);
    updateDashboard_();
  }

  ui.alert('Historical Import Complete', summary.join('\n') + '\n\nTotal new rows added: ' + allNewRows.length, ui.ButtonSet.OK);
}

// Builds a composite key for a row so exact duplicates can be detected
function rowKey_(row) {
  var date = row[0];
  var dateKey = (date instanceof Date)
    ? Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : (date || '').toString().trim();
  return [
    dateKey,
    (row[1] || '').toString().trim().toLowerCase(),
    (row[2] || '').toString().trim(),
    (row[3] || '').toString().trim(),
    (row[4] || '').toString().trim(),
    (row[5] || '').toString().trim().toLowerCase(),
    (row[6] || '').toString().trim().toLowerCase()
  ].join('|');
}

// --- Dashboard ---------------------------------------------------------

// Parses "8" or "8,6,6" into [8] or [8,6,6]
function parseNumbers_(str) {
  if (!str) return [];
  return str
    .toString()
    .split(',')
    .map(function (s) { return parseFloat(s.trim()); })
    .filter(function (n) { return !isNaN(n); });
}

// Any exercise name containing "iso" (case-insensitive) is treated as
// isometric — matches your existing naming convention, no separate tag needed.
function isIsometricExercise_(exerciseName) {
  return /iso/i.test((exerciseName || '').toString());
}

// Estimates total volume (kg) for one logged row. For isometric exercises,
// the Reps field holds hold-duration-in-seconds, not actual reps — treating
// duration as a rep count would wildly inflate volume (a 30-second hold
// isn't "30 reps"). Per-set count is preserved; each hold just counts as
// one "rep" for the math, same convention used for 1RM/Relative Intensity.
function computeVolume_(sets, reps, load, exerciseName) {
  var repsArr = parseNumbers_(reps);
  var loadArr = parseNumbers_(load);
  var setsNum = parseFloat(sets) || 0;

  if (isIsometricExercise_(exerciseName)) {
    repsArr = repsArr.length > 0 ? repsArr.map(function () { return 1; }) : [1];
  }

  if (repsArr.length > 1 || loadArr.length > 1) {
    // A load sequence also defines per-set entries; repeat a single rep count.
    var setCount = repsArr.length > 1 ? repsArr.length : loadArr.length;
    var vol = 0;
    for (var i = 0; i < setCount; i++) {
      var l = loadArr.length > 1 ? (loadArr[i] !== undefined ? loadArr[i] : loadArr[loadArr.length - 1])
                                  : (loadArr[0] || 0);
      var r = repsArr.length > 1 ? repsArr[i] : (repsArr[0] || 0);
      vol += r * l;
    }
    return vol;
  } else {
    var r = repsArr[0] || 0;
    var l = loadArr[0] || 0;
    var s = setsNum || 1;
    return s * r * l;
  }
}

// Coerces a sheet cell value (Date object or text) into a real Date, or null if unparseable
function toDateObj_(value) {
  if (value instanceof Date) return value;
  if (!value) return null;
  var parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

// "PR-style" representative load for a row: the heaviest single value logged
function computeRepresentativeLoad_(loadStr) {
  var arr = parseNumbers_(loadStr);
  return arr.length > 0 ? Math.max.apply(null, arr) : 0;
}

// Representative reps for a row: average reps per set (handles "8" or "8,6,6")
function computeRepresentativeReps_(repsStr) {
  var arr = parseNumbers_(repsStr);
  if (arr.length === 0) return 0;
  var sum = arr.reduce(function (a, b) { return a + b; }, 0);
  return Math.round((sum / arr.length) * 10) / 10;
}

// Estimates 1RM for a single (load, reps) pair using both formulas.
// Brzycki breaks down (goes negative) at reps >= 37, so it's omitted there.
// Both formulas lose real accuracy above ~12 reps — that's a formula
// limitation, not a bug, worth remembering when reading a high-rep estimate.
function estimate1RM_(load, reps) {
  if (!load || !reps || reps <= 0) return { epley: 0, brzycki: 0 };
  var epley = load * (1 + reps / 30);
  var brzycki = reps < 37 ? load * (36 / (37 - reps)) : 0;
  return { epley: epley, brzycki: brzycki };
}

// Best (highest) estimated 1RM across all sets in one logged row — mirrors
// the reps/load pairing convention used by computeVolume_, so a pyramid set
// like reps "8,6,4" / load "60,70,80" is evaluated set-by-set, not averaged.
// For isometrics, every hold counts as reps=1 (same convention as volume) —
// at reps=1 both formulas reduce to essentially the load itself, so the
// "1RM" for an isometric is roughly "heaviest load ever held," which is
// the sensible reading of a single maximal-effort hold.
function computeBestEstimated1RM_(repsStr, loadStr, exerciseName) {
  var repsArr = parseNumbers_(repsStr);
  var loadArr = parseNumbers_(loadStr);
  if (loadArr.length === 0) return { epley: 0, brzycki: 0 }; // no load, genuinely nothing to estimate

  if (isIsometricExercise_(exerciseName)) {
    // Unparseable reps ("failure", blank, etc.) still count as one hold at
    // the given load — same fallback computeVolume_ already uses, so a
    // non-numeric duration doesn't silently zero out 1RM/PR tracking.
    repsArr = repsArr.length > 0 ? repsArr.map(function () { return 1; }) : [1];
  } else if (repsArr.length === 0) {
    return { epley: 0, brzycki: 0 }; // non-isometric with unparseable reps — genuinely nothing to estimate
  }

  var bestEpley = 0, bestBrzycki = 0;
  for (var i = 0; i < repsArr.length; i++) {
    var l = loadArr.length > 1 ? (loadArr[i] !== undefined ? loadArr[i] : loadArr[loadArr.length - 1])
                                : (loadArr[0] || 0);
    var est = estimate1RM_(l, repsArr[i]);
    if (est.epley > bestEpley) bestEpley = est.epley;
    if (est.brzycki > bestBrzycki) bestBrzycki = est.brzycki;
  }
  return { epley: Math.round(bestEpley * 10) / 10, brzycki: Math.round(bestBrzycki * 10) / 10 };
}

// Classifies a single logged row's set structure from its load sequence.
// A single value (no per-set breakdown) or all-equal per-set loads is
// "Straight" — that's literally what happened (same weight every set), not
// an assumption. Ascending/descending load sequences are Pyramid/Reverse
// Pyramid; anything else (up-then-down, irregular) is "Mixed".
function classifySetType_(loadStr) {
  var loadArr = parseNumbers_(loadStr);
  if (loadArr.length <= 1) return 'Straight';

  var allEqual = loadArr.every(function (v) { return v === loadArr[0]; });
  if (allEqual) return 'Straight';

  var nonDecreasing = true, nonIncreasing = true;
  for (var i = 1; i < loadArr.length; i++) {
    if (loadArr[i] < loadArr[i - 1]) nonDecreasing = false;
    if (loadArr[i] > loadArr[i - 1]) nonIncreasing = false;
  }
  if (nonDecreasing) return 'Pyramid';
  if (nonIncreasing) return 'Reverse Pyramid';
  return 'Mixed';
}

// Most frequent set-type across an exercise's logged history
function modeSetType_(entries) {
  var counts = {};
  entries.forEach(function (e) { counts[e.setType] = (counts[e.setType] || 0) + 1; });
  var best = null, bestCount = -1;
  Object.keys(counts).forEach(function (k) {
    if (counts[k] > bestCount) { best = k; bestCount = counts[k]; }
  });
  return best || 'Straight';
}

/**
 * Public wrapper: returns { exerciseName: { epley, brzycki } } — the
 * all-time best estimated 1RM per exercise per formula. Powers the live
 * Relative Intensity hint in the app (today's set vs. this) and the
 * Dashboard's per-exercise summary.
 */
function getMaxEstimated1RMs() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var result = {};
  if (!sheet || sheet.getLastRow() < 2) return result;

  var data = sheet.getRange(2, 2, sheet.getLastRow() - 1, 4).getValues(); // B:E = Exercise, Sets, Reps, Load
  data.forEach(function (row) {
    var name = (row[0] || '').toString().trim();
    if (!name) return;
    var est = computeBestEstimated1RM_(row[2], row[3], row[0]); // Reps, Load, Exercise
    if (!result[name]) result[name] = { epley: 0, brzycki: 0 };
    if (est.epley > result[name].epley) result[name].epley = est.epley;
    if (est.brzycki > result[name].brzycki) result[name].brzycki = est.brzycki;
  });
  return result;
}

// Rolls granular Focus tags up into broader muscle groups for reporting
// (dashboard charts + weekly digest only — the raw Focus value logged on
// each row is never changed, this is purely an aggregation-time lookup).
//
// For combo tags ("Quads & Glutes", "Chest & Triceps", "Lats, Delts",
// "Back/Core"), only the FIRST-listed muscle determines the group — the
// whole row's volume counts toward that group, nothing is split.
//
// Matching is keyword/substring based (not exact-string) so typos, Spanish
// terms, and descriptive variants ("Rear Delts", "Glúteos", "Isquios
// Gluteos Espalda baja", "Calves (Iso-lateral)") all resolve correctly
// without needing every literal variant listed out. Accents are stripped
// before matching. Anything that matches nothing keeps its own name as an
// ungrouped bucket (e.g. "Stability" alone, if it ever appears first).
var FOCUS_KEYWORD_GROUPS = [
  ['glute', 'Glutes'],
  ['hamstring', 'Legs'],
  ['isquio', 'Legs'],
  ['quad', 'Legs'],
  ['adductor', 'Legs'],
  ['abductor', 'Legs'],
  ['thigh', 'Legs'],
  ['calv', 'Legs'],
  ['calf', 'Legs'],
  ['erector', 'Legs'],   // spinal erectors are near-exclusively hit by deadlifts (quad-focused)
  ['espalda', 'Back'],
  ['lats', 'Back'],
  ['delt', 'Back'],
  ['shoulder', 'Back'],
  ['back', 'Back'],
  ['chest', 'Chest'],
  ['bicep', 'Biceps'],
  ['tricep', 'Triceps'],
  ['forearm', 'Forearms'],
  ['core', 'Core']
];

function stripAccents_(s) {
  return s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function classifyMuscleWord_(word) {
  var w = stripAccents_(word).toLowerCase();
  for (var i = 0; i < FOCUS_KEYWORD_GROUPS.length; i++) {
    if (w.indexOf(FOCUS_KEYWORD_GROUPS[i][0]) !== -1) return FOCUS_KEYWORD_GROUPS[i][1];
  }
  return null;
}

function normalizeFocusGroup_(focus) {
  var trimmed = focus.toString().trim();
  if (!trimmed) return trimmed;

  // Only the first-listed muscle (before any ",", "&", "/", or " and ") counts
  var firstSegment = trimmed.split(/,|&|\/|\band\b/i)[0].trim();
  var words = firstSegment.split(/\s+/);

  for (var i = 0; i < words.length; i++) {
    var group = classifyMuscleWord_(words[i]);
    if (group) return group;
  }
  return trimmed; // unrecognized — keep as its own bucket under its original name
}

// Sums volume per Focus GROUP within an optional date window (sinceKey inclusive, 'yyyy-MM-dd' or null for all-time)
function aggregateVolumeByFocus_(data, tz, sinceKey) {
  var byFocus = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    if (!dateObj) return;
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    if (sinceKey && dateKey < sinceKey) return;

    var focusRaw = (row[5] || '').toString().trim();
    var focus = focusRaw ? normalizeFocusGroup_(focusRaw) : '(no focus)';
    var vol = computeVolume_(row[2], row[3], row[4], row[1]);
    byFocus[focus] = (byFocus[focus] || 0) + vol;
  });
  return Object.keys(byFocus)
    .map(function (f) { return [f, Math.round(byFocus[f] * 10) / 10]; })
    .sort(function (a, b) { return b[1] - a[1]; });
}

// Returns [{ focus, lastDateKey, daysSince }], sorted most-overdue first —
// shared by the live in-app panel, the Dashboard, and the weekly digest so
// all three always agree. daysSince is capped at FOCUS_GAP_CAP_DAYS: past
// that point the exact number isn't statistically meaningful and one wildly
// overdue muscle group (e.g. 81 days) would flatten every other bar on the chart.
var FOCUS_GAP_CAP_DAYS = 15;

function computeFocusGaps_(data, tz, today) {
  var lastTrainedByFocus = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    if (!dateObj) return;
    var focusRaw = (row[5] || '').toString().trim();
    if (!focusRaw) return;
    var focus = normalizeFocusGroup_(focusRaw);
    if (!lastTrainedByFocus[focus] || dateObj > lastTrainedByFocus[focus]) {
      lastTrainedByFocus[focus] = dateObj;
    }
  });

  return Object.keys(lastTrainedByFocus).map(function (focus) {
    var lastDate = lastTrainedByFocus[focus];
    var daysSince = Math.floor((today - lastDate) / (1000 * 60 * 60 * 24));
    return {
      focus: focus,
      lastDateKey: Utilities.formatDate(lastDate, tz, 'yyyy-MM-dd'),
      daysSince: Math.min(daysSince, FOCUS_GAP_CAP_DAYS)
    };
  }).sort(function (a, b) { return b.daysSince - a.daysSince; });
}

// Rest-day pattern stats: current no-training streak, average days between
// sessions historically, and which weekday gets trained least often
// (relative to how many of that weekday have occurred in your logging span).
function computeRestDayStats_(data, tz, today) {
  var trainingDatesSet = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    var exercise = (row[1] || '').toString().trim();
    if (!dateObj || !exercise) return;
    trainingDatesSet[Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd')] = true;
  });

  var trainingDateKeys = Object.keys(trainingDatesSet).sort();
  if (trainingDateKeys.length === 0) {
    return { longestStreakDays: null, avgDaysBetweenSessions: null, mostSkippedWeekday: null, weekdayBreakdown: [] };
  }

  var lastTrainingDate = toDateObj_(trainingDateKeys[trainingDateKeys.length - 1]);
  var longestStreakDays = Math.floor((today - lastTrainingDate) / (1000 * 60 * 60 * 24));

  var gaps = [];
  for (var i = 1; i < trainingDateKeys.length; i++) {
    var prev = toDateObj_(trainingDateKeys[i - 1]);
    var curr = toDateObj_(trainingDateKeys[i]);
    gaps.push(Math.round((curr - prev) / (1000 * 60 * 60 * 24)));
  }
  var avgDaysBetweenSessions = gaps.length > 0
    ? Math.round((gaps.reduce(function (a, b) { return a + b; }, 0) / gaps.length) * 10) / 10
    : null;

  // Weekday breakdown: what fraction of each weekday (within your logging
  // span) actually had a session, so "most skipped" accounts for how many
  // of that weekday have even occurred yet — not just a raw count.
  var weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var occurrenceCount = [0, 0, 0, 0, 0, 0, 0];
  var trainingCount = [0, 0, 0, 0, 0, 0, 0];

  var cursor = new Date(toDateObj_(trainingDateKeys[0]));
  while (cursor <= today) {
    occurrenceCount[cursor.getDay()]++;
    cursor.setDate(cursor.getDate() + 1);
  }
  trainingDateKeys.forEach(function (key) {
    trainingCount[toDateObj_(key).getDay()]++;
  });

  var weekdayBreakdown = weekdayNames.map(function (name, i) {
    var ratio = occurrenceCount[i] > 0 ? trainingCount[i] / occurrenceCount[i] : 0;
    return { day: name, trained: trainingCount[i], possible: occurrenceCount[i], pct: Math.round(ratio * 1000) / 10 };
  });
  var mostSkipped = weekdayBreakdown.slice().sort(function (a, b) { return a.pct - b.pct; })[0];

  return {
    longestStreakDays: longestStreakDays,
    avgDaysBetweenSessions: avgDaysBetweenSessions,
    mostSkippedWeekday: mostSkipped.day,
    weekdayBreakdown: weekdayBreakdown
  };
}

/**
 * Acute:Chronic Workload Ratio, rolling-average method:
 * acute = total volume in the last 7 days
 * chronic = average WEEKLY volume over the last 28 days (i.e. 28-day sum ÷ 4)
 * Returns null if there's not enough history yet (chronic would be 0).
 * "Sweet spot" is commonly cited as ~0.8–1.3, with >1.5 flagged as a real
 * spike — see the digest email for the caveats worth keeping in mind here.
 */
function computeACWR_(data, tz, today) {
  var volumeByDate = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    if (!dateObj) return;
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    volumeByDate[dateKey] = (volumeByDate[dateKey] || 0) + computeVolume_(row[2], row[3], row[4], row[1]);
  });

  var acute = 0, chronic28 = 0;
  for (var i = 0; i < 28; i++) {
    var d = new Date(today); d.setDate(d.getDate() - i);
    var key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    var vol = volumeByDate[key] || 0;
    chronic28 += vol;
    if (i < 7) acute += vol;
  }

  var chronicWeeklyAvg = chronic28 / 4;
  if (chronicWeeklyAvg <= 0) return null; // not enough history to mean anything yet

  return {
    acute: Math.round(acute * 10) / 10,
    chronicWeeklyAvg: Math.round(chronicWeeklyAvg * 10) / 10,
    ratio: Math.round((acute / chronicWeeklyAvg) * 100) / 100
  };
}

/**
 * Week-over-week per-exercise trend: for each exercise trained in BOTH the
 * current and prior 7-day windows, compares total volume and average
 * representative load, flagging increased/decreased/plateaued. A change
 * within +/-10% counts as "plateaued" — otherwise trivial rounding noise
 * (62.5kg -> 63kg) would get flagged as a real change. Exercises trained in
 * only one of the two weeks are skipped — nothing to compare.
 */
function computeProgressiveOverloadDelta_(data, tz, weekStartKey, todayKey, prevWeekStartKey, prevWeekEndKey) {
  var TOLERANCE = 0.10;

  function classify(current, prior) {
    if (prior <= 0) return null;
    var change = (current - prior) / prior;
    if (Math.abs(change) <= TOLERANCE) return 'plateaued';
    return change > 0 ? 'increased' : 'decreased';
  }

  var thisWeek = {}, lastWeek = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    var exercise = (row[1] || '').toString().trim();
    if (!dateObj || !exercise) return;
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    var vol = computeVolume_(row[2], row[3], row[4], row[1]);
    var load = computeRepresentativeLoad_(row[4]);

    var bucket = null;
    if (dateKey >= weekStartKey && dateKey <= todayKey) bucket = thisWeek;
    else if (dateKey >= prevWeekStartKey && dateKey <= prevWeekEndKey) bucket = lastWeek;
    if (!bucket) return;

    if (!bucket[exercise]) bucket[exercise] = { volume: 0, loadSum: 0, loadCount: 0 };
    bucket[exercise].volume += vol;
    if (load > 0) { bucket[exercise].loadSum += load; bucket[exercise].loadCount++; }
  });

  var results = [];
  Object.keys(thisWeek).forEach(function (exercise) {
    if (!lastWeek[exercise]) return; // only trained this week — nothing to compare

    var curr = thisWeek[exercise], prev = lastWeek[exercise];
    var currAvgLoad = curr.loadCount > 0 ? curr.loadSum / curr.loadCount : 0;
    var prevAvgLoad = prev.loadCount > 0 ? prev.loadSum / prev.loadCount : 0;

    results.push({
      exercise: exercise,
      volumeThisWeek: Math.round(curr.volume * 10) / 10,
      volumeLastWeek: Math.round(prev.volume * 10) / 10,
      volumeTrend: classify(curr.volume, prev.volume),
      avgLoadThisWeek: Math.round(currAvgLoad * 10) / 10,
      avgLoadLastWeek: Math.round(prevAvgLoad * 10) / 10,
      loadTrend: classify(currAvgLoad, prevAvgLoad)
    });
  });

  return results.sort(function (a, b) { return a.exercise.localeCompare(b.exercise); });
}

function updateDashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName(SHEET_NAME);
  if (!source || source.getLastRow() < 2) return;

  var tz = Session.getScriptTimeZone();
  var data = source.getRange(2, 1, source.getLastRow() - 1, 7).getValues();
  // columns: Date, Exercise, Sets, Reps, Load, Focus, Notes

  // Rolling windows, recalculated relative to today every time the dashboard rebuilds
  var today = new Date();
  var sevenDaysAgo = new Date(today); sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 6); // 7-day window incl. today
  var threeWeeksAgo = new Date(today); threeWeeksAgo.setDate(threeWeeksAgo.getDate() - 20); // 3-week window incl. today
  var threeMonthsAgo = new Date(today); threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
  var sevenDayKey = Utilities.formatDate(sevenDaysAgo, tz, 'yyyy-MM-dd');
  var threeWeekKey = Utilities.formatDate(threeWeeksAgo, tz, 'yyyy-MM-dd');
  var threeMonthKey = Utilities.formatDate(threeMonthsAgo, tz, 'yyyy-MM-dd');

  var volumeByDate = {};             // dateKey -> total volume
  var volumeByExercise = {};         // exercise -> total volume (all-time, used by Table 3)
  var entriesByExercise = {};        // exercise -> [{ dateObj, load, reps }] (all-time, used by Table 3)
  var volumeByExerciseWindow = {};   // exercise -> total volume, last 3 weeks only (used by Table 2 / bar chart)

  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    var exercise = (row[1] || '').toString().trim();
    if (!dateObj || !exercise) return;

    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    var vol = computeVolume_(row[2], row[3], row[4], row[1]);

    volumeByDate[dateKey] = (volumeByDate[dateKey] || 0) + vol;
    volumeByExercise[exercise] = (volumeByExercise[exercise] || 0) + vol;

    if (dateKey >= threeWeekKey) {
      volumeByExerciseWindow[exercise] = (volumeByExerciseWindow[exercise] || 0) + vol;
    }

    if (!entriesByExercise[exercise]) entriesByExercise[exercise] = [];
    var est1RM = computeBestEstimated1RM_(row[3], row[4], exercise);
    entriesByExercise[exercise].push({
      dateObj: dateObj,
      dateKey: dateKey,
      load: computeRepresentativeLoad_(row[4]),
      reps: computeRepresentativeReps_(row[3]),
      estEpley: est1RM.epley,
      estBrzycki: est1RM.brzycki,
      setType: classifySetType_(row[4])
    });
  });

  var dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!dashboard) {
    dashboard = ss.insertSheet(DASHBOARD_SHEET_NAME);
  } else {
    dashboard.clear();
    dashboard.getCharts().forEach(function (chart) { dashboard.removeChart(chart); });
  }

  // Table 1: Total volume by date, in chronological order (full history — data source, not directly charted)
  var dateRows = Object.keys(volumeByDate).sort().map(function (d) {
    return [d, Math.round(volumeByDate[d] * 10) / 10];
  });
  dashboard.getRange(1, 1, 1, 2).setValues([['Date', 'Total Volume (kg)']]).setFontWeight('bold');
  if (dateRows.length > 0) {
    dashboard.getRange(2, 1, dateRows.length, 2).setValues(dateRows);
    dashboard.getRange(2, 1, dateRows.length, 1).setNumberFormat('@'); // Date: plain text, guards against auto-date-conversion
  }

  var last7DaysRows = dateRows.filter(function (r) { return r[0] >= sevenDayKey; });
  var last3MonthsRows = dateRows.filter(function (r) { return r[0] >= threeMonthKey; });

  dashboard.getRange(1, 14, 1, 2).setValues([['Date', 'Volume (kg)']]).setFontWeight('bold'); // N1 — last 7 days data
  if (last7DaysRows.length > 0) {
    dashboard.getRange(2, 14, last7DaysRows.length, 2).setValues(last7DaysRows);
    dashboard.getRange(2, 14, last7DaysRows.length, 1).setNumberFormat('@');
  }
  dashboard.getRange(1, 17, 1, 2).setValues([['Date', 'Volume (kg)']]).setFontWeight('bold'); // Q1 — last 3 months data
  if (last3MonthsRows.length > 0) {
    dashboard.getRange(2, 17, last3MonthsRows.length, 2).setValues(last3MonthsRows);
    dashboard.getRange(2, 17, last3MonthsRows.length, 1).setNumberFormat('@');
  }

  // Muscle-group (Focus) balance, same rolling windows
  var focusLast7 = aggregateVolumeByFocus_(data, tz, sevenDayKey);
  var focusLast3Months = aggregateVolumeByFocus_(data, tz, threeMonthKey);

  dashboard.getRange(1, 20, 1, 2).setValues([['Focus', 'Volume (kg)']]).setFontWeight('bold'); // T1 — last 7 days by focus
  if (focusLast7.length > 0) {
    dashboard.getRange(2, 20, focusLast7.length, 2).setValues(focusLast7);
  }
  dashboard.getRange(1, 23, 1, 2).setValues([['Focus', 'Volume (kg)']]).setFontWeight('bold'); // W1 — last 3 months by focus
  if (focusLast3Months.length > 0) {
    dashboard.getRange(2, 23, focusLast3Months.length, 2).setValues(focusLast3Months);
  }

  // Table 2: Total volume by exercise, rolling last 3 weeks — top N by volume,
  // no minimum threshold (a 3-week window is already a strong enough filter
  // on its own; anything trained in that span at all is fair to show).
  var exerciseRows = Object.keys(volumeByExerciseWindow)
    .map(function (ex) { return [ex, Math.round(volumeByExerciseWindow[ex] * 10) / 10]; })
    .sort(function (a, b) { return b[1] - a[1]; })
    .slice(0, EXERCISE_CHART_MAX_COUNT);
  dashboard.getRange(1, 4, 1, 2).setValues([['Exercise', 'Volume (kg) — Last 3 Weeks']]).setFontWeight('bold');
  if (exerciseRows.length > 0) {
    dashboard.getRange(2, 4, exerciseRows.length, 2).setValues(exerciseRows);
  }

  // Table 3: Per-exercise summary — times performed, load/rep progress, first/last dates
  var summaryHeader = [
    'Exercise', 'Times Performed',
    'First Load (kg)', 'Last Load (kg)', 'Load Increase (kg)',
    'First Reps (avg/set)', 'Last Reps (avg/set)', 'Rep Increase',
    'First Performed', 'Last Performed',
    'Est. 1RM — Epley (kg)', 'Est. 1RM — Brzycki (kg)',
    'Most Common Set-Type'
  ];
  var summaryStartRow = dateRows.length + exerciseRows.length + 6; // clear of tables 1 & 2, whichever is longer
  var summaryRows = Object.keys(entriesByExercise).sort(function (a, b) { return a.localeCompare(b); }).map(function (exercise) {
    var entries = entriesByExercise[exercise].slice().sort(function (a, b) { return a.dateObj - b.dateObj; });
    var first = entries[0];
    var last = entries[entries.length - 1];
    var bestEpley = Math.max.apply(null, entries.map(function (e) { return e.estEpley; }));
    var bestBrzycki = Math.max.apply(null, entries.map(function (e) { return e.estBrzycki; }));
    return [
      exercise,
      entries.length,
      first.load, last.load, Math.round((last.load - first.load) * 10) / 10,
      first.reps, last.reps, Math.round((last.reps - first.reps) * 10) / 10,
      first.dateKey, last.dateKey,
      Math.round(bestEpley * 10) / 10, Math.round(bestBrzycki * 10) / 10,
      modeSetType_(entries)
    ];
  });
  dashboard.getRange(summaryStartRow, 1, 1, summaryHeader.length).setValues([summaryHeader]).setFontWeight('bold');
  if (summaryRows.length > 0) {
    var summaryDataRange = dashboard.getRange(summaryStartRow + 1, 1, summaryRows.length, summaryHeader.length);
    summaryDataRange.setValues(summaryRows);
    // Explicit formats guard against Sheets auto-detecting date-like strings
    // (First/Last Performed) as real dates, and that date formatting later
    // bleeding onto "Times Performed" if row positions shift between rebuilds
    // as your data grows. Hardcoded here so it can't drift again.
    dashboard.getRange(summaryStartRow + 1, 2, summaryRows.length, 1).setNumberFormat('0');       // Times Performed
    dashboard.getRange(summaryStartRow + 1, 9, summaryRows.length, 2).setNumberFormat('@');        // First/Last Performed (plain text, not real dates)
  }

  // Table 4: Muscle Group Gaps — same data/logic as the live in-app panel and the digest
  var focusGapsForSheet = computeFocusGaps_(data, tz, today);
  var focusGapsStartRow = summaryStartRow + summaryRows.length + 3;
  dashboard.getRange(focusGapsStartRow, 1, 1, 3)
    .setValues([['Muscle Group', 'Last Trained', 'Days Since']]).setFontWeight('bold');
  if (focusGapsForSheet.length > 0) {
    var focusGapsRows = focusGapsForSheet.map(function (g) { return [g.focus, g.lastDateKey, g.daysSince]; });
    dashboard.getRange(focusGapsStartRow + 1, 1, focusGapsRows.length, 3).setValues(focusGapsRows);
    dashboard.getRange(focusGapsStartRow + 1, 2, focusGapsRows.length, 1).setNumberFormat('@'); // Last Trained: plain text
    dashboard.getRange(focusGapsStartRow + 1, 3, focusGapsRows.length, 1).setNumberFormat('0');  // Days Since: plain number
  }

  dashboard.autoResizeColumns(1, summaryHeader.length);
  dashboard.autoResizeColumns(14, 11); // rolling-window date + focus-balance data columns
  dashboard.setFrozenRows(1);

  // Line chart: rolling 7-day volume trend
  if (last7DaysRows.length > 1) {
    var weekChart = dashboard.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dashboard.getRange(1, 14, last7DaysRows.length + 1, 2))
      .setPosition(1, 7, 0, 0)
      .setOption('title', 'Total Volume — Last 7 Days (rolling)')
      .setOption('width', 600)
      .setOption('height', 350)
      .build();
    dashboard.insertChart(weekChart);
  }

  // Line chart: rolling 3-month volume trend
  if (last3MonthsRows.length > 1) {
    var threeMonthChart = dashboard.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(dashboard.getRange(1, 17, last3MonthsRows.length + 1, 2))
      .setPosition(20, 7, 0, 0)
      .setOption('title', 'Total Volume — Last 3 Months (rolling)')
      .setOption('width', 600)
      .setOption('height', 350)
      .build();
    dashboard.insertChart(threeMonthChart);
  }

  // Bar chart: volume by exercise, all-time (top N passing the threshold)
  if (exerciseRows.length > 0) {
    var barChart = dashboard.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dashboard.getRange(1, 4, exerciseRows.length + 1, 2))
      .setPosition(40, 7, 0, 0)
      .setOption('title', 'Total Volume by Exercise (Last 3 Weeks, rolling)')
      .setOption('width', 600)
      .setOption('height', Math.max(350, exerciseRows.length * 22))
      .build();
    dashboard.insertChart(barChart);
  }

  // Pie chart: muscle-group (Focus) balance — last 7 days
  if (focusLast7.length > 0) {
    var focusWeekChart = dashboard.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dashboard.getRange(1, 20, focusLast7.length + 1, 2))
      .setPosition(60, 7, 0, 0)
      .setOption('title', 'Muscle-Group Balance — Last 7 Days')
      .setOption('width', 600)
      .setOption('height', 350)
      .setOption('pieSliceText', 'value-and-percentage') // show kg AND % on every slice, not just %
      .build();
    dashboard.insertChart(focusWeekChart);
  }

  // Pie chart: muscle-group (Focus) balance — last 3 months
  if (focusLast3Months.length > 0) {
    var focusQuarterChart = dashboard.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(dashboard.getRange(1, 23, focusLast3Months.length + 1, 2))
      .setPosition(80, 7, 0, 0)
      .setOption('title', 'Muscle-Group Balance — Last 3 Months')
      .setOption('width', 600)
      .setOption('height', 350)
      .setOption('pieSliceText', 'value-and-percentage')
      .build();
    dashboard.insertChart(focusQuarterChart);
  }

  // Bar chart: Muscle Group Gaps — days since last trained, most overdue first
  // (computeFocusGaps_ already returns it pre-sorted that way, so the tallest
  // bar is always the thing that most needs attention, top to bottom).
  if (focusGapsForSheet.length > 0) {
    var focusGapsChart = dashboard.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(dashboard.getRange(focusGapsStartRow, 1, focusGapsForSheet.length + 1, 1)) // names (skip Last Trained column)
      .addRange(dashboard.getRange(focusGapsStartRow, 3, focusGapsForSheet.length + 1, 1)) // days since
      .setPosition(100, 7, 0, 0)
      .setOption('title', 'Muscle Group Gaps (Days Since Last Trained)')
      .setOption('width', 600)
      .setOption('height', Math.max(300, focusGapsForSheet.length * 25))
      .build();
    dashboard.insertChart(focusGapsChart);
  }

  // Line chart: Workout Density over time (volume per minute)
  var densityData = computeWorkoutDensity_(tz);
  if (densityData.length > 0) {
    dashboard.getRange(1, 26, 1, 2).setValues([['Date', 'Density (kg/min)']]).setFontWeight('bold'); // Z1
    var densityRows = densityData.map(function (d) { return [d.dateKey, d.density]; });
    dashboard.getRange(2, 26, densityRows.length, 2).setValues(densityRows);
    dashboard.getRange(2, 26, densityRows.length, 1).setNumberFormat('@');
    dashboard.autoResizeColumns(26, 2);

    if (densityRows.length > 1) {
      var densityChart = dashboard.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(dashboard.getRange(1, 26, densityRows.length + 1, 2))
        .setPosition(120, 7, 0, 0)
        .setOption('title', 'Workout Density Over Time (kg/min)')
        .setOption('width', 600)
        .setOption('height', 350)
        .build();
      dashboard.insertChart(densityChart);
    }
  }
}

// --- Weekly Digest --------------------------------------------------------

// Installs a Monday-morning time trigger for sendWeeklyDigest(). Safe to run
// more than once — won't create a duplicate trigger.
function setUpWeeklyDigestTrigger() {
  var ui = SpreadsheetApp.getUi();
  var alreadyExists = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendWeeklyDigest';
  });
  if (alreadyExists) {
    ui.alert('A weekly digest trigger is already set up for Monday mornings.');
    return;
  }

  ScriptApp.newTrigger('sendWeeklyDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();

  ui.alert('Weekly digest scheduled for Monday mornings (Apps Script runs it sometime in the 8-9am window, not an exact minute).');
}

/**
 * Builds and sends the weekly training digest: this week's volume/exercises/
 * muscle-group split, which muscle groups have gone longest untrained,
 * new PRs this week, and any exercises that have plateaued.
 * "This week" = rolling last 7 days from whenever this runs.
 */
function sendWeeklyDigest() {
  var source = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!source || source.getLastRow() < 2) return;

  var tz = Session.getScriptTimeZone();
  var data = source.getRange(2, 1, source.getLastRow() - 1, 7).getValues();

  var today = new Date();
  var weekStart = new Date(today); weekStart.setDate(weekStart.getDate() - 6);
  var prevWeekEnd = new Date(weekStart); prevWeekEnd.setDate(prevWeekEnd.getDate() - 1);
  var prevWeekStart = new Date(prevWeekEnd); prevWeekStart.setDate(prevWeekStart.getDate() - 6);

  var weekStartKey = Utilities.formatDate(weekStart, tz, 'yyyy-MM-dd');
  var todayKey = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
  var prevWeekStartKey = Utilities.formatDate(prevWeekStart, tz, 'yyyy-MM-dd');
  var prevWeekEndKey = Utilities.formatDate(prevWeekEnd, tz, 'yyyy-MM-dd');

  // All-time per-exercise entries (chronological) — needed for PR detection
  // and plateau detection.
  var entriesByExercise = {};

  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    var exercise = (row[1] || '').toString().trim();
    if (!dateObj || !exercise) return;

    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    var load = computeRepresentativeLoad_(row[4]);
    var vol = computeVolume_(row[2], row[3], row[4], row[1]);

    if (!entriesByExercise[exercise]) entriesByExercise[exercise] = [];
    entriesByExercise[exercise].push({ dateObj: dateObj, dateKey: dateKey, load: load, vol: vol });
  });

  Object.keys(entriesByExercise).forEach(function (ex) {
    entriesByExercise[ex].sort(function (a, b) { return a.dateObj - b.dateObj; });
  });

  // --- This week's exercises + volume ---
  var thisWeekVolumeTotal = 0;
  var exerciseStatsThisWeek = {};
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    var exercise = (row[1] || '').toString().trim();
    if (!dateObj || !exercise) return;
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    if (dateKey < weekStartKey || dateKey > todayKey) return;

    var vol = computeVolume_(row[2], row[3], row[4], row[1]);
    thisWeekVolumeTotal += vol;
    if (!exerciseStatsThisWeek[exercise]) exerciseStatsThisWeek[exercise] = { times: 0, volume: 0 };
    exerciseStatsThisWeek[exercise].times++;
    exerciseStatsThisWeek[exercise].volume += vol;
  });
  var exerciseListThisWeek = Object.keys(exerciseStatsThisWeek)
    .map(function (ex) {
      return [ex, exerciseStatsThisWeek[ex].times, Math.round(exerciseStatsThisWeek[ex].volume * 10) / 10];
    })
    .sort(function (a, b) { return b[2] - a[2]; });

  // --- Volume vs last week ---
  var lastWeekVolumeTotal = 0;
  data.forEach(function (row) {
    var dateObj = toDateObj_(row[0]);
    if (!dateObj) return;
    var dateKey = Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd');
    if (dateKey < prevWeekStartKey || dateKey > prevWeekEndKey) return;
    lastWeekVolumeTotal += computeVolume_(row[2], row[3], row[4], row[1]);
  });
  var volumeDelta = thisWeekVolumeTotal - lastWeekVolumeTotal;
  var volumeDeltaPct = lastWeekVolumeTotal > 0 ? Math.round((volumeDelta / lastWeekVolumeTotal) * 1000) / 10 : null;

  // --- Muscle-group volume this week ---
  var focusThisWeek = aggregateVolumeByFocus_(data, tz, weekStartKey);

  // --- Muscle groups to prioritize next week: longest since last trained ---
  var focusGaps = computeFocusGaps_(data, tz, today);

  // --- Rest-day patterns: current streak, avg gap between sessions, most-skipped weekday ---
  var restDayStats = computeRestDayStats_(data, tz, today);

  // --- ACWR: acute (7-day) vs chronic (28-day weekly avg) load ---
  var acwr = computeACWR_(data, tz, today);

  // --- Progressive Overload Delta: week-over-week, exercises trained in both weeks only ---
  var overloadDelta = computeProgressiveOverloadDelta_(data, tz, weekStartKey, todayKey, prevWeekStartKey, prevWeekEndKey);

  // --- Workout Density: average this week, from the separate _Sessions tab ---
  var densityAll = computeWorkoutDensity_(tz);
  var densityThisWeek = densityAll.filter(function (d) { return d.dateKey >= weekStartKey && d.dateKey <= todayKey; });
  var avgDensityThisWeek = densityThisWeek.length > 0
    ? Math.round((densityThisWeek.reduce(function (s, d) { return s + d.density; }, 0) / densityThisWeek.length) * 100) / 100
    : null;

  // --- New PRs this week (must beat a REAL prior best — first-ever entries don't count) ---
  var prsThisWeek = [];
  Object.keys(entriesByExercise).forEach(function (ex) {
    var runningMax = null;
    entriesByExercise[ex].forEach(function (entry) {
      var inWindow = entry.dateKey >= weekStartKey && entry.dateKey <= todayKey;
      if (runningMax !== null && entry.load > runningMax && inWindow) {
        prsThisWeek.push({ exercise: ex, date: entry.dateKey, newLoad: entry.load, previousBest: runningMax });
      }
      if (runningMax === null || entry.load > runningMax) runningMax = entry.load;
    });
  });

  // --- Plateau alert: qualifying exercises whose all-time best hasn't been beaten recently ---
  var plateauAlerts = [];
  Object.keys(entriesByExercise).forEach(function (ex) {
    var entries = entriesByExercise[ex];
    var totalVolume = entries.reduce(function (sum, e) { return sum + e.vol; }, 0);
    var qualifies = entries.length >= EXERCISE_MIN_SESSIONS && totalVolume >= EXERCISE_MIN_VOLUME_KG;
    if (!qualifies || entries.length < PLATEAU_LOOKBACK_SESSIONS + 1) return;

    var allTimeMax = Math.max.apply(null, entries.map(function (e) { return e.load; }));
    var lastMaxIndex = -1;
    // Matching a PR is not beating it: count from its first achievement.
    entries.forEach(function (e, i) { if (lastMaxIndex === -1 && e.load === allTimeMax) lastMaxIndex = i; });
    var sessionsSincePR = entries.length - 1 - lastMaxIndex;

    if (sessionsSincePR >= PLATEAU_LOOKBACK_SESSIONS) {
      plateauAlerts.push({
        exercise: ex, bestLoad: allTimeMax, sessionsSincePR: sessionsSincePR,
        lastPRDate: entries[lastMaxIndex].dateKey
      });
    }
  });
  plateauAlerts.sort(function (a, b) { return b.sessionsSincePR - a.sessionsSincePR; });

  // --- Build and send the email ---
  var recipient = DIGEST_RECIPIENT_EMAIL || Session.getActiveUser().getEmail();
  if (!recipient) {
    Logger.log('sendWeeklyDigest: no recipient email available, aborting.');
    return;
  }

  var subject = 'Weekly Training Digest — ' + weekStartKey + ' to ' + todayKey;

  var html = '<div style="font-family:-apple-system,sans-serif;max-width:600px;color:#18181b;">';
  html += '<h2 style="margin-bottom:4px;">Weekly Training Digest</h2>';
  html += '<p style="color:#71717a;margin-top:0;">' + weekStartKey + ' → ' + todayKey + '</p>';

  html += '<h3>This Week</h3>';
  html += '<p>Total volume: <strong>' + (Math.round(thisWeekVolumeTotal * 10) / 10) + 'kg</strong>';
  if (volumeDeltaPct !== null) {
    html += ' (' + (volumeDelta >= 0 ? '+' : '') + (Math.round(volumeDelta * 10) / 10) + 'kg, ' +
      (volumeDeltaPct >= 0 ? '+' : '') + volumeDeltaPct + '% vs last week)';
  } else if (thisWeekVolumeTotal > 0) {
    html += ' (no training logged last week to compare against)';
  }
  html += '</p>';

  if (exerciseListThisWeek.length > 0) {
    html += '<table style="border-collapse:collapse;width:100%;font-size:14px;">';
    html += '<tr style="background:#f4f4f5;text-align:left;"><th style="padding:6px;">Exercise</th><th style="padding:6px;">Times</th><th style="padding:6px;">Volume (kg)</th></tr>';
    exerciseListThisWeek.forEach(function (row) {
      html += '<tr><td style="padding:6px;border-top:1px solid #e4e4e7;">' + row[0] + '</td>' +
        '<td style="padding:6px;border-top:1px solid #e4e4e7;">' + row[1] + '</td>' +
        '<td style="padding:6px;border-top:1px solid #e4e4e7;">' + row[2] + '</td></tr>';
    });
    html += '</table>';
  } else {
    html += '<p><em>No exercises logged this week.</em></p>';
  }

  if (focusThisWeek.length > 0) {
    html += '<h3>Muscle Group Volume This Week</h3><ul>';
    focusThisWeek.forEach(function (f) { html += '<li>' + f[0] + ': ' + f[1] + 'kg</li>'; });
    html += '</ul>';
  }

  html += '<h3>Prioritize Next Week</h3>';
  if (focusGaps.length > 0) {
    html += '<p>Muscle groups sorted by longest time since last trained:</p><ul>';
    focusGaps.slice(0, 5).forEach(function (g) {
      html += '<li><strong>' + g.focus + '</strong> — last trained ' + g.daysSince + ' day(s) ago (' + g.lastDateKey + ')</li>';
    });
    html += '</ul>';
  } else {
    html += '<p>No Focus data logged yet — tag your exercises to unlock this.</p>';
  }

  html += '<h3>Rest Day Patterns</h3>';
  if (restDayStats.longestStreakDays !== null) {
    html += '<ul>';
    html += '<li>Days since last session: <strong>' + restDayStats.longestStreakDays + '</strong></li>';
    html += '<li>Average gap between sessions (all-time): <strong>' + restDayStats.avgDaysBetweenSessions + ' days</strong></li>';
    html += '<li>Most often skipped weekday: <strong>' + restDayStats.mostSkippedWeekday + '</strong></li>';
    html += '</ul>';
  } else {
    html += '<p>Not enough data yet.</p>';
  }

  html += '<h3>Acute:Chronic Workload Ratio (ACWR)</h3>';
  if (acwr) {
    var acwrFlag = acwr.ratio > 1.5 ? ' — <strong style="color:#dc2626;">notably elevated</strong>'
      : (acwr.ratio > 1.3 ? ' — <strong style="color:#d97706;">a bit high, worth noting</strong>'
      : (acwr.ratio < 0.8 ? ' — on the low side (recent volume well below your baseline)' : ' — within the typical range'));
    html += '<p>Acute (last 7 days): <strong>' + acwr.acute + 'kg</strong> · Chronic (avg weekly, last 28 days): <strong>' + acwr.chronicWeeklyAvg + 'kg</strong></p>';
    html += '<p>Ratio: <strong>' + acwr.ratio + '</strong>' + acwrFlag + '</p>';
    html += '<p style="font-size:12px;color:#71717a;">Common reference range is ~0.8–1.3, with >1.5 flagged as a real spike — treat this as a load-monitoring heuristic, not a validated injury predictor (the research behind it is genuinely disputed, and it wasn\'t developed for solo strength training).</p>';
  } else {
    html += '<p>Not enough history yet (need at least a few weeks logged).</p>';
  }

  html += '<h3>Progressive Overload — Week over Week</h3>';
  if (overloadDelta.length > 0) {
    html += '<table style="border-collapse:collapse;width:100%;font-size:14px;">';
    html += '<tr style="background:#f4f4f5;text-align:left;"><th style="padding:6px;">Exercise</th><th style="padding:6px;">Volume</th><th style="padding:6px;">Avg Load</th></tr>';
    overloadDelta.forEach(function (o) {
      var volSymbol = o.volumeTrend === 'increased' ? '↑' : (o.volumeTrend === 'decreased' ? '↓' : '→');
      var loadSymbol = o.loadTrend === 'increased' ? '↑' : (o.loadTrend === 'decreased' ? '↓' : '→');
      html += '<tr><td style="padding:6px;border-top:1px solid #e4e4e7;">' + o.exercise + '</td>' +
        '<td style="padding:6px;border-top:1px solid #e4e4e7;">' + volSymbol + ' ' + o.volumeTrend + ' (' + o.volumeLastWeek + ' → ' + o.volumeThisWeek + 'kg)</td>' +
        '<td style="padding:6px;border-top:1px solid #e4e4e7;">' + loadSymbol + ' ' + o.loadTrend + ' (' + o.avgLoadLastWeek + ' → ' + o.avgLoadThisWeek + 'kg)</td></tr>';
    });
    html += '</table>';
  } else {
    html += '<p>No exercises trained in both this week and last week — nothing to compare yet.</p>';
  }

  html += '<h3>Workout Density</h3>';
  if (avgDensityThisWeek !== null) {
    html += '<p>Average this week: <strong>' + avgDensityThisWeek + ' kg/min</strong></p>';
  } else {
    html += '<p>No session timing recorded this week yet.</p>';
  }

  if (prsThisWeek.length > 0) {
    html += '<h3>🎉 New PRs This Week</h3><ul>';
    prsThisWeek.forEach(function (pr) {
      html += '<li><strong>' + pr.exercise + '</strong>: ' + pr.newLoad + 'kg (prev ' + pr.previousBest + 'kg) on ' + pr.date + '</li>';
    });
    html += '</ul>';
  }

  if (plateauAlerts.length > 0) {
    html += '<h3>⚠️ Plateau Alert</h3><ul>';
    plateauAlerts.forEach(function (p) {
      html += '<li><strong>' + p.exercise + '</strong>: stuck at ' + p.bestLoad + 'kg for ' + p.sessionsSincePR +
        ' sessions (last PR ' + p.lastPRDate + ')</li>';
    });
    html += '</ul>';
  }

  html += '</div>';

  var plainBody = 'Weekly Training Digest (' + weekStartKey + ' to ' + todayKey + ')\n\n' +
    'Total volume: ' + (Math.round(thisWeekVolumeTotal * 10) / 10) + 'kg\n' +
    'Open this email in HTML view for the full breakdown.';

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: plainBody,
    htmlBody: html
  });
}
