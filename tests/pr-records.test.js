// Run with: node --test tests/pr-records.test.js
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const { test } = require('node:test');

const root = path.join(__dirname, '..');
const backend = fs.readFileSync(path.join(root, 'Code.gs'), 'utf8');
const clientSource = fs.readFileSync(path.join(root, 'Index.html'), 'utf8')
  .match(/<script>([\s\S]*?)<\/script>/)[1];

function server(history = []) {
  const settings = [['Exercise', 'Bodyweight']];
  const settingsSheet = {
    getLastRow: () => settings.length,
    deleteRow: row => settings.splice(row - 1, 1),
    getRange(row, col, height, width) {
      return {
        getValues: () => settings.slice(row - 1, row - 1 + height).map(r => r.slice(col - 1, col - 1 + width)),
        setValues(values) { values.forEach((r, i) => { settings[row - 1 + i] = [...r]; }); }
      };
    }
  };
  const rows = [['Date', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes'], ...history];
  const sheet = {
    getMaxColumns: () => 8,
    getName: () => 'StrengthWorkoutINPUT',
    getLastRow: () => rows.length,
    getRange(row, col, height, width) {
      return {
        getValues: () => rows.slice(row - 1, row - 1 + height).map(r => r.slice(col - 1, col - 1 + width)),
        setValues(values) { values.forEach((r, i) => {
          if (!rows[row - 1 + i]) rows[row - 1 + i] = [];
          r.forEach((value, j) => { rows[row - 1 + i][col - 1 + j] = value; });
        }); }
      };
    }
  };
  const ctx = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'ExerciseSettings' ? settingsSheet : sheet }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) }
  });
  vm.runInContext(backend, ctx);
  ctx.recordSessionTiming_ = () => {};
  ctx.updateDashboard_ = () => {};
  ctx.settingsRows = settings;
  return ctx;
}

// Extract complete declarations, including nested helper functions.
function clientFunction(name) {
  const start = clientSource.indexOf('function ' + name + '(');
  assert.ok(start >= 0);
  let end = clientSource.indexOf('{', start) + 1, depth = 1;
  for (; depth; end++) {
    if (clientSource[end] === '{') depth++;
    else if (clientSource[end] === '}') depth--;
  }
  return clientSource.slice(start, end);
}

function client(records = {}) {
  const ctx = vm.createContext({ prRecordsByExercise: records, getEffectiveValues: card => card.values });
  ['getLoadRecordClient', 'isIsometricExercise', 'updatePRHint', 'formatPRMessage'].forEach(name => {
    vm.runInContext(clientFunction(name), ctx);
  });
  return ctx;
}

const plain = value => JSON.parse(JSON.stringify(value));
const row = (name, reps, load) => ['2026-09-24', name, 3, reps, load, '', ''];
const save = (ctx, exercises) => ctx.submitSession({ date: '2026-09-24', exercises: exercises.map(ex => ({ sets: 1, ...ex })) });

function brandSheet(name, rows, columns = 8) {
  return {
    getName: () => name, getLastRow: () => rows.length, getMaxColumns: () => columns,
    insertColumnsAfter(after, count) { columns += count; }, setFrozenRows() {},
    deleteRow(row) { rows.splice(row - 1, 1); },
    getRange(row, col, height = 1, width = 1) {
      const range = {
        getValues: () => Array.from({ length: height }, (_, i) => Array.from({ length: width }, (_, j) => (rows[row - 1 + i] || [])[col - 1 + j] ?? '')),
        setValues(values) {
          values.forEach((valuesRow, i) => {
            if (!rows[row - 1 + i]) rows[row - 1 + i] = [];
            valuesRow.forEach((value, j) => { rows[row - 1 + i][col - 1 + j] = value; });
          });
          return range;
        },
        setFontWeight() { return range; }
      };
      return range;
    }
  };
}

test('Brand setup appends headers, preserves history and templates, and is safe to repeat', () => {
  const ctx = server();
  const history = [['Date', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes'],
    row('Life Fitness Leg Press', '12', '65')];
  const templates = [['TemplateName', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes'],
    ['Leg day', 'Matrix Leg Press', 3, 12, 54, 'Legs', 'Original note']];
  const originalHistory = plain(history), originalTemplates = plain(templates);
  const sheets = { StrengthWorkoutINPUT: brandSheet('StrengthWorkoutINPUT', history, 7), Templates: brandSheet('Templates', templates, 7) };
  ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[name] }) };
  ctx.setupBrandSupport();
  ctx.setupBrandSupport();
  assert.deepEqual(history[1], originalHistory[1]);
  assert.deepEqual(templates[1], originalTemplates[1]);
  assert.equal(history[0][7], 'Brand');
  assert.equal(templates[0][7], 'Brand');
  assert.equal(ctx.getTemplateExercises('Leg day')[0].brand, '');
  ctx.saveTemplate('New leg day', [{ exercise: 'Leg Press', sets: 3, reps: 12, load: 0, brand: ' Matrix ' }]);
  assert.equal(ctx.getTemplateExercises('New leg day')[0].brand, 'Matrix');
  assert.equal(ctx.getTemplateExercises('New leg day')[0].load, 0);
});

test('Brand setup refuses to overwrite an occupied column H', () => {
  const ctx = server();
  for (const rows of [[['', '', '', '', '', '', '', 'Other data']], [[], ['', '', '', '', '', '', '', 'Unlabelled data']]]) {
    const before = plain(rows);
    assert.throws(() => ctx.ensureBrandColumn_(brandSheet('Workout', rows)), /already contains data/);
    assert.deepEqual(rows, before);
  }
});

test('session Brand is stored separately while PRs remain unified and old names stay intact', () => {
  const old = row('Life Fitness Leg Press', '12', '65');
  const ctx = server([old]);
  save(ctx, [{ exercise: 'Leg Press', reps: 12, load: 65, brand: 'Life Fitness' }]);
  const result = save(ctx, [{ exercise: 'Leg Press', reps: 12, load: 72, brand: ' Matrix ' }]);
  assert.equal(result.prs[0].exercise, 'Leg Press');
  assert.equal(result.prs[0].previousMax, 65);
  const sheet = ctx.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('StrengthWorkoutINPUT');
  const saved = sheet.getRange(2, 1, 3, 8).getValues();
  assert.deepEqual(saved[0], old);
  assert.equal(saved[1][7], 'Life Fitness');
  assert.equal(saved[2][7], 'Matrix');
  save(ctx, [{ exercise: 'Leg Press', reps: 8, load: 60 }]);
  assert.equal(sheet.getRange(5, 8, 1, 1).getValues()[0][0], '');
});

test('historical import carries explicit Brand and preserves old names without extracting Brand', () => {
  const ctx = server();
  const header = ['Date', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes'];
  const targetRows = [[...header]];
  const brandedRows = [[...header, 'Brand'], [...row('Leg Press', 12, 0), 'Matrix']];
  const legacyRows = [[...header], row('Life Fitness Leg Press', 12, 65)];
  const sheets = {
    StrengthWorkoutINPUT: brandSheet('StrengthWorkoutINPUT', targetRows),
    Branded: brandSheet('Branded', brandedRows),
    Legacy: brandSheet('Legacy', legacyRows, 7)
  };
  ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: name => sheets[name] }), getUi: () => ({ alert() {}, ButtonSet: { OK: 'OK' } }) };
  ctx.IMPORT_SOURCE_SHEETS = ['Branded', 'Legacy'];
  ctx.importHistoricalData();
  assert.equal(targetRows[1][4], 0);
  assert.equal(targetRows[1][7], 'Matrix');
  assert.equal(targetRows[2][1], 'Life Fitness Leg Press');
  assert.equal(targetRows[2][7], '');
  ctx.importHistoricalData();
  assert.equal(targetRows.length, 3);
});

test('Brand survives client collection and draft save/restore', () => {
  const fields = { '.f-exercise': 'Leg Press', '.f-brand': ' Matrix ', '.f-sets': '3', '.f-reps': '12', '.f-load': '72', '.f-focus-select': 'Legs', '.f-focus-other': '', '.f-notes': 'Test' };
  const card = { dataset: { mode: 'simple' }, querySelector: selector => ({ value: fields[selector] }) };
  const restored = [];
  const ctx = vm.createContext({
    sessionStartTime: new Date(),
    document: { querySelectorAll: () => [card], getElementById: () => ({ value: '', innerHTML: '' }) },
    getFocusValue: () => 'Legs', getEffectiveValues: () => ({ sets: 3, reps: '12', load: '72' }),
    todayISO: () => '2026-10-08', addExerciseBlock: values => restored.push(values)
  });
  ['serializeDraft', 'collectExercises', 'restoreDraftFromObject'].forEach(name => vm.runInContext(clientFunction(name), ctx));
  assert.equal(ctx.collectExercises()[0].brand, 'Matrix');
  const draft = ctx.serializeDraft();
  assert.equal(draft.exercises[0].brand, ' Matrix ');
  ctx.restoreDraftFromObject(draft);
  assert.equal(restored[0].brand, ' Matrix ');
  delete draft.exercises[0].brand;
  ctx.restoreDraftFromObject(draft);
  assert.equal(restored[1].brand, undefined);
});

test('muscle group gaps track combined focuses separately with or without spaces', () => {
  const ctx = server();
  ctx.Utilities = { formatDate: date => date.toISOString().slice(0, 10) };
  const history = [
    ['2026-10-01T12:00:00Z', 'Squat', 1, 8, 50, 'Legs,Glutes'],
    ['2026-10-05T12:00:00Z', 'Deadlift', 1, 8, 50, 'Legs,Glutes,Back'],
    ['2026-10-06T12:00:00Z', 'Press', 1, 8, 50, 'Chest, Triceps'],
    ['2026-10-04T12:00:00Z', 'Curl', 1, 8, 50, 'Biceps & Forearms'],
    ['2026-10-03T12:00:00Z', 'Other', 1, 8, 50, 'Stability,Core']
  ];
  const gaps = plain(ctx.computeFocusGaps_(history, 'UTC', new Date('2026-10-07T12:00:00Z')));
  const byFocus = Object.fromEntries(gaps.map(gap => [gap.focus, gap]));
  assert.deepEqual(Object.keys(byFocus).sort(), ['Back', 'Biceps', 'Chest', 'Core', 'Forearms', 'Glutes', 'Legs', 'Stability', 'Triceps']);
  ['Legs', 'Glutes', 'Back'].forEach(focus => {
    assert.equal(byFocus[focus].lastDateKey, '2026-10-05');
    assert.equal(byFocus[focus].daysSince, 2);
  });
  const spaced = history.map(row => row.map((value, index) => index === 5 ? value.replace(/,/g, ', ') : value));
  assert.deepEqual(plain(ctx.computeFocusGaps_(spaced, 'UTC', new Date('2026-10-07T12:00:00Z'))), gaps);
  assert.equal(ctx.normalizeFocusGroup_('Legs,Glutes,Back'), 'Legs');
  assert.equal(ctx.normalizeFocusGroup_('Stability,Core'), 'Stability');
  ctx.computeVolume_ = () => 100;
  assert.deepEqual(plain(ctx.aggregateVolumeByFocus_([history[1]], 'UTC', null)), [['Legs', 100]]);
});

test('focus normalization and catalog retain all historical associations and latest dates', () => {
  const ctx = server();
  assert.equal(ctx.normalizeFocus_(' chest, BICEPS ,chest,, upper back '), 'Chest,Biceps,Upper Back');
  const catalog = ctx.buildExerciseCatalog_([
    ['2025-01-01T12:00:00Z', ' Bench ', 1, 8, 50, 'chest'],
    ['2026-01-01T12:00:00Z', 'bench', 1, 8, 60, ' chest,TRICEPS'],
    ['invalid', 'Curl', 1, 8, 10, 'biceps']
  ]);
  assert.deepEqual(plain(catalog), [
    { exercise: 'Bench', focuses: ['Chest', 'Triceps'], lastLogged: '2026-01-01T12:00:00.000Z' },
    { exercise: 'Curl', focuses: ['Biceps'], lastLogged: '' }
  ]);
  save(ctx, [{ exercise: 'Bench', reps: 8, load: 50, focus: 'chest,TRICEPS,chest' }]);
  assert.equal(ctx.SpreadsheetApp.getActiveSpreadsheet().getSheetByName('StrengthWorkoutINPUT').getRange(2, 1, 1, 7).getValues()[0][5], 'Chest,Triceps');
});

test('catalog fetch reads the existing dashboard summary without workout history', () => {
  const ctx = server();
  const header = ['Exercise', 'Times Performed', '', '', '', '', '', '', '', 'Last Performed', '', '', '', 'Focus'];
  const entry = ['Bench', 2, '', '', '', '', '', '', '', '2026-01-01T12:00:00.000Z', '', '', '', 'chest,TRICEPS'];
  ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName(name) {
    assert.equal(name, 'Dashboard');
    return {
      getLastRow: () => 4,
      getRange: (row, col, height, width) => {
        assert.deepEqual([row, col, height, width], [1, 1, 4, 14]);
        return { getValues: () => [['Date', 'Volume'], header, entry, ['', '']] };
      }
    };
  } }) };
  assert.deepEqual(plain(ctx.getExerciseCatalog()), { maxAgeDays: 30, exercises: [
    { exercise: 'Bench', focuses: ['Chest', 'Triceps'], lastLogged: '2026-01-01T12:00:00.000Z' }
  ] });
});

test('focus suggestions filter by muscle group and recency independently of full matching', () => {
  const recent = new Date().toISOString();
  const ctx = vm.createContext({ suggestionMaxAgeDays: 30, exerciseCatalog: [
    { exercise: 'Bench', focuses: ['Chest', 'Triceps'], lastLogged: recent },
    { exercise: 'Curl', focuses: ['Biceps'], lastLogged: recent },
    { exercise: 'Old Press', focuses: ['Chest'], lastLogged: '2020-01-01' }
  ] });
  vm.runInContext(clientFunction('getExerciseSuggestions'), ctx);
  assert.deepEqual(plain(ctx.getExerciseSuggestions(' chest ')), ['Bench']);
  assert.deepEqual(plain(ctx.getExerciseSuggestions('Triceps')), ['Bench']);
  assert.deepEqual(plain(ctx.getExerciseSuggestions('Chest,Biceps')), ['Bench', 'Curl']);
  assert.deepEqual(plain(ctx.getExerciseSuggestions('')), ['Bench', 'Curl']);
  assert.deepEqual(plain(ctx.getExerciseSuggestions('Legs')), []);
});

test('dashboard rebuild adds Focus to its existing summary without creating a second catalog', () => {
  const ctx = server();
  vm.runInContext(backend, ctx);
  const data = [
    ['2020-01-01T12:00:00Z', 'Bench', 1, 8, 50, ' chest ', ''],
    ['2020-01-02T12:00:00Z', 'Bench', 1, 8, 60, 'TRICEPS,chest', '']
  ];
  const writes = [];
  const dashboard = {
    clear() {}, getCharts: () => [], autoResizeColumns() {}, setFrozenRows() {},
    getRange(row, col, height, width) {
      assert.ok(col < 31, 'must reuse existing summary, not write another catalog');
      const range = {
        setValues(values) { writes.push({ row, col, values: plain(values) }); return range; },
        setFontWeight() { return range; }, setNumberFormat() { return range; }
      };
      return range;
    }
  };
  const source = { getLastRow: () => 3, getRange: (row, col) => ({
    getValues: () => data,
    setValues(values) { assert.equal(col, 6); assert.deepEqual(plain(values), [['Chest'], ['Triceps,Chest']]); }
  }) };
  ctx.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getSheetByName: name => name === 'Dashboard' ? dashboard : source }) };
  ctx.Session = { getScriptTimeZone: () => 'UTC' };
  ctx.Utilities = { formatDate: date => date.toISOString().slice(0, 10) };
  ctx.applyBodyweightToData_ = rows => rows;
  ctx.computeVolume_ = () => 0;
  ctx.computeBestEstimated1RM_ = () => ({ epley: 0, brzycki: 0 });
  ctx.aggregateVolumeByFocus_ = () => [];
  ctx.computeFocusGaps_ = () => [];
  ctx.computeWorkoutDensity_ = () => [];
  ctx.updateDashboard_();
  const header = writes.find(write => write.values[0][1] === 'Times Performed');
  assert.equal(header.values[0][13], 'Focus');
  const summary = writes.find(write => write.row === header.row + 1 && write.col === 1);
  assert.equal(summary.values.length, 1);
  assert.equal(summary.values[0][13], 'Chest,Triceps');
  assert.equal(summary.values[0][9], '2020-01-02');
});

test('exercise suggestion cutoff filters recent names without changing all-time names', () => {
  const daysAgo = days => {
    const date = new Date();
    date.setDate(date.getDate() - days);
    return date.toISOString();
  };
  const ctx = server([
    [daysAgo(1), ' Squat '],
    [daysAgo(20), 'Squat'],
    [daysAgo(29), 'Bench'],
    [daysAgo(31), 'Deadlift'],
    ['invalid date', 'Old row'],
    [daysAgo(1), '']
  ]);
  assert.deepEqual(plain(ctx.getRecentExercises()), ['Bench', 'Squat']);
  assert.deepEqual(plain(ctx.getExistingExercises()), ['Bench', 'Deadlift', 'Old row', 'Squat']);
  ctx.EXERCISE_SUGGESTION_MAX_AGE_DAYS = 10;
  assert.deepEqual(plain(ctx.getRecentExercises()), ['Squat']);
  assert.deepEqual(plain(server().getRecentExercises()), []);
});

test('typo suggestions use recent names and suppress exact matches from all history', () => {
  const ctx = vm.createContext({
    exerciseOptions: ['Bench', 'Bunch', 'Deadlift'],
    recentExerciseOptions: ['Bench']
  });
  ['findCanonicalExercise', 'levenshtein', 'findCloseMatch'].forEach(name => {
    vm.runInContext(clientFunction(name), ctx);
  });
  assert.equal(ctx.findCloseMatch('bunch'), null);
  assert.equal(ctx.findCanonicalExercise('DEADLIFT'), 'Deadlift');
  assert.equal(ctx.findCloseMatch('Deadlft'), null);
  assert.equal(ctx.findCloseMatch('Benhc'), 'Bench');
});

test('bodyweight classification persists after negative history is removed', () => {
  const ctx = server();
  assert.equal(ctx.findBodyweightExercises_([row('Chin-up', '8', '-55')])['Chin-up'], true);
  assert.deepEqual(ctx.settingsRows, [['Exercise', 'Bodyweight'], ['Chin-up', true]]);
  assert.equal(ctx.findBodyweightExercises_([row('chin-up', '8', '0')])['chin-up'], true);
  assert.equal(ctx.settingsRows.length, 2);
  assert.equal(ctx.findBodyweightExercises_([])['Chin-up'], true);
});

test('editable settings override inference and support exercises without negative history', () => {
  const ctx = server();
  ctx.settingsRows.push(['Chin-up', false], ['Dips', 'TRUE']);
  ctx.BODYWEIGHT_EXERCISES = ['Chin-up'];
  const found = ctx.findBodyweightExercises_([row('CHIN-UP', '8', '55'), row('Dips', '8', '10')]);
  assert.equal(found['CHIN-UP'], undefined);
  assert.equal(found.Dips, true);
  assert.equal(ctx.settingsRows.length, 3);
  ctx.settingsRows[1][1] = true;
  assert.equal(ctx.findBodyweightExercises_([])['Chin-up'], true);
});

test('invalid and duplicate exercise settings fail visibly', () => {
  const ctx = server();
  ctx.settingsRows.push(['Dips', 'maybe']);
  assert.throws(() => ctx.findBodyweightExercises_([]), /must be TRUE or FALSE/);
  ctx.settingsRows[1][1] = true;
  ctx.settingsRows.push(['dips', false]);
  assert.throws(() => ctx.findBodyweightExercises_([]), /duplicate exercise/);
});

test('bodyweight accepts only positive decimals with at most two places', () => {
  const ctx = server();
  for (const value of ['100', '100.25', '100,25', 100.25]) {
    assert.equal(ctx.parseBodyweight_(value), Number(String(value).replace(',', '.')));
  }
  for (const value of ['80kg', '80.5.2', 'Infinity', '1e2', '100.123', '0', '-10']) {
    assert.throws(() => ctx.parseBodyweight_(value), /Bodyweight must/);
  }
  assert.equal(ctx.parseBodyweight_(''), null);
  assert.equal(ctx.parseBodyweight_(undefined), null);
});

test('effective loads include same-day weight, preserve raw rows, and classify history', () => {
  const ctx = server();
  ctx.getBodyweightSeries_ = () => [{ key: '2026-10-01', kg: 90 }, { key: '2026-10-02', kg: 110 }];
  const rows = [
    ['2026-10-02', 'Chin-up', 3, '8', '-55,-45,0', '', ''],
    ['2026-10-02', 'Chin-up', 1, '8', '10', '', ''],
    ['2026-10-02', 'Squat', 1, '8', '60', '', '']
  ];
  const result = ctx.applyBodyweightToData_(rows, 'America/Mexico_City');
  assert.equal(result[0][4], '45,55,100');
  assert.equal(result[1][4], '110');
  assert.equal(result[2][4], '60');
  assert.equal(rows[0][4], '-55,-45,0');
  assert.equal(ctx.computeVolume_(3, '8', result[0][4], 'Chin-up'), 1600);
  ctx.getBodyweightSeries_ = () => [];
  assert.equal(ctx.applyBodyweightToData_(rows, 'America/Mexico_City')[0][4], '');
});

test('smoothing uses the last five eligible weigh-ins and earliest fallback', () => {
  const ctx = server();
  const series = [90, 100, 101, 102, 103, 104, 200].map((kg, i) => ({ key: '2026-10-0' + (i + 1), kg }));
  assert.equal(ctx.bodyweightFor_(series, '2026-10-06'), 102);
  assert.equal(ctx.bodyweightFor_(series, '2026-09-01'), 90);
  assert.equal(ctx.bodyweightFor_([], '2026-10-06'), null);
});

test('invalid weight and weigh-in storage failures stop before appending workouts', () => {
  const ctx = server();
  const payload = { date: '2026-10-02', exercises: [{ exercise: 'Squat', sets: 1, reps: '8', load: '60' }] };
  assert.throws(() => ctx.submitSession({ ...payload, bodyweight: '100.123' }), /Bodyweight must/);
  ctx.saveBodyweight_ = () => { throw new Error('Weigh-in storage failed'); };
  assert.throws(() => ctx.submitSession({ ...payload, bodyweight: '100.25' }), /Weigh-in storage failed/);
  assert.deepEqual(plain(ctx.getPRRecords()), {});
  assert.throws(() => ctx.submitSession({ date: payload.date, exercises: [], bodyweight: '100' }), /at least one exercise/);
});

test('bodyweight exercise intensity is suppressed even with positive added-load history', () => {
  const ctx = server([row('Chin-up', '8', '-55'), row('Chin-up', '8', '10'), row('Squat', '8', '60')]);
  const estimates = ctx.getMaxEstimated1RMs();
  assert.equal(estimates['Chin-up'].bodyweight, true);
  assert.equal(estimates['Chin-up'].epley, 0);
  assert.ok(estimates.Squat.epley > 0);
  const front = vm.createContext({
    getEffectiveValues: () => ({ reps: '8', load: '10' }),
    getTopSetClient: () => ({ reps: 8, load: 10 }),
    isIsometricExercise: () => false,
    findCanonicalExercise: () => 'Chin-up',
    maxEst1RMByExercise: { 'Chin-up': { epley: 20, brzycki: 20, bodyweight: true } }
  });
  vm.runInContext(clientFunction('updateEst1RMHint'), front);
  const hint = { style: {} };
  front.updateEst1RMHint({ querySelector: s => s === '.est1rm-hint' ? hint : { value: 'Chin-up' } });
  assert.equal(hint.style.display, 'none');
});

test('ISO workout dates are parsed in script time zone rather than UTC', () => {
  const ctx = server();
  ctx.Session = { getScriptTimeZone: () => 'America/Mexico_City' };
  ctx.Utilities = { parseDate(value, tz, format) {
    assert.equal(value, '2026-10-02 12:00');
    assert.equal(tz, 'America/Mexico_City');
    assert.equal(format, 'yyyy-MM-dd HH:mm');
    return new Date('2026-10-02T18:00:00Z');
  } };
  assert.equal(ctx.toDateObj_('2026-10-02').toISOString(), '2026-10-02T18:00:00.000Z');
});

test('both scripts parse', () => {
  new vm.Script(backend);
  new vm.Script(clientSource);
});

test('assisted load and rep PRs work on server and client through zero and added weight', () => {
  const ctx = server([row('Chin-up', '8', '-55')]);
  const front = client({ 'Chin-up': { maxLoad: -55, bestReps: 8 } });
  assert.deepEqual(plain(front.getLoadRecordClient('8', '-45')), { maxLoad: -45, bestReps: 8 });
  assert.deepEqual(plain(ctx.getLoadRecord_('-2,8', '-45,-55')), { maxLoad: -45, bestReps: null });
  const hint = { style: {}, textContent: '' };
  front.updatePRHint({ values: { reps: '8', load: '-45' }, querySelector: s => s === '.pr-hint' ? hint : { value: 'Chin-up' } });
  assert.match(hint.textContent, /beats your PR of -55kg/);
  const load = save(ctx, [{ exercise: 'Chin-up', reps: 8, load: -45 }]).prs[0];
  assert.equal(load.previousMax, -55);
  assert.equal(load.newMax, -45);
  assert.equal(load.type, 'load');
  assert.equal(save(ctx, [{ exercise: 'Chin-up', reps: 9, load: -45 }]).prs[0].type, 'reps');
  assert.equal(save(ctx, [{ exercise: 'Chin-up', reps: 20, load: -55 }]).prs.length, 0);
  assert.equal(save(ctx, [{ exercise: 'Chin-up', reps: 5, load: 0 }]).prs[0].previousMax, -45);
  assert.equal(save(ctx, [{ exercise: 'Chin-up', reps: 5, load: 10 }]).prs[0].previousMax, 0);
});

test('FALSE/negative conflicts stop submission before weigh-in and workout writes', () => {
  const ctx = server();
  ctx.settingsRows.push(['Chin-up', 'false']);
  ctx.saveBodyweight_ = () => assert.fail('must not write bodyweight');
  assert.throws(() => ctx.submitSession({ date: '2026-10-07', bodyweight: '100', exercises: [
    { exercise: ' CHIN-UP ', sets: 1, reps: 8, load: -55 }
  ] }), /ExerciseSettings.*CHIN-UP.*FALSE.*negative/);
  assert.deepEqual(plain(ctx.getPRRecords()), {});
  assert.throws(() => ctx.applyBodyweightToData_([row('Chin-up', '8', '-55')], 'America/Mexico_City'), /FALSE/);
  ctx.settingsRows.push(['Unrelated', 'invalid']);
  assert.equal(save(ctx, [{ exercise: 'Squat', reps: 8, load: 60 }]).success, true);
});

test('list lengths allow scalar broadcast and report exercise and missing set', () => {
  const back = server(), front = vm.createContext({});
  vm.runInContext(clientFunction('validateWorkoutExercise'), front);
  for (const validate of [back.validateWorkoutExercise_, front.validateWorkoutExercise]) {
    assert.equal(validate({ exercise: 'Squat', sets: 3, reps: '8', load: '60,65,70' }), '');
    assert.equal(validate({ exercise: 'Squat', sets: 3, reps: '8,6,4', load: '60' }), '');
    assert.match(validate({ exercise: 'Squat', sets: 1, reps: '8', load: '62,5' }), /exactly 1.*decimal point/);
    assert.match(validate({ exercise: 'Squat', sets: 3, reps: '8,6', load: '60' }), /exactly 3/);
    assert.match(validate({ exercise: 'Squat', sets: 3, reps: '8,,4', load: '60' }), /Squat, set 2:.*reps/);
    assert.match(validate({ exercise: 'Squat', sets: 3, reps: '8', load: '60,65,' }), /Squat, set 3:.*load/);
  }
});

test('settings merge transfers source, deduplicates matches, and preserves unrelated entries', () => {
  const ctx = server();
  ctx.settingsRows.push(['Pullup', true], ['Squat', false]);
  ctx.mergeExerciseSettings_('Pullup', 'Chin-up', []);
  assert.deepEqual(ctx.settingsRows.slice(1), [['Chin-up', true], ['Squat', false]]);
  ctx.settingsRows.push(['Pullup', true]);
  ctx.mergeExerciseSettings_('Pullup', 'Chin-up', []);
  assert.deepEqual(ctx.settingsRows.slice(1), [['Chin-up', true], ['Squat', false]]);
  ctx.mergeExerciseSettings_('Chin-up', 'CHIN-UP', []);
  assert.deepEqual(ctx.settingsRows.slice(1), [['CHIN-UP', true], ['Squat', false]]);
});

test('settings merge rejects conflicts without mutation and preserves inferred classification', () => {
  const ctx = server();
  ctx.settingsRows.push(['Pullup', true], ['Chin-up', false]);
  const before = JSON.stringify(ctx.settingsRows);
  assert.throws(() => ctx.mergeExerciseSettings_('Pullup', 'Chin-up', []), /conflicting Bodyweight/);
  assert.equal(JSON.stringify(ctx.settingsRows), before);
  ctx.settingsRows.splice(1);
  ctx.mergeExerciseSettings_('Pullup', 'Chin-up', [row('Pullup', '8', '-55')]);
  assert.deepEqual(ctx.settingsRows.slice(1), [['Chin-up', true]]);
  ctx.settingsRows[1][1] = false;
  assert.throws(() => ctx.mergeExerciseSettings_('Pullup', 'Chin-up', [row('Pullup', '8', '-55')]), /conflicting Bodyweight/);
});

test('merge menu stops before renaming workouts or templates when settings conflict', () => {
  const ctx = server([row('Pullup', '8', '0'), row('Chin-up', '8', '10')]);
  ctx.settingsRows.push(['Pullup', true], ['Chin-up', false]);
  const responses = ['Pullup', 'Chin-up'], alerts = [];
  ctx.SpreadsheetApp.getUi = () => ({
    Button: { OK: 'OK', YES: 'YES' }, ButtonSet: { OK_CANCEL: 'OK_CANCEL', YES_NO: 'YES_NO', OK: 'OK' },
    prompt() { const text = responses.shift(); return { getSelectedButton: () => 'OK', getResponseText: () => text }; },
    alert(...args) { alerts.push(args); return 'YES'; }
  });
  ctx.renameExerciseInSheet_ = () => assert.fail('must not rename before settings are resolved');
  ctx.mergeExerciseNames();
  assert.match(alerts.at(-1)[0], /Merge cancelled:.*conflicting Bodyweight/);
  assert.deepEqual(ctx.settingsRows.slice(1), [['Pullup', true], ['Chin-up', false]]);
});

test('client and server require explicit reps and loads, while accepting signed loads and zero', () => {
  const ctx = server(), front = vm.createContext({});
  vm.runInContext(clientFunction('validateWorkoutExercise'), front);
  const valid = { exercise: 'Chin-up', sets: 1, reps: '8', load: 0 };
  const cases = [
    [valid, true],
    [{ ...valid, sets: 3, reps: '8,6,4', load: '-55,0,10.5' }, true],
    [{ ...valid, sets: '' }, false],
    [{ ...valid, sets: 0 }, false],
    [{ ...valid, sets: 1.5 }, false],
    [{ ...valid, reps: '' }, false],
    [{ ...valid, reps: '8,,4' }, false],
    [{ ...valid, reps: 0 }, false],
    [{ ...valid, reps: 'failure' }, false],
    [{ ...valid, load: '' }, false],
    [{ ...valid, load: '-55,,0' }, false],
    [{ ...valid, load: '55kg' }, false],
    [{ ...valid, load: 'Infinity' }, false]
  ];
  for (const [ex, accepted] of cases) {
    const message = ctx.validateWorkoutExercise_(ex);
    assert.equal(message === '', accepted);
    assert.equal(front.validateWorkoutExercise(ex), message);
  }
  assert.throws(() => save(ctx, [{ ...valid, load: '' }]), /numeric load/);
  assert.deepEqual(plain(ctx.getPRRecords()), {});
  save(ctx, [valid]);
  assert.equal(ctx.getPRRecords()['Chin-up'].maxLoad, 0);
});

test('per-set reading preserves incomplete and entirely blank rows for validation', () => {
  const ctx = vm.createContext({});
  vm.runInContext(clientFunction('readPerSetValues'), ctx);
  const inputs = [['8', '-55'], ['6', ''], ['', '']];
  const card = { querySelectorAll: () => inputs.map(([reps, load]) => ({
    querySelector: selector => ({ value: selector === '.set-reps' ? reps : load })
  })) };
  assert.deepEqual(plain(ctx.readPerSetValues(card)), { sets: 3, reps: '8,6,', load: '-55,,' });
});

test('templates preserve numeric zero loads when saved and loaded', () => {
  const ctx = server();
  const rows = [['Template', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes']];
  ctx.getTemplatesSheet_ = () => ({
    getLastRow: () => rows.length,
    getRange(row, col, height, width) {
      return {
        getValues: () => rows.slice(row - 1, row - 1 + height).map(r => r.slice(col - 1, col - 1 + width)),
        setValues(values) { values.forEach((r, i) => {
          if (!rows[row - 1 + i]) rows[row - 1 + i] = [];
          r.forEach((value, j) => { rows[row - 1 + i][col - 1 + j] = value; });
        }); }
      };
    }
  });
  ctx.saveTemplate('Bodyweight', [{ exercise: 'Chin-up', sets: 1, reps: 8, load: 0 }]);
  assert.equal(ctx.getTemplateExercises('Bodyweight')[0].load, 0);
});

test('client and server pair reps with loads, preserve missing slots, and handle zero', () => {
  const back = server(), front = client();
  const cases = [
    ['12,5,8', '60,80,80', { maxLoad: 80, bestReps: 8 }],
    ['12,6', '60,80', { maxLoad: 80, bestReps: 6 }],
    ['8', '60,70,80', { maxLoad: 80, bestReps: 8 }],
    ['8,12,10', '80', { maxLoad: 80, bestReps: 12 }],
    ['30,45', '60,60', { maxLoad: 60, bestReps: 45 }],
    ['', '80', { maxLoad: 80, bestReps: null }],
    ['failure,12', '80,60', { maxLoad: 80, bestReps: null }],
    [',12', '80,60', { maxLoad: 80, bestReps: null }],
    ['8,12', ',60', { maxLoad: 60, bestReps: 12 }],
    [8, 0, { maxLoad: 0, bestReps: 8 }],
    ['8', '', null]
  ];
  cases.forEach(([reps, load, expected]) => {
    assert.deepEqual(plain(back.getLoadRecord_(reps, load)), expected);
    assert.deepEqual(plain(front.getLoadRecordClient(reps, load)), expected);
  });
});

test('history uses only reps at maximum load, independent of row order', () => {
  const history = [row('Squat', 20, 60), row('Squat', 5, 80), row('Squat', 8, 80), row('Squat', 30, 70)];
  for (const rows of [history, [...history].reverse()]) {
    assert.deepEqual(plain(server(rows).getPRRecords()), { Squat: { maxLoad: 80, bestReps: 8 } });
  }
  assert.deepEqual(plain(server([...history, row('Squat', '', 90)]).getPRRecords()),
    { Squat: { maxLoad: 90, bestReps: null } });
});

test('save detects rep records at the load PR, but ignores lighter sets and ties', () => {
  const ctx = server([row('Squat', 8, 80)]);
  assert.equal(save(ctx, [{ exercise: 'Squat', reps: 30, load: 60 }]).prs.length, 0);
  assert.equal(save(ctx, [{ exercise: 'Squat', reps: 8, load: 80 }]).prs.length, 0);
  const result = save(ctx, [{ exercise: 'squat', reps: 10, load: 80 }]);
  assert.equal(result.prs[0].type, 'reps');
  assert.equal(result.prs[0].previousReps, 8);
  assert.equal(result.prs[0].bestReps, 10);
  assert.equal(result.prs[0].exercise, 'Squat');
  assert.equal(ctx.getPRRecords().Squat.bestReps, 10);
});

test('a heavier load resets the reps baseline; duplicate cards produce one record', () => {
  const ctx = server([row('Squat', 8, 80)]);
  const result = save(ctx, [
    { exercise: 'Squat', reps: 12, load: 80 },
    { exercise: 'Squat', reps: 3, load: 90 },
    { exercise: 'Squat', reps: 5, load: 90 }
  ]);
  assert.equal(result.prs.length, 1);
  assert.equal(result.prs[0].type, 'load');
  assert.equal(result.prs[0].bestReps, 5);
  assert.deepEqual(plain(ctx.getPRRecords().Squat), { maxLoad: 90, bestReps: 5 });
});

test('isometric records use seconds and rep-only saves announce hold PRs', () => {
  const ctx = server([row('Iso squat', '30,45', '60,60')]);
  const result = save(ctx, [{ exercise: 'Iso squat', sets: 2, reps: '90,50', load: '40,60' }]);
  assert.equal(result.prs[0].bestReps, 50);
  assert.equal(result.prs[0].previousReps, 45);
  assert.equal(result.prs[0].isIsometric, true);
  assert.match(client().formatPRMessage(result.prs[0]), /hold PR.*50 sec.*60kg.*45 sec/);
});

test('live hints gate reps by load and label isometric holds in seconds', () => {
  const ctx = client({ Squat: { maxLoad: 80, bestReps: 8 }, 'Iso squat': { maxLoad: 60, bestReps: 45 } });
  function hint(name, reps, load) {
    const el = { style: {}, textContent: '' };
    ctx.updatePRHint({ values: { reps, load }, querySelector: s => s === '.pr-hint' ? el : { value: name } });
    return el;
  }
  assert.equal(hint('Squat', '30', '60').textContent, 'Current PR: 80kg');
  assert.match(hint('squat', '', '80').textContent, /Most reps at max-ever load: 8 reps at 80kg/);
  assert.match(hint('Squat', '10', '80').textContent, /New rep PR: 10 reps/);
  assert.doesNotMatch(hint('Squat', '8', '80').textContent, /New rep PR/);
  assert.match(hint('Squat', '3', '90').textContent, /New max-load baseline: 3 reps at 90kg/);
  assert.match(hint('Iso squat', '50', '60').textContent, /Longest hold at max-ever load: 45 sec.*New hold PR: 50 sec/);
  assert.equal(hint('Iso squat', '100', '40').textContent, 'Current PR: 60kg');
  assert.equal(hint('Squat', '10', '').style.display, 'none');
});

test('first records and missing historical reps produce usable confirmations', () => {
  const ctx = server([row('Squat', '', 80)]), front = client();
  const firstReps = save(ctx, [{ exercise: 'Squat', reps: 5, load: 80 }]).prs[0];
  assert.match(front.formatPRMessage(firstReps), /5 reps.*first recorded at this load/);
  const firstLoad = save(ctx, [{ exercise: 'Curl', reps: 12, load: 10 }]).prs[0];
  assert.match(front.formatPRMessage(firstLoad), /New load PR.*12 reps.*first time logged/);
});
