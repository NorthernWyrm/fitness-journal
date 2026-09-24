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
  const rows = [['Date', 'Exercise', 'Sets', 'Reps', 'Load', 'Focus', 'Notes'], ...history];
  const sheet = {
    getLastRow: () => rows.length,
    getRange(row, col, height, width) {
      return {
        getValues: () => rows.slice(row - 1, row - 1 + height).map(r => r.slice(col - 1, col - 1 + width)),
        setValues(values) { values.forEach((r, i) => { rows[row - 1 + i] = [...r]; }); }
      };
    }
  };
  const ctx = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }) }
  });
  vm.runInContext(backend, ctx);
  ctx.recordSessionTiming_ = () => {};
  ctx.updateDashboard_ = () => {};
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
const save = (ctx, exercises) => ctx.submitSession({ date: '2026-09-24', exercises });

test('both scripts parse', () => {
  new vm.Script(backend);
  new vm.Script(clientSource);
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
  const result = save(ctx, [{ exercise: 'Iso squat', reps: '90,50', load: '40,60' }]);
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
