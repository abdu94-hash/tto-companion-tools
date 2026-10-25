#!/usr/bin/env node
/*
 * TTO Companion Tools v1.1 - automated verification of the calculator engine against CALC-SPEC 1.1
 * vectors V01-V38 (tests/tto_calc_test_vectors_v1.1.json; V01-V27 are the unchanged CALC-SPEC 1.0 vectors).
 *
 * Usage (from the repository root, Node.js 18 or later, no packages needed):
 *   node tests/run_vectors.js
 *   node tests/run_vectors.js --python <path/to/tto_calc_reference_check_v1.1.py>   (optional cross-check)
 *
 * Exits with a non-zero status if any numeric vector fails, if docs/vectors.js differs from the JSON,
 * or if the optional Python cross-check disagrees.
 */
'use strict';
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

const ROOT = path.join(__dirname, '..');
const E = require(path.join(ROOT, 'docs', 'tto_calc_engine.js'));
const V = require(path.join(ROOT, 'docs', 'tto_verify.js'));
const JSON_PATH = path.join(__dirname, 'tto_calc_test_vectors_v1.1.json');
const data = JSON.parse(fs.readFileSync(JSON_PATH, 'utf8'));

let exitCode = 0;
const out = (s) => process.stdout.write(s + '\n');

out(E.VERSION_LABEL + ' - verification vectors ' + data.spec + ' (' + path.basename(JSON_PATH) + ')');
out('Tolerance: relative ' + data.tolerance.relative + ', absolute ' + data.tolerance.absolute_for_zero + ' when the expected value is zero; units, sequence and flags exact.\n');

// 1. every vector
const run = V.runAll(E, data);
for (const r of run.results) {
  if (r.type === 'behavioral') {
    out(r.id.padEnd(4) + ' BEHAVIORAL  ' + r.status + '  ' + r.rule + (r.notes && r.notes.length ? '\n       ' + r.notes.join('\n       ') : ''));
  } else {
    out(r.id.padEnd(4) + ' ' + r.status.padEnd(10) + ' ' + r.rule + r.problems.map((p) => '\n       ' + p).join(''));
  }
}
out('\nNumeric vectors: ' + run.passed + '/' + run.numeric + ' PASS, ' + run.failed + ' FAIL. Behavioral: ' +
  run.behavioral.map((b) => b.id + ' ' + b.status).join(', ') + '.');
if (run.failed > 0 || run.numeric !== 36) exitCode = 1;
if (run.behavioral.some((b) => b.status !== 'DEMONSTRATED')) exitCode = 1;

// 2. the browser copy of the vectors must equal the JSON
const vjs = require(path.join(ROOT, 'docs', 'vectors.js'));
const same = JSON.stringify(vjs) === JSON.stringify(data);
out('docs/vectors.js matches the JSON vector file: ' + (same ? 'yes' : 'NO - run node tests/build_vectors_js.js'));
if (!same) exitCode = 1;

// 3. the checker must fail when a wrong value is planted (guards against a runner that always passes)
const mutated = JSON.parse(JSON.stringify(data));
const planted = [['V16', (v) => { v.expected.fraction_remaining = 0.1251; }], ['V06', (v) => { v.expected.flags = []; }],
  ['V10', (v) => { v.expected.unit = 'minutes'; }], ['V27', (v) => { v.expected.results_ml_min[0] = 68.06; }],
  ['V28', (v) => { v.expected.value = 20; }], ['V29', (v) => { v.expected.flags = []; }], ['V30', (v) => { v.expected.interventions_counted = ['TX-START']; }],
  ['V31', (v) => { v.expected.minimum = 15; }], ['V32', (v) => { v.expected.fraction_of_change_completed = 0.8751; }],
  ['V33', (v) => { v.expected.value = 1020; }], ['V34', (v) => { v.expected.flags = []; }],
  ['V35', (v) => { v.expected.flags[0] = ['E_NONPOSITIVE_PARAMETER', 'E_INVALID_UNIT', 'E_APPLICABILITY_UNCONFIRMED']; }],
  ['V36', (v) => { v.expected.flags[0] = ['E_MISSING_INPUT']; }], ['V37', (v) => { v.expected.flags = ['I_SCENARIO_NOT_CI']; }],
  ['V38', (v) => { v.expected.single_bsa_first_mg_m2 = 170.588235; }]];
let caught = 0;
for (const [id, f] of planted) {
  const v = mutated.vectors.find((x) => x.id === id); f(v);
  if (V.runVector(E, v, data.tolerance).status === 'FAIL') caught++;
}
out('Planted wrong values detected: ' + caught + '/' + planted.length);
if (caught !== planted.length) exitCode = 1;

// 4. V24 description: changing a parameter's status changes the Parameter_Set shown in the output
const q1Entry = (status) => ({ id: 'T1', rule: 'Q1', inputs: { t_half: '24', t: '72', unit: 'hours' },
  meta: { t_half: { status: status, source: 'Test source', version: 'v1', section: '12.3' }, t: { status: 'hypothetical example' } },
  gate: { model: true, half_life: true, half_life_type: 'terminal', half_life_population: 'adults', half_life_source: 'Test source', reference: true, reference_text: 'end of dosing', status: true } });
const ps1 = E.evaluateTechnical(q1Entry('label fact'))[0], ps2 = E.evaluateTechnical(q1Entry('model assumption'))[0];
const v24ok = ps1.Parameter_Set[0].Parameter_Status === 'label fact' && ps2.Parameter_Set[0].Parameter_Status === 'model assumption' &&
  JSON.stringify(ps1.Parameter_Set) !== JSON.stringify(ps2.Parameter_Set) && ps1.Results[0].value === 0.125 && !ps1.Blocked;
out('V24 Parameter_Set reflects a change of Parameter_Status: ' + (v24ok ? 'yes' : 'NO'));
if (!v24ok) exitCode = 1;

// 5. CSV formula-injection protection keeps negative numbers numeric and neutralizes formulas
const csv = E.toCSV([{ Result_ID: '=HYPERLINK("x")', Label: '+1+1', Interval_Minimum: -5, QC_Flags: [], Messages: [], Assumptions: ['@SUM(A1)'] }]);
const back = E.recordsFromCSV(csv)[0];
const csvOk = /'=HYPERLINK/.test(csv) && /'\+1\+1/.test(csv) && /'@SUM/.test(csv) && /,-5,/.test(csv) &&
  back.Result_ID === '=HYPERLINK("x")' && back.Interval_Minimum === -5;
out('CSV formula-injection protection and round trip: ' + (csvOk ? 'yes' : 'NO'));
if (!csvOk) exitCode = 1;

// 6. optional cross-check with the independent Python reference implementation
function findReference() {
  const i = process.argv.indexOf('--python');
  if (i > 0 && process.argv[i + 1]) return process.argv[i + 1];
  if (process.env.TTO_REFERENCE) return process.env.TTO_REFERENCE;
  const guess = path.join(ROOT, '..', 'Stage1', 'tto_calc_reference_check_v1.1.py');
  return fs.existsSync(guess) ? guess : null;
}
function runPython(args) {
  const cands = process.env.TTO_PYTHON ? [[process.env.TTO_PYTHON, []]] :
    process.platform === 'win32' ? [['py', ['-3.12']], ['py', ['-3']], ['python', []]] : [['python3', []], ['python', []]];
  for (const [cmd, pre] of cands) {
    const r = spawnSync(cmd, pre.concat(['-I']).concat(args), { encoding: 'utf8' });
    if (!r.error && r.status !== null && !(r.status === 9009 || /not found|No suitable Python/i.test(r.stderr || ''))) return r;
  }
  return null;
}
const ref = findReference();
if (!ref) {
  out('\nPython cross-check: skipped (reference script not found; pass --python <path to tto_calc_reference_check_v1.1.py>).');
} else {
  out('\nPython cross-check against ' + ref);
  const own = runPython([ref, JSON_PATH]);
  if (!own) { out('  Python not available: cross-check skipped.'); }
  else {
    const last = (own.stdout || '').trim().split(/\r?\n/).pop();
    out('  Reference script: exit ' + own.status + ' - "' + last + '"');
    if (own.status !== 0) exitCode = 1;
    const cap = runPython([path.join(__dirname, 'python_reference_capture.py'), ref, JSON_PATH]);
    const py = JSON.parse(cap.stdout);
    let agree = 0, disagree = 0;
    for (const r of run.numeric ? run.results.filter((x) => x.type !== 'behavioral') : []) {
      const p = py[r.id];
      if (!p) { disagree++; out('  ' + r.id + ' missing from the Python capture'); continue; }
      const js = r.pairs;
      let ok = p.pairs.length === js.length && p.problems.length === 0;
      for (let k = 0; ok && k < js.length; k++) {
        if (!V.close(js[k][1], p.pairs[k][0], data.tolerance) || !V.close(p.pairs[k][0], js[k][1], data.tolerance)) ok = false;
      }
      if (ok) agree++; else { disagree++; out('  ' + r.id + ' DISAGREES: JS ' + JSON.stringify(js.map((x) => x[1])) + ' vs Python ' + JSON.stringify(p.pairs.map((x) => x[0]))); }
    }
    out('  JavaScript and Python agree value by value on ' + agree + '/' + (agree + disagree) + ' numeric vectors.');
    if (disagree) exitCode = 1;
  }
}

out('\n' + (exitCode === 0 ? 'RESULT: PASS' : 'RESULT: FAIL'));
process.exit(exitCode);
