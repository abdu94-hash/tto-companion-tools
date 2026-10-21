#!/usr/bin/env node
/*
 * TTO Companion Tools v1.0 - checks that the three fictional example cases loaded by the calculator
 * reproduce the numbers printed in their home chapters.
 *   C01 version 2: Chapter 3, Section 3.7 (Tables 5 and 8) and Chapter 24, Worked exercises 24.1 and 24.5(a)
 *   C04 version 2: Chapter 11, Section 11.8 (and Chapter 23, Section 23.11 for LAST_DOSE and the Q8 rows)
 *   C28 version 2: Chapter 20, Section 20.4 (and Chapter 23, Section 23.12 for the Q8 rows)
 * Usage: node tests/run_examples.js
 */
'use strict';
const path = require('path');
const ROOT = path.join(__dirname, '..');
const E = require(path.join(ROOT, 'docs', 'tto_calc_engine.js'));
const X = require(path.join(ROOT, 'docs', 'examples.js'));

let pass = 0, fail = 0;
const near = (a, b, rel) => (b === 0 ? Math.abs(a) <= 1e-9 : Math.abs(a - b) / Math.abs(b) <= (rel || 1e-6));
function check(label, cond, detail) {
  if (cond) { pass++; console.log('  PASS  ' + label); }
  else { fail++; console.log('  FAIL  ' + label + (detail ? '  -> ' + detail : '')); }
}
function flags(r) { return JSON.stringify(r.QC_Flags); }
function rec(ev, id) { return ev.records.find((r) => r.Result_ID === id); }
function interval(ev, id, want) {
  const r = rec(ev, id);
  if (!r) return check(id + ' present', false);
  const parts = [];
  let ok = !r.Blocked;
  if ('value' in want) { ok = ok && r.Interval_Value === want.value; parts.push('value ' + r.Interval_Value); }
  if ('min' in want) { ok = ok && r.Interval_Minimum === want.min && r.Interval_Maximum === want.max; parts.push('bounds ' + r.Interval_Minimum + '..' + r.Interval_Maximum); }
  if (want.unit) { ok = ok && r.Unit === want.unit; parts.push(r.Unit); }
  if (want.seq) { ok = ok && r.Sequence_Status === want.seq; parts.push(r.Sequence_Status); }
  if (want.flags) { ok = ok && flags(r) === JSON.stringify(want.flags); parts.push(flags(r)); }
  if (want.anchor) { ok = ok && r.Anchor_Record_ID === want.anchor; parts.push('anchor ' + r.Anchor_Record_ID); }
  if ('day' in want) { ok = ok && r.Treatment_Day === want.day; parts.push('Day ' + r.Treatment_Day); }
  if (want.display) { ok = ok && r.Display === want.display; parts.push('"' + r.Display + '"'); }
  check(id + ' ' + want.label + ' = ' + parts.join(', '), ok, r.Messages.map((m) => m.code).join(' '));
}
function technical(ev, id, want) {
  const r = rec(ev, id);
  if (!r) return check(id + ' present', false);
  if (want.blocked) {
    return check(id + ' ' + want.label + ' blocked with ' + flags(r), r.Blocked && flags(r) === JSON.stringify(want.blocked) && r.Results.length === 0 && r.Display === 'Not calculable');
  }
  let ok = !r.Blocked && flags(r) === JSON.stringify(want.flags || []);
  const shown = [];
  want.values.forEach(([key, value, display]) => {
    const x = r.Results.find((y) => y.key === key);
    ok = ok && !!x && near(x.value, value, want.rel) && (!display || x.display === display);
    shown.push(key + '=' + (x ? x.value + ' (' + x.display + ')' : 'missing'));
  });
  check(id + ' ' + want.label + ': ' + shown.join('; ') + ' ' + flags(r), ok);
}

// ---------------------------------------------------------------- C01
console.log('C01 (version 2) - Chapter 3, Section 3.7');
let c = E.clone(X.cases.C01), ev = E.evaluateCase(c);
check('no validation messages', ev.validation.length === 0, JSON.stringify(ev.validation));
interval(ev, 'I1', { label: 'Drug A start to onset (C1, V01)', value: 15, unit: 'minutes', seq: 'AFTER', flags: [], anchor: 'A-START', display: '15 minutes' });
interval(ev, 'I1-C2', { label: 'same pair at date level (C2; order decided by C1)', value: 0, unit: 'calendar days', seq: 'AFTER', flags: [], day: 1 });
interval(ev, 'I2', { label: 'LAST_DOSE_BEFORE_ONSET is the 09:00 infusion', value: 15, anchor: 'A-START', flags: [] });
interval(ev, 'I3', { label: 'Drug B start (CONCOMITANT_CHANGE)', value: 55, unit: 'minutes', seq: 'AFTER' });
interval(ev, 'I4', { label: 'Drug B end (OTHER)', value: 25, unit: 'minutes', seq: 'AFTER' });
interval(ev, 'I5', { label: 'Skin antiseptic (OTHER)', value: 60, unit: 'minutes', seq: 'AFTER' });
interval(ev, 'I6', { label: 'Drug A stopped (CESSATION)', value: -1, unit: 'minutes', seq: 'BEFORE', display: '−1 minute' });
interval(ev, 'R1', { label: 'recovery, stop to improvement (C7)', value: 19, unit: 'minutes', flags: ['W_INTERVENTION_BEFORE_RECOVERY'] });
technical(ev, 'T1', { label: 'Q9 with a rate in mL/hour', blocked: ['E_INVALID_UNIT'] });
// Worked exercise 24.1: offset left unknown
c = E.clone(X.cases.C01);
c.exposures[0].records[0].timing.offset = 'unknown'; c.events[0].onset.offset = 'unknown';
ev = E.evaluateCase(c);
interval(ev, 'I1', { label: 'offsets unknown (Worked exercise 24.1)', value: 15, unit: 'minutes (local clock)', flags: ['W_TZ_UNKNOWN'] });
// Worked exercise 24.5(a): with the drug A link removed, the interval is blocked with E_MISSING_LINK
c = E.clone(X.cases.C01); c.links = c.links.filter((l) => l.id !== 'L1'); ev = E.evaluateCase(c);
check('I1 without the Drug_Event_Link is blocked with E_MISSING_LINK', rec(ev, 'I1').Blocked && flags(rec(ev, 'I1')) === '["E_MISSING_LINK"]');

// ---------------------------------------------------------------- C04
console.log('\nC04 (version 2) - Chapter 11, Section 11.8');
ev = E.evaluateCase(E.clone(X.cases.C04));
check('no validation messages', ev.validation.length === 0, JSON.stringify(ev.validation));
interval(ev, 'I1', { label: 'FIRST_ADMIN (C2)', value: 185, unit: 'calendar days', seq: 'AFTER', flags: [] });
interval(ev, 'I2', { label: 'first documented renal deterioration (OTHER)', value: 4, unit: 'calendar days', seq: 'AFTER' });
interval(ev, 'I3', { label: 'ibuprofen start (CONCOMITANT_CHANGE)', value: 9, unit: 'calendar days', seq: 'AFTER' });
interval(ev, 'I4', { label: 'LAST_DOSE_BEFORE_ONSET (Chapter 23)', value: 1, unit: 'calendar days', anchor: 'DIG-LAST', flags: [] });
interval(ev, 'I5', { label: 'last dose to digoxin sample (C1, local clock)', value: 1660, unit: 'minutes (local clock)', flags: ['W_TZ_UNKNOWN'], display: '1,660 minutes (local clock)' });
check('I5 derived display about 27.7 hours', rec(ev, 'I5').Derived_Displays.hours === '27.7 hours');
interval(ev, 'R1', { label: 'recovery, withholding to improvement (C7)', value: 3, unit: 'calendar days', flags: ['W_INTERVENTION_BEFORE_RECOVERY'] });
technical(ev, 'T1', { label: 'Q11 baseline, female 81 y, 58 kg, 1.0 mg/dL', values: [['crcl_ml_min', (140 - 81) * 58 / 72 * 0.85, '40 mL/minute']] });
check('T1 unrounded 40.399 (Chapter 11)', Math.abs(rec(ev, 'T1').Results[0].value - 40.399) < 0.0005);
technical(ev, 'T2', { label: 'Q11 on 1 July (creatinine rising)', blocked: ['E_CREATININE_STABILITY_UNCONFIRMED'] });
technical(ev, 'T3', { label: 'Q11 on 5 July (creatinine rising)', blocked: ['E_CREATININE_STABILITY_UNCONFIRMED'] });
// a reviewer who ignored the gate would calculate 24 and 17 mL/minute (Chapter 11): check the arithmetic only
const ignored = [1.7, 2.4].map((s) => E.q11({ age: 81, weight_kg: 58, scr: s, scr_unit: 'mg/dL', sex: 'female', creatinine_stable: true }).out.crcl_ml_min);
check('gate-ignored values would be 24 and 17 mL/minute (Chapter 11 text)', Math.round(ignored[0]) === 24 && Math.round(ignored[1]) === 17, ignored.join(', '));
technical(ev, 'T4', { label: 'Q8 normal renal function 1.5–2 days at 3 days', flags: ['I_SCENARIO_NOT_CI'], values: [['fraction_lower', 0.25, '25.0%'], ['fraction_upper', Math.pow(2, -1.5), '35.4%']] });
technical(ev, 'T5', { label: 'Q8 anuric 3.5–5 days at 3 days', flags: ['I_SCENARIO_NOT_CI'], values: [['fraction_lower', Math.pow(2, -3 / 3.5), '55.2%'], ['fraction_upper', Math.pow(2, -0.6), '66.0%']] });

// ---------------------------------------------------------------- C28
console.log('\nC28 (version 2) - Chapter 20, Section 20.4');
ev = E.evaluateCase(E.clone(X.cases.C28));
check('no validation messages', ev.validation.length === 0, JSON.stringify(ev.validation));
interval(ev, 'I1', { label: 'last infusion to first positive test (C2) with reported text', value: 435, unit: 'calendar days', seq: 'AFTER', anchor: 'RTX-LAST', flags: ['I_APPROXIMATE_NOT_COMPARED'] });
check('I1 keeps "fourteen months after the last dose" as reported', rec(ev, 'I1').Reported_Timing.text === 'fourteen months after the last dose');
interval(ev, 'I2', { label: 'last infusion to onset bounds (C4)', min: 343, max: 435, unit: 'calendar days', seq: 'AFTER', flags: [], display: '343 to 435 calendar days' });
interval(ev, 'I3', { label: 'FIRST_ADMIN to first positive test', value: 540, seq: 'AFTER' });
interval(ev, 'I4', { label: 'FIRST_ADMIN to onset bounds', min: 448, max: 540, seq: 'AFTER' });
interval(ev, 'I5', { label: 'end of prophylaxis to first positive test', value: 251, seq: 'AFTER' });
interval(ev, 'I6', { label: 'end of prophylaxis to onset bounds', min: 159, max: 251, seq: 'AFTER' });
interval(ev, 'I7', { label: 'prednisolone start to first positive test', value: 71, seq: 'AFTER' });
interval(ev, 'I8', { label: 'prednisolone start to onset bounds', min: -21, max: 71, seq: 'SEQUENCE_UNCERTAIN', flags: ['W_SEQUENCE_UNCERTAIN'], display: '−21 to +71 calendar days' });
interval(ev, 'I9', { label: 'last infusion to biochemical hepatitis', value: 449, seq: 'AFTER' });
technical(ev, 'T1', { label: 'Q8 antibody 22–52 days at 435 days', flags: ['W_BIOLOGIC_MODEL_CAUTION', 'I_SCENARIO_NOT_CI'],
  values: [['fraction_lower', Math.pow(2, -435 / 22), '0.000112%'], ['fraction_upper', Math.pow(2, -435 / 52), '0.303%']] });
technical(ev, 'T2', { label: 'Q8 antibody 22–52 days at 343 days', flags: ['W_BIOLOGIC_MODEL_CAUTION', 'I_SCENARIO_NOT_CI'],
  values: [['fraction_lower', Math.pow(2, -343 / 22), '0.00203%'], ['fraction_upper', Math.pow(2, -343 / 52), '1.03%']] });
const t1 = rec(ev, 'T1');
check('Q8 shows no midpoint or average', !t1.Results.some((r) => /mid|average|mean/i.test(r.name)));

// ---------------------------------------------------------------- exports carry the limitation text and versions
console.log('\nExport records');
const b = E.buildExport(E.clone(X.cases.C28), { timestamp: '2026-10-20T12:00:00.000Z' });
check('JSON export carries tool version, CALC-SPEC version, timestamp and limitation text',
  b.Tool_Version === 'TTO Companion Tools v1.0' && b.Calculation_Version === 'CALC-SPEC 1.0' && b.Export_Timestamp && /not validated for regulated production use/.test(b.Limitation_Text));
check('every record carries Calculation_Version, Tool_Version, Export_Timestamp and Limitations',
  b.records.every((r) => r.Calculation_Version === 'CALC-SPEC 1.0' && r.Tool_Version === 'TTO Companion Tools v1.0' && r.Export_Timestamp && r.Limitations));
const txt = E.toText(b), html = E.toHTML(b), csv = E.toCSV(b.records);
check('text and HTML exports carry the limitation text and version label', txt.indexOf(E.LIMITATION_TEXT) >= 0 && html.indexOf('not validated for regulated production use') >= 0 && txt.indexOf(E.VERSION_LABEL) >= 0);
check('CSV carries Tool_Version and Calculation_Version on every row', E.recordsFromCSV(csv).every((r) => r.Tool_Version === E.TOOL_VERSION && r.Calculation_Version === E.SPEC_VERSION));
const evil = E.clone(X.cases.C01); evil.case_id = '<script>alert(1)</script>';
check('free text is escaped in the HTML export', E.toHTML(E.buildExport(evil)).indexOf('<script>alert') < 0);

console.log('\nExample checks: ' + pass + ' PASS, ' + fail + ' FAIL');
process.exit(fail ? 1 : 0);
