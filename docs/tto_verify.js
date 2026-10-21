/*
 * TTO Companion Tools v1.0 - verification runner for CALC-SPEC 1.0 vectors V01-V27.
 * Shared by the calculator's Verification tab and tests/run_vectors.js (Node).
 * Compares unrounded values with the JSON tolerance (relative 1e-6; absolute 1e-9 when the expected value is 0),
 * and compares units, sequence status and flags exactly.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) { module.exports = factory(); }
  else { root.TTOVerify = factory(); }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function close(a, b, tol) {
    if (a === null || a === undefined || b === null || b === undefined) return (a === null || a === undefined) && (b === null || b === undefined);
    if (typeof a === 'boolean' || typeof b === 'boolean' || typeof a === 'string' || typeof b === 'string') return a === b;
    if (Array.isArray(b)) return Array.isArray(a) && a.length === b.length && a.every(function (x, i) { return close(x, b[i], tol); });
    if (Array.isArray(a)) return false;
    if (typeof a !== 'number' || typeof b !== 'number') return JSON.stringify(a) === JSON.stringify(b);
    if (b === 0) return Math.abs(a) <= tol.absolute_for_zero;
    return Math.abs(a - b) / Math.abs(b) <= tol.relative * 1.0001;
  }

  function codesOf(E, fn) {
    try { fn(); return []; } catch (e) { if (e instanceof E.CalcError) return e.codes.slice(); throw e; }
  }

  // Build a minimal case for the case-level path (C5, link and ID validation).
  function miniCase(E, spec) {
    var c = E.emptyCase();
    c.case_id = 'VECTOR'; c.case_version = '1';
    for (var k in spec) c[k] = spec[k];
    return c;
  }

  function runC(E, v) {
    var inp = v.inputs, R = E.resolveTiming, T = E.timingFromIso, got = {}, res;
    switch (v.rule) {
      case 'C1':
        res = E.c1Exact(R(T(inp.anchor)), R(T(inp.onset)));
        got = { value: res.value, unit: res.unit, sequence: res.sequence, flags: res.flags };
        break;
      case 'C2':
        res = E.c2Calendar(R(T(inp.anchor)), R(T(inp.onset)));
        got = { value: res.value, unit: res.unit, sequence: res.sequence, flags: res.flags };
        break;
      case 'C2+C3':
        res = E.c2Calendar(R(T(inp.anchor)), R(T(inp.onset)));
        var td = E.c3TreatmentDay(res.value);
        got = { value: res.value, unit: res.unit, sequence: res.sequence, treatment_day: td.treatment_day, flags: res.flags.concat(td.flags) };
        break;
      case 'C4':
        var a = R({ type: 'partial', precision: 'range', lower: inp.anchor.lower, upper: inp.anchor.upper });
        var o = R({ type: 'partial', precision: 'range', lower: inp.onset.lower, upper: inp.onset.upper });
        res = E.c4Bounded(a.lo, a.hi, o.lo, o.hi);
        got = { minimum: res.minimum, maximum: res.maximum, unit: res.unit, sequence: res.sequence, flags: res.flags };
        break;
      case 'C5':
        if (inp.administrations) {
          var recs = inp.administrations.map(function (s, i) { return { id: 'AD' + (i + 1), kind: i === 0 ? 'start' : 'administration', label: 'Dose ' + (i + 1), timing: T(s) }; });
          var c = miniCase(E, {
            exposures: [{ id: 'EX1', drug: 'Drug A', role: 'suspect', records: recs }],
            events: [{ id: 'EV1', term: 'Event', onset: T(inp.onset) }],
            links: [{ id: 'L1', exposure_id: 'EX1', event_id: 'EV1', reason: 'vector' }],
            intervals: [{ id: 'I1', exposure_id: 'EX1', event_id: 'EV1', anchor_type: inp.anchor_type }]
          });
          var r = E.evaluateCase(c).records[0];
          var chosen = recs.filter(function (x) { return x.id === r.Anchor_Record_ID; })[0];
          got = { selected_anchor: chosen ? chosen.timing.original : null, value: r.Interval_Value, unit: r.Unit, sequence: r.Sequence_Status, flags: r.QC_Flags };
        } else {
          var c2 = miniCase(E, {
            exposures: [{ id: 'EX1', drug: 'Drug A', role: 'suspect', records: [
              { id: 'START', kind: 'start', label: 'Start', timing: T(inp.first_administration) },
              { id: 'STOP', kind: 'stop', label: 'Stop', timing: T(inp.stop) },
              { id: 'RESTART', kind: 'restart', label: 'Restart', timing: T(inp.restart) }] }],
            events: [{ id: 'EV1', term: 'Event', onset: T(inp.onset) }],
            links: [{ id: 'L1', exposure_id: 'EX1', event_id: 'EV1', reason: 'vector' }],
            intervals: [{ id: 'I1', exposure_id: 'EX1', event_id: 'EV1', anchor_type: 'FIRST_ADMIN' },
              { id: 'I2', exposure_id: 'EX1', event_id: 'EV1', anchor_type: 'RESTART', anchor_ref: 'RESTART' }]
          });
          var rr = E.evaluateCase(c2).records;
          var flags = []; rr.forEach(function (x) { x.QC_Flags.forEach(function (f) { if (flags.indexOf(f) < 0) flags.push(f); }); });
          got = { intervals: { FIRST_ADMIN: rr[0].Interval_Value, RESTART: rr[1].Interval_Value }, unit: rr[0].Unit === rr[1].Unit ? rr[0].Unit : null, flags: flags };
        }
        break;
      case 'C6':
        if (inp.anchor) {
          res = E.c1Exact(R(T(inp.anchor)), R(T(inp.onset)));
          var cmp = E.c6Compare(res, inp.reported);
          got = { value: res.value, unit: res.unit, reported_retained: true, flags: res.flags.concat(cmp.flags) };
        } else {
          var cmp2 = E.c6Compare(null, inp.reported);
          got = { value: null, flags: cmp2.flags };
        }
        break;
    }
    return got;
  }

  function runValidation(E, v) {
    var flags = [];
    v.inputs.cases.forEach(function (cs) {
      var codes = [];
      if (cs.check === 'date') codes = codesOf(E, function () { E.parseDateStr(cs.value); });
      else if (cs.check === 'bounds') codes = codesOf(E, function () { E.resolveTiming({ type: 'partial', precision: 'range', lower: cs.lower, upper: cs.upper }); });
      else if (cs.check === 'ids') {
        var c = miniCase(E, { exposures: cs.exposure_episode_ids.map(function (id) { return { id: id, drug: 'Drug', role: 'suspect', records: [] }; }) });
        E.validateCase(c).forEach(function (i) { if (codes.indexOf(i.code) < 0) codes.push(i.code); });
      } else if (cs.check === 'link') {
        var c2 = miniCase(E, {
          exposures: [{ id: cs.drug_id, drug: 'Drug', role: 'suspect', records: [{ id: cs.drug_id + '-S', kind: 'start', label: 'Start', timing: { type: 'date', date: '2026-01-01' } }] }],
          events: [{ id: cs.event_id, term: 'Event', onset: { type: 'date', date: '2026-01-05' } }],
          links: cs.links.map(function (l, i) { return { id: 'L' + i, exposure_id: l[0], event_id: l[1] }; }),
          intervals: [{ id: 'I1', exposure_id: cs.drug_id, event_id: cs.event_id, anchor_type: 'FIRST_ADMIN' }]
        });
        codes = E.evaluateCase(c2).records[0].QC_Flags.filter(function (f) { return f.charAt(0) === 'E'; });
      }
      codes.forEach(function (x) { flags.push(x); });
    });
    return { flags: flags };
  }

  function runTechnical(E, v) {
    var inp = v.inputs, got = {};
    if (v.rule === 'Technical validation') {
      var flags = [];
      inp.cases.forEach(function (cs) { codesOf(E, function () { E.technical[cs.rule](cs); }).forEach(function (x) { flags.push(x); }); });
      return { flags: flags };
    }
    if (v.rule === 'Q11') {
      var results = [], fl = [];
      inp.cases.forEach(function (cs) {
        try { var r = E.q11(cs); results.push(r.out.crcl_ml_min); fl.push(r.flags); }
        catch (e) { if (!(e instanceof E.CalcError)) throw e; results.push(null); fl.push(e.codes.slice()); }
      });
      return { results_ml_min: results, flags: fl };
    }
    var x = E.technical[v.rule];
    var input = JSON.parse(JSON.stringify(inp));
    if (v.rule === 'Q10') input.method_b = ['first', 'last'];
    var res = x(input);
    for (var k in res.out) got[k] = res.out[k];
    got.flags = res.flags;
    return got;
  }

  /** Run one vector. Returns {id, rule, type, status, problems, got, pairs}. pairs mirror the Python reference's comparisons. */
  function runVector(E, v, tol) {
    var out = { id: v.id, rule: v.rule, type: v.type, status: 'PASS', problems: [], got: null, pairs: [] };
    if (v.type === 'behavioral') { out.status = 'BEHAVIORAL'; return out; }
    var got;
    try {
      if (/^C/.test(v.rule)) got = runC(E, v);
      else if (v.rule === 'Validation') got = runValidation(E, v);
      else got = runTechnical(E, v);
    } catch (e) {
      out.status = 'FAIL'; out.problems.push('exception: ' + (e && e.message)); return out;
    }
    out.got = got;
    var exp = v.expected;
    Object.keys(exp).forEach(function (key) {
      if (key === 'description') return;
      var want = exp[key];
      if (key === 'unit') {                         // stricter than the Python reference: units must match exactly
        if (got.unit !== want) out.problems.push('unit: got ' + JSON.stringify(got.unit) + ', want ' + JSON.stringify(want));
        return;
      }
      if (!(key in got)) { out.problems.push(key + ': not produced'); return; }
      if (want !== null && typeof want === 'object' && !Array.isArray(want)) {
        Object.keys(want).forEach(function (k2) {
          var g = got[key] ? got[key][k2] : undefined;
          out.pairs.push([key + '.' + k2, g === undefined ? null : g, want[k2]]);
          if (!close(g, want[k2], tol)) out.problems.push(key + '.' + k2 + ': got ' + JSON.stringify(g) + ', want ' + JSON.stringify(want[k2]));
        });
      } else {
        out.pairs.push([key, got[key], want]);
        if (!close(got[key], want, tol)) out.problems.push(key + ': got ' + JSON.stringify(got[key]) + ', want ' + JSON.stringify(want));
      }
    });
    out.status = out.problems.length ? 'FAIL' : 'PASS';
    return out;
  }

  // ------------------------------------------------------------------ behavioral vectors, demonstrated automatically
  /** V12: follow-up creates a new case version; version 1 (9 calendar days) is preserved; version 2 gives 7. */
  function demoV12(E, v) {
    var i = v.inputs, notes = [];
    var c1 = miniCase(E, {
      case_id: 'V12', case_version: '1',
      exposures: [{ id: 'EX1', drug: 'Drug A', role: 'suspect', source: 'initial report', records: [{ id: 'EX1-S', kind: 'start', label: 'Start', timing: { type: 'date', date: i.version_1.anchor, original: i.version_1.anchor } }] }],
      events: [{ id: 'EV1', term: 'Event', source: 'initial report', onset: { type: 'date', date: i.version_1.onset, original: i.version_1.onset } }],
      links: [{ id: 'L1', exposure_id: 'EX1', event_id: 'EV1', reason: 'suspect drug' }],
      intervals: [{ id: 'I1', exposure_id: 'EX1', event_id: 'EV1', anchor_type: 'FIRST_ADMIN' }]
    });
    var versions = [];
    var snap = snapshotVersion(E, c1, 'Initial report');
    versions.push(snap);
    var c2 = newVersionFrom(E, c1);
    c2.events[0].onset = { type: 'date', date: i.version_2.onset, original: i.version_2.onset };
    c2.events[0].source = 'follow-up report';
    var r2 = E.evaluateCase(c2).records[0];
    var r1 = versions[0].records[0];
    var ok = r1.Interval_Value === 9 && r2.Interval_Value === 7 && versions[0].case.case_version === '1' && c2.case_version === '2' &&
      versions[0].case.events[0].onset.date === i.version_1.onset;
    notes.push('Version 1 (preserved snapshot): ' + r1.Display + '; onset ' + versions[0].case.events[0].onset.date + '.');
    notes.push('Version 2: ' + r2.Display + '; onset ' + c2.events[0].onset.date + '; Case_Version ' + c2.case_version + '.');
    return { id: v.id, rule: v.rule, type: 'behavioral', status: ok ? 'DEMONSTRATED' : 'NOT DEMONSTRATED', notes: notes };
  }
  /** V14: export V06 to JSON and CSV and re-import; bounds, unit, original text and precision, link, status and flags preserved. */
  function demoV14(E, v, vectors) {
    var v06 = vectors.filter(function (x) { return x.id === 'V06'; })[0].inputs, notes = [];
    var c = miniCase(E, {
      case_id: 'V14', case_version: '1',
      exposures: [{ id: 'EX1', drug: 'Drug A', role: 'suspect', source: 'consumer report', records: [{ id: 'EX1-S', kind: 'start', label: 'Start',
        timing: { type: 'partial', precision: 'range', lower: v06.anchor.lower, upper: v06.anchor.upper, original: 'first week or so of January, definitely before the 10th' } }] }],
      events: [{ id: 'EV1', term: 'Rash', source: 'consumer report and clinic note', onset: { type: 'date', date: v06.onset.lower, original: '5 January 2026' } }],
      links: [{ id: 'L1', exposure_id: 'EX1', event_id: 'EV1', reason: 'suspect drug' }],
      intervals: [{ id: 'I1', exposure_id: 'EX1', event_id: 'EV1', anchor_type: 'FIRST_ADMIN' }]
    });
    var bundle = E.buildExport(c, { timestamp: '2026-10-20T00:00:00.000Z' });
    var before = bundle.records[0];
    // JSON round trip
    var reloaded = E.importBundle(JSON.parse(JSON.stringify(bundle)));
    var after = E.evaluateCase(reloaded.case).records[0];
    var anc = reloaded.case.exposures[0].records[0].timing;
    var jsonOk = after.Interval_Minimum === before.Interval_Minimum && after.Interval_Maximum === before.Interval_Maximum &&
      after.Unit === before.Unit && after.Sequence_Status === before.Sequence_Status && after.Drug_Event_Link === before.Drug_Event_Link &&
      JSON.stringify(after.QC_Flags) === JSON.stringify(before.QC_Flags) && anc.original === 'first week or so of January, definitely before the 10th' &&
      anc.precision === 'range' && after.Interval_Minimum === -5 && after.Interval_Maximum === 4;
    // CSV round trip
    var csv = E.toCSV(bundle.records), back = E.recordsFromCSV(csv)[0];
    var pAnchor = back.Parameter_Set[0];
    var csvOk = back.Interval_Minimum === -5 && back.Interval_Maximum === 4 && back.Interval_Value === null && back.Unit === 'calendar days' &&
      back.Sequence_Status === 'SEQUENCE_UNCERTAIN' && back.Drug_Event_Link === 'L1' && JSON.stringify(back.QC_Flags) === JSON.stringify(['W_SEQUENCE_UNCERTAIN']) &&
      pAnchor.original_text === 'first week or so of January, definitely before the 10th' && pAnchor.precision === 'range' &&
      back.Tool_Version === E.TOOL_VERSION && back.Calculation_Version === E.SPEC_VERSION;
    notes.push('JSON export and re-import: ' + (jsonOk ? 'bounds −5 to +4, unit, original text and precision, link L1, SEQUENCE_UNCERTAIN and W_SEQUENCE_UNCERTAIN preserved.' : 'mismatch.'));
    notes.push('CSV export and re-import: ' + (csvOk ? 'the same fields preserved; Tool_Version and Calculation_Version present.' : 'mismatch.'));
    return { id: v.id, rule: v.rule, type: 'behavioral', status: jsonOk && csvOk ? 'DEMONSTRATED' : 'NOT DEMONSTRATED', notes: notes };
  }

  /** Snapshot a case version with its computed records (used by the app and by V12). */
  function snapshotVersion(E, c, label) {
    var ev = E.evaluateCase(c);
    return { label: label || '', saved_at: new Date().toISOString(), case: E.clone(c), records: E.clone(ev.records) };
  }
  function newVersionFrom(E, c) {
    var n = E.clone(c), v = parseInt(c.case_version, 10);
    n.case_version = isFinite(v) ? String(v + 1) : (c.case_version || '1') + '+1';
    return n;
  }

  /** Run all vectors. Returns {results, numeric, passed, failed, behavioral}. */
  function runAll(E, data) {
    if (data.spec !== E.SPEC_VERSION) throw new Error('Vector file is for ' + data.spec + ', engine implements ' + E.SPEC_VERSION);
    var results = data.vectors.map(function (v) {
      if (v.type === 'behavioral') {
        try { return v.id === 'V12' ? demoV12(E, v) : v.id === 'V14' ? demoV14(E, v, data.vectors) : { id: v.id, rule: v.rule, type: 'behavioral', status: 'INSPECT', notes: [] }; }
        catch (e) { return { id: v.id, rule: v.rule, type: 'behavioral', status: 'NOT DEMONSTRATED', notes: ['exception: ' + e.message] }; }
      }
      return runVector(E, v, data.tolerance);
    });
    var numeric = results.filter(function (r) { return r.type !== 'behavioral'; });
    var passed = numeric.filter(function (r) { return r.status === 'PASS'; }).length;
    return { results: results, numeric: numeric.length, passed: passed, failed: numeric.length - passed,
      behavioral: results.filter(function (r) { return r.type === 'behavioral'; }) };
  }

  return { close: close, runVector: runVector, runAll: runAll, demoV12: demoV12, demoV14: demoV14, snapshotVersion: snapshotVersion, newVersionFrom: newVersionFrom };
}));
