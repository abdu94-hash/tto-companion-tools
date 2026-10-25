/*
 * TTO Companion Tools v1.1 - calculator user interface (Chapter 24). CALC-SPEC 1.1.
 * All calculation is done by tto_calc_engine.js in this browser. No network requests, no analytics,
 * no browser storage: a case is saved only when the user downloads a file.
 */
(function () {
  'use strict';
  var E = window.TTOCalc, VF = window.TTOVerify, VEC = window.TTO_VECTORS, EX = window.TTO_EXAMPLES;

  var state = { c: blankCase(), versions: [], tab: 'facts', dirty: false, verify: null, importedCSV: null };
  var TABS = ['facts', 'arith', 'tech', 'assess', 'export', 'verify', 'about'];
  var TAB_IDS = { facts: 'tab-facts', arith: 'tab-arith', tech: 'tab-tech', assess: 'tab-assess', export: 'tab-export', verify: 'tab-verify', about: 'tab-about' };

  var RECORD_KINDS = [['start', 'First administration or start'], ['administration', 'Administration (dose)'], ['dose_change', 'Dose change'], ['stop', 'Stop or cessation'], ['restart', 'Restart']];
  var MILESTONE_KINDS = [['recovery', 'Recovery or improvement'], ['intervention', 'Intervention or treatment'], ['test', 'Test or sample'], ['organ_function', 'Organ function'], ['diagnosis', 'Diagnosis'], ['detection', 'First detection'], ['baseline', 'Baseline status'], ['other', 'Other']];
  var ROLES = [['suspect', 'Suspect'], ['concomitant', 'Concomitant'], ['prophylaxis', 'Prophylaxis'], ['treatment of the event', 'Treatment of the event'], ['other', 'Other']];
  var FIT_OPTIONS = [['compatible', 'Compatible'], ['atypical but plausible', 'Atypical but plausible'], ['inconsistent with the stated hypothesis', 'Inconsistent with the stated hypothesis'], ['indeterminate', 'Indeterminate']];
  var PROMPTS = [['mechanism_hypothesis', 'Mechanism hypothesis'], ['expected_exposure_course', 'Expected exposure course'], ['expected_pd_course', 'Expected PD course'],
    ['clinical_latency_evidence', 'Clinical latency evidence (state first which kind of timing statement it is)'], ['observed_course', 'Observed course'], ['temporal_fit_rationale', 'Temporal-fit rationale']];
  var OFFSETS = (function () {
    var o = [['', '— choose —'], ['unknown', 'offset unknown']];
    var list = ['-12:00', '-11:00', '-10:00', '-09:30', '-09:00', '-08:00', '-07:00', '-06:00', '-05:00', '-04:00', '-03:30', '-03:00', '-02:00', '-01:00', '+00:00',
      '+01:00', '+02:00', '+03:00', '+03:30', '+04:00', '+04:30', '+05:00', '+05:30', '+05:45', '+06:00', '+06:30', '+07:00', '+08:00', '+08:45', '+09:00', '+09:30',
      '+10:00', '+10:30', '+11:00', '+12:00', '+12:45', '+13:00', '+14:00'];
    list.forEach(function (x) { o.push([x, 'UTC' + x.replace('-', '−')]); });
    return o;
  })();
  var TIME_UNIT_OPTS = [['', '— choose —'], ['minutes', 'minutes'], ['hours', 'hours'], ['days', 'days']];
  var STATUS_OPTS = [['', '— status —']].concat(E.PARAMETER_STATUSES.map(function (s) { return [s, s]; }));

  function blankCase() {
    var c = E.emptyCase(); c.case_id = 'CASE-1'; c.case_version = '1'; c.technical_mode = false; return c;
  }

  // ------------------------------------------------------------------ small helpers
  function h(s) { return E.escapeHTML(s); }
  function $(sel, root) { return (root || document).querySelector(sel); }
  function $all(sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); }
  function fid(path) { return 'f-' + String(path).replace(/[^A-Za-z0-9]/g, '-'); }
  function getPath(obj, path) {
    var parts = String(path).split('.'), o = obj;
    for (var i = 0; i < parts.length; i++) { if (o === null || o === undefined) return undefined; o = o[parts[i]]; }
    return o;
  }
  function setPath(obj, path, value) {
    var parts = String(path).split('.'), o = obj;
    for (var i = 0; i < parts.length - 1; i++) {
      if (o[parts[i]] === null || o[parts[i]] === undefined || typeof o[parts[i]] !== 'object') o[parts[i]] = /^\d+$/.test(parts[i + 1]) ? [] : {};
      o = o[parts[i]];
    }
    o[parts[parts.length - 1]] = value;
  }
  function opts(list, value) {
    return list.map(function (o) { var v = o[0], t = o.length > 1 ? o[1] : o[0]; return '<option value="' + h(v) + '"' + (String(value == null ? '' : value) === String(v) ? ' selected' : '') + '>' + h(t) + '</option>'; }).join('');
  }
  function inputF(label, path, o) {
    o = o || {};
    var v = getPath(state.c, path), id = fid(path);
    return '<label class="field' + (o.cls ? ' ' + o.cls : '') + '" for="' + id + '"><span class="lbl">' + h(label) + '</span>' +
      '<input type="text" id="' + id + '" data-path="' + h(path) + '" value="' + h(v == null ? '' : v) + '"' +
      (o.placeholder ? ' placeholder="' + h(o.placeholder) + '"' : '') + (o.inputmode ? ' inputmode="' + o.inputmode + '"' : '') +
      (o.rerender ? ' data-rerender="1"' : '') + (o.readonly ? ' readonly' : '') + ' autocomplete="off" spellcheck="false"></label>';
  }
  function selectF(label, path, list, o) {
    o = o || {};
    var v = getPath(state.c, path), id = fid(path);
    return '<label class="field' + (o.cls ? ' ' + o.cls : '') + '" for="' + id + '"><span class="lbl">' + h(label) + '</span>' +
      '<select id="' + id + '" data-path="' + h(path) + '"' + (o.rerender ? ' data-rerender="1"' : '') + '>' + (o.raw || opts(list, v)) + '</select></label>';
  }
  function checkF(label, path, o) {
    o = o || {};
    var v = getPath(state.c, path) === true, id = fid(path);
    return '<label class="field check' + (o.cls ? ' ' + o.cls : '') + '" for="' + id + '"><input type="checkbox" id="' + id + '" data-path="' + h(path) + '"' +
      (v ? ' checked' : '') + (o.rerender ? ' data-rerender="1"' : '') + '><span class="lbl">' + h(label) + '</span></label>';
  }
  function textareaF(label, path, o) {
    o = o || {};
    var v = getPath(state.c, path), id = fid(path);
    return '<label class="field' + (o.cls ? ' ' + o.cls : '') + '" for="' + id + '"><span class="lbl">' + h(label) + '</span>' +
      '<textarea id="' + id + '" data-path="' + h(path) + '" rows="' + (o.rows || 2) + '">' + h(v == null ? '' : v) + '</textarea></label>';
  }
  function btn(text, action, data, cls) {
    var attrs = ''; for (var k in (data || {})) attrs += ' data-' + k + '="' + h(data[k]) + '"';
    return '<button type="button" class="' + (cls || 'small') + '" data-action="' + action + '"' + attrs + '>' + h(text) + '</button>';
  }
  function allIds() {
    var idx = E.indexCase(state.c); return Object.keys(idx.counts);
  }
  function nextId(prefix) {
    var used = allIds(), n = 1;
    while (used.indexOf(prefix + n) >= 0) n++;
    return prefix + n;
  }
  function toast(msg) {
    var t = $('#toast'); t.textContent = msg; t.hidden = false;
    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 9000);
  }
  function emptyTs() { return { type: 'timestamp', date: '', time: '', offset: '', original: '' }; }

  // ------------------------------------------------------------------ timing editor (Section 2 data types)
  function timingEditor(label, path, o) {
    o = o || {};
    var t = getPath(state.c, path);
    if (!t) { t = o.tsOnly ? emptyTs() : { type: 'none', original: '' }; setPath(state.c, path, t); }
    var s = '<fieldset class="timing"><legend class="small"><strong>' + h(label) + '</strong></legend><div class="grid">';
    if (!o.tsOnly) {
      var types = [['timestamp', 'Date and time'], ['date', 'Date only'], ['partial', 'Partial date or range'], ['none', 'Not recorded']];
      if (o.allowLab) types.splice(3, 0, ['lab', 'Laboratory-defined (between two tests)']);
      s += selectF('Entered as', path + '.type', types, { rerender: true });
    }
    if (t.type === 'lab') {
      var msOpts = [['', '— choose a test milestone —']].concat(state.c.milestones.map(function (m) { return [m.id, m.id + ' – ' + (m.label || m.kind || '') + ' (' + E.fmtTiming(m.timing) + ')']; }));
      s += selectF('Last normal test (lower bound)', path + '.last_normal_ref', msOpts);
      s += selectF('First abnormal test (upper bound)', path + '.first_abnormal_ref', msOpts);
    }
    if (t.type === 'timestamp') {
      s += inputF('Local date (YYYY-MM-DD)', path + '.date', { placeholder: 'YYYY-MM-DD', inputmode: 'numeric' });
      s += inputF('Local time (HH:MM, 24-hour)', path + '.time', { placeholder: 'HH:MM', inputmode: 'numeric' });
      s += selectF('UTC offset', path + '.offset', OFFSETS);
    } else if (t.type === 'date') {
      s += inputF('Date (YYYY-MM-DD)', path + '.date', { placeholder: 'YYYY-MM-DD', inputmode: 'numeric' });
    } else if (t.type === 'partial') {
      s += selectF('Precision', path + '.precision', [['', '— choose —'], ['year', 'Year'], ['month', 'Month'], ['range', 'Explicit range']], { rerender: true });
      if (t.precision === 'year') s += inputF('Year (YYYY)', path + '.value', { placeholder: 'YYYY', inputmode: 'numeric' });
      if (t.precision === 'month') s += inputF('Month (YYYY-MM)', path + '.value', { placeholder: 'YYYY-MM', inputmode: 'numeric' });
      if (t.precision === 'range') {
        s += inputF('Earliest (YYYY-MM-DD)', path + '.lower', { placeholder: 'YYYY-MM-DD', inputmode: 'numeric' });
        s += inputF('Latest (YYYY-MM-DD)', path + '.upper', { placeholder: 'YYYY-MM-DD', inputmode: 'numeric' });
      }
    }
    if (!o.noOriginal) s += inputF('Original text, as reported', path + '.original', { cls: 'wide-field' });
    s += '</div><div class="readout" data-readout="' + h(path) + '"></div></fieldset>';
    return s;
  }
  function readoutFor(t) {
    if (!t || !t.type || t.type === 'none') return { text: 'Not recorded. Nothing is imputed.', bad: false };
    if (t.type === 'lab') {
      try {
        var lt = E.labOnsetTiming(E.indexCase(state.c), t);
        return { text: 'Reads as: ' + E.fmtRangeDays(E.parseDateStr(lt.lower), E.parseDateStr(lt.upper)) + ' (laboratory-defined range: lower bound = date of the last normal test ' + lt.last_normal_ref +
          ', upper bound = date of the first abnormal test ' + lt.first_abnormal_ref + '; C4 applies). You decide which results are normal; nothing is imputed.', bad: false };
      } catch (e) {
        if (!(e instanceof E.CalcError)) throw e;
        return { text: e.codes.map(function (c) { return c + ': ' + E.message(c, e.details[c]); }).join(' '), bad: true };
      }
    }
    try {
      E.resolveTiming(t, 'this date');
      return { text: 'Reads as: ' + E.fmtTiming(t) + (t.type === 'partial' ? '. The bounds are kept; no single date is imputed.' : ''), bad: false };
    } catch (e) {
      if (!(e instanceof E.CalcError)) throw e;
      return { text: e.codes.map(function (c) { return c + ': ' + E.message(c, e.details[c]); }).join(' '), bad: true };
    }
  }

  // ------------------------------------------------------------------ result display
  function flagHTML(m) {
    return '<div class="flag ' + m.code.charAt(0) + '" role="' + (m.code.charAt(0) === 'E' ? 'alert' : 'note') + '"><code>' + h(m.code) + '</code>' +
      '<span class="sr-only">' + h(m.type) + ': </span>' + h(m.message) + '</div>';
  }
  function raw(v) { return Array.isArray(v) ? v.join(', ') : String(v); }
  function paramList(ps) {
    if (!ps || !ps.length) return '<p class="small muted">No inputs.</p>';
    return '<dl class="kv">' + ps.map(function (p) {
      var bits = [];
      bits.push(h(p.display || (p.value === null || p.value === undefined ? 'blank' : p.value)) + (p.unit ? ' [' + h(p.unit) + ']' : ''));
      if (p.precision) bits.push('precision: ' + h(p.precision));
      if (p.original_text) bits.push('as reported: “' + h(p.original_text) + '”');
      bits.push('Parameter_Status: ' + (p.Parameter_Status ? h(p.Parameter_Status) : '<em>not recorded</em>'));
      bits.push('source: ' + (p.source ? h(p.source) : '<em>not recorded</em>'));
      return '<dt>' + h(p.name) + (p.record_id ? ' (' + h(p.record_id) + ')' : '') + '</dt><dd>' + bits.join('; ') + '</dd>';
    }).join('') + '</dl>';
  }
  function resultHTML(r) {
    var tech = /^Q/.test(r.Calculation_Mode);
    var s = '<div class="result' + (r.Blocked ? ' blocked' : '') + '" aria-live="polite">';
    s += '<div class="r-head"><span class="r-title">' + h(r.Result_ID) + ' · ' + h(r.Label) + '</span><span class="pill mode" title="Calculation_Mode">' + h(r.Calculation_Mode) + '</span>' +
      (r.Anchor_Type ? '<span class="pill" title="Anchor_Type">' + h(r.Anchor_Type) + '</span>' : '') +
      (r.Sequence_Status ? '<span class="seq ' + h(r.Sequence_Status) + '" title="Sequence_Status">' + h(r.Sequence_Status) + '</span>' : '') + '</div>';
    if (r.Blocked) {
      s += '<div class="value na">Not calculable</div><p class="small muted">Left blank rather than estimated (rule G2). Interval_Value, Interval_Minimum and Interval_Maximum are blank in the export.</p>';
    } else if (tech) {
      if (r.Calculation_Mode === 'Q8') {
        s += '<div class="value">' + h(r.Results[0].display) + ' <span class="unit">to</span> ' + h(r.Results[1].display) + '</div>' +
          '<p class="small">Scenario range: fraction remaining at the lower half-life to the fraction at the upper half-life. No midpoint or average is shown.</p>';
      }
      s += '<div class="table-wrap"><table><thead><tr><th scope="col">Result</th><th scope="col">Displayed</th><th scope="col">Unrounded</th></tr></thead><tbody>' +
        r.Results.map(function (x) { return '<tr><td>' + h(x.name) + '</td><td class="num"><strong>' + h(x.display) + '</strong></td><td class="num">' + h(raw(x.value)) + '</td></tr>'; }).join('') +
        '</tbody></table></div>';
      if (r.Calculation_Mode === 'Q5') s += '<p class="small">Each percentage is normalized to its own starting value. Their sum is never calculated or displayed.</p>';
    } else {
      s += '<div class="value">' + h(r.Display) + '</div>';
      var un = r.Interval_Value !== null ? r.Interval_Value : r.Interval_Minimum + ' to ' + r.Interval_Maximum;
      s += '<div class="derived">Unrounded: ' + h(un) + ' ' + h(r.Unit) + (r.Derived_Displays ? ' · derived displays: ' + h(r.Derived_Displays.hours) + '; ' + h(r.Derived_Displays.elapsed_days) : '') + '</div>';
      if (r.Treatment_Day !== null && r.Treatment_Day !== undefined) s += '<div class="derived">Treatment day (C3, inclusive convention): <strong>Day ' + h(r.Treatment_Day) + '</strong> (never written to E2B fields automatically)</div>';
    }
    if (r.Reported_Timing) {
      var rt = r.Reported_Timing;
      s += '<div class="reported"><strong>Reported timing (kept as stated, not calculated):</strong> “' + h(rt.text) + '”' +
        (rt.lower !== null && rt.lower !== undefined ? ' (bounds ' + h(rt.lower) + '–' + h(rt.upper) + ' ' + h(rt.unit || '') + ')' : '') + ' – ' + h(rt.comparison || '') + '</div>';
    }
    if (r.Constraint_Source) s += '<div class="reported"><strong>Constrained by a documented fact.</strong> Source: ' + h(r.Constraint_Source) + '</div>';
    s += r.Messages.map(flagHTML).join('');
    if (tech) s += '<div class="mandatory"><strong>Mandatory text.</strong> ' + h(r.Limitations) + '</div>';
    s += '<details><summary>Inputs, assumptions and equation (Parameter_Set)</summary>' + paramList(r.Parameter_Set) +
      (r.Assumptions.length ? '<h4>Assumptions</h4><ul class="small">' + r.Assumptions.map(function (a) { return '<li>' + h(a) + '</li>'; }).join('') + '</ul>' : '') +
      (r.Equation ? '<p class="small"><strong>Equation.</strong> ' + h(r.Equation) + '</p>' : '') +
      (!tech && r.Limitations ? '<p class="small"><strong>Limitations.</strong> ' + h(r.Limitations) + '</p>' : '') +
      '<p class="small muted">' + h(r.Calculation_Version) + ' · Drug_Event_Link ' + h(r.Drug_Event_Link || '—') + ' · Anchor_Record_ID ' + h(r.Anchor_Record_ID || '—') + '</p></details>';
    return s + '</div>';
  }
  function summaryTable(records) {
    if (!records.length) return '<p class="small muted">No results yet.</p>';
    return '<div class="table-wrap"><table><thead><tr><th scope="col">ID</th><th scope="col">Interval or calculation</th><th scope="col">Displayed</th><th scope="col">Unrounded</th><th scope="col">Rule</th><th scope="col">Anchor</th><th scope="col">Sequence</th><th scope="col">Flags</th></tr></thead><tbody>' +
      records.map(function (r) {
        var un = r.Blocked ? 'blank' : r.Results && r.Results.length ? r.Results.map(function (x) { return raw(x.value); }).join('; ') :
          (r.Interval_Value !== null ? r.Interval_Value : r.Interval_Minimum + ' to ' + r.Interval_Maximum) + ' ' + r.Unit;
        return '<tr><td>' + h(r.Result_ID) + '</td><td>' + h(r.Label) + '</td><td class="num"><strong>' + h(r.Blocked ? 'Not calculable' : r.Display) + '</strong></td><td class="num">' + h(un) + '</td><td>' + h(r.Calculation_Mode) +
          '</td><td>' + h(r.Anchor_Type ? r.Anchor_Type + (r.Anchor_Record_ID ? ' (' + r.Anchor_Record_ID + ')' : '') : '') + '</td><td>' + (r.Sequence_Status ? '<span class="seq ' + h(r.Sequence_Status) + '">' + h(r.Sequence_Status) + '</span>' : '') +
          '</td><td>' + (r.QC_Flags.length ? r.QC_Flags.map(function (f) { return '<code class="flagcode ' + f.charAt(0) + '">' + h(f) + '</code>'; }).join(' ') : '<span class="muted">none</span>') + '</td></tr>';
      }).join('') + '</tbody></table></div>';
  }

  // ------------------------------------------------------------------ timeline (fact display)
  function tlPoint(t) {
    try {
      var r = E.resolveTiming(t);
      if (r.kind === 'ts') return { a: r.localMin, b: r.localMin, exact: true };
      if (r.kind === 'date') return { a: r.day * 1440 + 720, b: r.day * 1440 + 720, exact: true, dateOnly: true };
      return { a: r.lo * 1440, b: r.hi * 1440 + 1440, exact: false };
    } catch (e) { return null; }
  }
  function timelineHTML() {
    var c = state.c, rows = [], undated = [];
    (c.exposures || []).forEach(function (x) {
      var marks = [];
      (x.records || []).forEach(function (r) { var p = tlPoint(r.timing); if (p) marks.push({ p: p, title: (r.label || r.kind) + ': ' + E.fmtTiming(r.timing) }); else undated.push(x.drug + ': ' + (r.label || r.kind) + (r.timing && r.timing.original ? ' (“' + r.timing.original + '”)' : '')); });
      if (marks.length) rows.push({ cls: 'exposure', label: x.drug || x.id, kind: 'exposure', marks: marks, span: true });
    });
    (c.events || []).forEach(function (ev) {
      var on = ev.onset;
      if (on && on.type === 'lab') { try { on = E.eventOnset(c, ev); } catch (e) { on = null; } }
      var p = on ? tlPoint(on) : null;
      if (p) rows.push({ cls: 'event', label: (ev.term || ev.id) + ' (onset)', kind: 'event', marks: [{ p: p, title: 'Onset: ' + E.fmtTiming(on) + (on.lab ? ' (laboratory-defined)' : '') }] });
      else undated.push((ev.term || ev.id) + ' onset');
    });
    (c.milestones || []).forEach(function (m) {
      var p = tlPoint(m.timing);
      if (p) rows.push({ cls: 'milestone', label: m.label || m.id, kind: 'milestone', marks: [{ p: p, title: (m.label || m.id) + ': ' + E.fmtTiming(m.timing) }] });
      else undated.push((m.label || m.id) + (m.timing && m.timing.original ? ' (“' + m.timing.original + '”)' : '') + (m.value ? ': ' + m.value : ''));
    });
    var reported = [];
    (c.intervals || []).forEach(function (q) { if (q.reported && String(q.reported.text || '').trim()) reported.push('“' + q.reported.text + '” (interval ' + q.id + ')'); });
    if (!rows.length) return '<p class="small muted">The timeline appears when dates are entered. Exact dates and times are drawn as points; ranges as bars.</p>' +
      (undated.length ? '<p class="small"><strong>Not dated:</strong> ' + h(undated.join('; ')) + '</p>' : '');
    var lo = Infinity, hi = -Infinity;
    rows.forEach(function (r) { r.marks.forEach(function (m) { lo = Math.min(lo, m.p.a); hi = Math.max(hi, m.p.b); }); });
    var span = hi - lo; if (span <= 0) { span = 60; lo -= 30; hi += 30; }
    lo -= span * 0.03; hi += span * 0.03; span = hi - lo;
    function pct(x) { return ((x - lo) / span * 100).toFixed(3); }
    var s = '<div class="timeline" role="img" aria-label="Case timeline: exposures, event onset and milestones">';
    rows.forEach(function (r) {
      s += '<div class="tl-row ' + r.cls + '"><div class="tl-label" title="' + h(r.label) + '"><span class="k">' + h(r.kind) + '</span> ' + h(r.label) + '</div><div class="tl-track">';
      if (r.span && r.marks.length > 1) {
        var a = Math.min.apply(null, r.marks.map(function (m) { return m.p.a; })), b = Math.max.apply(null, r.marks.map(function (m) { return m.p.b; }));
        s += '<span class="tl-span" data-left="' + pct(a) + '" data-width="' + (Math.max(0, (b - a) / span * 100)).toFixed(3) + '"></span>';
      }
      r.marks.forEach(function (m) {
        if (m.p.exact) s += '<span class="tl-pt" data-left="' + pct(m.p.a) + '" title="' + h(m.title) + '"></span>';
        else s += '<span class="tl-bar" data-left="' + pct(m.p.a) + '" data-width="' + Math.max(0.6, (m.p.b - m.p.a) / span * 100).toFixed(3) + '" title="' + h(m.title) + '"></span>';
        s += '<span class="sr-only">' + h(m.title) + '. </span>';
      });
      s += '</div></div>';
    });
    // axis
    var ticks = [], n = 5, byTime = span < 3 * 1440;
    for (var i = 0; i <= n; i++) ticks.push(lo + span * i / n);
    s += '<div class="tl-axis"><div></div><div class="tl-ticks" aria-hidden="true">' + ticks.map(function (t, i) {
      var day = Math.floor(t / 1440), mod = Math.round(t - day * 1440); if (mod >= 1440) { day++; mod -= 1440; }
      var lab = byTime ? (('0' + Math.floor(mod / 60)).slice(-2) + ':' + ('0' + (mod % 60)).slice(-2)) + (i === 0 ? ' ' + E.fmtDay(day) : '') : E.fmtDay(day);
      return '<span class="tl-tick' + (i === 0 ? ' first' : i === n ? ' last' : '') + '" data-left="' + pct(t) + '"><span>' + h(lab) + '</span></span>';
    }).join('') + '</div></div></div>';
    s += '<div class="legend"><span><i class="x"></i>exposure record</span><span><i class="e"></i>event onset</span><span><i class="m"></i>milestone</span><span><i class="r"></i>range (bounds, not a date)</span><span>Times on the local clock as recorded.</span></div>';
    if (reported.length) s += '<p class="small"><strong>Reported relative timing (as stated; not calculated, not drawn as dates):</strong> ' + h(reported.join('; ')) + '</p>';
    if (undated.length) s += '<p class="small"><strong>Not dated:</strong> ' + h(undated.join('; ')) + '</p>';
    return s;
  }
  function applyPositions(root) {
    $all('[data-left]', root).forEach(function (el) {
      el.style.left = el.getAttribute('data-left') + '%';
      if (el.hasAttribute('data-width')) el.style.width = el.getAttribute('data-width') + '%';
    });
  }

  // ------------------------------------------------------------------ panel 1: source facts
  function renderFacts() {
    var c = state.c, s = '';
    s += '<h2>Source facts</h2><p class="lead">Enter facts as reported: each date keeps its original wording, precision and source. Nothing is imputed, and no input has a default value. Do not enter direct patient identifiers.</p>';
    s += '<div class="card"><div class="card-head"><h3>Case setup</h3></div><div class="grid">' +
      inputF('Case_ID (pseudonymous)', 'case_id') + inputF('Case_Version', 'case_version') + inputF('Reviewer (optional)', 'reviewer') + inputF('Setting', 'setting') + '</div>' +
      '<div class="grid wide">' + textareaF('Synthetic or fictional data note', 'synthetic_note') + textareaF('Case notes', 'notes') + '</div>' +
      '<p class="small muted">When follow-up arrives, use <strong>New case version</strong>: the current version is kept read-only with its results, and you edit the next one (V12).</p>' +
      '<div id="version-history"></div></div>';
    s += '<div class="card"><div class="card-head"><h3>Timeline</h3></div><div data-out="timeline"></div></div>';
    // events
    s += '<div class="card"><div class="card-head"><h3>Events</h3>' + btn('Add event', 'add', { list: 'events' }) + '</div>';
    if (!c.events.length) s += '<p class="small muted">No event yet.</p>';
    c.events.forEach(function (ev, i) {
      var p = 'events.' + i;
      s += '<div class="card sub"><div class="grid">' + inputF('Event_ID', p + '.id', { rerender: true }) + inputF('Event term', p + '.term', { rerender: true }) + inputF('Source', p + '.source') + '</div>' +
        timingEditor('Onset', p + '.onset', { allowLab: true }) + '<div class="row-actions">' + btn('Remove event', 'remove', { path: 'events', index: i }, 'small danger') + '</div></div>';
    });
    s += '</div>';
    // exposures
    s += '<div class="card"><div class="card-head"><h3>Drug exposure episodes</h3>' + btn('Add exposure', 'add', { list: 'exposures' }) + '</div>';
    if (!c.exposures.length) s += '<p class="small muted">No exposure yet.</p>';
    c.exposures.forEach(function (x, i) {
      var p = 'exposures.' + i;
      s += '<div class="card sub"><div class="grid">' + inputF('Exposure_Episode_ID', p + '.id', { rerender: true }) + inputF('Drug or product', p + '.drug', { rerender: true }) +
        selectF('Role', p + '.role', [['', '— choose —']].concat(ROLES)) + inputF('Source', p + '.source') + '</div>' +
        '<div class="grid wide">' + inputF('Dose, route and details as reported', p + '.details') + '</div>';
      (x.records || []).forEach(function (r, j) {
        var q = p + '.records.' + j;
        s += '<div class="card"><div class="grid">' + inputF('Record ID', q + '.id', { rerender: true }) + selectF('Kind', q + '.kind', [['', '— choose —']].concat(RECORD_KINDS), { rerender: true }) +
          inputF('Label', q + '.label', { rerender: true }) + '</div>' + timingEditor('Date or time', q + '.timing') +
          '<div class="row-actions">' + btn('Remove record', 'remove', { path: p + '.records', index: j }, 'small danger') + '</div></div>';
      });
      s += '<div class="row-actions">' + btn('Add administration or change', 'add-record', { index: i }) + btn('Remove exposure', 'remove', { path: 'exposures', index: i }, 'small danger') + '</div></div>';
    });
    s += '</div>';
    // milestones
    s += '<div class="card"><div class="card-head"><h3>Clinical milestones (optional)</h3>' + btn('Add milestone', 'add', { list: 'milestones' }) + '</div>' +
      '<p class="small muted">Recovery, interventions, tests, organ-function values, diagnosis and first detection are separate milestones. A diagnosis or detection date is never an onset substitute.</p>';
    c.milestones.forEach(function (m, i) {
      var p = 'milestones.' + i;
      s += '<div class="card sub"><div class="grid">' + inputF('Milestone ID', p + '.id', { rerender: true }) + selectF('Kind', p + '.kind', [['', '— choose —']].concat(MILESTONE_KINDS), { rerender: true }) +
        inputF('Label', p + '.label', { rerender: true }) + inputF('Value as reported', p + '.value') + inputF('Source', p + '.source') + '</div>' +
        timingEditor('Date or time', p + '.timing') + '<div class="row-actions">' + btn('Remove milestone', 'remove', { path: 'milestones', index: i }, 'small danger') + '</div></div>';
    });
    s += '</div>';
    return s;
  }
  function renderVersionHistory() {
    var el = $('#version-history'); if (!el) return;
    if (!state.versions.length) { el.innerHTML = '<p class="small muted">Version history: none saved in this session.</p>'; return; }
    el.innerHTML = '<h4>Version history (read only)</h4>' + state.versions.map(function (v, i) {
      return '<details><summary>Case_Version ' + h(v.case.case_version) + (v.label ? ' – ' + h(v.label) : '') + ' (kept ' + h(v.saved_at) + '; ' + v.records.length + ' results)</summary>' + summaryTable(v.records) + '</details>';
    }).join('');
  }

  // ------------------------------------------------------------------ panel 2: arithmetic
  function exposureOptions() { return [['', '— choose —']].concat(state.c.exposures.map(function (x) { return [x.id, x.id + ' – ' + (x.drug || '')]; })); }
  function eventOptions() { return [['', '— choose —']].concat(state.c.events.map(function (e) { return [e.id, e.id + ' – ' + (e.term || '')]; })); }
  function recordOptionsRaw(value, firstLabel, onlyExposure, adminOnly, includeMilestones) {
    var s = '<option value=""' + (!value ? ' selected' : '') + '>' + h(firstLabel) + '</option>';
    state.c.exposures.forEach(function (x) {
      if (onlyExposure && x.id !== onlyExposure) return;
      var recs = (x.records || []).filter(function (r) { return !adminOnly || ['start', 'administration', 'restart', 'dose_change'].indexOf(r.kind) >= 0; });
      if (!recs.length) return;
      s += '<optgroup label="' + h(x.drug || x.id) + '">' + recs.map(function (r) { return '<option value="' + h(r.id) + '"' + (value === r.id ? ' selected' : '') + '>' + h(r.id + ' – ' + (r.label || r.kind) + ' (' + E.fmtTiming(r.timing) + ')') + '</option>'; }).join('') + '</optgroup>';
    });
    if (includeMilestones && state.c.milestones.length) {
      s += '<optgroup label="Milestones">' + state.c.milestones.map(function (m) { return '<option value="' + h(m.id) + '"' + (value === m.id ? ' selected' : '') + '>' + h(m.id + ' – ' + (m.label || m.kind) + ' (' + E.fmtTiming(m.timing) + ')') + '</option>'; }).join('') + '</optgroup>';
    }
    return s;
  }
  function renderArith() {
    var c = state.c, s = '';
    s += '<h2>Arithmetic</h2><p class="lead">Intervals are calculated only for a drug episode and an event that you have linked. Choose each anchor yourself: every requested anchor gives its own labeled interval, the first-administration interval is always kept, and the tool never picks the shortest one.</p>';
    s += '<div data-out="validation"></div>';
    s += '<div class="card"><div class="card-head"><h3>Interval review</h3></div><div data-out="summary"></div></div>';
    // links
    s += '<div class="card"><div class="card-head"><h3>Drug_Event_Links</h3>' + btn('Add link', 'add', { list: 'links' }) + '</div>';
    if (!c.links.length) s += '<p class="small muted">No link yet. Without a link, an interval is blocked with E_MISSING_LINK.</p>';
    c.links.forEach(function (l, i) {
      var p = 'links.' + i;
      s += '<div class="grid">' + inputF('Link ID', p + '.id', { rerender: true }) + selectF('Drug exposure episode', p + '.exposure_id', exposureOptions(), { rerender: true }) +
        selectF('Event episode', p + '.event_id', eventOptions(), { rerender: true }) + inputF('Reason for the pairing', p + '.reason') +
        '<div class="field"><span class="lbl">&nbsp;</span>' + btn('Remove link', 'remove', { path: 'links', index: i }, 'small danger') + '</div></div>';
    });
    s += '</div>';
    // interval requests
    s += '<div class="card"><div class="card-head"><h3>Interval requests (C1–C6)</h3>' + btn('Add interval', 'add', { list: 'intervals' }) + '</div>' +
      '<p class="small muted">The rule is chosen from the precision of the two dates: two timestamps give C1 (minutes), two dates give C2 (calendar days), any partial date or range gives C4 (minimum and maximum).</p>';
    c.intervals.forEach(function (q, i) {
      var p = 'intervals.' + i, auto = q.anchor_type === 'FIRST_ADMIN' || q.anchor_type === 'LAST_DOSE_BEFORE_ONSET';
      var anchorRaw = recordOptionsRaw(q.anchor_ref || '', auto ? (q.anchor_type === 'FIRST_ADMIN' ? '(automatic: earliest start of this exposure)' : '(automatic under C5: latest administration at or before the end point)') : '— choose —',
        null, q.anchor_type === 'LAST_DOSE_BEFORE_ONSET', true);
      var endRaw = '<option value="onset"' + (!q.end_ref || q.end_ref === 'onset' ? ' selected' : '') + '>Event onset</option>' +
        state.c.milestones.map(function (m) { return '<option value="' + h(m.id) + '"' + (q.end_ref === m.id ? ' selected' : '') + '>' + h('Milestone ' + m.id + ' – ' + (m.label || '')) + '</option>'; }).join('');
      s += '<div class="card sub"><div class="grid">' + inputF('Interval ID', p + '.id', { rerender: true }) + selectF('Drug exposure episode', p + '.exposure_id', exposureOptions(), { rerender: true }) +
        selectF('Event episode', p + '.event_id', eventOptions(), { rerender: true }) +
        selectF('Anchor type (C5)', p + '.anchor_type', [['', '— choose —']].concat(E.ANCHOR_TYPES.map(function (a) { return [a, a]; })), { rerender: true }) +
        selectF('Anchor record', p + '.anchor_ref', null, { raw: anchorRaw }) +
        (q.anchor_type === 'OTHER' ? inputF('Description of the OTHER anchor', p + '.anchor_description') : '') +
        selectF('End point', p + '.end_ref', null, { raw: endRaw }) + inputF('Label (optional)', p + '.label') + '</div>' +
        '<div class="switch-row">' + checkF('Also show the calendar-date difference (C2) when both are timestamps', p + '.date_level') + checkF('Treatment-day number, inclusive convention (C3)', p + '.treatment_day') + '</div>' +
        '<details' + (q.reported && (q.reported.text || q.reported.lower) ? ' open' : '') + '><summary>Reported relative timing (C6)</summary><div class="grid">' +
        inputF('Reported text, as stated', p + '.reported.text') + inputF('Lower bound (only if the source states it)', p + '.reported.lower', { inputmode: 'decimal' }) +
        inputF('Upper bound', p + '.reported.upper', { inputmode: 'decimal' }) + selectF('Unit of the bounds', p + '.reported.unit', [['', '— choose —']].concat(E.REPORTED_UNITS.map(function (u) { return [u, u]; }))) + '</div>' +
        '<p class="small muted">Without numeric bounds nothing is compared (I_APPROXIMATE_NOT_COMPARED). Weeks, months and years are never converted to days.</p></details>' +
        '<details' + (q.constraint && q.constraint.target ? ' open' : '') + '><summary>Documented constraint on the bounds (C4)</summary><div class="grid">' +
        selectF('Constrains', p + '.constraint.target', [['', 'No constraint'], ['anchor', 'Anchor bounds'], ['onset', 'End-point bounds']]) +
        inputF('Not before (YYYY-MM-DD)', p + '.constraint.lower', { placeholder: 'YYYY-MM-DD' }) + inputF('Not after (YYYY-MM-DD)', p + '.constraint.upper', { placeholder: 'YYYY-MM-DD' }) +
        inputF('Source of the constraint', p + '.constraint.source') + '</div><p class="small muted">The constrained result is shown beside the unconstrained one, never instead of it.</p></details>' +
        '<div data-out="interval-' + i + '"></div><div class="row-actions">' + btn('Remove interval', 'remove', { path: 'intervals', index: i }, 'small danger') + '</div></div>';
    });
    s += '</div>';
    // recovery
    s += '<div class="card"><div class="card-head"><h3>Recovery intervals (C7)</h3>' + btn('Add recovery interval', 'add', { list: 'recoveries' }) + '</div>' +
      '<p class="small muted">From CESSATION to a recovery milestone (C1, C2 or C4 by the precision of the two records). Only milestones of kind &ldquo;Intervention or treatment&rdquo; count: one that can lie between cessation and recovery, ends included, raises W_INTERVENTION_BEFORE_RECOVERY. Nothing is inferred from other records, such as an exposure whose role is treatment of the event (CALC-SPEC 1.1; vectors V28&ndash;V30).</p>';
    c.recoveries.forEach(function (q, i) {
      var p = 'recoveries.' + i;
      s += '<div class="card sub"><div class="grid">' + inputF('Recovery ID', p + '.id', { rerender: true }) + selectF('Drug exposure episode', p + '.exposure_id', exposureOptions(), { rerender: true }) +
        selectF('Event episode', p + '.event_id', eventOptions(), { rerender: true }) +
        selectF('Cessation record', p + '.cessation_ref', null, { raw: recordOptionsRaw(q.cessation_ref || '', '— choose —', null, false, true) }) +
        selectF('Recovery milestone', p + '.recovery_ref', null, { raw: '<option value="">— choose —</option>' + state.c.milestones.map(function (m) { return '<option value="' + h(m.id) + '"' + (q.recovery_ref === m.id ? ' selected' : '') + '>' + h(m.id + ' – ' + (m.label || '')) + '</option>'; }).join('') }) +
        '</div><div data-out="recovery-' + i + '"></div><div class="row-actions">' + btn('Remove recovery interval', 'remove', { path: 'recoveries', index: i }, 'small danger') + '</div></div>';
    });
    s += '</div>';
    return s;
  }

  // ------------------------------------------------------------------ panel 3: technical mode
  function newTechnical(rule) {
    var t = { id: nextId('T'), rule: rule, label: '', inputs: {}, meta: {}, gate: { model: false, half_life: false, half_life_type: '', half_life_population: '', half_life_source: '', reference: false, reference_text: '', status: false },
      teaching_demonstration: false, creatinine_stable: false, stability_basis: '', assumptions: '' };
    (E.Q_PARAMS[rule] || []).forEach(function (p) { t.inputs[p.key] = ''; t.meta[p.key] = { status: '', source: '', version: '', section: '' }; });
    if (['Q1', 'Q2', 'Q3', 'Q5', 'Q6', 'Q8'].indexOf(rule) >= 0) t.inputs.unit = '';
    if (rule === 'Q4') t.inputs.cl_unit = '';
    if (rule === 'Q8') t.inputs.antibody = false;
    if (rule === 'Q9') { t.inputs = { segments: [{ start: emptyTs(), end: emptyTs(), rate: '', rate_unit: '' }], onset: emptyTs(), planned_dose_mg: '' }; t.meta = { segments: { status: '', source: '' }, onset: { status: '', source: '' } }; }
    if (rule === 'Q10') { t.inputs = { administrations: [{ date: '', dose: '', dose_unit: '', bsa: '' }], method_b: '', method_b_bsa: '' }; t.meta = { administrations: { status: '', source: '' } }; }
    if (rule === 'Q11') { t.inputs.scr_unit = ''; t.inputs.sex = ''; t.inputs.weight_type = ''; t.inputs.sample_date = ''; }
    if (rule === 'Q6') t.inputs.use = '';
    return t;
  }
  function paramRow(base, p, unitText) {
    return '<div class="grid">' + inputF(p.label + (unitText ? ' (' + unitText + ')' : ''), base + '.inputs.' + p.key, { inputmode: 'decimal' }) +
      selectF('Parameter_Status', base + '.meta.' + p.key + '.status', STATUS_OPTS) + inputF('Source document', base + '.meta.' + p.key + '.source') +
      inputF('Version', base + '.meta.' + p.key + '.version') + inputF('Section', base + '.meta.' + p.key + '.section') + '</div>';
  }
  function techCard(t, i) {
    var base = 'technical.' + i, rule = t.rule, s = '';
    s += '<div class="card sub"><div class="card-head"><h3>' + h(rule + ' – ' + E.RULE_TITLES[rule]) + '</h3>' + btn('Remove', 'remove', { path: 'technical', index: i }, 'small danger') + '</div>';
    s += '<div class="grid">' + inputF('Calculation ID', base + '.id', { rerender: true }) + inputF('Label', base + '.label') + '</div>';
    s += '<p class="small"><strong>Equation.</strong> ' + h(E.EQUATIONS[rule]) + '</p>';
    var hasTime = (E.Q_PARAMS[rule] || []).some(function (p) { return p.time; });
    if (hasTime) s += '<div class="grid">' + selectF('Time unit (one unit for every time input; mixed units are not converted)', base + '.inputs.unit', TIME_UNIT_OPTS) + '</div>';
    if (rule === 'Q6') s += '<div class="grid">' + selectF('Use (CALC-SPEC 1.1 names two; the formula is the same)', base + '.inputs.use', [['', '— choose —'], ['replacement or recovery', 'Replacement or recovery after withdrawal'], ['approach to a new level', 'Approach to a new level after a step change']]) + '</div>';
    if (rule === 'Q4') s += '<div class="grid">' + selectF('Clearance unit (both periods)', base + '.inputs.cl_unit', [['', '— choose —'], ['L/hour', 'L/hour'], ['mL/minute', 'mL/minute']]) + '</div>';
    if (rule === 'Q11') s += '<div class="grid">' + selectF('Serum creatinine unit', base + '.inputs.scr_unit', [['', '— choose —'], ['mg/dL', 'mg/dL'], ['µmol/L', 'µmol/L']]) +
      selectF('Sex as recorded', base + '.inputs.sex', [['', '— choose —'], ['male', 'male'], ['female', 'female'], ['not recorded', 'not recorded']]) +
      selectF('Weight type', base + '.inputs.weight_type', [['', '— choose —'], ['actual body weight', 'actual body weight'], ['ideal body weight', 'ideal body weight'], ['adjusted body weight', 'adjusted body weight'], ['other (state in assumptions)', 'other (state in assumptions)']]) +
      inputF('Creatinine sample date (YYYY-MM-DD)', base + '.inputs.sample_date') + '</div>';
    (E.Q_PARAMS[rule] || []).forEach(function (p) {
      var unitText = p.time ? (t.inputs.unit || 'time unit') : p.cl ? (t.inputs.cl_unit || 'clearance unit') : p.scr ? (t.inputs.scr_unit || 'unit') : p.unit;
      s += paramRow(base, p, unitText);
    });
    if (rule === 'Q8') s += checkF('The molecule is a monoclonal antibody: each bound carries the Q7 caution (W_BIOLOGIC_MODEL_CAUTION)', base + '.inputs.antibody');
    if (rule === 'Q9') {
      (t.inputs.segments || []).forEach(function (sg, j) {
        var sp = base + '.inputs.segments.' + j;
        s += '<div class="card"><h4>Segment ' + (j + 1) + '</h4>' + timingEditor('Start', sp + '.start', { tsOnly: true, noOriginal: true }) + timingEditor('End', sp + '.end', { tsOnly: true, noOriginal: true }) +
          '<div class="grid">' + inputF('Rate', sp + '.rate', { inputmode: 'decimal' }) + selectF('Rate unit', sp + '.rate_unit', [['', '— choose —'], ['mg/hour', 'mg/hour'], ['mL/hour', 'mL/hour (not supported in version 1.1)']]) + '</div>' +
          '<div class="row-actions">' + btn('Remove segment', 'remove', { path: sp.replace(/\.\d+$/, ''), index: j }, 'small danger') + '</div></div>';
      });
      s += '<div class="row-actions">' + btn('Add segment', 'add-segment', { index: i }) + '</div>';
      s += timingEditor('Onset', base + '.inputs.onset', { tsOnly: true, noOriginal: true });
      s += '<div class="grid">' + inputF('Planned dose (mg, optional)', base + '.inputs.planned_dose_mg', { inputmode: 'decimal' }) +
        selectF('Parameter_Status of the segments', base + '.meta.segments.status', STATUS_OPTS) + inputF('Source of the segments', base + '.meta.segments.source') + '</div>';
    }
    if (rule === 'Q10') {
      (t.inputs.administrations || []).forEach(function (a, j) {
        var ap = base + '.inputs.administrations.' + j;
        s += '<div class="grid">' + inputF('Date of administration ' + (j + 1) + ' (YYYY-MM-DD; for Method B by first or last)', ap + '.date', { placeholder: 'YYYY-MM-DD', inputmode: 'numeric' }) + inputF('Dose ' + (j + 1), ap + '.dose', { inputmode: 'decimal' }) + selectF('Dose unit', ap + '.dose_unit', [['', '— choose —'], ['mg', 'mg'], ['g', 'g (not supported)'], ['micrograms', 'micrograms (not supported)']]) +
          inputF('BSA at this administration (m²; blank if not recorded)', ap + '.bsa', { inputmode: 'decimal' }) +
          '<div class="field"><span class="lbl">&nbsp;</span>' + btn('Remove', 'remove', { path: base + '.inputs.administrations', index: j }, 'small danger') + '</div></div>';
      });
      s += '<div class="row-actions">' + btn('Add administration', 'add-admin', { index: i }) + '</div>';
      s += '<div class="grid">' + selectF('Method B (only on request)', base + '.inputs.method_b', [['', 'Not requested'], ['first', 'Total dose / first BSA'], ['last', 'Total dose / last BSA'], ['stated', 'Total dose / a stated BSA']], { rerender: true }) +
        (t.inputs.method_b === 'stated' ? inputF('Stated BSA (m²)', base + '.inputs.method_b_bsa', { inputmode: 'decimal' }) : '') +
        selectF('Parameter_Status of the administrations', base + '.meta.administrations.status', STATUS_OPTS) + inputF('Source', base + '.meta.administrations.source') + '</div>' +
        '<p class="small muted">Enter the BSA recorded at each administration. The tools never calculate BSA from height and weight or otherwise (CALC-SPEC 1.1). Method B by first or last uses the earliest and latest administration by date.</p>';
    }
    // gates
    if (E.COMMON_GATE_RULES.indexOf(rule) >= 0) {
      var hl = rule === 'Q4' ? 'The parameters used are identified' : 'The half-life used is identified';
      s += '<fieldset class="timing"><legend class="small"><strong>Common gate</strong> (each confirmation is your assertion and is exported in Assumptions)</legend>' +
        checkF('The simple model applies: one compartment, linear first-order elimination, constant parameters over the period', base + '.gate.model') +
        checkF(hl + ', with its type, population and source', base + '.gate.half_life') +
        '<div class="grid">' + inputF('Type (for example terminal half-life)', base + '.gate.half_life_type') + inputF('Population', base + '.gate.half_life_population') + inputF('Source (document, version, section)', base + '.gate.half_life_source') + '</div>' +
        checkF('The reference time is stated: the moment whose amount or concentration counts as 100%', base + '.gate.reference') +
        '<div class="grid">' + inputF('Reference time', base + '.gate.reference_text') + '</div>' +
        checkF('Each input’s Parameter_Status is recorded', base + '.gate.status') + '</fieldset>';
    }
    if (rule === 'Q6') s += checkF('Teaching-demonstration switch: hypothetical values only, never to date a recovery in a real case', base + '.teaching_demonstration');
    if (rule === 'Q11') s += checkF('Creatinine was stable when measured (confirm only on evidence)', base + '.creatinine_stable') + '<div class="grid">' + inputF('Basis for the stability confirmation', base + '.stability_basis') + '</div>';
    s += '<div class="grid wide">' + textareaF('Stated assumptions (exported)', base + '.assumptions') + '</div>';
    s += '<div data-out="technical-' + i + '"></div></div>';
    return s;
  }
  function renderTech() {
    var c = state.c, s = '<h2>Technical mode</h2><p class="lead">Optional, gated arithmetic (Q1–Q11). Use it only for a question the chronology cannot answer. Every result carries its equation, units, Parameter_Set, flags and mandatory text; a scenario range is never a confidence interval and never shows a midpoint.</p>';
    s += '<div class="card"><div class="switch-row">' + checkF('Technical mode on', 'technical_mode', { rerender: true }) + '</div>';
    if (!c.technical_mode) {
      s += '<p class="small muted">Technical mode is off. Chronology mode (Source facts and Arithmetic) needs no model. Technical results are neither shown nor exported while it is off.</p></div>';
      return s;
    }
    var ruleOpts = Object.keys(E.technical).map(function (r) { return [r, r + ' – ' + E.RULE_TITLES[r]]; });
    s += '<div class="grid"><label class="field" for="add-rule"><span class="lbl">Add a calculation</span><select id="add-rule">' + opts(ruleOpts, 'Q1') + '</select></label>' +
      '<div class="field"><span class="lbl">&nbsp;</span><button type="button" class="primary" data-action="add-technical">Add calculation</button></div></div></div>';
    if (!c.technical.length) s += '<p class="small muted">No calculation yet.</p>';
    c.technical.forEach(function (t, i) { s += techCard(t, i); });
    return s;
  }

  // ------------------------------------------------------------------ panel 4: assessment
  function ensureAssessment(linkId) {
    var a = state.c.assessment.filter(function (x) { return x && x.link_id === linkId; })[0];
    if (!a) {
      a = { link_id: linkId, fit: '', rationale: '', prompts: {}, followup: '' };
      PROMPTS.forEach(function (p) { a.prompts[p[0]] = { text: '', source: '' }; });
      state.c.assessment.push(a);
    }
    PROMPTS.forEach(function (p) { if (!a.prompts[p[0]]) a.prompts[p[0]] = { text: '', source: '' }; });
    return state.c.assessment.indexOf(a);
  }
  function renderAssess() {
    var c = state.c, s = '<h2>Assessment</h2><p class="lead">Reviewer judgment, written by you. The tool never pre-fills these fields from a result, and the biological-fit options are not relatedness categories. The causality conclusion stays outside the tool.</p>';
    if (!c.links.length) return s + '<p class="small muted">Add a Drug_Event_Link on the Arithmetic tab first.</p>';
    c.links.forEach(function (l) {
      var ai = ensureAssessment(l.id), p = 'assessment.' + ai, a = c.assessment[ai];
      var ex = c.exposures.filter(function (x) { return x.id === l.exposure_id; })[0], ev = c.events.filter(function (x) { return x.id === l.event_id; })[0];
      s += '<div class="card"><div class="card-head"><h3>' + h(l.id + ': ' + (ex ? ex.drug : l.exposure_id) + ' – ' + (ev ? ev.term : l.event_id)) + '</h3></div>';
      s += '<fieldset class="timing"><legend class="small"><strong>Biological fit</strong> (fit between the observed course and the stated hypothesis)</legend><div class="switch-row">' +
        [['', 'Not yet recorded']].concat(FIT_OPTIONS).map(function (o, k) {
          var id = fid(p + '.fit') + '-' + k;
          return '<label class="field check" for="' + id + '"><input type="radio" name="' + fid(p + '.fit') + '" id="' + id + '" data-path="' + h(p + '.fit') + '" value="' + h(o[0]) + '"' + ((a.fit || '') === o[0] ? ' checked' : '') + '><span class="lbl">' + h(o[1]) + '</span></label>';
        }).join('') + '</div></fieldset>';
      s += '<div class="grid wide">' + textareaF('Rationale (required when an option is chosen)', p + '.rationale', { rows: 3 }) + '</div><div data-out="assess-warn-' + ai + '"></div>';
      s += '<details><summary>PD and mechanism prompts (free text with a source)</summary>';
      PROMPTS.forEach(function (q) { s += '<div class="grid">' + textareaF(q[1], p + '.prompts.' + q[0] + '.text') + inputF('Source', p + '.prompts.' + q[0] + '.source') + '</div>'; });
      s += '<p class="small muted">Kinds of timing statement: typical onset or latency; risk window; maximum or distribution reported; monitoring or precaution window; definitional threshold. Only the first and third describe latency.</p></details>';
      s += '<div class="grid wide">' + textareaF('Follow-up questions and notes', p + '.followup') + '</div></div>';
    });
    return s;
  }

  // ------------------------------------------------------------------ panel 5: export
  function renderExportStatic() {
    return '<h2>Export</h2><p class="lead">Every export carries the tool version, the CALC-SPEC version, the export timestamp, the assumptions, the unrounded values beside the displayed ones, every flag and the limitation text.</p>' +
      '<div class="card"><div class="banner" role="note"><strong>Limitation (included in every export)</strong><p>' + h(E.LIMITATION_TEXT) + '</p></div>' +
      '<p class="small">' + h(E.VERSION_LABEL) + ' · Calculation_Version ' + h(E.SPEC_VERSION) + ' · Export_Timestamp is set when you export.</p>' +
      '<div class="row-actions no-print">' +
      '<button type="button" class="primary" data-action="copy-text">Copy as text</button>' +
      '<button type="button" data-action="dl-json">Download JSON</button>' +
      '<button type="button" data-action="dl-csv">Download CSV</button>' +
      '<button type="button" data-action="dl-html">Download HTML summary</button>' +
      '<button type="button" data-action="print">Print</button></div>' +
      '<p class="small muted no-print">Files are generated in this browser and saved by your browser&rsquo;s download. Free text is escaped, and CSV cells that begin with = + &minus; @ are prefixed with an apostrophe so that a spreadsheet does not run them as formulas. On a shared device, save through an approved route and press Reset before you leave.</p></div>' +
      '<div data-out="export"></div>';
  }

  // ------------------------------------------------------------------ verification
  function renderVerifyStatic() {
    return '<h2>Verification</h2><p class="lead">Runs CALC-SPEC 1.1 verification vectors V01–V38 against this calculator&rsquo;s engine, in this browser (V01–V27 are the unchanged CALC-SPEC 1.0 vectors; V28–V38 are new in 1.1). Unrounded values are compared with a relative tolerance of 1 × 10⁻⁶ (absolute 1 × 10⁻⁹ when the expected value is zero); units, sequence status and flags must match exactly. V12 and V14 describe behavior and are demonstrated automatically.</p>' +
      '<div class="card"><button type="button" class="primary" data-action="run-verify">Run all verification vectors</button> <span class="small muted">Vector file: ' + h(VEC ? VEC.spec + ', ' + VEC.vectors.length + ' vectors, dated ' + VEC.date : 'not loaded') + '</span></div>' +
      '<div data-out="verify"></div>';
  }
  function runVerify() {
    var res = VF.runAll(E, VEC);
    state.verify = res;
    renderVerifyOut();
    return res;
  }
  function renderVerifyOut() {
    var el = $('[data-out="verify"]'); if (!el) return;
    var res = state.verify;
    if (!res) { el.innerHTML = '<p class="small muted">Not run yet.</p>'; return; }
    var behav = res.behavioral.map(function (b) { return b.id + ' ' + b.status; }).join('; ');
    var ok = res.failed === 0 && res.behavioral.every(function (b) { return b.status === 'DEMONSTRATED'; });
    var s = '<p id="selftest-summary" class="summary-box ' + (ok ? 'ok' : 'bad') + '" data-passed="' + res.passed + '" data-numeric="' + res.numeric + '" data-result="' + (ok ? 'PASS' : 'FAIL') + '">' +
      'SELFTEST ' + (ok ? 'PASS' : 'FAIL') + ': ' + res.passed + '/' + res.numeric + ' numeric vectors PASS; behavioral ' + h(behav) + '. ' + h(E.VERSION_LABEL) + '.</p>';
    s += '<div class="table-wrap"><table><thead><tr><th scope="col">Vector</th><th scope="col">Rule</th><th scope="col">Status</th><th scope="col">Details</th></tr></thead><tbody>';
    res.results.forEach(function (r) {
      var det;
      if (r.type === 'behavioral') det = (r.notes || []).map(h).join('<br>');
      else if (r.problems.length) det = r.problems.map(h).join('<br>');
      else det = h(r.pairs.map(function (p) { return p[0] + ' = ' + JSON.stringify(p[1]); }).join('; ')) + (r.got && r.got.unit ? h('; unit = ' + r.got.unit) : '');
      s += '<tr><td>' + h(r.id) + '</td><td>' + h(r.rule) + '</td><td class="status-' + h(String(r.status).split(' ')[0]) + '">' + h(r.status) + '</td><td class="small">' + det + '</td></tr>';
    });
    el.innerHTML = s + '</tbody></table></div>';
    document.body.setAttribute('data-selftest', (ok ? 'PASS ' : 'FAIL ') + res.passed + '/' + res.numeric);
  }

  // ------------------------------------------------------------------ about
  function renderAbout() {
    var flags = Object.keys(E.MESSAGES).map(function (k) { return '<tr><td><code>' + h(k) + '</code></td><td>' + h(E.flagType(k)) + '</td><td>' + h(E.MESSAGES[k]) + '</td></tr>'; }).join('');
    return '<h2>About this calculator</h2>' +
      '<div class="two-col"><div class="card"><h3>What it does</h3><ul class="small"><li>Chronology intervals and bounds (C1–C7) with data-quality and sequence flags.</li><li>Optional technical arithmetic (Q1–Q5, Q7–Q11) with declared equations, units and assumptions, and a clearly labeled Q6 teaching demonstration (replacement or recovery, or the approach to a new level).</li><li>Reviewer-authored biological-fit judgment with a required rationale.</li><li>Text, HTML, CSV and JSON export; fictional demonstration cases; a reset function.</li></ul></div>' +
      '<div class="card"><h3>What it does not do</h3><ul class="small"><li>No causality scoring, and no category produced from timing (rule G8).</li><li>No patient-specific PK or PD simulation, no dosing recommendation, no diagnosis.</li><li>Not for regulatory submissions; not validated for regulated production use.</li><li>No built-in product numbers presented as applying to a patient.</li></ul></div></div>' +
      '<div class="card"><h3>Local operation and privacy</h3><ul class="small"><li>Calculation runs in this browser. The page makes no network requests (its content security policy forbids them), loads nothing from other sites and has no analytics.</li><li>It does not request direct patient identifiers. Do not enter them in free-text fields.</li><li>Nothing is saved silently: there is no browser storage. A case is kept only if you download a file (Save case file or Export). Reset clears everything on screen.</li><li>On a shared device, anyone using this browser window can see what is on screen until you reset or close it, and a downloaded file stays in the downloads folder.</li><li>Local operation does not remove your workplace data-handling obligations. Use fictional or synthetic data unless your organization&rsquo;s rules cover this tool.</li></ul></div>' +
      '<div class="card"><h3>Four states of a result</h3><ul class="small"><li><strong>Result</strong>: value with its unit, rule, anchor and sequence status.</li><li><strong>Result with a W-code</strong>: the warning is part of the result and goes wherever the number goes.</li><li><strong>Result with an I-code</strong>: an annotation, such as approximate timing kept as reported.</li><li><strong>Blocked by an E-code</strong>: no value; the message names what is missing or wrong. A blank is not a zero.</li></ul></div>' +
      '<div class="card"><h3>Flags and messages (CALC-SPEC 1.1, Section 5)</h3><div class="table-wrap"><table><thead><tr><th scope="col">Code</th><th scope="col">Type</th><th scope="col">Message shown</th></tr></thead><tbody>' + flags + '</tbody></table></div></div>' +
      '<div class="card"><h3>Units and display rounding (Sections 6 and 7)</h3><div class="table-wrap"><table><thead><tr><th scope="col">Output</th><th scope="col">Display</th></tr></thead><tbody>' +
      '<tr><td>Minutes (C1)</td><td>Whole minutes; hours (one decimal) and elapsed days (minutes ÷ 1,440, two decimals) are derived displays</td></tr>' +
      '<tr><td>Calendar days and treatment days</td><td>Whole days; never converted to hours</td></tr><tr><td>Fraction remaining or replaced</td><td>Percent to three significant figures; below 0.0001% in powers of ten</td></tr>' +
      '<tr><td>Accumulation factor and ratios</td><td>Two decimal places</td></tr><tr><td>Amounts in mg and mg/m²</td><td>One decimal place</td></tr>' +
      '<tr><td>Percent of the planned dose (Q9)</td><td>One decimal place (CALC-SPEC 1.1 Section 7; the unrounded value is always exported)</td></tr>' +
      '<tr><td>Estimated creatinine clearance</td><td>Whole mL/minute</td></tr><tr><td>Weeks, months, years</td><td>Stored as reported; never converted</td></tr>' +
      '<tr><td>Clearance</td><td>L/hour or mL/minute (1 L/hour = 16.667 mL/minute)</td></tr><tr><td>Serum creatinine</td><td>mg/dL or µmol/L (÷ 88.4)</td></tr></tbody></table></div>' +
      '<p class="small">Values are kept at full precision and rounded only for display; exports carry the unrounded value beside the displayed one.</p></div>' +
      '<div class="card"><h3>What is new in version 1.1 (CALC-SPEC 1.1)</h3><ul class="small">' +
      '<li><strong>Laboratory-defined onset (C4).</strong> An event onset can be entered as &ldquo;Laboratory-defined&rdquo;: the lower bound is the date of the last normal test and the upper bound the date of the first abnormal test (vector V31). You choose which results are normal. Case C28 now uses it.</li>' +
      '<li><strong>C7 is verified</strong> by V28 (minute level, as in C01), V29 (date level, as in C04) and V30, a labeled no-flag variant.</li>' +
      '<li><strong>Q6 has two named uses</strong> with one formula: replacement or recovery, and the approach to a new level (V32). You choose the use.</li>' +
      '<li><strong>BSA is never calculated.</strong> Q10 takes the BSA recorded at each administration.</li></ul></div>' +
      '<div class="card"><h3>Agreed rules where the calculator and the workbook differed (CALC-SPEC 1.1, Section 9)</h3><p class="small">The stricter behavior was chosen: block rather than guess, list every error, compare in UTC when offsets are known, and never infer an unrecorded intervention. Both tools at version 1.1 follow these rules.</p><ul class="small">' +
      '<li><strong>Rule choice.</strong> The finest rule both records support (C1, C2 or C4); a coarser one on request (the date-level C2 box); never added precision.</li>' +
      '<li><strong>LAST_DOSE_BEFORE_ONSET (C5).</strong> Only administrations with an exact date or time are candidates, ordered in UTC when offsets are known (V33). The result is blocked with E_MISSING_INPUT, so that you choose the anchor, when the last dose cannot be identified: a partial-date administration that could be the last dose (V34), an administration whose order relative to onset is uncertain, or two that cannot be told apart.</li>' +
      '<li><strong>C7 interventions.</strong> Only milestones of kind &ldquo;Intervention or treatment&rdquo; count, when they can lie between cessation and recovery, ends included (UTC when offsets are known). An exposure whose role is treatment of the event is pointed out in Assumptions but never counted (V30).</li>' +
      '<li><strong>Common gate.</strong> Confirmed only with all four confirmations and every detail (half-life type, population and source; reference time; each input&rsquo;s Parameter_Status). Each blank detail gives E_MISSING_INPUT, and the unconfirmed gate gives E_APPLICABILITY_UNCONFIRMED; both are listed (V36).</li>' +
      '<li><strong>Every error is listed</strong>, each code once, E-codes first, then W- and I-codes, in the order of the flags table below (V35).</li>' +
      '<li><strong>Q8 for a monoclonal antibody</strong> (tick box) carries W_BIOLOGIC_MODEL_CAUTION beside I_SCENARIO_NOT_CI (V37).</li>' +
      '<li><strong>Q10 Method B</strong> by first or last administration uses the earliest and latest administration by date; blocked when only some dates are given or when two administrations share that date with different BSAs; with no dates at all, the order entered is used and stated (V38).</li>' +
      '<li><strong>Units.</strong> A blank unit gives E_MISSING_INPUT; an unsupported unit gives E_INVALID_UNIT.</li>' +
      '<li><strong>Display.</strong> Percentages below 0.0001% are shown in powers of ten (for example 3.41 × 10⁻²⁰%).</li>' +
      '<li><strong>Export.</strong> Every record carries Tool_Version, Calculation_Version and Export_Timestamp, in the CALC-SPEC 1.1 Section 8 columns.</li></ul></div>' +
      '<div class="card"><h3>Conventions recorded in CALC-SPEC 1.1</h3><ul class="small">' +
      '<li>Each technical calculation takes one time unit for all its time inputs; mixed units are not converted. Q7 accepts days only.</li>' +
      '<li>Reported timing in minutes, hours or days is compared with a C1 interval; against a calendar-day interval only calendar days are compared; weeks, months or years are never converted. Otherwise it is kept as reported with I_APPROXIMATE_NOT_COMPARED.</li>' +
      '<li>A documented constraint is shown as a second, &ldquo;-constrained&rdquo; record with its Constraint_Source, beside the unconstrained one.</li>' +
      '<li>Interval records carry the CALC-SPEC 1.1 limitations text: the interval describes sequence and spacing, not causation.</li>' +
      '<li>Not yet in the specification (future work): END_OF_ADMINISTRATION, INTERVENTION and interruption or restart anchor types; a dose-reduction input; comparison with label windows stated in weeks. Use OTHER with a description, a DOSE_CHANGE record, or your own judgment meanwhile.</li></ul></div>' +
      '<div class="card"><h3>Version</h3><p class="small">' + h(E.VERSION_LABEL) + '. Companion to <em>Time to Onset in Drug Safety: A Technical Guide to ICSR Analysis Using Mechanism of Action, Pharmacokinetics and Pharmacodynamics</em> by Dr. Hafez Selim. License CC BY-NC 4.0.</p></div>';
  }

  // ------------------------------------------------------------------ rendering
  function renderAll() {
    var active = document.activeElement, activeId = active && active.id;
    $('#panel-facts').innerHTML = renderFacts();
    $('#panel-arith').innerHTML = renderArith();
    $('#panel-tech').innerHTML = renderTech();
    $('#panel-assess').innerHTML = renderAssess();
    if (!$('#panel-export').innerHTML) $('#panel-export').innerHTML = renderExportStatic();
    if (!$('#panel-verify').innerHTML) { $('#panel-verify').innerHTML = renderVerifyStatic(); renderVerifyOut(); }
    if (!$('#panel-about').innerHTML) $('#panel-about').innerHTML = renderAbout();
    renderOutputs();
    if (activeId) { var el = document.getElementById(activeId); if (el && el !== document.activeElement) { try { el.focus({ preventScroll: true }); } catch (e) { el.focus(); } } }
  }
  function renderOutputs() {
    var c = state.c, idx = E.indexCase(c);
    var ev = E.evaluateCase(c, { technical: c.technical_mode === true });
    // readouts
    $all('[data-readout]').forEach(function (el) { var r = readoutFor(getPath(c, el.getAttribute('data-readout'))); el.textContent = r.text; el.classList.toggle('bad', r.bad); });
    // timeline
    $all('[data-out="timeline"]').forEach(function (el) { el.innerHTML = timelineHTML(); applyPositions(el); });
    renderVersionHistory();
    // validation
    var vEl = $('[data-out="validation"]');
    if (vEl) vEl.innerHTML = ev.validation.length ? '<div class="card"><h3>Validation messages</h3>' + ev.validation.map(function (v) { return flagHTML({ code: v.code, type: v.type, message: v.message + ' (' + v.where + ')' }); }).join('') + '</div>' : '';
    // per request results
    c.intervals.forEach(function (q, i) { var el = $('[data-out="interval-' + i + '"]'); if (el) el.innerHTML = E.evaluateInterval(c, q, idx).map(resultHTML).join(''); });
    c.recoveries.forEach(function (q, i) { var el = $('[data-out="recovery-' + i + '"]'); if (el) el.innerHTML = E.evaluateRecovery(c, q, idx).map(resultHTML).join(''); });
    c.technical.forEach(function (t, i) { var el = $('[data-out="technical-' + i + '"]'); if (el) el.innerHTML = E.evaluateTechnical(t).map(resultHTML).join(''); });
    var sEl = $('[data-out="summary"]');
    if (sEl) sEl.innerHTML = summaryTable(ev.records.filter(function (r) { return !/^Q/.test(r.Calculation_Mode); }));
    // assessment warnings
    c.assessment.forEach(function (a, i) {
      var el = $('[data-out="assess-warn-' + i + '"]'); if (!el) return;
      el.innerHTML = a.fit && !String(a.rationale || '').trim() ? flagHTML({ code: 'RATIONALE', type: 'Warning', message: 'A rationale is required for the biological-fit option. It is exported as missing until written.' }).replace('flag R', 'flag W') : '';
    });
    renderExportOut(ev);
  }
  function renderExportOut(ev) {
    var el = $('[data-out="export"]'); if (!el) return;
    var bundle = currentBundle('(set at export)');
    var unresolved = bundle.unresolved_flags;
    var s = '<div class="card"><h3>Records (Section 8 output record)</h3>' + summaryTable(bundle.records) + '</div>';
    s += '<div class="card"><h3>Unresolved flags</h3>' + (unresolved.length ? '<ul class="small">' + unresolved.map(function (u) { return '<li><code>' + h(u.code) + '</code> (' + h(u.result) + '): ' + h(u.message) + '</li>'; }).join('') + '</ul>' : '<p class="small muted">None.</p>') + '</div>';
    s += '<div class="card"><h3>Text preview</h3><pre class="export" tabindex="0">' + h(E.toText(bundle)) + '</pre></div>';
    if (state.importedCSV) {
      s += '<div class="card"><h3>CSV read back: ' + h(state.importedCSV.name) + '</h3><p class="small muted">Records re-imported from a CSV export, shown as stored (read only).</p><div class="table-wrap"><table><thead><tr><th>Result_ID</th><th>Label</th><th>Min / value / max</th><th>Unit</th><th>Sequence</th><th>Flags</th><th>Tool_Version</th></tr></thead><tbody>' +
        state.importedCSV.records.map(function (r) { return '<tr><td>' + h(r.Result_ID) + '</td><td>' + h(r.Label) + '</td><td class="num">' + h([r.Interval_Minimum, r.Interval_Value, r.Interval_Maximum].map(function (x) { return x === null ? 'blank' : x; }).join(' / ')) + '</td><td>' + h(r.Unit) + '</td><td>' + h(r.Sequence_Status) + '</td><td>' + h((r.QC_Flags || []).join(', ')) + '</td><td>' + h(r.Tool_Version) + '</td></tr>'; }).join('') +
        '</tbody></table></div></div>';
    }
    el.innerHTML = s;
  }
  function currentBundle(ts) {
    var c = E.clone(state.c);
    if (!c.technical_mode) c.technical = [];
    return E.buildExport(c, { timestamp: ts || new Date().toISOString(), versions: state.versions });
  }

  // ------------------------------------------------------------------ actions
  function fileBase() { return (String(state.c.case_id || 'case').replace(/[^A-Za-z0-9_-]+/g, '_') || 'case') + '_v' + String(state.c.case_version || '1').replace(/[^A-Za-z0-9_-]+/g, '_'); }
  function download(name, mime, text) {
    var blob = new Blob([text], { type: mime }), url = URL.createObjectURL(blob), a = document.createElement('a');
    a.href = url; a.download = name; a.rel = 'noopener'; document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.parentNode.removeChild(a); }, 1500);
  }
  function copyText(text) {
    function fallback() {
      var ta = document.createElement('textarea'); ta.value = text; ta.setAttribute('readonly', ''); ta.className = 'sr-only'; document.body.appendChild(ta); ta.select();
      var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      document.body.removeChild(ta); toast(ok ? 'Copied the text record, with the limitation text and versions.' : 'Copy is not available here; select the text preview and copy it.');
    }
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(text).then(function () { toast('Copied the text record, with the limitation text and versions.'); }, fallback);
    else fallback();
  }
  function confirmReplace(what) {
    if (!state.dirty) return true;
    return window.confirm('Replace the case on screen with ' + what + '? Unsaved entries will be lost. Use Save case file first to keep them.');
  }
  function loadExample(id, quiet) {
    var ex = EX.cases[id]; if (!ex) return;
    if (!quiet && !confirmReplace('the fictional example ' + id)) return;
    state.c = E.clone(ex); state.c.technical_mode = (state.c.technical || []).length > 0; state.versions = []; state.dirty = false; state.importedCSV = null;
    renderAll();
    var meta = EX.list.filter(function (x) { return x.id === id; })[0];
    if (!quiet) toast('Loaded fictional example ' + (meta ? meta.title + ' (' + meta.chapter + ')' : id) + '. Dates are synthetic or fictional.');
  }
  function addItem(list) {
    var c = state.c, item;
    if (list === 'events') item = { id: nextId('EV'), term: '', source: '', onset: { type: 'none', original: '' } };
    if (list === 'milestones') item = { id: nextId('MS'), kind: '', label: '', value: '', source: '', timing: { type: 'none', original: '' } };
    if (list === 'exposures') item = { id: nextId('EX'), drug: '', role: '', details: '', source: '', records: [] };
    if (list === 'links') item = { id: nextId('L'), exposure_id: '', event_id: '', reason: '' };
    if (list === 'intervals') item = { id: nextId('I'), exposure_id: c.exposures.length === 1 ? c.exposures[0].id : '', event_id: c.events.length === 1 ? c.events[0].id : '', anchor_type: '', anchor_ref: '', end_ref: 'onset', date_level: false, treatment_day: false,
      reported: { text: '', lower: '', upper: '', unit: '' }, constraint: { target: '', lower: '', upper: '', source: '' } };
    if (list === 'recoveries') item = { id: nextId('R'), exposure_id: c.exposures.length === 1 ? c.exposures[0].id : '', event_id: c.events.length === 1 ? c.events[0].id : '', cessation_ref: '', recovery_ref: '' };
    c[list].push(item);
  }
  function onAction(btnEl) {
    var a = btnEl.getAttribute('data-action'), c = state.c;
    switch (a) {
      case 'add': addItem(btnEl.getAttribute('data-list')); state.dirty = true; renderAll(); break;
      case 'add-record': {
        var x = c.exposures[+btnEl.getAttribute('data-index')];
        x.records = x.records || []; x.records.push({ id: nextId(x.id + '-R'), kind: '', label: '', timing: { type: 'none', original: '' } });
        state.dirty = true; renderAll(); break;
      }
      case 'add-segment': c.technical[+btnEl.getAttribute('data-index')].inputs.segments.push({ start: emptyTs(), end: emptyTs(), rate: '', rate_unit: '' }); renderAll(); break;
      case 'add-admin': c.technical[+btnEl.getAttribute('data-index')].inputs.administrations.push({ date: '', dose: '', dose_unit: '', bsa: '' }); renderAll(); break;
      case 'remove': {
        var arr = getPath(c, btnEl.getAttribute('data-path')); arr.splice(+btnEl.getAttribute('data-index'), 1);
        state.dirty = true; renderAll(); break;
      }
      case 'add-technical': c.technical.push(newTechnical($('#add-rule').value)); state.dirty = true; renderAll(); break;
      case 'open-file': $('#file-input').click(); break;
      case 'save-case': download(fileBase() + '_tto_case.json', 'application/json', JSON.stringify(currentBundle(), null, 2)); state.dirty = false; toast('Case file generated for download. It contains the case inputs, results, version history and the limitation text.'); break;
      case 'new-version': {
        var snap = VF.snapshotVersion(E, c, ''); state.versions.push(snap);
        state.c = VF.newVersionFrom(E, c); renderAll();
        toast('Case_Version ' + snap.case.case_version + ' is kept read only with its results (Source facts, Version history). You are now editing Case_Version ' + state.c.case_version + '.');
        break;
      }
      case 'reset':
        if (!window.confirm('Reset clears the case, its versions and any imported file from this page. Nothing is stored in the browser. Continue?')) return;
        state.c = blankCase(); state.versions = []; state.dirty = false; state.importedCSV = null; $('#file-input').value = ''; renderAll(); toast('Reset: the page holds no case data now.'); break;
      case 'copy-text': copyText(E.toText(currentBundle())); break;
      case 'dl-json': download(fileBase() + '_tto_export.json', 'application/json', JSON.stringify(currentBundle(), null, 2)); break;
      case 'dl-csv': download(fileBase() + '_tto_export.csv', 'text/csv;charset=utf-8', '﻿' + E.toCSV(currentBundle().records)); break;
      case 'dl-html': download(fileBase() + '_tto_summary.html', 'text/html;charset=utf-8', E.toHTML(currentBundle())); break;
      case 'print': selectTab('export'); window.print(); break;
      case 'run-verify': runVerify(); break;
    }
  }
  function onFile(file) {
    if (!file) return;
    var reader = new FileReader();
    reader.onload = function () {
      var text = String(reader.result || '');
      try {
        if (/\.csv$/i.test(file.name) || /^﻿?Case_ID,/.test(text)) {
          state.importedCSV = { name: file.name, records: E.recordsFromCSV(text.replace(/^﻿/, '')) };
          renderOutputs(); selectTab('export'); toast('Read ' + state.importedCSV.records.length + ' records from ' + file.name + ' (shown on the Export tab).');
        } else {
          var obj = JSON.parse(text);
          if (!confirmReplace('the case in ' + file.name)) return;
          var imp = E.importBundle(obj);
          state.c = imp.case; if (state.c.technical_mode === undefined) state.c.technical_mode = (state.c.technical || []).length > 0;
          state.versions = imp.versions || []; state.dirty = false; renderAll(); toast('Opened ' + file.name + ': case ' + (state.c.case_id || '') + ', version ' + (state.c.case_version || '') + '.');
        }
      } catch (e) { toast('Could not read ' + file.name + ': ' + e.message); }
      $('#file-input').value = '';
    };
    reader.readAsText(file);
  }
  function selectTab(name, focus) {
    if (TABS.indexOf(name) < 0) name = 'facts';
    state.tab = name;
    TABS.forEach(function (t) {
      var b = document.getElementById(TAB_IDS[t]), p = document.getElementById('panel-' + t), on = t === name;
      b.setAttribute('aria-selected', on ? 'true' : 'false'); b.tabIndex = on ? 0 : -1; p.hidden = !on;
    });
    if (focus) document.getElementById(TAB_IDS[name]).focus();
  }

  // ------------------------------------------------------------------ events
  var timer = null;
  function valueOf(el) {
    if (el.type === 'checkbox') return el.checked;
    return el.value;
  }
  document.addEventListener('input', function (e) {
    var el = e.target; if (!el.getAttribute || !el.getAttribute('data-path')) return;
    if (el.type === 'radio' && !el.checked) return;
    setPath(state.c, el.getAttribute('data-path'), valueOf(el)); state.dirty = true;
    if (el.getAttribute('data-rerender') && el.tagName === 'INPUT' && el.type === 'text') return;   // rerender on change, not on every keystroke
    if (el.getAttribute('data-rerender')) { renderAll(); return; }
    clearTimeout(timer); timer = setTimeout(renderOutputs, 120);
  });
  document.addEventListener('change', function (e) {
    var el = e.target;
    if (el.id === 'file-input') { onFile(el.files && el.files[0]); return; }
    if (el.id === 'example-select') { var v = el.value; el.value = ''; if (v) loadExample(v); return; }
    if (!el.getAttribute || !el.getAttribute('data-path')) return;
    if (el.type === 'radio' && !el.checked) return;
    setPath(state.c, el.getAttribute('data-path'), valueOf(el)); state.dirty = true;
    if (el.getAttribute('data-rerender')) setTimeout(renderAll, 0); else renderOutputs();
  });
  document.addEventListener('click', function (e) {
    var b = e.target.closest ? e.target.closest('[data-action]') : null;
    if (b) { onAction(b); return; }
    var tab = e.target.closest ? e.target.closest('[role=tab]') : null;
    if (tab) { var name = tab.id.replace('tab-', ''); selectTab(name); }
  });
  document.querySelector('nav.tabs').addEventListener('keydown', function (e) {
    var i = TABS.indexOf(state.tab);
    if (e.key === 'ArrowRight') { selectTab(TABS[(i + 1) % TABS.length], true); e.preventDefault(); }
    if (e.key === 'ArrowLeft') { selectTab(TABS[(i - 1 + TABS.length) % TABS.length], true); e.preventDefault(); }
    if (e.key === 'Home') { selectTab(TABS[0], true); e.preventDefault(); }
    if (e.key === 'End') { selectTab(TABS[TABS.length - 1], true); e.preventDefault(); }
  });

  // ------------------------------------------------------------------ start
  function init() {
    $('#version-label').textContent = E.VERSION_LABEL;
    var sel = $('#example-select');
    EX.list.forEach(function (x) { var o = document.createElement('option'); o.value = x.id; o.textContent = x.title + ' – ' + x.chapter; sel.appendChild(o); });
    var params = {};
    (location.search || '').replace(/^\?/, '').split('&').forEach(function (kv) { if (!kv) return; var p = kv.split('='); params[decodeURIComponent(p[0])] = decodeURIComponent(p[1] || ''); });
    var hash = (location.hash || '').replace('#', '');
    if (params.example && EX.cases[params.example]) loadExample(params.example, true); else renderAll();
    var tab = params.tab || (TABS.indexOf(hash) >= 0 ? hash : 'facts');
    if (params.selftest !== undefined || hash === 'verify') { tab = params.tab || 'verify'; runVerify(); }
    selectTab(tab);
  }
  init();
}());
