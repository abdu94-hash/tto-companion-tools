/*
 * TTO Companion Tools v1.0 - calculation engine implementing CALC-SPEC 1.0
 * Companion to "Time to Onset in Drug Safety" (Dr. Hafez Selim). License: CC BY-NC 4.0.
 *
 * Pure functions, no dependencies, no network access. Loads in the browser (global TTOCalc)
 * and in Node (require). Interval rules C1-C7, technical calculations Q1-Q11, flags (Section 5),
 * units (Section 6), display rounding (Section 7) and the output record (Section 8).
 *
 * Educational tool. Not validated for regulated production use. Not a dosing calculator.
 * It never assigns causality or produces a causality category (rule G8).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) { module.exports = factory(); }
  else { root.TTOCalc = factory(); }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ------------------------------------------------------------------ versions and fixed text
  var SPEC_VERSION = 'CALC-SPEC 1.0';
  var TOOL_VERSION = 'TTO Companion Tools v1.0';
  var VERSION_LABEL = 'TTO Companion Tools v1.0 (CALC-SPEC 1.0)';
  var LIMITATION_TEXT = 'Educational tool, not validated for regulated production use. It is not a dosing ' +
    'calculator and it does not assess causality: no result is a causality category, a probability, or a ' +
    'statement that an event is related to a drug. Results are arithmetic on the facts entered, under the ' +
    'stated rules and assumptions, and must be read with their flags and mandatory text.';

  // Section 5, verbatim. [input] is filled for E_MISSING_INPUT.
  var MESSAGES = {
    E_INVALID_DATE: 'This date does not exist. Check the source, or record it as a partial date.',
    E_REVERSED_BOUNDS: 'The earliest value is after the latest value. Check the order of the bounds.',
    E_DUPLICATE_ID: 'This identifier is already used in this case. Each episode needs its own ID.',
    E_MISSING_LINK: 'Link this drug episode to this event before calculating an interval.',
    E_MISSING_INPUT: 'This result needs [input]. It is left blank rather than estimated.',
    E_INVALID_UNIT: 'This unit is not supported here. Use one of the listed units.',
    E_NONPOSITIVE_PARAMETER: 'This value must be greater than zero.',
    E_OUT_OF_RANGE: 'This value is outside the range the equation accepts.',
    E_APPLICABILITY_UNCONFIRMED: 'Confirm the model, the half-life used, the reference time and each input’s status before a result is shown.',
    E_CREATININE_STABILITY_UNCONFIRMED: 'Cockcroft–Gault assumes stable creatinine. Confirm stability before an estimate is shown.',
    E_OVERLAPPING_SEGMENTS: 'Two infusion segments overlap. Check the recorded start and stop times.',
    W_TZ_UNKNOWN: 'Offset unknown: this interval uses local clock times and could be wrong across a clock change or time zone.',
    W_SAME_DATE_ORDER_UNKNOWN: 'Same calendar date: the order within the day is unknown.',
    W_SEQUENCE_UNCERTAIN: 'The available dates allow onset before or after this anchor. Sequence is not established.',
    W_ADMIN_SAME_DATE_AS_ONSET: 'This dose was on the onset date; whether it came before onset is unknown.',
    W_REPORTED_TIMING_CONFLICT: 'The reported timing conflicts with the recorded dates. Both are kept; query the source.',
    W_INTERVENTION_BEFORE_RECOVERY: 'Treatment was given before recovery, so recovery cannot be attributed to withdrawal alone.',
    W_BIOLOGIC_MODEL_CAUTION: 'Antibody disposition can be nonlinear or time-varying; this simple model may not apply.',
    W_SEX_NOT_RECORDED: 'Sex not recorded: both estimates are shown.',
    W_PARTIAL_BSA: 'BSA missing for some administrations: the value shown is a lower bound.',
    I_TREATMENT_DAY_UNDEFINED: 'No treatment-day number exists before the first dose.',
    I_APPROXIMATE_NOT_COMPARED: 'Approximate timing is kept as reported and not converted.',
    I_SCENARIO_NOT_CI: 'Scenario range from assumed parameters, not a confidence interval.',
    I_TEACHING_DEMONSTRATION: 'Teaching demonstration with hypothetical values.'
  };

  // Section 4, mandatory text for each technical rule.
  var MANDATORY_TEXT = {
    Q1: 'Fraction of the chosen reference amount or concentration under the stated model. It is not a causal or pharmacological threshold.',
    Q2: 'Assumes constant input from zero. Intermittent dosing compares corresponding points in the dosing cycle. Toxicity need not wait for steady state.',
    Q3: 'Model-based repeated-dose comparison. It is not a toxicity multiplier and does not give the time of clinical onset.',
    Q4: 'Describes the eventual average steady state. It does not date the patient’s onset; the transition depends on prior concentrations and when clearance changed.',
    Q5: 'Assumes no further formation or input. Continued metabolite formation, potency, distribution and nonlinear kinetics can make this calculation invalid.',
    Q6: 'Hypothetical turnover illustration. It is not a measured recovery curve for platelets or any other clinical target.',
    Q7: 'Antibody disposition can depend on dose, time, target burden and immunogenicity. Exposure persistence is not receptor occupancy, immune activity or biological persistence.',
    Q8: 'Assumption-based scenario range, not a confidence interval.',
    Q9: 'Amount infused, not systemic or effect-site exposure. Assumes the recorded times without interruption and ignores line dead volume.',
    Q10: 'Describes cumulative exposure. It is not a toxicity threshold.',
    Q11: 'Assumes stable creatinine. It is not drug clearance and not a dosing recommendation.'
  };

  var RULE_TITLES = {
    C1: 'Exact elapsed interval', C2: 'Calendar-date difference', C3: 'Treatment-day number',
    C4: 'Bounded interval', C5: 'Anchor selection', C6: 'Reported timing comparison', C7: 'Recovery interval',
    Q1: 'Elimination and fractional persistence', Q2: 'Time to approach steady state',
    Q3: 'Accumulation during regular dosing', Q4: 'A change in clearance', Q5: 'Parent and metabolite persistence',
    Q6: 'Biological replacement and recovery (teaching demonstration)', Q7: 'Long persistence of a monoclonal antibody',
    Q8: 'Parameter uncertainty and scenarios', Q9: 'Dose received before onset during an infusion',
    Q10: 'Cumulative dose normalized to body surface area', Q11: 'Estimated renal function (Cockcroft–Gault)'
  };

  var EQUATIONS = {
    C1: 'interval = onset (UTC) − anchor (UTC), in minutes; local clock times when an offset is unknown',
    C2: 'interval = onset date − anchor date, in calendar days',
    C3: 'treatment day = C2 + 1 (inclusive convention), only when C2 ≥ 0',
    C4: 'minimum = O_lower − A_upper; maximum = O_upper − A_lower, in calendar days',
    C7: 'recovery interval = recovery milestone − cessation (C1 or C2)',
    Q1: 'k = ln 2 / t½; fraction remaining = 2^(−t / t½) = exp(−k × t)',
    Q2: 'n = −log2(1 − f); time = n × t½',
    Q3: 'R = 1 / (1 − 2^(−τ / t½))',
    Q4: 'half-life ratio = average Css ratio = CL_before / CL_after; with V: t½ = ln 2 × V / CL',
    Q5: 'parent and metabolite each: fraction = 2^(−t / t½), each normalized to its own starting value',
    Q6: 'fraction replaced = 1 − 2^(−t / replacement half-time)',
    Q7: 'fraction remaining = 2^(−t / t½), t and t½ in days',
    Q8: 'fraction at each bound = 2^(−t / t½_lower) and 2^(−t / t½_upper); no midpoint',
    Q9: 'amount = Σ rate × hours of each segment before onset',
    Q10: 'Method A: Σ dose_i / BSA_i; Method B (on request): total dose / one stated BSA',
    Q11: 'CrCl = (140 − age) × weight / (72 × SCr), × 0.85 when female sex is recorded; µmol/L ÷ 88.4 = mg/dL'
  };

  var ANCHOR_TYPES = ['FIRST_ADMIN', 'LAST_DOSE_BEFORE_ONSET', 'DOSE_CHANGE', 'RESTART', 'CONCOMITANT_CHANGE', 'CESSATION', 'OTHER'];
  var SEQUENCE_STATUSES = ['AFTER', 'BEFORE', 'SIMULTANEOUS', 'SAME_DATE_ORDER_UNKNOWN', 'SEQUENCE_UNCERTAIN'];
  var PARAMETER_STATUSES = ['label fact', 'study observation', 'case measurement', 'model assumption', 'hypothetical example'];
  var TIME_UNITS = { minutes: 1 / 60, hours: 1, days: 24 };           // factor to hours (Section 6)
  var REPORTED_UNITS = ['minutes', 'hours', 'days', 'calendar days', 'weeks', 'months', 'years'];
  var CLEARANCE_UNITS = { 'L/hour': 1, 'mL/minute': 60 / 1000 };      // factor to L/hour; 1 L/hour = 16.667 mL/minute
  var SCR_UNITS = { 'mg/dL': 1, 'umol/L': 1 / 88.4, 'µmol/L': 1 / 88.4, 'μmol/L': 1 / 88.4 };
  var MINUS = '−';

  function flagType(code) {
    return code.charAt(0) === 'E' ? 'Error' : code.charAt(0) === 'W' ? 'Warning' : 'Info';
  }
  function message(code, detail) {
    var m = MESSAGES[code] || code;
    return m.replace('[input]', detail || 'an input');
  }

  // ------------------------------------------------------------------ errors
  function CalcError(codes, details) {
    this.name = 'CalcError';
    this.codes = Array.isArray(codes) ? codes : [codes];
    this.details = details || {};
    this.message = this.codes.join(', ');
  }
  CalcError.prototype = Object.create(Error.prototype);
  CalcError.prototype.constructor = CalcError;

  function fail(code, detail) {
    var d = {}; if (detail) d[code] = detail;
    throw new CalcError([code], d);
  }
  function Errs() { this.codes = []; this.details = {}; }
  Errs.prototype.add = function (code, detail) {
    if (this.codes.indexOf(code) < 0) { this.codes.push(code); if (detail) this.details[code] = detail; }
    else if (detail && code === 'E_MISSING_INPUT' && this.details[code].indexOf(detail) < 0) {
      this.details[code] += ', ' + detail;
    }
  };
  Errs.prototype.absorb = function (e) {
    if (!(e instanceof CalcError)) throw e;
    for (var i = 0; i < e.codes.length; i++) this.add(e.codes[i], e.details[e.codes[i]]);
  };
  Errs.prototype.any = function () { return this.codes.length > 0; };
  Errs.prototype.raise = function () { if (this.codes.length) throw new CalcError(this.codes.slice(), this.details); };

  // ------------------------------------------------------------------ dates and times (Section 2)
  var MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September',
    'October', 'November', 'December'];

  function dayNumber(y, m, d) { return Math.round(Date.UTC(y, m - 1, d) / 86400000); }
  function fromDayNumber(n) {
    var dt = new Date(n * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  }
  function isValidYMD(y, m, d) {
    if (m < 1 || m > 12 || d < 1 || y < 1 || y > 9999) return false;
    var x = fromDayNumber(dayNumber(y, m, d));
    return x.y === y && x.m === m && x.d === d;
  }
  function daysInMonth(y, m) { return fromDayNumber(dayNumber(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1, 1) - 1).d; }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function isoDay(n) { var x = fromDayNumber(n); return x.y + '-' + pad2(x.m) + '-' + pad2(x.d); }

  /** Parse a strict ISO calendar date YYYY-MM-DD to a day number. */
  function parseDateStr(s) {
    var m = /^\s*(\d{4})-(\d{2})-(\d{2})\s*$/.exec(String(s == null ? '' : s));
    if (!m) fail('E_INVALID_DATE');
    var y = +m[1], mo = +m[2], d = +m[3];
    if (!isValidYMD(y, mo, d)) fail('E_INVALID_DATE');
    return dayNumber(y, mo, d);
  }
  function parseTimeStr(s) {
    var m = /^\s*(\d{1,2}):(\d{2})\s*$/.exec(String(s == null ? '' : s));
    if (!m) fail('E_INVALID_DATE');
    var h = +m[1], mi = +m[2];
    if (h > 23 || mi > 59) fail('E_INVALID_DATE');
    return h * 60 + mi;
  }
  /** Offset text to minutes east of UTC; null when "unknown". */
  function parseOffset(s) {
    var t = String(s == null ? '' : s).trim().replace(MINUS, '-');
    if (t === '') fail('E_MISSING_INPUT', 'the UTC offset, or “offset unknown”');
    if (/^unknown$/i.test(t) || /^offset unknown$/i.test(t)) return null;
    if (t === 'Z' || t === '+00:00' || t === '-00:00') return 0;
    var m = /^([+-])(\d{2}):?(\d{2})$/.exec(t);
    if (!m || +m[2] > 14 || +m[3] > 59) fail('E_INVALID_DATE');
    return (m[1] === '-' ? -1 : 1) * (+m[2] * 60 + +m[3]);
  }
  function fmtOffset(min) {
    if (min === null || min === undefined) return 'offset unknown';
    var s = min < 0 ? MINUS : '+'; var a = Math.abs(min);
    return 'UTC' + s + pad2(Math.floor(a / 60)) + ':' + pad2(a % 60);
  }

  /**
   * A Timing object is the stored form of a date or time as entered (Section 2):
   *  {type:'timestamp', date:'YYYY-MM-DD', time:'HH:MM', offset:'-05:00'|'unknown', original}
   *  {type:'date', date:'YYYY-MM-DD', original}
   *  {type:'partial', precision:'year'|'month'|'range', value:'YYYY'|'YYYY-MM', lower, upper, original}
   *  {type:'none', original}   (not recorded)
   */
  function timingFromIso(s) {
    var str = String(s);
    var m = /^(\d{4}-\d{2}-\d{2})(?:T(\d{2}:\d{2})(?::(\d{2}))?(Z|[+\-−]\d{2}:\d{2})?)?$/.exec(str);
    if (!m) return { type: 'date', date: str, original: str };   // will raise E_INVALID_DATE on resolve
    if (!m[2]) return { type: 'date', date: m[1], original: str };
    return { type: 'timestamp', date: m[1], time: m[2], offset: m[4] ? m[4] : 'unknown', original: str };
  }
  function timingToIso(t) {
    if (!t) return '';
    if (t.type === 'timestamp') {
      var off = String(t.offset || '');
      return t.date + 'T' + t.time + (/^unknown$/i.test(off) ? '' : off.replace(MINUS, '-'));
    }
    if (t.type === 'date') return t.date;
    if (t.type === 'partial') {
      if (t.precision === 'range') return t.lower + '/' + t.upper;
      return String(t.value || '');
    }
    return '';
  }

  /** Resolve a Timing object for arithmetic. Throws CalcError on invalid or missing values. */
  function resolveTiming(t, label) {
    label = label || 'a date';
    if (!t || !t.type || t.type === 'none') fail('E_MISSING_INPUT', label);
    if (t.type === 'timestamp') {
      if (!t.date) fail('E_MISSING_INPUT', label + ' (date)');
      if (!t.time) fail('E_MISSING_INPUT', label + ' (time)');
      var day = parseDateStr(t.date), mod = parseTimeStr(t.time), off = parseOffset(t.offset);
      var local = day * 1440 + mod;
      return { kind: 'ts', day: day, minuteOfDay: mod, localMin: local, offsetMin: off,
        utcMin: off === null ? null : local - off };
    }
    if (t.type === 'date') {
      if (!t.date) fail('E_MISSING_INPUT', label);
      return { kind: 'date', day: parseDateStr(t.date) };
    }
    if (t.type === 'partial') {
      var p = t.precision, lo, hi, mm;
      if (p === 'year') {
        mm = /^\s*(\d{4})\s*$/.exec(String(t.value || ''));
        if (!mm) fail(t.value ? 'E_INVALID_DATE' : 'E_MISSING_INPUT', label + ' (year)');
        lo = dayNumber(+mm[1], 1, 1); hi = dayNumber(+mm[1], 12, 31);
      } else if (p === 'month') {
        mm = /^\s*(\d{4})-(\d{2})\s*$/.exec(String(t.value || ''));
        if (!mm) fail(t.value ? 'E_INVALID_DATE' : 'E_MISSING_INPUT', label + ' (month)');
        if (+mm[2] < 1 || +mm[2] > 12) fail('E_INVALID_DATE');
        lo = dayNumber(+mm[1], +mm[2], 1); hi = dayNumber(+mm[1], +mm[2], daysInMonth(+mm[1], +mm[2]));
      } else if (p === 'range') {
        if (!t.lower || !t.upper) fail('E_MISSING_INPUT', label + ' (both bounds)');
        lo = parseDateStr(t.lower); hi = parseDateStr(t.upper);
        if (lo > hi) fail('E_REVERSED_BOUNDS');
      } else {
        fail('E_MISSING_INPUT', label + ' (precision)');
      }
      return { kind: 'range', lo: lo, hi: hi, precision: p };
    }
    fail('E_MISSING_INPUT', label);
  }
  function dayBounds(r) { return r.kind === 'range' ? [r.lo, r.hi] : [r.day, r.day]; }
  function precisionLabel(t) {
    if (!t || !t.type || t.type === 'none') return 'not recorded';
    if (t.type === 'timestamp') return /^unknown$/i.test(String(t.offset)) ? 'timestamp, offset unknown' : 'timestamp with UTC offset';
    if (t.type === 'date') return 'day';
    return t.precision || 'partial';
  }

  // ------------------------------------------------------------------ display (Section 7)
  function signed(n, withPlus) {
    if (n < 0) return MINUS + String(Math.abs(n));
    return (withPlus && n > 0 ? '+' : '') + String(n);
  }
  function fixed(x, dp) {
    var s = (Math.abs(x) < 0.5 * Math.pow(10, -dp) ? 0 : x).toFixed(dp);
    return s.charAt(0) === '-' ? MINUS + s.slice(1) : s;
  }
  function groupThousands(s) {
    var neg = s.charAt(0) === MINUS; if (neg) s = s.slice(1);
    var parts = s.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (neg ? MINUS : '') + parts.join('.');
  }
  /** Percent to three significant figures: 12.5%, 0.391%, 63.0%. */
  function percent3(fraction) {
    var p = fraction * 100;
    if (p === 0) return '0%';
    var mag = Math.floor(Math.log10(Math.abs(p)));
    var d = Math.max(0, 2 - mag), s = p.toFixed(d);
    var mag2 = Math.floor(Math.log10(Math.abs(parseFloat(s))));
    if (mag2 !== mag) s = parseFloat(s).toFixed(Math.max(0, 2 - mag2));
    return groupThousands(s) + '%';
  }
  function sig3(x) {
    if (x === 0) return '0';
    var mag = Math.floor(Math.log10(Math.abs(x)));
    return fixed(x, Math.max(0, 2 - mag));
  }
  function fmtDay(n) { var x = fromDayNumber(n); return x.d + ' ' + MONTHS[x.m - 1] + ' ' + x.y; }
  function fmtDateStr(s) { try { return fmtDay(parseDateStr(s)); } catch (e) { return String(s); } }
  function fmtRangeDays(lo, hi) {
    if (lo === hi) return fmtDay(lo);
    var a = fromDayNumber(lo), b = fromDayNumber(hi);
    if (a.y === b.y && a.m === b.m) return a.d + '–' + b.d + ' ' + MONTHS[a.m - 1] + ' ' + a.y;
    if (a.y === b.y) return a.d + ' ' + MONTHS[a.m - 1] + ' – ' + b.d + ' ' + MONTHS[b.m - 1] + ' ' + a.y;
    return fmtDay(lo) + ' – ' + fmtDay(hi);
  }
  /** Human display of a Timing object, day month year (US spelling). */
  function fmtTiming(t) {
    if (!t || !t.type || t.type === 'none') return 'not recorded';
    try {
      if (t.type === 'timestamp') {
        var r = resolveTiming(t);
        return fmtDay(r.day) + ' ' + pad2(Math.floor(r.minuteOfDay / 60)) + ':' + pad2(r.minuteOfDay % 60) +
          ' (' + fmtOffset(r.offsetMin) + ')';
      }
      if (t.type === 'date') return fmtDay(parseDateStr(t.date));
      var p = resolveTiming(t);
      if (t.precision === 'month') { var x = fromDayNumber(p.lo); return MONTHS[x.m - 1] + ' ' + x.y + ' (month: ' + fmtRangeDays(p.lo, p.hi) + ')'; }
      if (t.precision === 'year') return fromDayNumber(p.lo).y + ' (year: ' + fmtRangeDays(p.lo, p.hi) + ')';
      return fmtRangeDays(p.lo, p.hi) + ' (range)';
    } catch (e) {
      return timingToIso(t) + ' (not valid)';
    }
  }
  function fmtMinutes(v) { return groupThousands(signed(Math.round(v))); }
  /** Display of an interval result, by unit. */
  function displayInterval(res) {
    if (res.unit === 'calendar days') {
      if (res.value !== null && res.value !== undefined) return signed(res.value) + ' calendar ' + (Math.abs(res.value) === 1 ? 'day' : 'days');
      var cross = res.minimum < 0 && res.maximum > 0;
      return signed(res.minimum, cross) + ' to ' + signed(res.maximum, cross) + ' calendar days';
    }
    var u = res.unit;
    if (Math.abs(res.value) === 1) u = u.replace('minutes', 'minute');
    return fmtMinutes(res.value) + ' ' + u;
  }

  // ------------------------------------------------------------------ interval rules C1-C6
  function seqFromNumber(v) { return v > 0 ? 'AFTER' : v < 0 ? 'BEFORE' : 'SIMULTANEOUS'; }

  /** C1 exact elapsed interval between two resolved timestamps. */
  function c1Exact(a, o) {
    if (a.kind !== 'ts' || o.kind !== 'ts') fail('E_MISSING_INPUT', 'a time on both the anchor and the end point');
    var flags = [], minutes, unit;
    if (a.offsetMin !== null && o.offsetMin !== null) { minutes = o.utcMin - a.utcMin; unit = 'minutes'; }
    else { minutes = o.localMin - a.localMin; unit = 'minutes (local clock)'; flags.push('W_TZ_UNKNOWN'); }
    return { mode: 'C1', value: minutes, minimum: null, maximum: null, unit: unit, sequence: seqFromNumber(minutes),
      flags: flags, derived: { hours: minutes / 60, elapsed_days: minutes / 1440 } };
  }
  /** C2 calendar-date difference; uses the recorded local date of a timestamp. */
  function c2Calendar(a, o) {
    if (a.kind === 'range' || o.kind === 'range') fail('E_MISSING_INPUT', 'an exact date on both the anchor and the end point');
    var days = o.day - a.day, seq, flags = [], note = null;
    if (days > 0) seq = 'AFTER';
    else if (days < 0) seq = 'BEFORE';
    else if (a.kind === 'ts' && o.kind === 'ts') {
      var c1 = c1Exact(a, o); seq = c1.sequence; flags = c1.flags.slice();
      note = 'Same calendar date; both times are known, so C1 decides the order (' + seq + ').';
    } else { seq = 'SAME_DATE_ORDER_UNKNOWN'; flags.push('W_SAME_DATE_ORDER_UNKNOWN'); }
    return { mode: 'C2', value: days, minimum: null, maximum: null, unit: 'calendar days', sequence: seq, flags: flags, note: note };
  }
  /** C3 treatment-day number, inclusive convention only. */
  function c3TreatmentDay(calendarDays) {
    if (calendarDays >= 0) return { treatment_day: calendarDays + 1, flags: [] };
    return { treatment_day: null, flags: ['I_TREATMENT_DAY_UNDEFINED'] };
  }
  /** C4 bounded interval from day bounds. */
  function c4Bounded(aLo, aHi, oLo, oHi) {
    if (aLo > aHi || oLo > oHi) fail('E_REVERSED_BOUNDS');
    var lo = oLo - aHi, hi = oHi - aLo, seq, flags = [];
    if (lo > 0) seq = 'AFTER';
    else if (hi < 0) seq = 'BEFORE';
    else if (lo === 0 && hi === 0) { seq = 'SAME_DATE_ORDER_UNKNOWN'; flags.push('W_SAME_DATE_ORDER_UNKNOWN'); }
    else { seq = 'SEQUENCE_UNCERTAIN'; flags.push('W_SEQUENCE_UNCERTAIN'); }
    return { mode: 'C4', value: null, minimum: lo, maximum: hi, unit: 'calendar days', sequence: seq, flags: flags };
  }
  /** Choose C1, C2 or C4 from the precision of the two resolved timings. */
  function computeInterval(a, o, opts) {
    opts = opts || {};
    if (a.kind === 'ts' && o.kind === 'ts' && !opts.dateLevel) return c1Exact(a, o);
    if (a.kind !== 'range' && o.kind !== 'range') return c2Calendar(a, o);
    var ab = dayBounds(a), ob = dayBounds(o);
    return c4Bounded(ab[0], ab[1], ob[0], ob[1]);
  }
  function compareInstant(a, b) {
    // both timestamps: compare at UTC when both offsets are known, otherwise on the local clock
    if (a.offsetMin !== null && b.offsetMin !== null) return a.utcMin - b.utcMin;
    return a.localMin - b.localMin;
  }
  /**
   * C5 LAST_DOSE_BEFORE_ONSET: the latest administration at or before onset. An administration after
   * onset is never selected; a same-date administration without times is selected with W_ADMIN_SAME_DATE_AS_ONSET.
   * candidates: [{id, r (resolved timing)}]; onset: resolved timing.
   */
  function selectLastDose(candidates, onset) {
    var eligible = [], uncertain = [];
    for (var i = 0; i < candidates.length; i++) {
      var c = candidates[i], r = c.r, status;
      if (r.kind === 'ts' && onset.kind === 'ts') status = compareInstant(r, onset) <= 0 ? 'before' : 'after';
      else {
        var ab = dayBounds(r), ob = dayBounds(onset);
        if (ab[1] < ob[0]) status = 'before';
        else if (ab[0] > ob[1]) status = 'after';
        else if (ab[0] === ab[1] && ob[0] === ob[1]) status = 'same_date';
        else status = 'uncertain';
      }
      if (status === 'before' || status === 'same_date') eligible.push({ c: c, status: status });
      else if (status === 'uncertain') uncertain.push(c.id);
    }
    if (!eligible.length) fail('E_MISSING_INPUT', 'an administration at or before onset');
    function key(e) {
      var r = e.c.r;
      if (r.kind === 'ts') return r.day * 1440 + r.minuteOfDay;
      return dayBounds(r)[1] * 1440;
    }
    eligible.sort(function (x, y) { return key(y) - key(x); });
    var chosen = eligible[0];
    return { chosen: chosen.c, flags: chosen.status === 'same_date' ? ['W_ADMIN_SAME_DATE_AS_ONSET'] : [], uncertain: uncertain };
  }
  /**
   * C6 reported timing comparison. result: interval result or null. reported: {text, lower, upper, unit}.
   * Returns {flags, comparison, lower_converted, upper_converted, unit}.
   */
  function c6Compare(result, reported) {
    if (!reported) return { flags: [], comparison: null };
    var has = function (v) { return v !== null && v !== undefined && v !== '' && !isNaN(+v); };
    var unit = reported.unit;
    if (!has(reported.lower) || !has(reported.upper)) {
      return { flags: ['I_APPROXIMATE_NOT_COMPARED'], comparison: 'not compared (no numeric bounds)' };
    }
    if (unit && REPORTED_UNITS.indexOf(unit) < 0) fail('E_INVALID_UNIT');
    if (!unit) fail('E_MISSING_INPUT', 'the unit of the reported timing');
    var lo = +reported.lower, hi = +reported.upper;
    if (lo > hi) fail('E_REVERSED_BOUNDS');
    if (unit === 'weeks' || unit === 'months' || unit === 'years') {
      return { flags: ['I_APPROXIMATE_NOT_COMPARED'], comparison: 'not compared (' + unit + ' are never converted)' };
    }
    if (!result) return { flags: [], comparison: 'not compared (no computed interval)' };
    var cLo, cHi;
    if (result.unit === 'calendar days') {
      if (unit !== 'calendar days') return { flags: ['I_APPROXIMATE_NOT_COMPARED'], comparison: 'not compared (calendar days are never converted to ' + unit + ')' };
      cLo = lo; cHi = hi;
    } else {
      if (!(unit in TIME_UNITS)) return { flags: ['I_APPROXIMATE_NOT_COMPARED'], comparison: 'not compared (' + unit + ' cannot be converted to minutes)' };
      cLo = lo * TIME_UNITS[unit] * 60; cHi = hi * TIME_UNITS[unit] * 60;
    }
    var rMin = result.value !== null && result.value !== undefined ? result.value : result.minimum;
    var rMax = result.value !== null && result.value !== undefined ? result.value : result.maximum;
    var overlap = rMin <= cHi && cLo <= rMax;
    return { flags: overlap ? [] : ['W_REPORTED_TIMING_CONFLICT'], comparison: overlap ? 'overlaps the computed interval' : 'conflicts with the computed interval',
      lower_converted: cLo, upper_converted: cHi, unit: result.unit === 'calendar days' ? 'calendar days' : 'minutes' };
  }

  // ------------------------------------------------------------------ technical rules Q1-Q11
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  function req(errs, inp, key, label) {
    var v = inp[key];
    if (v === null || v === undefined || v === '' || (typeof v === 'number' && !isFinite(v))) { errs.add('E_MISSING_INPUT', label || key); return false; }
    if (typeof v !== 'number') { errs.add('E_MISSING_INPUT', (label || key) + ' as a number'); return false; }
    return true;
  }
  function positive(errs, inp, key, label) {
    if (req(errs, inp, key, label) && !(inp[key] > 0)) { errs.add('E_NONPOSITIVE_PARAMETER'); return false; }
    return isNum(inp[key]) && inp[key] > 0;
  }
  function timeUnit(errs, u) {
    if (u === null || u === undefined || u === '') { errs.add('E_MISSING_INPUT', 'a time unit'); return false; }
    if (!(u in TIME_UNITS)) { errs.add('E_INVALID_UNIT'); return false; }
    return true;
  }
  function commonGate(errs, inp) { if (inp.applicability_confirmed !== true) errs.add('E_APPLICABILITY_UNCONFIRMED'); }
  function frac(tHalf, t) { return Math.pow(2, -t / tHalf); }

  function q1(inp) {
    var e = new Errs();
    timeUnit(e, inp.unit); positive(e, inp, 't_half', 'the half-life');
    if (req(e, inp, 't', 'the time') && inp.t < 0) e.add('E_OUT_OF_RANGE');
    commonGate(e, inp); e.raise();
    var k = Math.LN2 / inp.t_half;
    return { out: { fraction_remaining: frac(inp.t_half, inp.t), k_per_unit: k, k_per_hour: Math.LN2 / (inp.t_half * TIME_UNITS[inp.unit]) }, flags: [] };
  }
  function q2(inp) {
    var e = new Errs();
    timeUnit(e, inp.unit); positive(e, inp, 't_half', 'the half-life');
    if (req(e, inp, 'target_fraction', 'the target fraction') && !(inp.target_fraction > 0 && inp.target_fraction < 1)) e.add('E_OUT_OF_RANGE');
    commonGate(e, inp); e.raise();
    var n = -Math.log2(1 - inp.target_fraction);
    return { out: { time_to_fraction: n * inp.t_half, half_lives: n }, flags: [] };
  }
  function q3(inp) {
    var e = new Errs();
    timeUnit(e, inp.unit); positive(e, inp, 't_half', 'the half-life'); positive(e, inp, 'tau', 'the dosing interval');
    commonGate(e, inp); e.raise();
    return { out: { accumulation_factor: 1 / (1 - Math.pow(2, -inp.tau / inp.t_half)) }, flags: [] };
  }
  function q4(inp) {
    var e = new Errs();
    var unit = inp.cl_unit;
    if (!unit) e.add('E_MISSING_INPUT', 'the clearance unit'); else if (!(unit in CLEARANCE_UNITS)) e.add('E_INVALID_UNIT');
    positive(e, inp, 'cl_before', 'the clearance before'); positive(e, inp, 'cl_after', 'the clearance after');
    var hasV = inp.volume_l !== null && inp.volume_l !== undefined && inp.volume_l !== '';
    if (hasV) {
      positive(e, inp, 'volume_l', 'the volume');
      if (inp.volume_unit && inp.volume_unit !== 'L') e.add('E_INVALID_UNIT');
    }
    commonGate(e, inp); e.raise();
    var r = inp.cl_before / inp.cl_after, out = { half_life_ratio: r, average_css_ratio: r };
    if (hasV) {
      var f = CLEARANCE_UNITS[unit];
      out.t_half_before_hours = Math.LN2 * inp.volume_l / (inp.cl_before * f);
      out.t_half_after_hours = Math.LN2 * inp.volume_l / (inp.cl_after * f);
    }
    return { out: out, flags: [] };
  }
  function q5(inp) {
    var e = new Errs();
    timeUnit(e, inp.unit); positive(e, inp, 'parent_t_half', 'the parent half-life'); positive(e, inp, 'metabolite_t_half', 'the metabolite half-life');
    if (req(e, inp, 't', 'the time') && inp.t < 0) e.add('E_OUT_OF_RANGE');
    commonGate(e, inp); e.raise();
    return { out: { parent_percent_remaining: 100 * frac(inp.parent_t_half, inp.t),
      metabolite_percent_remaining: 100 * frac(inp.metabolite_t_half, inp.t), sum_displayed: false }, flags: [] };
  }
  function q6(inp) {
    var e = new Errs();
    timeUnit(e, inp.unit); positive(e, inp, 'replacement_half_time', 'the replacement half-time');
    if (req(e, inp, 't', 'the time') && inp.t < 0) e.add('E_OUT_OF_RANGE');
    if (inp.teaching_demonstration !== true) e.add('E_APPLICABILITY_UNCONFIRMED');
    e.raise();
    return { out: { fraction_replaced: 1 - Math.pow(2, -inp.t / inp.replacement_half_time) }, flags: ['I_TEACHING_DEMONSTRATION'] };
  }
  function q7(inp) {
    var e = new Errs();
    if (!inp.unit) e.add('E_MISSING_INPUT', 'a time unit'); else if (inp.unit !== 'days') e.add('E_INVALID_UNIT');
    positive(e, inp, 't_half', 'the half-life');
    if (req(e, inp, 't', 'the time') && inp.t < 0) e.add('E_OUT_OF_RANGE');
    commonGate(e, inp); e.raise();
    return { out: { fraction_remaining: frac(inp.t_half, inp.t), k_per_unit: Math.LN2 / inp.t_half }, flags: ['W_BIOLOGIC_MODEL_CAUTION'] };
  }
  function q8(inp) {
    var e = new Errs();
    timeUnit(e, inp.unit);
    var okLo = positive(e, inp, 't_half_lower', 'the lower half-life'), okHi = positive(e, inp, 't_half_upper', 'the upper half-life');
    if (okLo && okHi && inp.t_half_lower > inp.t_half_upper) e.add('E_REVERSED_BOUNDS');
    if (req(e, inp, 't', 'the time') && inp.t < 0) e.add('E_OUT_OF_RANGE');
    commonGate(e, inp); e.raise();
    var flags = inp.antibody === true ? ['W_BIOLOGIC_MODEL_CAUTION', 'I_SCENARIO_NOT_CI'] : ['I_SCENARIO_NOT_CI'];
    return { out: { fraction_lower: frac(inp.t_half_lower, inp.t), fraction_upper: frac(inp.t_half_upper, inp.t), midpoint_displayed: false }, flags: flags };
  }
  function asTiming(x) { return typeof x === 'string' ? timingFromIso(x) : x; }
  function q9(inp) {
    var e = new Errs(), segs = [], onset = null, flags = [];
    try { onset = resolveTiming(asTiming(inp.onset), 'the onset time'); if (onset.kind !== 'ts') { onset = null; e.add('E_MISSING_INPUT', 'the onset time'); } }
    catch (x) { e.absorb(x); }
    var list = inp.segments || [];
    if (!list.length) e.add('E_MISSING_INPUT', 'at least one infusion segment');
    for (var i = 0; i < list.length; i++) {
      var s = list[i], st = null, en = null;
      try { st = resolveTiming(asTiming(s.start), 'the start of segment ' + (i + 1)); en = resolveTiming(asTiming(s.end), 'the end of segment ' + (i + 1));
        if (st.kind !== 'ts' || en.kind !== 'ts') { e.add('E_MISSING_INPUT', 'start and end times of segment ' + (i + 1)); st = en = null; } }
      catch (x) { e.absorb(x); st = en = null; }
      var rate = s.rate_mg_per_hour !== undefined ? s.rate_mg_per_hour : s.rate;
      var rUnit = s.rate_mg_per_hour !== undefined ? 'mg/hour' : s.rate_unit;
      if (!rUnit) e.add('E_MISSING_INPUT', 'the rate unit of segment ' + (i + 1));
      else if (rUnit !== 'mg/hour') e.add('E_INVALID_UNIT');
      if (rate === null || rate === undefined || rate === '' || !isNum(rate)) e.add('E_MISSING_INPUT', 'the rate of segment ' + (i + 1));
      else if (!(rate > 0)) e.add('E_NONPOSITIVE_PARAMETER');
      if (st && en) segs.push({ st: st, en: en, rate: rate });
    }
    var hasPlanned = inp.planned_dose_mg !== null && inp.planned_dose_mg !== undefined && inp.planned_dose_mg !== '';
    if (hasPlanned) {
      if (!isNum(inp.planned_dose_mg)) e.add('E_MISSING_INPUT', 'the planned dose as a number');
      else if (!(inp.planned_dose_mg > 0)) e.add('E_NONPOSITIVE_PARAMETER');
      if (inp.planned_dose_unit && inp.planned_dose_unit !== 'mg') e.add('E_INVALID_UNIT');
    }
    // clock: UTC when every offset is known, otherwise the local clock with W_TZ_UNKNOWN
    var all = segs.reduce(function (acc, s) { return acc.concat([s.st, s.en]); }, onset ? [onset] : []);
    var useUtc = all.every(function (r) { return r.offsetMin !== null; });
    if (!useUtc) flags.push('W_TZ_UNKNOWN');
    function t(r) { return useUtc ? r.utcMin : r.localMin; }
    segs.sort(function (a, b) { return t(a.st) - t(b.st); });
    for (var j = 0; j < segs.length; j++) {
      if (t(segs[j].st) > t(segs[j].en)) e.add('E_REVERSED_BOUNDS');
      if (j > 0 && t(segs[j].st) < t(segs[j - 1].en)) e.add('E_OVERLAPPING_SEGMENTS');
    }
    e.raise();
    var amount = 0;
    for (var k = 0; k < segs.length; k++) {
      var hours = Math.max(0, (Math.min(t(segs[k].en), t(onset)) - t(segs[k].st)) / 60);
      amount += segs[k].rate * hours;
    }
    var out = { amount_mg: amount };
    if (hasPlanned) out.percent_of_planned = 100 * amount / inp.planned_dose_mg;
    return { out: out, flags: flags };
  }
  function q10(inp) {
    var e = new Errs(), admins = inp.administrations || [];
    if (!admins.length) e.add('E_MISSING_INPUT', 'at least one administration');
    var known = [], total = 0, missing = 0;
    for (var i = 0; i < admins.length; i++) {
      var a = admins[i];
      if (a.dose_unit !== undefined && a.dose_unit !== 'mg') { if (!a.dose_unit) e.add('E_MISSING_INPUT', 'the dose unit of administration ' + (i + 1)); else e.add('E_INVALID_UNIT'); }
      if (a.dose_mg === null || a.dose_mg === undefined || a.dose_mg === '' || !isNum(a.dose_mg)) e.add('E_MISSING_INPUT', 'the dose of administration ' + (i + 1));
      else if (!(a.dose_mg > 0)) e.add('E_NONPOSITIVE_PARAMETER');
      else total += a.dose_mg;
      var hasB = a.bsa_m2 !== null && a.bsa_m2 !== undefined && a.bsa_m2 !== '';
      if (hasB) {
        if (!isNum(a.bsa_m2)) e.add('E_MISSING_INPUT', 'the BSA of administration ' + (i + 1) + ' as a number');
        else if (!(a.bsa_m2 > 0)) e.add('E_NONPOSITIVE_PARAMETER');
        else if (isNum(a.dose_mg)) known.push(a);
      } else missing++;
    }
    if (admins.length && known.length === 0 && !e.any()) e.add('E_MISSING_INPUT', 'the BSA recorded at one or more administrations');
    var methodB = inp.method_b ? (Array.isArray(inp.method_b) ? inp.method_b : [inp.method_b]) : [];
    var bValues = [];
    for (var b = 0; b < methodB.length; b++) {
      var m = methodB[b];
      if (m === 'first' || m === 'last') {
        var src = admins.length ? admins[m === 'first' ? 0 : admins.length - 1] : null;
        if (!src || !isNum(src.bsa_m2)) e.add('E_MISSING_INPUT', 'the ' + m + ' BSA for Method B');
        else if (src.bsa_m2 > 0) bValues.push({ key: m, bsa: src.bsa_m2 });
      } else if (isNum(m)) {
        if (!(m > 0)) e.add('E_NONPOSITIVE_PARAMETER'); else bValues.push({ key: 'stated', bsa: m });
      } else if (m !== '' && m !== null) e.add('E_MISSING_INPUT', 'the stated BSA for Method B');
    }
    e.raise();
    var sumA = known.reduce(function (s, a) { return s + a.dose_mg / a.bsa_m2; }, 0);
    var out = {}, flags = [];
    if (missing > 0) { out.per_administration_lower_bound_mg_m2 = sumA; flags.push('W_PARTIAL_BSA'); }
    else out.per_administration_mg_m2 = sumA;
    out.total_dose_mg = total;
    for (var c = 0; c < bValues.length; c++) {
      var v = bValues[c];
      out[v.key === 'first' ? 'single_bsa_first_mg_m2' : v.key === 'last' ? 'single_bsa_last_mg_m2' : 'single_bsa_stated_mg_m2'] = total / v.bsa;
      if (v.key === 'stated') out.stated_bsa_m2 = v.bsa;
    }
    return { out: out, flags: flags };
  }
  function q11(inp) {
    if (inp.creatinine_stable !== true) fail('E_CREATININE_STABILITY_UNCONFIRMED');
    var e = new Errs();
    positive(e, inp, 'weight_kg', 'the weight');
    positive(e, inp, 'scr', 'the serum creatinine');
    if (req(e, inp, 'age', 'the age') && !(inp.age >= 18 && inp.age <= 120)) e.add('E_OUT_OF_RANGE');
    if (!inp.scr_unit) e.add('E_MISSING_INPUT', 'the creatinine unit'); else if (!(inp.scr_unit in SCR_UNITS)) e.add('E_INVALID_UNIT');
    if (inp.weight_unit && inp.weight_unit !== 'kg') e.add('E_INVALID_UNIT');
    var sex = inp.sex;
    if (sex !== 'male' && sex !== 'female' && sex !== 'not recorded') e.add('E_MISSING_INPUT', 'sex as recorded (male, female or not recorded)');
    e.raise();
    var scr = inp.scr * SCR_UNITS[inp.scr_unit];
    var male = (140 - inp.age) * inp.weight_kg / (72 * scr);
    if (sex === 'male') return { out: { crcl_ml_min: male, scr_mg_dl: scr }, flags: [] };
    if (sex === 'female') return { out: { crcl_ml_min: 0.85 * male, scr_mg_dl: scr }, flags: [] };
    return { out: { crcl_ml_min: [0.85 * male, male], scr_mg_dl: scr }, flags: ['W_SEX_NOT_RECORDED'] };
  }
  var TECHNICAL = { Q1: q1, Q2: q2, Q3: q3, Q4: q4, Q5: q5, Q6: q6, Q7: q7, Q8: q8, Q9: q9, Q10: q10, Q11: q11 };
  var COMMON_GATE_RULES = ['Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q7', 'Q8'];

  // ------------------------------------------------------------------ case model
  function emptyCase() {
    return { schema: 'tto-case/1', case_id: '', case_version: '1', reviewer: '', setting: '', notes: '', synthetic_note: '',
      events: [], milestones: [], exposures: [], links: [], intervals: [], recoveries: [], technical: [], assessment: [] };
  }
  function clone(x) { return x === undefined ? undefined : JSON.parse(JSON.stringify(x)); }

  /** All records that can serve as an anchor or end point: exposure records and milestones. */
  function indexCase(c) {
    var idx = { byId: {}, counts: {}, exposureOf: {} };
    function reg(id, obj, kind) {
      var k = String(id == null ? '' : id).trim();
      idx.counts[k] = (idx.counts[k] || 0) + 1;
      if (!idx.byId[k]) idx.byId[k] = { kind: kind, obj: obj };
    }
    (c.events || []).forEach(function (ev) { reg(ev.id, ev, 'event'); });
    (c.milestones || []).forEach(function (m) { reg(m.id, m, 'milestone'); });
    (c.exposures || []).forEach(function (x) {
      reg(x.id, x, 'exposure');
      (x.records || []).forEach(function (r) { reg(r.id, r, 'exposure_record'); idx.exposureOf[String(r.id).trim()] = x; });
    });
    (c.links || []).forEach(function (l) { reg(l.id, l, 'link'); });
    (c.intervals || []).forEach(function (q) { reg(q.id, q, 'interval'); });
    (c.recoveries || []).forEach(function (q) { reg(q.id, q, 'recovery'); });
    (c.technical || []).forEach(function (q) { reg(q.id, q, 'technical'); });
    return idx;
  }
  function timingOf(idx, id) {
    var e = idx.byId[String(id).trim()];
    if (!e) return null;
    if (e.kind === 'event') return e.obj.onset;
    return e.obj.timing;
  }
  function recordLabel(idx, id) {
    var e = idx.byId[String(id).trim()];
    if (!e) return String(id);
    if (e.kind === 'event') return (e.obj.term || e.obj.id) + ' (onset)';
    if (e.kind === 'exposure_record') { var x = idx.exposureOf[String(id).trim()]; return (x ? x.drug + ': ' : '') + (e.obj.label || e.obj.kind); }
    return e.obj.label || e.obj.term || e.obj.id;
  }

  /** Case-level validation: E_DUPLICATE_ID, E_INVALID_DATE, E_REVERSED_BOUNDS, missing identifiers and references. */
  function validateCase(c) {
    var idx = indexCase(c), issues = [];
    Object.keys(idx.counts).forEach(function (k) {
      if (k === '') issues.push({ code: 'E_MISSING_INPUT', where: 'identifiers', detail: 'an identifier for every record' });
      else if (idx.counts[k] > 1) issues.push({ code: 'E_DUPLICATE_ID', where: k, detail: k });
    });
    function chk(t, where) {
      if (!t || !t.type || t.type === 'none') return;
      try { resolveTiming(t, where); } catch (e) {
        if (!(e instanceof CalcError)) throw e;
        e.codes.forEach(function (code) { issues.push({ code: code, where: where, detail: e.details[code] }); });
      }
    }
    (c.events || []).forEach(function (ev) { chk(ev.onset, (ev.id || '?') + ' onset'); });
    (c.milestones || []).forEach(function (m) { chk(m.timing, (m.id || '?')); });
    (c.exposures || []).forEach(function (x) { (x.records || []).forEach(function (r) { chk(r.timing, (r.id || '?')); }); });
    (c.links || []).forEach(function (l) {
      var ex = idx.byId[String(l.exposure_id).trim()], ev = idx.byId[String(l.event_id).trim()];
      if (!ex || ex.kind !== 'exposure') issues.push({ code: 'E_MISSING_INPUT', where: l.id, detail: 'an existing drug exposure episode for link ' + l.id });
      if (!ev || ev.kind !== 'event') issues.push({ code: 'E_MISSING_INPUT', where: l.id, detail: 'an existing event episode for link ' + l.id });
    });
    return issues;
  }
  function findLink(c, exposureId, eventId) {
    var links = c.links || [];
    for (var i = 0; i < links.length; i++) {
      if (String(links[i].exposure_id).trim() === String(exposureId).trim() && String(links[i].event_id).trim() === String(eventId).trim()) return links[i];
    }
    return null;
  }
  function paramEntry(name, id, idx, t, extra) {
    var e = idx.byId[String(id).trim()], src = e ? (e.obj.source || '') : '';
    if (!src && e && e.kind === 'exposure_record') { var xo = idx.exposureOf[String(id).trim()]; src = xo ? (xo.source || '') : ''; }
    var p = { name: name, record_id: id, value: timingToIso(t), display: fmtTiming(t), unit: t && t.type === 'timestamp' ? 'timestamp' : 'date',
      precision: precisionLabel(t), original_text: (t && t.original) || '', Parameter_Status: 'case measurement', source: src };
    if (extra) for (var k in extra) p[k] = extra[k];
    return p;
  }
  function blankRecord(base) {
    var r = { Result_ID: '', Label: '', Drug_Event_Link: '', Exposure_Episode_ID: '', Event_ID: '', End_Point: '',
      Interval_Value: null, Interval_Minimum: null, Interval_Maximum: null, Unit: '', Calculation_Mode: '',
      Anchor_Type: '', Anchor_Record_ID: '', Anchor_Description: '', Sequence_Status: '', Treatment_Day: null,
      Derived_Displays: null, Display: '', Reported_Timing: null, Constraint_Source: '', Results: [],
      Parameter_Set: [], Assumptions: [], Calculation_Version: SPEC_VERSION, QC_Flags: [], Messages: [],
      Blocked: false, Limitations: '', Equation: '' };
    for (var k in base) r[k] = base[k];
    return r;
  }
  function addFlags(rec, codes, details) {
    for (var i = 0; i < codes.length; i++) {
      if (rec.QC_Flags.indexOf(codes[i]) < 0) {
        rec.QC_Flags.push(codes[i]);
        rec.Messages.push({ code: codes[i], type: flagType(codes[i]), message: message(codes[i], details && details[codes[i]]) });
      }
    }
    if (rec.QC_Flags.some(function (c) { return c.charAt(0) === 'E'; })) rec.Blocked = true;
  }
  function block(rec, err) {
    if (!(err instanceof CalcError)) throw err;
    addFlags(rec, err.codes, err.details);
    rec.Interval_Value = rec.Interval_Minimum = rec.Interval_Maximum = null;
    rec.Sequence_Status = ''; rec.Treatment_Day = null; rec.Results = []; rec.Display = 'Not calculable';
    rec.Blocked = true;
    return rec;
  }
  function fillInterval(rec, res) {
    rec.Interval_Value = res.value; rec.Interval_Minimum = res.minimum; rec.Interval_Maximum = res.maximum;
    rec.Unit = res.unit; rec.Sequence_Status = res.sequence;
    if (rec.Calculation_Mode === '') rec.Calculation_Mode = res.mode;
    rec.Equation = EQUATIONS[res.mode];
    rec.Display = displayInterval(res);
    if (res.mode === 'C1') rec.Derived_Displays = { hours: fixed(res.value / 60, 1) + ' hours', elapsed_days: fixed(res.value / 1440, 2) + ' elapsed days' };
    if (res.note) rec.Assumptions.push(res.note);
    addFlags(rec, res.flags);
  }
  var C_LIMITATION = function (mode) {
    return 'No mandatory text is defined for ' + mode + ' in CALC-SPEC 1.0. The interval describes sequence and spacing of recorded times; it is not evidence of causation.';
  };

  function involvesDuplicate(idx, ids) {
    for (var i = 0; i < ids.length; i++) { var k = String(ids[i] == null ? '' : ids[i]).trim(); if (k && idx.counts[k] > 1) return true; }
    return false;
  }

  /** Evaluate one interval request (C1-C6 with C5 anchor selection). Returns one or more output records. */
  function evaluateInterval(c, q, idx) {
    idx = idx || indexCase(c);
    var rec = blankRecord({ Result_ID: q.id, Label: q.label || '', Exposure_Episode_ID: q.exposure_id || '', Event_ID: q.event_id || '',
      Anchor_Type: q.anchor_type || '', Anchor_Description: q.anchor_description || '' });
    var out = [rec];
    try {
      if (!q.exposure_id) fail('E_MISSING_INPUT', 'a drug exposure episode');
      if (!q.event_id) fail('E_MISSING_INPUT', 'an event episode');
      var exE = idx.byId[String(q.exposure_id).trim()], evE = idx.byId[String(q.event_id).trim()];
      if (!exE || exE.kind !== 'exposure') fail('E_MISSING_INPUT', 'an existing drug exposure episode');
      if (!evE || evE.kind !== 'event') fail('E_MISSING_INPUT', 'an existing event episode');
      var link = findLink(c, q.exposure_id, q.event_id);
      if (!link) fail('E_MISSING_LINK');
      rec.Drug_Event_Link = link.id;
      var endRef = q.end_ref && q.end_ref !== 'onset' ? q.end_ref : q.event_id;
      var endIsOnset = endRef === q.event_id;
      rec.End_Point = endIsOnset ? 'onset of ' + (evE.obj.term || q.event_id) : recordLabel(idx, endRef);
      if (!q.anchor_type || ANCHOR_TYPES.indexOf(q.anchor_type) < 0) fail('E_MISSING_INPUT', 'an anchor type');
      if (q.anchor_type === 'OTHER' && !String(q.anchor_description || '').trim()) fail('E_MISSING_INPUT', 'a description of the OTHER anchor');
      var endT = endIsOnset ? evE.obj.onset : timingOf(idx, endRef);
      if (!idx.byId[String(endRef).trim()]) fail('E_MISSING_INPUT', 'an existing end-point record');
      if (involvesDuplicate(idx, [q.exposure_id, q.event_id, endRef, q.anchor_ref])) fail('E_DUPLICATE_ID');
      var endR = resolveTiming(endT, endIsOnset ? 'the onset' : 'the end point');
      var exposure = exE.obj, anchorId = q.anchor_ref, anchorR, sel = null;
      if (q.anchor_type === 'LAST_DOSE_BEFORE_ONSET' && !anchorId) {
        var cands = [];
        (exposure.records || []).forEach(function (r) {
          if (['start', 'administration', 'restart', 'dose_change'].indexOf(r.kind) >= 0 && r.timing && r.timing.type && r.timing.type !== 'none') {
            cands.push({ id: r.id, r: resolveTiming(r.timing, r.id) });
          }
        });
        if (!cands.length) fail('E_MISSING_INPUT', 'an administration record (last dose before onset)');
        sel = selectLastDose(cands, endR);
        anchorId = sel.chosen.id; anchorR = sel.chosen.r;
        rec.Assumptions.push('Anchor selected under C5: latest administration at or before ' + (endIsOnset ? 'onset' : 'the end point') + ' (' + anchorId + '); later administrations are never selected.');
        if (sel.uncertain.length) rec.Assumptions.push('Not selected because their order relative to onset is uncertain: ' + sel.uncertain.join(', ') + '.');
      } else if (q.anchor_type === 'FIRST_ADMIN' && !anchorId) {
        var starts = (exposure.records || []).filter(function (r) { return r.kind === 'start'; });
        if (!starts.length) fail('E_MISSING_INPUT', 'the first administration');
        var best = null;
        starts.forEach(function (r) { var rr = resolveTiming(r.timing, r.id); var lo = dayBounds(rr)[0]; if (!best || lo < best.lo) best = { id: r.id, r: rr, lo: lo }; });
        anchorId = best.id; anchorR = best.r;
      } else {
        if (!anchorId) fail('E_MISSING_INPUT', 'the anchor record');
        if (!idx.byId[String(anchorId).trim()]) fail('E_MISSING_INPUT', 'an existing anchor record');
        anchorR = resolveTiming(timingOf(idx, anchorId), 'the anchor');
      }
      rec.Anchor_Record_ID = anchorId;
      var anchorT = timingOf(idx, anchorId);
      rec.Parameter_Set.push(paramEntry('Anchor (' + q.anchor_type + ')', anchorId, idx, anchorT, { label: recordLabel(idx, anchorId) }));
      rec.Parameter_Set.push(paramEntry('End point', endRef, idx, endT, { label: endIsOnset ? 'onset' : recordLabel(idx, endRef) }));
      if (!rec.Label) rec.Label = recordLabel(idx, anchorId) + ' → ' + (endIsOnset ? (evE.obj.term || q.event_id) + ' onset' : recordLabel(idx, endRef));
      var res = computeInterval(anchorR, endR);
      fillInterval(rec, res);
      if (sel) addFlags(rec, sel.flags);
      if (res.mode === 'C4') rec.Assumptions.push('Bounds are treated as independent; no single date is imputed from them.');
      if (q.treatment_day && res.mode === 'C2') {
        var td = c3TreatmentDay(res.value);
        rec.Treatment_Day = td.treatment_day; rec.Calculation_Mode = 'C2; C3';
        rec.Assumptions.push('Treatment day under the inclusive convention (C3); never written to E2B fields automatically.');
        addFlags(rec, td.flags);
      }
      if (q.reported && (String(q.reported.text || '').trim() || q.reported.lower !== '' && q.reported.lower != null)) {
        var cmp = c6Compare(res, q.reported);
        rec.Reported_Timing = { text: q.reported.text || '', lower: q.reported.lower === '' ? null : q.reported.lower,
          upper: q.reported.upper === '' ? null : q.reported.upper, unit: q.reported.unit || null, comparison: cmp.comparison };
        rec.Calculation_Mode += '; C6';
        addFlags(rec, cmp.flags);
      }
      rec.Limitations = C_LIMITATION(res.mode);
      if (rec.Constraint_Source === '' && res.mode === 'C4') rec.Assumptions.push('No constraint applied.');
      // date-level companion (C2) when both are timestamps
      if (q.date_level && res.mode === 'C1') {
        var r2 = blankRecord({ Result_ID: q.id + '-C2', Label: rec.Label + ' (date level)', Drug_Event_Link: rec.Drug_Event_Link,
          Exposure_Episode_ID: rec.Exposure_Episode_ID, Event_ID: rec.Event_ID, End_Point: rec.End_Point,
          Anchor_Type: rec.Anchor_Type, Anchor_Record_ID: rec.Anchor_Record_ID, Anchor_Description: rec.Anchor_Description,
          Parameter_Set: clone(rec.Parameter_Set) });
        var res2 = computeInterval(anchorR, endR, { dateLevel: true });
        fillInterval(r2, res2);
        if (q.treatment_day) { var td2 = c3TreatmentDay(res2.value); r2.Treatment_Day = td2.treatment_day; r2.Calculation_Mode = 'C2; C3'; addFlags(r2, td2.flags); }
        r2.Assumptions.push('Calendar-date difference from the recorded local dates; it is never described as elapsed days.');
        r2.Limitations = C_LIMITATION('C2');
        out.push(r2);
      }
      // documented constraint (C4): constrained result shown beside the unconstrained one
      if (q.constraint && q.constraint.target && (q.constraint.lower || q.constraint.upper) && res.mode === 'C4') {
        var r3 = blankRecord({ Result_ID: q.id + '-constrained', Label: rec.Label + ' (constrained)', Drug_Event_Link: rec.Drug_Event_Link,
          Exposure_Episode_ID: rec.Exposure_Episode_ID, Event_ID: rec.Event_ID, End_Point: rec.End_Point,
          Anchor_Type: rec.Anchor_Type, Anchor_Record_ID: rec.Anchor_Record_ID, Anchor_Description: rec.Anchor_Description,
          Parameter_Set: clone(rec.Parameter_Set), Constraint_Source: q.constraint.source || '' });
        try {
          if (!String(q.constraint.source || '').trim()) fail('E_MISSING_INPUT', 'the source of the constraint');
          var ab = dayBounds(anchorR), ob = dayBounds(endR), cl = q.constraint.lower ? parseDateStr(q.constraint.lower) : null,
            cu = q.constraint.upper ? parseDateStr(q.constraint.upper) : null, tgt = q.constraint.target === 'anchor' ? ab : ob;
          if (cl !== null) tgt[0] = Math.max(tgt[0], cl);
          if (cu !== null) tgt[1] = Math.min(tgt[1], cu);
          var res3 = c4Bounded(ab[0], ab[1], ob[0], ob[1]);
          fillInterval(r3, res3);
          r3.Calculation_Mode = 'C4 (constrained)';
          r3.Assumptions.push('Constraint applied to the ' + (q.constraint.target === 'anchor' ? 'anchor' : 'end-point') + ' bounds: ' +
            (cl !== null ? 'not before ' + fmtDay(cl) : '') + (cl !== null && cu !== null ? ', ' : '') + (cu !== null ? 'not after ' + fmtDay(cu) : '') +
            '. Source: ' + q.constraint.source + '. Constrained bounds: ' + fmtRangeDays(tgt[0], tgt[1]) + '.');
          r3.Limitations = C_LIMITATION('C4');
        } catch (e3) { block(r3, e3); }
        out.push(r3);
      }
    } catch (err) {
      block(rec, err);
      if (!rec.Label) rec.Label = q.label || ('Interval ' + q.id);
      rec.Calculation_Mode = rec.Calculation_Mode || 'C1–C6';
    }
    return out;
  }

  /** C7 recovery interval: CESSATION to a recovery milestone, with W_INTERVENTION_BEFORE_RECOVERY. */
  function evaluateRecovery(c, q, idx) {
    idx = idx || indexCase(c);
    var rec = blankRecord({ Result_ID: q.id, Label: q.label || '', Exposure_Episode_ID: q.exposure_id || '', Event_ID: q.event_id || '',
      Anchor_Type: 'CESSATION', Calculation_Mode: 'C7' });
    try {
      var link = findLink(c, q.exposure_id, q.event_id);
      if (!q.exposure_id || !q.event_id) fail('E_MISSING_INPUT', 'a drug exposure episode and an event episode');
      if (!link) fail('E_MISSING_LINK');
      rec.Drug_Event_Link = link.id;
      if (!q.cessation_ref) fail('E_MISSING_INPUT', 'the cessation record');
      if (!q.recovery_ref) fail('E_MISSING_INPUT', 'the recovery milestone');
      if (!idx.byId[String(q.cessation_ref).trim()]) fail('E_MISSING_INPUT', 'an existing cessation record');
      if (!idx.byId[String(q.recovery_ref).trim()]) fail('E_MISSING_INPUT', 'an existing recovery milestone');
      if (involvesDuplicate(idx, [q.exposure_id, q.event_id, q.cessation_ref, q.recovery_ref])) fail('E_DUPLICATE_ID');
      var cT = timingOf(idx, q.cessation_ref), rT = timingOf(idx, q.recovery_ref);
      var cR = resolveTiming(cT, 'the cessation'), rR = resolveTiming(rT, 'the recovery milestone');
      rec.Anchor_Record_ID = q.cessation_ref;
      rec.End_Point = recordLabel(idx, q.recovery_ref);
      if (!rec.Label) rec.Label = recordLabel(idx, q.cessation_ref) + ' → ' + recordLabel(idx, q.recovery_ref) + ' (recovery)';
      rec.Parameter_Set.push(paramEntry('Cessation (CESSATION)', q.cessation_ref, idx, cT, { label: recordLabel(idx, q.cessation_ref) }));
      rec.Parameter_Set.push(paramEntry('Recovery milestone', q.recovery_ref, idx, rT, { label: recordLabel(idx, q.recovery_ref) }));
      var res = computeInterval(cR, rR);
      fillInterval(rec, res);
      rec.Calculation_Mode = 'C7';
      rec.Equation = EQUATIONS.C7 + ' (' + res.mode + ': ' + EQUATIONS[res.mode] + ')';
      rec.Assumptions.unshift('Recovery interval (C7) computed with ' + res.mode + '.');
      // interventions recorded between cessation and recovery
      var found = [];
      var cands = (c.milestones || []).filter(function (m) { return m.kind === 'intervention'; }).map(function (m) { return { id: m.id, t: m.timing, label: m.label }; });
      (c.exposures || []).forEach(function (x) {
        if (x.role === 'treatment of the event') (x.records || []).forEach(function (r) { if (r.kind === 'start' || r.kind === 'administration') cands.push({ id: r.id, t: r.timing, label: x.drug }); });
      });
      cands.forEach(function (m) {
        if (!m.t || !m.t.type || m.t.type === 'none') return;
        var x; try { x = resolveTiming(m.t); } catch (e) { return; }
        var between;
        if (cR.kind === 'ts' && rR.kind === 'ts' && x.kind === 'ts') between = compareInstant(x, cR) >= 0 && compareInstant(x, rR) <= 0;
        else { var xb = dayBounds(x), cb = dayBounds(cR), rb = dayBounds(rR); between = xb[1] >= cb[0] && xb[0] <= rb[1]; }
        if (between) found.push(m.id + (m.label ? ' (' + m.label + ')' : ''));
      });
      if (found.length) { addFlags(rec, ['W_INTERVENTION_BEFORE_RECOVERY']); rec.Assumptions.push('Intervention recorded between cessation and recovery: ' + found.join('; ') + '.'); }
      else rec.Assumptions.push('No intervention is recorded between cessation and recovery; absence of a record is not evidence that none was given.');
      rec.Limitations = 'No mandatory text is defined for C7 in CALC-SPEC 1.0, and C7 has no verification vector yet; check the result by hand. A recovery interval is not, by itself, evidence of positive dechallenge.';
    } catch (err) {
      block(rec, err); if (!rec.Label) rec.Label = q.label || ('Recovery ' + q.id);
    }
    return [rec];
  }

  // ------------------------------------------------------------------ technical entries (UI path)
  /**
   * Parameter definitions for each Q rule. Each param: key, label, kind ('num'|'unit'|...), unit hint.
   * The UI builds its forms from these; evaluateTechnical maps them to the engine inputs.
   */
  var Q_PARAMS = {
    Q1: [{ key: 't_half', label: 'Half-life (t½)', time: true }, { key: 't', label: 'Time since the reference (t)', time: true }],
    Q2: [{ key: 't_half', label: 'Half-life (t½)', time: true }, { key: 'target_fraction', label: 'Target fraction of steady state (f, between 0 and 1)', unit: 'fraction' }],
    Q3: [{ key: 't_half', label: 'Half-life (t½)', time: true }, { key: 'tau', label: 'Dosing interval (τ)', time: true }],
    Q4: [{ key: 'cl_before', label: 'Clearance before', cl: true }, { key: 'cl_after', label: 'Clearance after', cl: true }, { key: 'volume_l', label: 'Apparent volume V (optional)', unit: 'L', optional: true }],
    Q5: [{ key: 'parent_t_half', label: 'Parent half-life', time: true }, { key: 'metabolite_t_half', label: 'Metabolite half-life', time: true }, { key: 't', label: 'Time since the reference (t)', time: true }],
    Q6: [{ key: 'replacement_half_time', label: 'Replacement half-time (hypothetical)', time: true }, { key: 't', label: 'Time since withdrawal (t)', time: true }],
    Q7: [{ key: 't_half', label: 'Half-life (t½)', unit: 'days' }, { key: 't', label: 'Time since the reference (t)', unit: 'days' }],
    Q8: [{ key: 't_half_lower', label: 'Lower half-life', time: true }, { key: 't_half_upper', label: 'Upper half-life', time: true }, { key: 't', label: 'Time since the reference (t)', time: true }],
    Q9: [], Q10: [],
    Q11: [{ key: 'age', label: 'Age', unit: 'years' }, { key: 'weight_kg', label: 'Weight', unit: 'kg' }, { key: 'scr', label: 'Serum creatinine', scr: true }]
  };
  function toNumber(v) {
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    if (v === null || v === undefined) return null;
    var s = String(v).trim().replace(MINUS, '-').replace(/,/g, '');
    if (s === '') return null;
    return /^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s) ? +s : 'NaN:' + s;
  }
  function numOrMissing(v) { var n = toNumber(v); return typeof n === 'string' ? n : n; }

  function evaluateTechnical(entry) {
    var rule = entry.rule, inp = entry.inputs || {}, meta = entry.meta || {}, gate = entry.gate || {};
    var rec = blankRecord({ Result_ID: entry.id, Label: entry.label || (rule + ' ' + (RULE_TITLES[rule] || '')), Calculation_Mode: rule,
      Equation: EQUATIONS[rule] || '', Limitations: MANDATORY_TEXT[rule] || '' });
    var pre = new Errs();
    if (!TECHNICAL[rule]) { block(rec, new CalcError(['E_MISSING_INPUT'], { E_MISSING_INPUT: 'a technical rule (Q1–Q11)' })); return [rec]; }
    var e = {}, k;
    // numeric fields
    (Q_PARAMS[rule] || []).forEach(function (p) {
      var raw = inp[p.key], n = toNumber(raw);
      if (typeof n === 'string') { pre.add('E_MISSING_INPUT', p.label.toLowerCase() + ' as a number'); n = null; }
      if (n === null && p.optional) { e[p.key] = null; }
      else e[p.key] = n;
      var unit = p.time ? inp.unit : p.cl ? inp.cl_unit : p.scr ? inp.scr_unit : p.unit;
      var m = meta[p.key] || {};
      if (!(n === null && p.optional)) {
        rec.Parameter_Set.push({ name: p.label, key: p.key, value: n, unit: unit || '', Parameter_Status: m.status || '',
          source: [m.source, m.version, m.section].filter(function (x) { return x; }).join(', '), source_document: m.source || '', source_version: m.version || '', source_section: m.section || '' });
      }
    });
    if (rule === 'Q1' || rule === 'Q2' || rule === 'Q3' || rule === 'Q5' || rule === 'Q6' || rule === 'Q8') e.unit = inp.unit;
    if (rule === 'Q7') e.unit = 'days';
    if (rule === 'Q4') { e.cl_unit = inp.cl_unit; e.volume_unit = 'L'; }
    if (rule === 'Q8') e.antibody = inp.antibody === true;
    if (rule === 'Q6') e.teaching_demonstration = entry.teaching_demonstration === true;
    if (rule === 'Q9') {
      e.segments = (inp.segments || []).map(function (s) {
        var r = toNumber(s.rate); if (typeof r === 'string') { pre.add('E_MISSING_INPUT', 'the segment rate as a number'); r = null; }
        return { start: s.start, end: s.end, rate: r, rate_unit: s.rate_unit };
      });
      e.onset = inp.onset;
      var pd = toNumber(inp.planned_dose_mg); if (typeof pd === 'string') { pre.add('E_MISSING_INPUT', 'the planned dose as a number'); pd = null; }
      e.planned_dose_mg = pd; e.planned_dose_unit = inp.planned_dose_unit || 'mg';
      (inp.segments || []).forEach(function (s, i) {
        rec.Parameter_Set.push({ name: 'Segment ' + (i + 1), key: 'segment_' + (i + 1), value: fmtTiming(s.start) + ' to ' + fmtTiming(s.end) + ', ' + s.rate + ' ' + (s.rate_unit || ''),
          unit: s.rate_unit || '', Parameter_Status: (meta.segments || {}).status || '', source: [(meta.segments || {}).source, (meta.segments || {}).version, (meta.segments || {}).section].filter(Boolean).join(', ') });
      });
      rec.Parameter_Set.push({ name: 'Onset', key: 'onset', value: fmtTiming(inp.onset), unit: 'timestamp', Parameter_Status: (meta.onset || {}).status || 'case measurement', source: (meta.onset || {}).source || '' });
      if (pd !== null) rec.Parameter_Set.push({ name: 'Planned dose', key: 'planned_dose_mg', value: pd, unit: 'mg', Parameter_Status: (meta.planned_dose_mg || {}).status || '', source: (meta.planned_dose_mg || {}).source || '' });
    }
    if (rule === 'Q10') {
      e.administrations = (inp.administrations || []).map(function (a, i) {
        var d = toNumber(a.dose), b = toNumber(a.bsa);
        if (typeof d === 'string') { pre.add('E_MISSING_INPUT', 'dose ' + (i + 1) + ' as a number'); d = null; }
        if (typeof b === 'string') { pre.add('E_MISSING_INPUT', 'BSA ' + (i + 1) + ' as a number'); b = null; }
        rec.Parameter_Set.push({ name: 'Administration ' + (i + 1), key: 'admin_' + (i + 1), value: (d === null ? '' : d) + ' ' + (a.dose_unit || '(unit not recorded)') + ' at BSA ' + (b === null ? 'not recorded' : b + ' m²'),
          unit: (a.dose_unit || '') + '; m²', Parameter_Status: (meta.administrations || {}).status || '', source: (meta.administrations || {}).source || '' });
        return { dose_mg: d, dose_unit: a.dose_unit === undefined ? '' : a.dose_unit, bsa_m2: b };
      });
      if (inp.method_b === 'first' || inp.method_b === 'last') e.method_b = inp.method_b;
      else if (inp.method_b === 'stated') { var sb = toNumber(inp.method_b_bsa); e.method_b = typeof sb === 'string' || sb === null ? '' : sb; if (sb === null || typeof sb === 'string') pre.add('E_MISSING_INPUT', 'the stated BSA for Method B'); }
    }
    if (rule === 'Q11') {
      e.scr_unit = inp.scr_unit; e.sex = inp.sex; e.creatinine_stable = entry.creatinine_stable === true; e.weight_unit = 'kg';
      if (!String(inp.weight_type || '').trim()) pre.add('E_MISSING_INPUT', 'the weight type (for example actual body weight)');
      rec.Parameter_Set.push({ name: 'Weight type', key: 'weight_type', value: inp.weight_type || '', unit: '', Parameter_Status: '', source: '' });
      rec.Parameter_Set.push({ name: 'Sex', key: 'sex', value: inp.sex || '', unit: '', Parameter_Status: (meta.sex || {}).status || 'case measurement', source: (meta.sex || {}).source || '' });
      if (inp.sample_date) rec.Parameter_Set.push({ name: 'Creatinine sample date', key: 'sample_date', value: inp.sample_date, unit: 'date', Parameter_Status: 'case measurement', source: (meta.scr || {}).source || '' });
      rec.Assumptions.push(entry.creatinine_stable === true ? 'Stable creatinine confirmed by the reviewer' + (entry.stability_basis ? ': ' + entry.stability_basis : '.') : 'Creatinine stability not confirmed.');
    }
    // gates
    if (COMMON_GATE_RULES.indexOf(rule) >= 0) {
      var allTicked = gate.model === true && gate.half_life === true && gate.reference === true && gate.status === true;
      if (allTicked) {
        var hlName = rule === 'Q4' ? 'the parameters used' : 'the half-life used';
        if (!String(gate.half_life_type || '').trim()) pre.add('E_MISSING_INPUT', 'the type of ' + hlName);
        if (!String(gate.half_life_population || '').trim()) pre.add('E_MISSING_INPUT', 'the population of ' + hlName);
        if (!String(gate.half_life_source || '').trim()) pre.add('E_MISSING_INPUT', 'the source of ' + hlName);
        if (!String(gate.reference_text || '').trim()) pre.add('E_MISSING_INPUT', 'the reference time');
        (Q_PARAMS[rule] || []).forEach(function (p) {
          var v = inp[p.key]; if (p.optional && (v === '' || v === null || v === undefined)) return;
          if (!(meta[p.key] || {}).status) pre.add('E_MISSING_INPUT', 'the Parameter_Status of ' + p.label.toLowerCase());
        });
        e.applicability_confirmed = true;
        rec.Assumptions.push('Gate confirmed: the simple model applies (one compartment, linear first-order elimination, constant parameters over the period).');
        rec.Assumptions.push('Gate confirmed: ' + hlName + ' – type: ' + (gate.half_life_type || '') + '; population: ' + (gate.half_life_population || '') + '; source: ' + (gate.half_life_source || '') + '.');
        rec.Assumptions.push('Gate confirmed: reference time – ' + (gate.reference_text || '') + '.');
        rec.Assumptions.push('Gate confirmed: each input’s Parameter_Status is recorded.');
      } else {
        e.applicability_confirmed = false;
        rec.Assumptions.push('Common gate not confirmed.');
      }
    }
    if (rule === 'Q6') rec.Assumptions.push(entry.teaching_demonstration === true ? 'Teaching-demonstration switch on.' : 'Teaching-demonstration switch off.');
    if (rule === 'Q8' && inp.antibody === true) rec.Assumptions.push('The molecule is a monoclonal antibody, so the Q7 caution applies to each bound.');
    if (entry.assumptions) rec.Assumptions.push(String(entry.assumptions));
    for (k in e) if (typeof e[k] === 'string' && /^NaN:/.test(e[k])) e[k] = null;
    try {
      var res;
      try { res = TECHNICAL[rule](e); }
      catch (x) { if (!(x instanceof CalcError)) throw x; pre.absorb(x); }
      pre.raise();
      addFlags(rec, res.flags);
      rec.Results = technicalResults(rule, res.out, e);
      rec.Display = rec.Results.map(function (r) { return r.name + ': ' + r.display; }).join('; ');
      rec.Unit = rec.Results.length ? rec.Results[0].unit : '';
    } catch (err) { block(rec, err); }
    return [rec];
  }
  function tDisplay(v, unit) {
    if (unit === 'minutes') return fmtMinutes(v) + ' minutes';
    return groupThousands(fixed(v, 1)) + ' ' + unit;
  }
  /** Build displayable results (unrounded value beside the displayed value) for each rule. */
  function technicalResults(rule, o, inp) {
    var R = [];
    function add(name, value, display, unit, key) { R.push({ name: name, key: key, value: value, display: display, unit: unit }); }
    switch (rule) {
      case 'Q1':
        add('Fraction remaining', o.fraction_remaining, percent3(o.fraction_remaining), 'fraction (shown as percent)', 'fraction_remaining');
        add('Elimination rate constant k', o.k_per_unit, sig3(o.k_per_unit) + ' per ' + inp.unit.replace(/s$/, ''), 'per ' + inp.unit.replace(/s$/, ''), 'k');
        break;
      case 'Q2':
        add('Time to reach ' + percent3(inp.target_fraction), o.time_to_fraction, tDisplay(o.time_to_fraction, inp.unit), inp.unit, 'time_to_fraction');
        add('Half-lives needed n', o.half_lives, fixed(o.half_lives, 2) + ' half-lives', 'half-lives', 'half_lives');
        break;
      case 'Q3':
        add('Accumulation factor R', o.accumulation_factor, fixed(o.accumulation_factor, 2), 'ratio', 'accumulation_factor'); break;
      case 'Q4':
        add('Half-life ratio (after / before)', o.half_life_ratio, fixed(o.half_life_ratio, 2), 'ratio', 'half_life_ratio');
        add('Average steady-state concentration ratio', o.average_css_ratio, fixed(o.average_css_ratio, 2), 'ratio', 'average_css_ratio');
        if (o.t_half_before_hours !== undefined) {
          add('Half-life before', o.t_half_before_hours, tDisplay(o.t_half_before_hours, 'hours'), 'hours', 't_half_before_hours');
          add('Half-life after', o.t_half_after_hours, tDisplay(o.t_half_after_hours, 'hours'), 'hours', 't_half_after_hours');
        }
        break;
      case 'Q5':
        add('Parent remaining (own reference)', o.parent_percent_remaining, percent3(o.parent_percent_remaining / 100), 'percent', 'parent_percent_remaining');
        add('Metabolite remaining (own reference)', o.metabolite_percent_remaining, percent3(o.metabolite_percent_remaining / 100), 'percent', 'metabolite_percent_remaining');
        break;
      case 'Q6':
        add('Fraction replaced', o.fraction_replaced, percent3(o.fraction_replaced), 'fraction (shown as percent)', 'fraction_replaced'); break;
      case 'Q7':
        add('Fraction remaining', o.fraction_remaining, percent3(o.fraction_remaining), 'fraction (shown as percent)', 'fraction_remaining');
        add('Elimination rate constant k', o.k_per_unit, sig3(o.k_per_unit) + ' per day', 'per day', 'k');
        break;
      case 'Q8':
        add('Fraction remaining at the lower half-life', o.fraction_lower, percent3(o.fraction_lower), 'fraction (shown as percent)', 'fraction_lower');
        add('Fraction remaining at the upper half-life', o.fraction_upper, percent3(o.fraction_upper), 'fraction (shown as percent)', 'fraction_upper');
        break;
      case 'Q9':
        add('Amount infused before onset', o.amount_mg, groupThousands(fixed(o.amount_mg, 1)) + ' mg', 'mg', 'amount_mg');
        if (o.percent_of_planned !== undefined) add('Percent of the planned dose', o.percent_of_planned, fixed(o.percent_of_planned, 1) + '%', 'percent', 'percent_of_planned');
        break;
      case 'Q10':
        if (o.per_administration_mg_m2 !== undefined) add('Cumulative dose, Method A (sum of dose / BSA at each administration)', o.per_administration_mg_m2, groupThousands(fixed(o.per_administration_mg_m2, 1)) + ' mg/m²', 'mg/m²', 'per_administration_mg_m2');
        else add('Cumulative dose, Method A, lower bound from the known administrations', o.per_administration_lower_bound_mg_m2, groupThousands(fixed(o.per_administration_lower_bound_mg_m2, 1)) + ' mg/m² (lower bound)', 'mg/m²', 'per_administration_lower_bound_mg_m2');
        if (o.single_bsa_first_mg_m2 !== undefined) add('Method B (total dose / first BSA)', o.single_bsa_first_mg_m2, groupThousands(fixed(o.single_bsa_first_mg_m2, 1)) + ' mg/m²', 'mg/m²', 'single_bsa_first_mg_m2');
        if (o.single_bsa_last_mg_m2 !== undefined) add('Method B (total dose / last BSA)', o.single_bsa_last_mg_m2, groupThousands(fixed(o.single_bsa_last_mg_m2, 1)) + ' mg/m²', 'mg/m²', 'single_bsa_last_mg_m2');
        if (o.single_bsa_stated_mg_m2 !== undefined) add('Method B (total dose / stated BSA ' + o.stated_bsa_m2 + ' m²)', o.single_bsa_stated_mg_m2, groupThousands(fixed(o.single_bsa_stated_mg_m2, 1)) + ' mg/m²', 'mg/m²', 'single_bsa_stated_mg_m2');
        break;
      case 'Q11':
        if (Array.isArray(o.crcl_ml_min)) {
          add('Estimated CrCl with the female factor (0.85)', o.crcl_ml_min[0], fixed(o.crcl_ml_min[0], 0) + ' mL/minute', 'mL/minute', 'crcl_female');
          add('Estimated CrCl without the female factor', o.crcl_ml_min[1], fixed(o.crcl_ml_min[1], 0) + ' mL/minute', 'mL/minute', 'crcl_male');
        } else add('Estimated creatinine clearance', o.crcl_ml_min, fixed(o.crcl_ml_min, 0) + ' mL/minute', 'mL/minute', 'crcl_ml_min');
        if (inp.scr_unit !== 'mg/dL') add('Serum creatinine in mg/dL', o.scr_mg_dl, fixed(o.scr_mg_dl, 2) + ' mg/dL', 'mg/dL', 'scr_mg_dl');
        break;
    }
    return R;
  }

  /** Evaluate a whole case: validation issues, interval records, recovery records, technical records. */
  function evaluateCase(c, opts) {
    opts = opts || {};
    var idx = indexCase(c), records = [];
    var validation = validateCase(c).map(function (i) { return { code: i.code, where: i.where, type: flagType(i.code), message: message(i.code, i.detail) }; });
    (c.intervals || []).forEach(function (q) { records = records.concat(evaluateInterval(c, q, idx)); });
    (c.recoveries || []).forEach(function (q) { records = records.concat(evaluateRecovery(c, q, idx)); });
    if (opts.technical !== false) (c.technical || []).forEach(function (t) { records = records.concat(evaluateTechnical(t)); });
    return { validation: validation, records: records };
  }

  // ------------------------------------------------------------------ export (Section 8) and re-import
  var CSV_FIELDS = ['Case_ID', 'Case_Version', 'Result_ID', 'Label', 'Drug_Event_Link', 'Exposure_Episode_ID', 'Event_ID', 'End_Point',
    'Interval_Value', 'Interval_Minimum', 'Interval_Maximum', 'Unit', 'Displayed_Value', 'Calculation_Mode', 'Anchor_Type',
    'Anchor_Record_ID', 'Anchor_Description', 'Sequence_Status', 'Treatment_Day', 'Reported_Timing', 'Constraint_Source', 'Results',
    'Parameter_Set', 'Assumptions', 'Calculation_Version', 'QC_Flags', 'Flag_Messages', 'Limitations', 'Export_Timestamp', 'Tool_Version'];

  function exportRecords(c, evaluation, timestamp) {
    return evaluation.records.map(function (r) {
      var x = clone(r);
      x.Case_ID = c.case_id || ''; x.Case_Version = c.case_version || '';
      x.Displayed_Value = r.Display;
      x.Export_Timestamp = timestamp; x.Tool_Version = TOOL_VERSION;
      return x;
    });
  }
  function buildExport(c, opts) {
    opts = opts || {};
    var ts = opts.timestamp || new Date().toISOString();
    var ev = evaluateCase(c);
    var records = exportRecords(c, ev, ts);
    var unresolved = [];
    records.forEach(function (r) { r.QC_Flags.forEach(function (f) { unresolved.push({ result: r.Result_ID, code: f, type: flagType(f), message: message(f) }); }); });
    ev.validation.forEach(function (v) { unresolved.push({ result: v.where, code: v.code, type: v.type, message: v.message }); });
    return {
      format: 'tto-companion-export/1', Tool_Version: TOOL_VERSION, Calculation_Version: SPEC_VERSION, Version_Label: VERSION_LABEL,
      Export_Timestamp: ts, Limitation_Text: LIMITATION_TEXT,
      case: clone(c), validation: ev.validation, records: records, unresolved_flags: unresolved,
      assessment: clone(c.assessment || []), versions: clone(opts.versions || [])
    };
  }
  function csvCell(v) {
    if (v === null || v === undefined) return '';
    var s;
    if (typeof v === 'number') s = String(v);
    else if (typeof v === 'string') s = v;
    else s = JSON.stringify(v);
    // formula-injection protection: a text cell starting with = + - @ tab or CR is prefixed with an apostrophe
    if (typeof v !== 'number' && /^[=+\-@\t\r]/.test(s) && !/^[+\-]?\d+(\.\d+)?$/.test(s)) s = "'" + s;
    if (/[",\r\n]/.test(s) || s !== s.trim()) s = '"' + s.replace(/"/g, '""') + '"';
    return s;
  }
  function toCSV(records) {
    var lines = [CSV_FIELDS.join(',')];
    records.forEach(function (r) {
      lines.push(CSV_FIELDS.map(function (f) {
        var v = r[f];
        if (f === 'QC_Flags') v = (r.QC_Flags || []).join('; ');
        if (f === 'Flag_Messages') v = (r.Messages || []).map(function (m) { return m.code + ': ' + m.message; }).join(' | ');
        if (f === 'Assumptions') v = (r.Assumptions || []).join(' | ');
        return csvCell(v);
      }).join(','));
    });
    return lines.join('\r\n') + '\r\n';
  }
  function parseCSV(text) {
    var rows = [], row = [], cell = '', i = 0, q = false, s = String(text).replace(/^﻿/, '');
    while (i < s.length) {
      var ch = s.charAt(i);
      if (q) {
        if (ch === '"') { if (s.charAt(i + 1) === '"') { cell += '"'; i++; } else q = false; }
        else cell += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') { if (ch === '\r' && s.charAt(i + 1) === '\n') i++; row.push(cell); rows.push(row); row = []; cell = ''; }
      else cell += ch;
      i++;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows;
  }
  /** Re-import an exported CSV into record objects (inverse of toCSV for the stored fields). */
  function recordsFromCSV(text) {
    var rows = parseCSV(text); if (!rows.length) return [];
    var head = rows[0], out = [];
    for (var r = 1; r < rows.length; r++) {
      if (rows[r].length === 1 && rows[r][0] === '') continue;
      var o = {};
      head.forEach(function (h, j) {
        var v = rows[r][j] === undefined ? '' : rows[r][j];
        if (/^'[=+\-@\t\r]/.test(v)) v = v.slice(1);
        if (['Interval_Value', 'Interval_Minimum', 'Interval_Maximum', 'Treatment_Day'].indexOf(h) >= 0) v = v === '' ? null : +v;
        else if (['Reported_Timing', 'Results', 'Parameter_Set'].indexOf(h) >= 0) v = v === '' ? null : JSON.parse(v);
        else if (h === 'QC_Flags') v = v === '' ? [] : v.split('; ');
        else if (h === 'Assumptions') v = v === '' ? [] : v.split(' | ');
        o[h] = v;
      });
      out.push(o);
    }
    return out;
  }
  /** Accept an exported bundle or a saved case file; return the case object. */
  function importBundle(obj) {
    if (!obj || typeof obj !== 'object') throw new Error('Not a TTO companion file.');
    var c = obj.format === 'tto-companion-export/1' ? obj.case : obj.schema === 'tto-case/1' ? obj : null;
    if (!c) throw new Error('Not a TTO companion case or export file.');
    var base = emptyCase();
    for (var k in base) if (c[k] === undefined) c[k] = base[k];
    return { case: c, versions: obj.versions || [] };
  }
  function escapeHTML(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function recordText(r) {
    var lines = [];
    lines.push('[' + r.Result_ID + '] ' + r.Label);
    lines.push('  Result: ' + (r.Blocked ? 'not calculable' : r.Display) + (r.Unit && !r.Blocked ? '' : ''));
    if (!r.Blocked && (r.Interval_Value !== null || r.Interval_Minimum !== null)) {
      lines.push('  Unrounded: ' + (r.Interval_Value !== null ? r.Interval_Value : r.Interval_Minimum + ' to ' + r.Interval_Maximum) + ' ' + r.Unit);
    }
    (r.Results || []).forEach(function (x) { lines.push('  ' + x.name + ': ' + x.display + ' (unrounded ' + (Array.isArray(x.value) ? x.value.join(', ') : x.value) + ')'); });
    if (r.Treatment_Day !== null && r.Treatment_Day !== undefined) lines.push('  Treatment day (inclusive): Day ' + r.Treatment_Day);
    lines.push('  Rule: ' + r.Calculation_Mode + (r.Anchor_Type ? '; anchor ' + r.Anchor_Type + (r.Anchor_Record_ID ? ' (' + r.Anchor_Record_ID + ')' : '') : '') +
      (r.Drug_Event_Link ? '; link ' + r.Drug_Event_Link : '') + (r.Sequence_Status ? '; sequence ' + r.Sequence_Status : ''));
    if (r.Reported_Timing) lines.push('  Reported timing: “' + r.Reported_Timing.text + '” – ' + r.Reported_Timing.comparison);
    if (r.Constraint_Source) lines.push('  Constraint source: ' + r.Constraint_Source);
    (r.Parameter_Set || []).forEach(function (p) {
      lines.push('  Input ' + p.name + ': ' + (p.display || p.value) + (p.unit ? ' [' + p.unit + ']' : '') + (p.precision ? '; precision ' + p.precision : '') +
        (p.original_text ? '; as reported “' + p.original_text + '”' : '') + (p.Parameter_Status ? '; status ' + p.Parameter_Status : '') + (p.source ? '; source ' + p.source : ''));
    });
    (r.Assumptions || []).forEach(function (a) { lines.push('  Assumption: ' + a); });
    lines.push('  Flags: ' + (r.QC_Flags.length ? r.Messages.map(function (m) { return m.code + ' – ' + m.message; }).join(' | ') : 'none'));
    if (r.Limitations) lines.push('  Limitations: ' + r.Limitations);
    return lines.join('\n');
  }
  function toText(bundle) {
    var c = bundle.case, L = [];
    L.push(VERSION_LABEL);
    L.push('LIMITATION: ' + LIMITATION_TEXT);
    L.push('Case ' + (c.case_id || '(no ID)') + ', version ' + (c.case_version || '') + '. Exported ' + bundle.Export_Timestamp + '.');
    L.push('Calculation_Version: ' + SPEC_VERSION + '; Tool_Version: ' + TOOL_VERSION);
    if (bundle.validation.length) { L.push(''); L.push('Validation messages:'); bundle.validation.forEach(function (v) { L.push('  ' + v.code + ' (' + v.where + '): ' + v.message); }); }
    L.push('');
    bundle.records.forEach(function (r) { L.push(recordText(r)); L.push(''); });
    (bundle.assessment || []).forEach(function (a) {
      if (!a) return;
      L.push('Assessment for link ' + a.link_id + ' (reviewer judgment, not a causality category): biological fit ' + (a.fit || 'not recorded') + '. Rationale: ' + (a.rationale || '(none)'));
    });
    L.push('Unresolved flags: ' + (bundle.unresolved_flags.length ? bundle.unresolved_flags.map(function (u) { return u.result + ' ' + u.code; }).join(', ') : 'none'));
    return L.join('\n');
  }
  function toHTML(bundle) {
    var c = bundle.case, h = [];
    h.push('<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">');
    h.push('<title>TTO result record ' + escapeHTML(c.case_id) + '</title><style>body{font:14px/1.5 system-ui,sans-serif;margin:24px;color:#1b1f24;background:#fff}' +
      'h1{font-size:20px}h2{font-size:16px;margin-top:20px}.lim{border:2px solid #9a5b00;background:#fff7e6;padding:10px}table{border-collapse:collapse;width:100%}' +
      'td,th{border:1px solid #bbb;padding:4px 6px;vertical-align:top;text-align:left}.E{color:#a40000}.W{color:#7a4a00}.I{color:#1f4e8c}</style></head><body>');
    h.push('<h1>Result record: case ' + escapeHTML(c.case_id) + ', version ' + escapeHTML(c.case_version) + '</h1>');
    h.push('<p class="lim"><strong>Limitation.</strong> ' + escapeHTML(LIMITATION_TEXT) + '</p>');
    h.push('<p>' + escapeHTML(VERSION_LABEL) + ' &middot; Calculation_Version ' + escapeHTML(SPEC_VERSION) + ' &middot; Export_Timestamp ' + escapeHTML(bundle.Export_Timestamp) + '</p>');
    if (bundle.validation.length) {
      h.push('<h2>Validation messages</h2><ul>');
      bundle.validation.forEach(function (v) { h.push('<li class="' + v.code.charAt(0) + '">' + escapeHTML(v.code + ' (' + v.where + '): ' + v.message) + '</li>'); });
      h.push('</ul>');
    }
    bundle.records.forEach(function (r) {
      h.push('<h2>' + escapeHTML(r.Result_ID + ' – ' + r.Label) + '</h2><table>');
      function row(k, v) { h.push('<tr><th>' + escapeHTML(k) + '</th><td>' + v + '</td></tr>'); }
      row('Displayed value', escapeHTML(r.Blocked ? 'Not calculable' : r.Display));
      row('Interval_Value / Minimum / Maximum', escapeHTML([r.Interval_Value, r.Interval_Minimum, r.Interval_Maximum].map(function (x) { return x === null || x === undefined ? 'blank' : x; }).join(' / ')));
      if ((r.Results || []).length) row('Results (unrounded)', escapeHTML(r.Results.map(function (x) { return x.name + ': ' + (Array.isArray(x.value) ? x.value.join(', ') : x.value) + ' (shown ' + x.display + ')'; }).join('; ')));
      row('Unit', escapeHTML(r.Unit));
      row('Calculation_Mode', escapeHTML(r.Calculation_Mode));
      row('Anchor_Type, Anchor_Record_ID', escapeHTML((r.Anchor_Type || '') + ', ' + (r.Anchor_Record_ID || '')));
      row('Drug_Event_Link', escapeHTML(r.Drug_Event_Link));
      row('Sequence_Status', escapeHTML(r.Sequence_Status));
      if (r.Treatment_Day !== null && r.Treatment_Day !== undefined) row('Treatment day (C3, inclusive)', escapeHTML('Day ' + r.Treatment_Day));
      if (r.Reported_Timing) row('Reported timing', escapeHTML('“' + r.Reported_Timing.text + '” – ' + r.Reported_Timing.comparison));
      if (r.Constraint_Source) row('Constraint source', escapeHTML(r.Constraint_Source));
      row('Parameter_Set', (r.Parameter_Set || []).map(function (p) {
        return escapeHTML(p.name + ': ' + (p.display || p.value) + (p.unit ? ' [' + p.unit + ']' : '') + (p.precision ? '; precision ' + p.precision : '') +
          (p.original_text ? '; as reported “' + p.original_text + '”' : '') + (p.Parameter_Status ? '; status ' + p.Parameter_Status : '') + (p.source ? '; source ' + p.source : ''));
      }).join('<br>'));
      row('Assumptions', (r.Assumptions || []).map(escapeHTML).join('<br>'));
      row('QC_Flags', r.QC_Flags.length ? r.Messages.map(function (m) { return '<span class="' + m.code.charAt(0) + '">' + escapeHTML(m.code + ': ' + m.message) + '</span>'; }).join('<br>') : 'none');
      row('Limitations', escapeHTML(r.Limitations));
      row('Calculation_Version, Tool_Version', escapeHTML(SPEC_VERSION + ', ' + TOOL_VERSION));
      h.push('</table>');
    });
    if ((bundle.assessment || []).length) {
      h.push('<h2>Assessment (reviewer judgment; not a causality category)</h2>');
      bundle.assessment.forEach(function (a) {
        if (!a) return;
        h.push('<p><strong>Link ' + escapeHTML(a.link_id) + '</strong>: biological fit ' + escapeHTML(a.fit || 'not recorded') + '. Rationale: ' + escapeHTML(a.rationale || '(none)') + '</p>');
        var pr = a.prompts || {};
        Object.keys(pr).forEach(function (k) { if (pr[k] && (pr[k].text || pr[k].source)) h.push('<p>' + escapeHTML(k.replace(/_/g, ' ') + ': ' + (pr[k].text || '') + (pr[k].source ? ' [source: ' + pr[k].source + ']' : '')) + '</p>'); });
      });
    }
    h.push('<p class="lim">' + escapeHTML(LIMITATION_TEXT) + '</p></body></html>');
    return h.join('\n');
  }

  return {
    SPEC_VERSION: SPEC_VERSION, TOOL_VERSION: TOOL_VERSION, VERSION_LABEL: VERSION_LABEL, LIMITATION_TEXT: LIMITATION_TEXT,
    MESSAGES: MESSAGES, MANDATORY_TEXT: MANDATORY_TEXT, RULE_TITLES: RULE_TITLES, EQUATIONS: EQUATIONS,
    ANCHOR_TYPES: ANCHOR_TYPES, SEQUENCE_STATUSES: SEQUENCE_STATUSES, PARAMETER_STATUSES: PARAMETER_STATUSES,
    TIME_UNITS: TIME_UNITS, REPORTED_UNITS: REPORTED_UNITS, CLEARANCE_UNITS: CLEARANCE_UNITS, SCR_UNITS: SCR_UNITS,
    Q_PARAMS: Q_PARAMS, COMMON_GATE_RULES: COMMON_GATE_RULES, CSV_FIELDS: CSV_FIELDS,
    CalcError: CalcError, flagType: flagType, message: message,
    parseDateStr: parseDateStr, parseTimeStr: parseTimeStr, parseOffset: parseOffset, timingFromIso: timingFromIso,
    timingToIso: timingToIso, resolveTiming: resolveTiming, dayBounds: dayBounds, isoDay: isoDay, dayNumber: dayNumber,
    fmtDay: fmtDay, fmtDateStr: fmtDateStr, fmtTiming: fmtTiming, fmtOffset: fmtOffset, fmtRangeDays: fmtRangeDays, percent3: percent3, fixed: fixed,
    displayInterval: displayInterval, precisionLabel: precisionLabel,
    c1Exact: c1Exact, c2Calendar: c2Calendar, c3TreatmentDay: c3TreatmentDay, c4Bounded: c4Bounded, computeInterval: computeInterval,
    selectLastDose: selectLastDose, c6Compare: c6Compare,
    technical: TECHNICAL, q1: q1, q2: q2, q3: q3, q4: q4, q5: q5, q6: q6, q7: q7, q8: q8, q9: q9, q10: q10, q11: q11,
    emptyCase: emptyCase, clone: clone, indexCase: indexCase, validateCase: validateCase, findLink: findLink,
    evaluateInterval: evaluateInterval, evaluateRecovery: evaluateRecovery, evaluateTechnical: evaluateTechnical, evaluateCase: evaluateCase,
    recordLabel: recordLabel, buildExport: buildExport, toCSV: toCSV, parseCSV: parseCSV, recordsFromCSV: recordsFromCSV, importBundle: importBundle,
    toText: toText, toHTML: toHTML, escapeHTML: escapeHTML, recordText: recordText
  };
}));
