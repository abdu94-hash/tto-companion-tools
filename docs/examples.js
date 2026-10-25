/*
 * TTO Companion Tools v1.1 - fictional teaching cases from the book, as calculator inputs (CALC-SPEC 1.1).
 * Every date, time and value is taken unchanged from the case's home chapter:
 *   C01 (version 2): Chapter 3, Section 3.7     C04 (version 2): Chapter 11, Section 11.8
 *   C28 (version 2): Chapter 20, Section 20.4   (the C04 Q8 rows follow Chapter 23, Section 23.11)
 * C28's onset is entered as a laboratory-defined onset (CALC-SPEC 1.1, C4): after the last undetectable HBV DNA
 * (MS-NEG3) and on or before the first positive result (MS-DET); the bounds are the chapter's 9 February to 12 May 2026.
 * All cases are fictional. No real patient data.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) { module.exports = factory(); }
  else { root.TTO_EXAMPLES = factory(); }
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  function ts(date, time, offset, original) { return { type: 'timestamp', date: date, time: time, offset: offset, original: original || '' }; }
  function d(date, original) { return { type: 'date', date: date, original: original || '' }; }
  function range(lower, upper, original) { return { type: 'partial', precision: 'range', lower: lower, upper: upper, original: original || '' }; }
  var GATE_OFF = { model: false, half_life: false, half_life_type: '', half_life_population: '', half_life_source: '', reference: false, reference_text: '', status: false };

  var C01 = {
    schema: 'tto-case/1', case_id: 'C01', case_version: '2', reviewer: '', setting: 'Hospital infusion; local time recorded at UTC−05:00',
    synthetic_note: 'Fictional teaching case (Chapter 3, Section 3.7). Version 2 = initial report plus follow-up received 14 January 2026.',
    notes: 'Drug A and Drug B are placeholders. The Drug A concentration in the bag is not reported, so the amount infused (Q9) is not calculable.',
    events: [
      { id: 'EV1', term: 'Urticaria, wheeze and hypotension', source: 'Nursing observation chart, 10 January 2026',
        onset: ts('2026-01-10', '09:15', '-05:00', '09:15 pt reports itching; urticaria chest and arms; audible wheeze; BP 84/52') }
    ],
    milestones: [
      { id: 'MS-BASE', kind: 'baseline', label: 'Baseline observations stable', value: 'BP 124/78, HR 76, SpO2 98%, skin clear', source: 'Nursing observation chart', timing: ts('2026-01-10', '08:50', '-05:00', '08:50 pre-infusion obs') },
      { id: 'MS-RX', kind: 'intervention', label: 'Emergency treatment begins', value: 'Agents not given in the source', source: 'Medication administration record', timing: ts('2026-01-10', '09:18', '-05:00', 'Emergency medicines administered from 09:18') },
      { id: 'MS-IMP', kind: 'recovery', label: 'Improvement documented', value: 'Wheeze settling, BP 108/70', source: 'Nursing observation chart', timing: ts('2026-01-10', '09:35', '-05:00', '09:35 wheeze settling, BP 108/70') },
      { id: 'MS-DX', kind: 'diagnosis', label: 'Diagnosis recorded (version 2)', value: 'Acute reaction during first Drug A infusion', source: 'Physician note, 09:50 (version 2)', timing: ts('2026-01-10', '09:50', '-05:00', 'written at 09:50') },
      { id: 'MS-TRY', kind: 'test', label: 'Tryptase sample (version 2)', value: 'Result pending', source: 'Follow-up, version 2', timing: ts('2026-01-10', '10:05', '-05:00', '10:05 tryptase sample') }
    ],
    exposures: [
      { id: 'EX-A', drug: 'Drug A', role: 'suspect', details: 'IV infusion, first exposure (version 2); pump 50 mL/hour from 09:00 (version 2); bag concentration not reported', source: 'Medication administration record; nursing chart',
        records: [
          { id: 'A-START', kind: 'start', label: 'Infusion started', timing: ts('2026-01-10', '09:00', '-05:00', 'Drug A IV, started 09:00') },
          { id: 'A-STOP', kind: 'stop', label: 'Infusion stopped', timing: ts('2026-01-10', '09:16', '-05:00', 'stopped 09:16 (reaction)') }
        ] },
      { id: 'EX-B', drug: 'Drug B', role: 'concomitant', details: 'IV infusion via the same cannula (version 2)', source: 'Follow-up, version 2',
        records: [
          { id: 'B-START', kind: 'start', label: 'Infusion started', timing: ts('2026-01-10', '08:20', '-05:00', '08:20') },
          { id: 'B-END', kind: 'stop', label: 'Infusion ended', timing: ts('2026-01-10', '08:50', '-05:00', '08:50') }
        ] },
      { id: 'EX-ANT', drug: 'Skin antiseptic', role: 'concomitant', details: 'Topical, at cannulation (version 2)', source: 'Follow-up, version 2',
        records: [{ id: 'ANT-1', kind: 'administration', label: 'Applied at cannulation', timing: ts('2026-01-10', '08:15', '-05:00', '08:15') }] }
    ],
    links: [
      { id: 'L1', exposure_id: 'EX-A', event_id: 'EV1', reason: 'Suspect drug' },
      { id: 'L2', exposure_id: 'EX-B', event_id: 'EV1', reason: 'Competing exposure (version 2)' },
      { id: 'L3', exposure_id: 'EX-ANT', event_id: 'EV1', reason: 'Competing exposure (version 2)' }
    ],
    intervals: [
      { id: 'I1', exposure_id: 'EX-A', event_id: 'EV1', anchor_type: 'FIRST_ADMIN', anchor_ref: '', end_ref: 'onset', date_level: true, treatment_day: true, note: 'Primary TTO' },
      { id: 'I2', exposure_id: 'EX-A', event_id: 'EV1', anchor_type: 'LAST_DOSE_BEFORE_ONSET', anchor_ref: '', end_ref: 'onset' },
      { id: 'I3', exposure_id: 'EX-B', event_id: 'EV1', anchor_type: 'CONCOMITANT_CHANGE', anchor_ref: 'B-START', end_ref: 'onset' },
      { id: 'I4', exposure_id: 'EX-B', event_id: 'EV1', anchor_type: 'OTHER', anchor_description: 'End of Drug B infusion', anchor_ref: 'B-END', end_ref: 'onset' },
      { id: 'I5', exposure_id: 'EX-ANT', event_id: 'EV1', anchor_type: 'OTHER', anchor_description: 'Topical exposure at cannulation', anchor_ref: 'ANT-1', end_ref: 'onset' },
      { id: 'I6', exposure_id: 'EX-A', event_id: 'EV1', anchor_type: 'CESSATION', anchor_ref: 'A-STOP', end_ref: 'onset', note: 'Exposure continued 1 minute after onset' }
    ],
    recoveries: [
      { id: 'R1', exposure_id: 'EX-A', event_id: 'EV1', cessation_ref: 'A-STOP', recovery_ref: 'MS-IMP' }
    ],
    technical: [
      { id: 'T1', rule: 'Q9', label: 'Drug A infused before onset (version 2 pump rate)',
        inputs: { segments: [{ start: ts('2026-01-10', '09:00', '-05:00'), end: ts('2026-01-10', '09:16', '-05:00'), rate: '50', rate_unit: 'mL/hour' }],
          onset: ts('2026-01-10', '09:15', '-05:00'), planned_dose_mg: '' },
        meta: { segments: { status: 'case measurement', source: 'Follow-up, version 2' }, onset: { status: 'case measurement', source: 'Nursing observation chart' } },
        assumptions: 'The bag concentration is not reported, so mL/hour cannot be expressed in mg/hour; the concentration is the targeted query.' }
    ],
    assessment: []
  };

  var C04 = {
    schema: 'tto-case/1', case_id: 'C04', case_version: '2', reviewer: '', setting: 'Primary care and emergency department; times recorded without UTC offset',
    synthetic_note: 'Fictional teaching case (Chapter 11, Section 11.8); digoxin is named because the teaching depends on its label facts. Version 2 = GP letter received 20 July 2026 and ED and ward notes.',
    notes: 'Q4 can be run only as a labeled scenario (Chapter 11, Worked exercise 11.2), because no measured clearance exists for either period; it is not included here.',
    events: [
      { id: 'EV1', term: 'Nausea and symptomatic bradycardia', source: 'Emergency department note, 5 July 2026',
        onset: d('2026-07-05', 'Nausea since this morning, dizzy, pulse 38') }
    ],
    milestones: [
      { id: 'MS-CR1', kind: 'organ_function', label: 'Creatinine, baseline', value: '1.0 mg/dL', source: 'GP letter', timing: d('2025-11-14', '14 November 2025') },
      { id: 'MS-CR2', kind: 'organ_function', label: 'Creatinine, baseline', value: '1.0 mg/dL', source: 'GP letter', timing: d('2026-03-12', '12 March 2026') },
      { id: 'MS-FLUID', kind: 'other', label: 'Reduced fluid intake', value: 'drinking less “for a few days” before 1 July', source: 'GP letter (seen 1 July)', timing: { type: 'none', original: 'for a few days' } },
      { id: 'MS-REN', kind: 'organ_function', label: 'First documented renal deterioration', value: 'Creatinine 1.7 mg/dL', source: 'GP letter', timing: d('2026-07-01', 'Seen 1 July ... creatinine 1.7') },
      { id: 'MS-CR3', kind: 'organ_function', label: 'Creatinine at onset (still rising)', value: '2.4 mg/dL; potassium 4.8 mmol/L', source: 'ED note', timing: d('2026-07-05', '5 July 2026') },
      { id: 'MS-DIG', kind: 'test', label: 'Serum digoxin sample', value: '2.9 ng/mL (laboratory reference 0.5–2.0)', source: 'ED note', timing: ts('2026-07-05', '11:40', 'unknown', 'sample 11:40') },
      { id: 'MS-FLU', kind: 'intervention', label: 'IV fluids and cardiac monitoring', value: 'Digoxin and ibuprofen withheld; IV fluids', source: 'ED note', timing: d('2026-07-05', '5 July 2026') },
      { id: 'MS-IMP', kind: 'recovery', label: 'Improvement documented', value: 'Nausea resolved, ventricular rate 64; creatinine 1.9', source: 'Ward note, 8 July 2026', timing: d('2026-07-08', '8 July 2026') }
    ],
    exposures: [
      { id: 'EX-DIG', drug: 'Digoxin', role: 'suspect', details: '125 micrograms each morning since 1 January 2026, dose never changed', source: 'GP letter; ED note',
        records: [
          { id: 'DIG-START', kind: 'start', label: 'Started', timing: d('2026-01-01', 'since 1 January 2026') },
          { id: 'DIG-LAST', kind: 'administration', label: 'Last dose (reported)', timing: ts('2026-07-04', '08:00', 'unknown', 'Took last digoxin yesterday morning, about 08:00') },
          { id: 'DIG-STOP', kind: 'stop', label: 'Withheld', timing: d('2026-07-05', 'Digoxin and ibuprofen withheld') }
        ] },
      { id: 'EX-IBU', drug: 'Ibuprofen', role: 'concomitant', details: '400 mg three times a day for back pain (version 2 only)', source: 'GP letter',
        records: [
          { id: 'IBU-START', kind: 'start', label: 'Started', timing: d('2026-06-26', 'Bought ibuprofen ... on 26 June') },
          { id: 'IBU-STOP', kind: 'stop', label: 'Last taken', timing: d('2026-07-04', 'last taken 4 July') }
        ] },
      { id: 'EX-BB', drug: 'Beta-blocker (product not named)', role: 'concomitant', details: 'Long term; dose unchanged since 2024', source: 'GP letter', records: [] }
    ],
    links: [{ id: 'L1', exposure_id: 'EX-DIG', event_id: 'EV1', reason: 'Suspect drug' }],
    intervals: [
      { id: 'I1', exposure_id: 'EX-DIG', event_id: 'EV1', anchor_type: 'FIRST_ADMIN', anchor_ref: '', end_ref: 'onset' },
      { id: 'I2', exposure_id: 'EX-DIG', event_id: 'EV1', anchor_type: 'OTHER', anchor_description: 'First documented renal deterioration', anchor_ref: 'MS-REN', end_ref: 'onset' },
      { id: 'I3', exposure_id: 'EX-DIG', event_id: 'EV1', anchor_type: 'CONCOMITANT_CHANGE', anchor_ref: 'IBU-START', end_ref: 'onset' },
      { id: 'I4', exposure_id: 'EX-DIG', event_id: 'EV1', anchor_type: 'LAST_DOSE_BEFORE_ONSET', anchor_ref: '', end_ref: 'onset' },
      { id: 'I5', exposure_id: 'EX-DIG', event_id: 'EV1', anchor_type: 'OTHER', anchor_description: 'Reported last dose (about 08:00)', anchor_ref: 'DIG-LAST', end_ref: 'MS-DIG', note: 'Sample time after the last dose' }
    ],
    recoveries: [{ id: 'R1', exposure_id: 'EX-DIG', event_id: 'EV1', cessation_ref: 'DIG-STOP', recovery_ref: 'MS-IMP' }],
    technical: [
      { id: 'T1', rule: 'Q11', label: 'Estimated CrCl at baseline (creatinine stable)', creatinine_stable: true, stability_basis: '1.0 mg/dL on 14 November 2025 and 12 March 2026',
        inputs: { age: '81', weight_kg: '58', weight_type: 'actual body weight', scr: '1.0', scr_unit: 'mg/dL', sex: 'female', sample_date: '2026-03-12' },
        meta: { age: { status: 'case measurement', source: 'GP letter' }, weight_kg: { status: 'case measurement', source: 'GP letter' }, scr: { status: 'case measurement', source: 'GP letter' }, sex: { status: 'case measurement', source: 'GP letter' } } },
      { id: 'T2', rule: 'Q11', label: 'Estimated CrCl on 1 July (creatinine rising)', creatinine_stable: false,
        inputs: { age: '81', weight_kg: '58', weight_type: 'actual body weight', scr: '1.7', scr_unit: 'mg/dL', sex: 'female', sample_date: '2026-07-01' },
        meta: { age: { status: 'case measurement', source: 'GP letter' }, weight_kg: { status: 'case measurement', source: 'GP letter' }, scr: { status: 'case measurement', source: 'GP letter' } } },
      { id: 'T3', rule: 'Q11', label: 'Estimated CrCl on 5 July (creatinine rising)', creatinine_stable: false,
        inputs: { age: '81', weight_kg: '58', weight_type: 'actual body weight', scr: '2.4', scr_unit: 'mg/dL', sex: 'female', sample_date: '2026-07-05' },
        meta: { age: { status: 'case measurement', source: 'GP letter' }, weight_kg: { status: 'case measurement', source: 'ED note' }, scr: { status: 'case measurement', source: 'ED note' } } },
      { id: 'T4', rule: 'Q8', label: 'Fall after withholding: normal renal function half-life (scenario)',
        inputs: { t_half_lower: '1.5', t_half_upper: '2', t: '3', unit: 'days', antibody: false },
        meta: { t_half_lower: { status: 'label fact', source: 'LANOXIN US', section: '12.3' }, t_half_upper: { status: 'label fact', source: 'LANOXIN US', section: '12.3' }, t: { status: 'case measurement', source: 'ED and ward notes (5 to 8 July)' } },
        gate: { model: true, half_life: true, half_life_type: 'elimination half-life', half_life_population: 'healthy volunteers with normal renal function (not this patient)', half_life_source: 'LANOXIN US 12.3',
          reference: true, reference_text: 'concentration at withholding on 5 July 2026; no further doses', status: true },
        assumptions: 'Chapter 23, Section 23.11. The patient’s renal function lies between the two labeled populations.' },
      { id: 'T5', rule: 'Q8', label: 'Fall after withholding: anuric half-life (scenario)',
        inputs: { t_half_lower: '3.5', t_half_upper: '5', t: '3', unit: 'days', antibody: false },
        meta: { t_half_lower: { status: 'label fact', source: 'LANOXIN US', section: '12.3' }, t_half_upper: { status: 'label fact', source: 'LANOXIN US', section: '12.3' }, t: { status: 'case measurement', source: 'ED and ward notes (5 to 8 July)' } },
        gate: { model: true, half_life: true, half_life_type: 'elimination half-life', half_life_population: 'anuric patients (not this patient)', half_life_source: 'LANOXIN US 12.3',
          reference: true, reference_text: 'concentration at withholding on 5 July 2026; no further doses', status: true },
        assumptions: 'Chapter 23, Section 23.11.' }
    ],
    assessment: [
      { link_id: 'L1', fit: 'compatible',
        rationale: 'Example entry from Chapters 11 and 23 (reviewer judgment, not a causality category): onset followed documented renal deterioration during unchanged dosing, with a well-timed concentration of 2.9 ng/mL; six months of prior tolerance does not exclude increased exposure after a clearance change. Bradycardia from the unchanged beta-blocker and other cardiac causes remain to be reviewed.',
        prompts: {
          mechanism_hypothesis: { text: 'Concentration-related digoxin toxicity after reduced renal clearance', source: 'LANOXIN US 8.6; 12.3; 5.3' },
          expected_exposure_course: { text: 'Exposure rises at an unchanged dose when renal clearance falls, at a rate set by the lengthening half-life', source: 'LANOXIN US 12.3' },
          expected_pd_course: { text: '', source: '' },
          clinical_latency_evidence: { text: '', source: '' },
          observed_course: { text: '185 calendar days from the start; 4 from the first documented renal deterioration; 9 from the ibuprofen start', source: 'Arithmetic display' },
          temporal_fit_rationale: { text: '', source: '' } },
        followup: 'Creatinine between March and July; magnesium, calcium and thyroid function; earlier ECGs; the beta-blocker and dose; adherence; the full treatment record for 5–8 July; any repeat digoxin level with its sample time.' }
    ]
  };

  var C28 = {
    schema: 'tto-case/1', case_id: 'C28', case_version: '2', reviewer: '', setting: 'Hematology and hepatology follow-up',
    synthetic_note: 'Fictional teaching case (Chapter 20, Section 20.4); dates are hypothetical. Version 2 = source excerpts received 15 June 2026.',
    notes: 'The US label reports HBV reactivation up to 24 months after completion of therapy (RITUXAN US 5.3); the EU SmPC gives no timing (MabThera EU 4.4). Months are compared on the calendar, never converted to days.',
    events: [
      { id: 'EV1', term: 'Hepatitis B virus reactivation (virological)', source: 'Hematology summary; hepatology letter',
        onset: { type: 'lab', last_normal_ref: 'MS-NEG3', first_abnormal_ref: 'MS-DET', original: 'Between the last undetectable HBV DNA (9 February 2026) and the first positive result (12 May 2026)' } }
    ],
    milestones: [
      { id: 'MS-BASE', kind: 'baseline', label: 'Baseline serology', value: 'HBsAg negative, anti-HBc positive, HBV DNA not detected', source: 'Hematology summary', timing: d('2024-11-11', '11/11/2024') },
      { id: 'MS-NEG1', kind: 'test', label: 'HBV DNA not detected', value: 'not detected', source: 'Hematology summary', timing: d('2025-09-01', '01/09/2025') },
      { id: 'MS-NEG2', kind: 'test', label: 'HBV DNA not detected', value: 'not detected', source: 'Hematology summary', timing: d('2025-12-01', '01/12/2025') },
      { id: 'MS-NEG3', kind: 'test', label: 'Last undetectable HBV DNA', value: 'not detected', source: 'Hematology summary', timing: d('2026-02-09', '09/02/2026') },
      { id: 'MS-DET', kind: 'detection', label: 'First positive HBV DNA (first detection)', value: 'HBV DNA 1,850 IU/mL; ALT 28 U/L', source: 'Hepatology letter', timing: d('2026-05-12', '12/05/2026') },
      { id: 'MS-BCELL', kind: 'test', label: 'CD19 B cells and IgG below reference range', value: 'values not reported', source: 'Hepatology letter', timing: d('2026-05-12', '12/05/2026') },
      { id: 'MS-HEP', kind: 'other', label: 'Biochemical hepatitis', value: 'HBV DNA 64,000 IU/mL; HBsAg positive; ALT 142 U/L (ULN 40)', source: 'Hepatology letter', timing: d('2026-05-26', '26/05/2026') }
    ],
    exposures: [
      { id: 'EX-RTX', drug: 'Rituximab', role: 'suspect', details: 'IV, six cycles with Regimen X (R-Regimen X × 6); intermediate cycle dates not itemized in the source', source: 'Hematology summary (version 1; 2)',
        records: [
          { id: 'RTX-FIRST', kind: 'start', label: 'First infusion', timing: d('2024-11-18', '18/11/2024') },
          { id: 'RTX-LAST', kind: 'administration', label: 'Last infusion', timing: d('2025-03-03', '03/03/2025') }
        ] },
      { id: 'EX-RGX', drug: 'Regimen X (placeholder; includes a glucocorticoid)', role: 'concomitant', details: 'Combination chemotherapy, same cycles', source: 'Hematology summary (version 2)',
        records: [
          { id: 'RGX-START', kind: 'start', label: 'First cycle', timing: d('2024-11-18', '18/11/2024') },
          { id: 'RGX-END', kind: 'stop', label: 'Last cycle', timing: d('2025-03-03', '03/03/2025') }
        ] },
      { id: 'EX-AVP', drug: 'Antiviral A (placeholder), prophylaxis', role: 'prophylaxis', details: 'Oral, daily', source: 'Hematology summary (version 2)',
        records: [
          { id: 'AVP-START', kind: 'start', label: 'Prophylaxis started', timing: d('2024-11-11', '11/11/2024') },
          { id: 'AVP-STOP', kind: 'stop', label: 'Prophylaxis stopped', timing: d('2025-09-03', '03/09/2025') }
        ] },
      { id: 'EX-PRED', drug: 'Prednisolone, 10-day oral course', role: 'concomitant', details: 'For a chest infection, prescribed by GP; dose not reported', source: 'Hepatology letter (version 2)',
        records: [
          { id: 'PRED-START', kind: 'start', label: 'Course started', timing: d('2026-03-02', '02–11/03/2026') },
          { id: 'PRED-STOP', kind: 'stop', label: 'Course ended', timing: d('2026-03-11', '02–11/03/2026') }
        ] },
      { id: 'EX-AVT', drug: 'Antiviral A, treatment', role: 'treatment of the event', details: 'Oral', source: 'Hepatology letter (version 2)',
        records: [{ id: 'AVT-START', kind: 'start', label: 'Antiviral A restarted', timing: d('2026-05-27', '27/05/2026') }] }
    ],
    links: [{ id: 'L1', exposure_id: 'EX-RTX', event_id: 'EV1', reason: 'Suspect drug' }],
    intervals: [
      { id: 'I1', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'LAST_DOSE_BEFORE_ONSET', anchor_ref: '', end_ref: 'MS-DET',
        reported: { text: 'fourteen months after the last dose', lower: '', upper: '', unit: '' } },
      { id: 'I2', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'LAST_DOSE_BEFORE_ONSET', anchor_ref: '', end_ref: 'onset' },
      { id: 'I3', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'FIRST_ADMIN', anchor_ref: '', end_ref: 'MS-DET' },
      { id: 'I4', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'FIRST_ADMIN', anchor_ref: '', end_ref: 'onset' },
      { id: 'I5', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'CONCOMITANT_CHANGE', anchor_ref: 'AVP-STOP', end_ref: 'MS-DET' },
      { id: 'I6', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'CONCOMITANT_CHANGE', anchor_ref: 'AVP-STOP', end_ref: 'onset' },
      { id: 'I7', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'CONCOMITANT_CHANGE', anchor_ref: 'PRED-START', end_ref: 'MS-DET' },
      { id: 'I8', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'CONCOMITANT_CHANGE', anchor_ref: 'PRED-START', end_ref: 'onset' },
      { id: 'I9', exposure_id: 'EX-RTX', event_id: 'EV1', anchor_type: 'LAST_DOSE_BEFORE_ONSET', anchor_ref: '', end_ref: 'MS-HEP' }
    ],
    recoveries: [],
    technical: [
      { id: 'T1', rule: 'Q8', label: 'Rituximab remaining at the first positive test (435 days)',
        inputs: { t_half_lower: '22', t_half_upper: '52', t: '435', unit: 'days', antibody: true },
        meta: { t_half_lower: { status: 'label fact', source: 'RITUXAN US', section: '12.3 (NHL median)' }, t_half_upper: { status: 'label fact', source: 'RITUXAN US', section: '12.3 (NHL longest)' }, t: { status: 'case measurement', source: 'Calendar days from the last infusion to the first positive test (interval I1)' } },
        gate: { model: true, half_life: true, half_life_type: 'terminal half-life; labeled median and longest value of the range 6.1–52 days', half_life_population: 'non-Hodgkin lymphoma', half_life_source: 'RITUXAN US 12.3; MabThera EU 5.2',
          reference: true, reference_text: 'concentration after the last infusion, 3 March 2025; accumulation over earlier cycles ignored', status: true } },
      { id: 'T2', rule: 'Q8', label: 'Rituximab remaining at the earliest onset bound (343 days)',
        inputs: { t_half_lower: '22', t_half_upper: '52', t: '343', unit: 'days', antibody: true },
        meta: { t_half_lower: { status: 'label fact', source: 'RITUXAN US', section: '12.3 (NHL median)' }, t_half_upper: { status: 'label fact', source: 'RITUXAN US', section: '12.3 (NHL longest)' }, t: { status: 'case measurement', source: 'Minimum of interval I2 (C4)' } },
        gate: { model: true, half_life: true, half_life_type: 'terminal half-life; labeled median and longest value of the range 6.1–52 days', half_life_population: 'non-Hodgkin lymphoma', half_life_source: 'RITUXAN US 12.3; MabThera EU 5.2',
          reference: true, reference_text: 'concentration after the last infusion, 3 March 2025; accumulation over earlier cycles ignored', status: true } }
    ],
    assessment: [
      { link_id: 'L1', fit: 'compatible',
        rationale: 'Example entry from Chapters 20 and 23 (reviewer judgment, not a causality category): exposure persistence is negligible on every labeled half-life; the patient was anti-HBc positive at baseline; B cells and IgG were still low at detection; the timing falls within the period for which the US label reports cases. Regimen X and the March 2026 prednisolone course are contributing exposures that timing cannot separate.',
        prompts: {
          mechanism_hypothesis: { text: 'Loss of immune control during B-cell depletion (hypothesis; Stage 1 HBV profile, not a label statement)', source: 'Stage 1 temporal profile' },
          expected_exposure_course: { text: 'Rituximab detectable for 3–6 months after the last dose', source: 'RITUXAN US 12.3; MabThera EU 5.2' },
          expected_pd_course: { text: 'B-cell recovery begins about 6 months after treatment; median levels normal by 12 months (NHL); up to a median of 23 months in some', source: 'RITUXAN US 12.2; MabThera EU 5.1' },
          clinical_latency_evidence: { text: 'Maximum reported: reactivation up to 24 months after completion of therapy (US); EU gives no timing', source: 'RITUXAN US 5.3; MabThera EU 4.4' },
          observed_course: { text: 'Onset bounded at 343 to 435 calendar days after the last infusion; B cells and IgG low at detection', source: 'Arithmetic display' },
          temporal_fit_rationale: { text: '', source: '' } },
        followup: 'Planned length of prophylaxis and why it stopped; prednisolone dose; any lymphoma relapse or further treatment; risk factors for new infection; course of the hepatitis on treatment.' }
    ]
  };

  return {
    list: [
      { id: 'C01', title: 'C01 (version 2): immediate infusion reaction', chapter: 'Chapter 3, Section 3.7' },
      { id: 'C04', title: 'C04 (version 2): digoxin toxicity after renal deterioration', chapter: 'Chapter 11, Section 11.8' },
      { id: 'C28', title: 'C28 (version 2): hepatitis B reactivation after rituximab', chapter: 'Chapter 20, Section 20.4' }
    ],
    cases: { C01: C01, C04: C04, C28: C28 },
    GATE_OFF: GATE_OFF
  };
}));
