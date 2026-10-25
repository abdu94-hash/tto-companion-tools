# TTO Companion Tools

Companion calculator and Excel workbook for the book **_Time to Onset in Drug Safety: A Technical Guide to
ICSR Analysis Using Mechanism of Action, Pharmacokinetics and Pharmacodynamics_** by Dr. Hafez Selim.

**Version:** TTO Companion Tools v1.1, implementing CALC-SPEC 1.1 (25 October 2026)

**Open the calculator online:** https://abdu94-hash.github.io/tto-companion-tools/

> **Educational use only.** These tools teach the calculations described in the book. They are not
> validated for regulated production use, are not dosing calculators, and do not assign causality. A
> calculated interval or exposure estimate is one input to a medical assessment, never the conclusion.

## What is included

| Tool | Where | Book chapter |
|---|---|---|
| Interactive calculator (runs in your browser; no data leaves your computer) | [`docs/`](docs/) and the link above | Chapter 24 |
| Excel workbook (plain `.xlsx`, formulas only, no macros) | [`workbook/TTO_Companion_Workbook_v1.1.xlsx`](workbook/) | Chapter 23 |
| Verification vectors V01–V38 and automated tests | [`tests/`](tests/) | Chapter 25 |

Both tools implement the book's shared calculation specification, **CALC-SPEC 1.1**:

- **Interval rules C1–C7:** exact elapsed interval, calendar-date difference, treatment-day number,
  bounded interval for partial dates (including a laboratory-defined onset), anchor selection, comparison with
  reported timing, and recovery interval.
- **Technical calculations Q1–Q11:** fractional persistence, approach to steady state, accumulation, a change in
  clearance, parent and metabolite persistence, a biological replacement demonstration (also used for the approach
  to a new level), monoclonal antibody persistence, parameter scenarios, dose received before onset during an
  infusion, cumulative dose per recorded body surface area, and estimated renal function.
- **Flags:** `E_` codes block a result; `W_` and `I_` flags are shown with the result and carried into every export.
  Every error that applies is listed, never only the first.

Each tool includes fictional teaching cases from the book (C01, C04 and C28) and a verification page that
reproduces every numeric verification case.

## What changed in version 1.1

CALC-SPEC 1.1 is a focused revision. No formula changed, and every version 1.0 verification result is reproduced
unchanged (V01–V27).

- **Laboratory-defined onset (C4):** an onset known only from tests is bounded below by the date of the last normal
  test and above by the date of the first abnormal test (V31). Case C28 uses it.
- **C7 recovery interval** now has verification vectors: V28 (minute level, as in case C01), V29 (date level, as in
  case C04) and V30, a labeled no-flag variant.
- **Q6** has a second named use, the approach to a new level (V32).
- **BSA** is never calculated; the BSA recorded at each administration is entered.
- **One agreed rule wherever the calculator and the workbook differed**, choosing the stricter behavior: block rather
  than guess, list every error, compare in UTC when offsets are known, and never infer an unrecorded intervention
  (V33–V38). In the calculator this changes: last-dose selection (UTC order; a partial-date or uncertain
  administration blocks automatic selection), C7 counts only intervention milestones, an incomplete gate lists both
  E_MISSING_INPUT and E_APPLICABILITY_UNCONFIRMED, flags are listed in a fixed order, Q10 Method B uses the first and
  last administrations by date, and very small percentages are shown in powers of ten.
- **Future work** (not in 1.1): END_OF_ADMINISTRATION, INTERVENTION and interruption or restart anchor types; a
  dose-reduction input; comparison with label windows stated in weeks.

The calculator's **About** tab lists every rule in full.

## Using the calculator

Open https://abdu94-hash.github.io/tto-companion-tools/ in any modern browser. To use it offline, download this
repository (green **Code** button, then **Download ZIP**), unzip it, and open `docs/index.html`.

All calculations run locally in your browser. The calculator makes no network requests and stores nothing on a
server.

## Using the workbook

Download `workbook/TTO_Companion_Workbook_v1.1.xlsx` and open it in Microsoft Excel. Start with the **Read_Me**
sheet. Input cells are yellow and unlocked; formula cells are grey and protected (no password) to prevent accidental
edits. The **Examples** and **Verification** sheets show the worked cases and the verification status. The workbook
was built and tested in Microsoft 365 Excel for Windows; Excel for the web and for Mac have not been tested.

`workbook/build_workbook.ps1` regenerates the workbook through Excel, and `workbook/verify_workbook.ps1`
recalculates it and reports the verification and example results (Windows PowerShell with Excel installed).

## Verifying the tools

The verification cases are in `tests/tto_calc_test_vectors_v1.1.json` (V01–V27 are the unchanged CALC-SPEC 1.0
cases, also kept as `tests/tto_calc_test_vectors_v1.0.json` for the record). With Node.js installed:

```
node tests/run_vectors.js
node tests/run_examples.js
```

`run_vectors.js` runs every vector, checks that a planted wrong value is caught, and, when the independent Python
reference check (`tto_calc_reference_check_v1.1.py`) is available, confirms that the calculator and the reference
agree value by value (`--python <path>`). `run_examples.js` checks the calculator's worked examples against the
book. The calculator's **Verification** tab and the workbook's **Verification** sheet run the same cases.

## Citation

Selim H. *TTO Companion Tools*, version 1.1 (CALC-SPEC 1.1). Companion to *Time to Onset in Drug Safety*.
https://github.com/abdu94-hash/tto-companion-tools

## License

[CC BY-NC 4.0](LICENSE): free to share and adapt for non-commercial purposes, with attribution. The text of the
book is not covered by this license.

## Reporting a problem

Open an issue on this repository with the inputs, the expected value, the value shown and the tool version. Do
not include patient-identifiable information.
