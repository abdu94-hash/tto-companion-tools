# TTO Companion Tools

Companion calculator and Excel workbook for the book **_Time to Onset in Drug Safety: A Technical Guide to
ICSR Analysis Using Mechanism of Action, Pharmacokinetics and Pharmacodynamics_** by Dr. Hafez Selim.

**Version:** TTO Companion Tools v1.0, implementing CALC-SPEC 1.0

**Open the calculator online:** https://abdu94-hash.github.io/tto-companion-tools/

> **Educational use only.** These tools teach the calculations described in the book. They are not
> validated for regulated production use, are not dosing calculators, and do not assign causality. A
> calculated interval or exposure estimate is one input to a medical assessment, never the conclusion.

## What is included

| Tool | Where | Book chapter |
|---|---|---|
| Interactive calculator (runs in your browser; no data leaves your computer) | [`docs/`](docs/) and the link above | Chapter 24 |
| Excel workbook (plain `.xlsx`, formulas only, no macros) | [`workbook/TTO_Companion_Workbook_v1.0.xlsx`](workbook/) | Chapter 23 |
| Verification vectors V01–V27 and an automated test | [`tests/`](tests/) | Chapter 25 |

Both tools implement the book's shared calculation specification, **CALC-SPEC 1.0**:

- **Interval rules C1–C7:** exact elapsed interval, calendar-date difference, treatment-day number,
  bounded interval for partial dates, anchor selection, comparison with reported timing, and recovery interval.
- **Technical calculations Q1–Q11:** fractional persistence, approach to steady state, accumulation, a change in
  clearance, parent and metabolite persistence, a biological replacement demonstration, monoclonal antibody
  persistence, parameter scenarios, dose received before onset during an infusion, cumulative dose per body
  surface area, and estimated renal function.
- **Flags:** `E_` codes block a result; `W_` and `I_` flags are shown with the result and carried into every export.

Each tool includes fictional teaching cases from the book (C01, C04 and C28) and a verification page that
reproduces every numeric verification case.

## Using the calculator

Open https://abdu94-hash.github.io/tto-companion-tools/ in any modern browser. To use it offline, download this
repository (green **Code** button, then **Download ZIP**), unzip it, and open `docs/index.html`.

All calculations run locally in your browser. The calculator makes no network requests and stores nothing on a
server.

## Using the workbook

Download `workbook/TTO_Companion_Workbook_v1.0.xlsx` and open it in Microsoft Excel. Start with the **Read_Me**
sheet. Input cells are yellow and unlocked; formula cells are grey and protected (no password) to prevent accidental
edits. The **Examples** and **Verification** sheets show the worked cases and the verification status. The workbook
was built and tested in Microsoft 365 Excel for Windows; Excel for the web and for Mac have not been tested.

`workbook/build_workbook.ps1` regenerates the workbook through Excel, and `workbook/verify_workbook.ps1`
recalculates it and reports the verification and example results (Windows PowerShell with Excel installed).

## Verifying the tools

The verification cases are in `tests/tto_calc_test_vectors_v1.0.json`. With Node.js installed:

```
node tests/run_vectors.js
```

The calculator's **Verification** tab and the workbook's **Verification** sheet run the same cases. `node tests/run_examples.js`
checks the calculator's worked examples against the book.

## Citation

Selim H. *TTO Companion Tools*, version 1.0 (CALC-SPEC 1.0). Companion to *Time to Onset in Drug Safety*.
https://github.com/abdu94-hash/tto-companion-tools

## License

[CC BY-NC 4.0](LICENSE): free to share and adapt for non-commercial purposes, with attribution. The text of the
book is not covered by this license.

## Reporting a problem

Open an issue on this repository with the inputs, the expected value, the value shown and the tool version. Do
not include patient-identifiable information.
