# Method

What the benchmark measures, and what a score does and doesn't mean.

## The reference

Excel is the reference. We took the test cases of eight open-source suites, built them into workbooks, had desktop Excel recalculate and save each one, and use the saved answers as the expected results. The suites' own expected values are kept as `upstreamLiteral` for comparison but aren't scored.

This is why engines don't reach 100% on their own suites: a suite tests what its authors decided an answer should be, and we test whether that matches Excel.

Excel version: Excel for Windows 16.0, build 20527, en-US, UTC.

## The cases

A case is one formula cell, with its spilled cells where it has them. It passes when the engine's value matches Excel's and, for a spilling formula, the spill range and every spilled value match too.

| Suite | Upstream licence | Counted cases |
|---|---|---:|
| IronCalc | MIT or Apache-2.0 | 72,678 |
| LibreOffice | MPL-2.0 | 11,454 |
| Univer | Apache-2.0 | 5,523 |
| HyperFormula | GPL-3.0 | 4,491 |
| Apache POI | Apache-2.0 | 3,178 |
| Gnumeric | GPL-2.0 or GPL-3.0 | 1,960 |
| IronCalc unit tests | MIT or Apache-2.0 | 1,767 |
| fast-formula-parser | MIT | 1,465 |

Counts are for GRID's report. Another engine's count can be a few hundred cases higher: a workbook it cannot load counts against it in full, including the cases that are skipped for every other engine (see below).

The HyperFormula suite is pinned at the 3.1.1 release, while the engine tested is 3.5.0: HyperFormula moved its tests to a private repository in March 2026, so 3.1.1 is the newest public suite.

**IronCalc's share.** Cases are pooled, so the IronCalc suite makes up about 71% of the score, and with the IronCalc unit tests about 73%. Most of that is two workbooks: `ironcalc/MATH_AND_TRIGONOMETRY/ARABIC_ROMAN.xlsx` holds 39% of all counted cases, and with `FLOOR_CEILING.xlsx` the two hold 51%. That lifts IronCalc, which scores highly on tests written for it.

We haven't weighted the headline: every case is a real test of Excel's behaviour, and any weighting would be our own opinion. To show what weighting does, the README also gives the unweighted mean of the eight suite scores. GRID is first either way. Others move: IronCalc is third on the pooled score and eighth on the mean of suites. Per-suite results are in each report and from `npm run score -- --json`, so you can weight them as you like.

## Scoring

An engine's score is its passes divided by the cases that count.

- **Numbers** match when they differ by less than 5×10⁻⁸, or when the difference is less than 5×10⁻⁸ of √(x² + y²): for two close values, a relative difference of about 7×10⁻⁸, so about seven significant digits. `src/compare.ts` has the rule.
- **Everything else** must match exactly. An empty string and an empty cell count as the same. That includes complex numbers, which Excel's functions return as text, so they get no numeric tolerance.
- **Errors in loading**, such as a crash or a timeout, count against the engine for every case in that workbook, including the cases that are otherwise skipped. The time limit is 60 seconds per workbook.

## What is skipped

A skipped case counts for no engine, unless the engine could not load its workbook: a workbook an engine cannot load counts against it in full. `src/model/classify.ts` is the full rule, in order:

| Reason | Why |
|---|---|
| Excel dropped the formula | Excel removed it on load, so there's no answer to compare |
| Read as a user function | The file named a function without the prefix Excel needs, so Excel answered `#NAME?` |
| Not in Excel | The function comes from another spreadsheet, such as `REGEXMATCH`, `ISEMAIL`, `WEEKS` or `HF.ADD`, and Excel doesn't have it |
| Volatile | `NOW`, `TODAY`, `RAND`, `RANDBETWEEN` or `RANDARRAY` anywhere in the formula |
| Depends on a volatile cell | The formula reads such a cell, directly or through other cells |
| Environment | `CELL("filename")` and `INFO` of the system, directory or version |
| Excel unstable | `WEEKNUM` return types that Excel itself answers differently from one recalculation to the next |
| Abandoned array | Arrays so large that Excel gives up and shows 0 |

## Which engines

Every engine on the board is tested at its newest release on the date it ran. Aspose.Cells is not on the board: its licence does not allow publishing benchmark results without Aspose's written consent. We will add it if Aspose agrees.

## How engines run

Each engine reads the same Excel-saved files and recalculates them. There are three designs: engines without an xlsx reader get the cells through an open-source reader, `@borgar/xlsx-convert`, the same for every such engine (it is written by a GRID engineer, and GRID's own engine does not use it here), others open and re-save the file, and the rest are asked for each cell directly. Where an engine reads the file itself, its reader's gaps count in its score. Runners check that the engine really recalculates.

Run in en-US and UTC. Each report records its environment in `manifest.json`.

## What a result means

A result is three things: the version of the test (v1 is this case set, tagged `compatibility-v1`), the date the engine ran, and the engine's version, which was its latest at that date. The engines don't have to run on the same day. A new case set would be v2.

## Known limits

- **It measures functions, not workbooks.** These are mostly small tests of edge cases. A high score doesn't mean every real model calculates correctly, and it says nothing about features.
- **It does not compare speed.** Reports include basic time and memory readings, which are not like for like across engines: some include starting a process or writing a workbook.
- **It's Excel on one platform.** Excel for Mac answers differently on a few hundred cases, mostly `CHAR` and `ENCODEURL`. We use Windows because most serious Excel work is done there.
- **Excel changes.** An old expected value in a suite may be right for an old Excel. Results are for the build above.
- **Dates without a year.** Seven cases, such as `DATEVALUE("5-Jul")`, take the current year, and the saved answers are from 2026. They will fail for correct engines from 1 January 2027 until they're handled.
- **Timeouts cost whole workbooks.** A slow engine can lose every case in a large file. See the `error` column of the reports.
- **GRID's own suite isn't here.** It's excluded because it changes with every GRID release and would favour GRID.
- **GRID has learned from these suites.** We have used them to find and fix bugs in the GRID engine, and its score reflects that work. The suites themselves were all written by others.
