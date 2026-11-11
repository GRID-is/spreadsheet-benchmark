# GRID Spreadsheet Benchmark: formula compatibility

Reference: Microsoft Excel for Windows 16.0, build 20527 · v1 · Published 10 November 2026

How closely does a spreadsheet engine match Excel? This repository holds about 102,500 formula test cases taken from eight open-source test suites, Excel's answer to each, every engine's answer, and everything needed to check them. Results and commentary are at [grid.is/benchmark/compatibility](https://grid.is/benchmark/compatibility). This release is tagged `compatibility-v1`: check it out to get exactly the numbers on that page.

GRID makes one of the engines measured here, so don't take our word for it. That's what this repository is for.

## Results

Compatibility is the share of test cases where the engine's result matches Excel's. Each report folder holds the engine's answer to every case and a readable summary per suite, such as [`univer.md`](reports/hyperformula-3.5.0/univer.md), with the functions that fail most.

| | Engine | Version | Tested | Compatibility | Failures per 1,000 | Mean of suites | Report |
|---:|---|---|---|---:|---:|---:|---|
| 1 | GRID | 19.0.0 | 2026-11-09 | 99.88% | 1.2 | 99.77% | [grid-19.0.0](reports/grid-19.0.0) |
| 2 | GridCraft | 0.18.0 | 2026-11-10 | 95.81% | 42 | 93.29% | [gridcraft-0.18.0](reports/gridcraft-0.18.0) |
| 3 | IronCalc | 0.9.1 | 2026-11-09 | 95.22% | 48 | 84.32% | [ironcalc-0.9.1-xlsx](reports/ironcalc-0.9.1-xlsx) |
| 4 | SpreadsheetGear | 9.3.85 | 2026-11-09 | 94.36% | 56 | 91.06% | [spreadsheetgear-9.3.85-xlsx](reports/spreadsheetgear-9.3.85-xlsx) |
| 5 | Mog | 1.0.0 | 2026-11-09 | 91.52% | 85 | 84.57% | [mog-1.0.0-xlsx](reports/mog-1.0.0-xlsx) |
| 6 | SpreadJS | 19.2.5 | 2026-11-09 | 91.27% | 87 | 84.49% | [spread-sheets-19.2.5](reports/spread-sheets-19.2.5) |
| 7 | Univer | 1.4.0 | 2026-11-10 | 91.20% | 88 | 88.65% | [univer-1.4.0](reports/univer-1.4.0) |
| 8 | Google Sheets | Web | 2026-11-09 | 90.62% | 94 | 86.36% | [google-sheets-2026-11-09](reports/google-sheets-2026-11-09) |
| 9 | LibreOffice | 26.8.2.2 | 2026-11-09 | 87.39% | 126 | 79.90% | [libreoffice-26.8.2.2](reports/libreoffice-26.8.2.2) |
| 10 | HyperFormula | 3.5.0 | 2026-11-09 | 82.93% | 171 | 69.63% | [hyperformula-3.5.0](reports/hyperformula-3.5.0) |
| 11 | Gnumeric | 1.12.62 | 2026-11-09 | 82.80% | 172 | 83.07% | [gnumeric-1.12.62](reports/gnumeric-1.12.62) |
| 12 | Apache POI | 6.0.0 | 2026-11-09 | 52.93% | 471 | 59.49% | [poi-6.0.0](reports/poi-6.0.0) |
| 13 | fast-formula-parser | 1.0.19 | 2026-11-09 | 46.98% | 530 | 57.60% | [fast-formula-parser-1.0.19](reports/fast-formula-parser-1.0.19) |

The headline pools every case, so the IronCalc suite, about 71% of the cases, weighs most. The mean of suites weighs each of the eight suites equally instead. [METHOD.md](METHOD.md) says why the headline is pooled.

## Check the scores

Needs Node 24.17 or later and the `bzip2` command, which macOS and Linux have (on Windows, use WSL). No Excel and no account.

```sh
npm ci
npm run score
```

This scores the engine answers stored in `reports/` with the rules in [METHOD.md](METHOD.md) and prints the scores in the table above. `npm run score -- --json` adds each suite's numbers.

## Run an engine yourself

Replay any engine against the stored Excel answers:

```sh
node replay/hyperformula.ts --source reports/grid-19.0.0 --out reports/local/hyperformula
npm run score -- reports/local/hyperformula
```

The npm engines work after `npm ci`: GRID (`replay/spreadsheet-engine.ts`, the public `@grid-is/spreadsheet-engine` package; the published GRID report, `reports/grid-19.0.0`, is the same engine version), HyperFormula, fast-formula-parser, Univer and SpreadJS (which also needs Google Chrome installed: its runner drives the system Chrome). The others need their own install: LibreOffice, Gnumeric, GridCraft, Mog, IronCalc (Python package), Apache POI (Java) and SpreadsheetGear (.NET, with a licence in `SPREADSHEETGEAR_LICENSE`, as its no-cost mode refuses large workbooks). Google Sheets needs your own Google account. The header of each runner says what it expects.

Set `LC_ALL=en_US.UTF-8` and `TZ=UTC` first. Some engines answer differently in other locales and time zones.

## Add an engine

Copy the closest runner in `replay/` and implement `buildFile`. It gets a fixture path and returns the engine's value for each case cell. Before trusting a score, confirm the engine really recalculates and doesn't pass Excel's saved values through. `replay/verify-recalc.ts` does this with a poisoned cached value.

## Reproducing the published numbers

`npm run score` reproduces the leaderboard and the per-suite table on that page from the stored answers exactly. The page's function table is counted from the same case records, though no script here prints it. Rerunning the engines yourself should give the same numbers, with these exceptions:

- **SpreadsheetGear** needs a licence or a trial key. Its no-cost mode refuses large workbooks and scores about half as much.
- **Google Sheets** can't be rerun exactly. Google changes it continuously, so its row is a dated snapshot.
- **Dates without a year**: seven cases depend on the current year (see [METHOD.md](METHOD.md)).
- **Engine versions**: each report records the version it ran. A newer release of an engine is a new result, not a reproduction.

## What's here

| Path | What it is |
|---|---|
| `suites/<suite>/excel/` | The test workbooks, recalculated and saved by Excel. These hold the reference answers |
| `suites/<suite>/import.ts` | How each suite was built from its upstream repository, at a pinned commit |
| `reports/<engine>-<version>/` | One engine's answer to every case, and its timing, run date and environment. In the case records, `apiaryValue` holds the measured engine's value whichever engine it is: Apiary is the GRID engine's internal name, and the field keeps it |
| `replay/` | The runner for each engine |
| `src/` | The harness and scoring. `src/model/classify.ts` decides pass, fail or skip |

Regenerating the Excel files needs desktop Excel and GRID-internal tooling, so `import.ts` and `src/case-workbook.ts` are published to show how the files were made, not to be run. You can open any workbook in your own Excel and compare.

## Found a problem?

Open an issue with the case: suite, file, cell and what you expected. If an engine's score is wrong because of how we run it, we want to fix that, whichever engine it is. The same goes for an engine you would like to see on the board.

## Licence

- Code: [Apache-2.0](LICENSE). This grants no rights to the GRID name.
- Results in `reports/`: [CC BY 4.0](reports/LICENSE). The case records quote formulas and values from the test workbooks, which stay under their upstream licences.
- Test workbooks: each suite keeps its upstream licence (MIT, Apache-2.0, MPL-2.0 or GPL), see [THIRD-PARTY-NOTICES](THIRD-PARTY-NOTICES).

Microsoft and Excel are trademarks of Microsoft Corporation. This project isn't affiliated with or endorsed by Microsoft, or by any engine's maker.
