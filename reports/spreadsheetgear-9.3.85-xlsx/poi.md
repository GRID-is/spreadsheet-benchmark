# poi vs spreadsheetgear@9.3.85 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 14. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2429 | 10 | 12 | 0 | 2 | 99.1% |
| functions | 693 | 665 | 25 | 1 | 0 | 2 | 96.2% |
| matrix | 34 | 34 | 0 | 0 | 0 | 0 | 100.0% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 3130 | 35 | 13 | 0 | 4 | 98.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 3130 | 35 | 13 | 0 | 4 | 98.5% |
| **total** | 3182 | 3130 | 35 | 13 | 0 | 4 | 98.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| CONCAT | 12 | unsupported 12 | `eval/FormulaEvalTestData.xls!D268` =CONCAT(B7,B15) → Excel "0-1", SpreadsheetGear "#NAME?" |
| VLOOKUP | 7 | fail 7 | `functions/LookupFunctionsTestCaseData.xls!B21` = VLOOKUP(F21, J$2:K$13, G21,H21) → Excel "#REF!", SpreadsheetGear "#VALUE!" |
| ROUND | 7 | fail 7 | `functions/Trend.xls!B102` =ROUND(I55,6) → Excel 23.119431, SpreadsheetGear 16.152719 |
| MODE | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", SpreadsheetGear 2 |
| DGET | 5 | fail 4, unsupported 1 | `functions/DGet.xls!B104` =DGET($G$94:$N$95,"Three",AH94:AI95) → Excel "h", SpreadsheetGear "#VALUE!" |
| WEEKNUM | 4 | skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", SpreadsheetGear "#VALUE!" |
| INDEX | 4 | fail 4 | `functions/IndexFunctionTestCaseData.xls!B18` =INDEX(threeByThree,E18) → Excel "#REF!", SpreadsheetGear "#VALUE!" |
| SUMPRODUCT | 2 | fail 2 | `eval/FormulaEvalTestData.xls!J1368` =SUMPRODUCT(G1370,I1370) → Excel 0, SpreadsheetGear "#VALUE!" |
| LOG | 2 | fail 2 | `eval/FormulaEvalTestData.xls!K884` =LOG(F17,E15) → Excel "#NUM!", SpreadsheetGear "#DIV/0!" |
| TREND | 2 | fail 2 | `functions/Trend.xls!B84` =TREND(F2:I5,F6:I10) → Excel "#REF!", SpreadsheetGear 12.000000000000043 |
| HLOOKUP | 1 | fail 1 | `functions/LookupFunctionsTestCaseData.xls!B19` = HLOOKUP(F19, J$2:S$3, G19,H19) → Excel "#REF!", SpreadsheetGear "#VALUE!" |
