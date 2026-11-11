# poi vs fast-formula-parser@1.0.19 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 73. Fails where the case's own cell matches but its spill does not: 24.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 1601 | 443 | 407 | 0 | 2 | 65.3% |
| functions | 693 | 265 | 164 | 262 | 0 | 2 | 38.4% |
| matrix | 34 | 4 | 24 | 6 | 0 | 0 | 11.8% |
| multisheet | 2 | 0 | 0 | 0 | 2 | 0 | 0.0% |
| **total** | 3182 | 1870 | 631 | 675 | 2 | 4 | 58.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 1870 | 631 | 675 | 2 | 4 | 58.8% |
| **total** | 3182 | 1870 | 631 | 675 | 2 | 4 | 58.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 89 | fail 16, unsupported 73 | `eval/FormulaEvalTestData.xls!E768` =INDEX(B14:C15,D7,2) → Excel "#VALUE!", fast-formula-parser "#ERROR!" |
| DGET | 86 | unsupported 86 | `functions/DGet.xls!B123` =DGET($G$122:$I$124,4,$K$122:$K$123) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| SUBSTITUTE | 42 | unsupported 42 | `eval/FormulaEvalTestData.xls!W1352` =SUBSTITUTE("ABCBBDBEB", "B", "x", -1) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| VLOOKUP | 40 | fail 40 | `functions/LookupFunctionsTestCaseData.xls!B86` =VLOOKUP("3~*"&F86,J38:K50,G86,H86) → Excel "something11", fast-formula-parser "#N/A" |
| SUM | 37 | fail 20, unsupported 15, error 2 | `eval/FormulaEvalTestData.xls!N1044` =SUM(OFFSET(K7:L8, 0,0,)) → Excel 7.0009, fast-formula-parser "#NAME?" |
| ROUND | 36 | fail 19, unsupported 17 | `eval/FormulaEvalTestData.xls!N1220` =ROUND(M9:O9,K10) → Excel 345, fast-formula-parser "#VALUE!" |
| ROMAN | 31 | fail 31 | `functions/RomanFunctionTestCaseData.xls!B29` =ROMAN(49,1) → Excel "VLIV", fast-formula-parser "#ERROR!" |
| UPPER | 30 | unsupported 30 | `eval/FormulaEvalTestData.xls!O1476` =UPPER(N7:P8) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| LOOKUP | 25 | unsupported 25 | `eval/FormulaEvalTestData.xls!E904` =LOOKUP("DOUBLE",C6:G6,C9:G9) → Excel 1.1, fast-formula-parser "#NAME?" |
| MIN | 22 | unsupported 22 | `eval/FormulaEvalTestData.xls!M940` =MIN(E10:E12) → Excel 0, fast-formula-parser "#NAME?" |
| LARGE | 21 | unsupported 21 | `eval/FormulaEvalTestData.xls!M848` =LARGE(F10:I13,3) → Excel 345, fast-formula-parser "#NAME?" |
| MODE | 20 | unsupported 20 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| MAX | 20 | unsupported 20 | `eval/FormulaEvalTestData.xls!K916` =MAX(F11,E11,E13:E15,H9:I11,L8) → Excel 345, fast-formula-parser "#NAME?" |
| MATCH | 20 | unsupported 20 | `eval/FormulaEvalTestData.xls!F912` =MATCH(-5, B7:B10) → Excel "#N/A", fast-formula-parser "#NAME?" |
| VARP | 18 | unsupported 18 | `eval/FormulaEvalTestData.xls!L1496` =VARP(1,2,3,E28) → Excel 0.6666666666666666, fast-formula-parser "#NAME?" |
| VAR | 18 | unsupported 18 | `eval/FormulaEvalTestData.xls!L1488` =VAR(1,2,3,E20) → Excel 1, fast-formula-parser "#NAME?" |
| STDEV | 18 | unsupported 18 | `eval/FormulaEvalTestData.xls!L1328` =STDEV(1,2,3,E20) → Excel 1, fast-formula-parser "#NAME?" |
| MEDIAN | 18 | unsupported 18 | `eval/FormulaEvalTestData.xls!L928` =MEDIAN(E9:G10) → Excel 2.3456, fast-formula-parser "#NAME?" |
| SMALL | 17 | unsupported 17 | `eval/FormulaEvalTestData.xls!K1312` =SMALL(C7:C10,100) → Excel "#NUM!", fast-formula-parser "#NAME?" |
| INDIRECT | 17 | unsupported 17 | `eval/FormulaEvalTestData.xls!F772` =INDIRECT(D10) → Excel "#REF!", fast-formula-parser "#NAME?" |
| TREND | 15 | unsupported 15 | `functions/Trend.xls!B88` =TREND(F13:F16) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| MINA | 14 | unsupported 14 | `eval/FormulaEvalTestData.xls!J944` =MINA(F12:F14,F10) → Excel 0, fast-formula-parser "#NAME?" |
| MAXA | 14 | unsupported 14 | `eval/FormulaEvalTestData.xls!I920` =MAXA(I8,G8:G9,F12) → Excel 3.5, fast-formula-parser "#NAME?" |
| NPER | 12 | unsupported 12 | `eval/FormulaEvalTestData.xls!I1024` =NPER(12,4500,100000,100000) → Excel "#NUM!", fast-formula-parser "#NAME?" |
| LN | 12 | fail 12 | `eval/FormulaEvalTestData.xls!P880` =LN(O878:O881) → Excel 1.8112511328273146, fast-formula-parser "#ERROR!" |

## Errors

- `multisheet/FormulaSheetRange.xlsx`: timeout after 60000 ms
- `multisheet/FormulaSheetRange.xlsx`: timeout after 60000 ms
