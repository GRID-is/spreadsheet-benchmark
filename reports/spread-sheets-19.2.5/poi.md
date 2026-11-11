# poi vs spread-sheets@19.2.5 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 186. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 1904 | 547 | 0 | 0 | 2 | 77.7% |
| functions | 693 | 601 | 90 | 0 | 0 | 2 | 87.0% |
| matrix | 34 | 28 | 6 | 0 | 0 | 0 | 82.4% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 2535 | 643 | 0 | 0 | 4 | 79.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2535 | 643 | 0 | 0 | 4 | 79.8% |
| **total** | 3182 | 2535 | 643 | 0 | 0 | 4 | 79.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 25 | fail 25 | `eval/FormulaEvalTestData.xls!E768` =INDEX(B14:C15,D7,2) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| ROUND | 20 | fail 20 | `eval/FormulaEvalTestData.xls!N1220` =ROUND(M9:O9,K10) → Excel 345, SpreadJS "#SPILL!" |
| DOLLAR | 18 | fail 18 | `eval/FormulaEvalTestData.xls!E436` =DOLLAR(B8) → Excel "$1.00", SpreadJS "$1" |
| DGET | 17 | fail 17 | `functions/DGet.xls!B110` =DGET($G$108:$H$118,"Two",O108:O109) → Excel "#VALUE!", SpreadJS 0 |
| ODD | 14 | fail 14 | `eval/FormulaEvalTestData.xls!N1040` =ODD(K8:L9) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| TANH | 12 | fail 12 | `eval/FormulaEvalTestData.xls!N1400` =TANH(I8:K8) → Excel "#VALUE!", SpreadJS 0.9981778976111987 |
| SIGN | 12 | fail 12 | `eval/FormulaEvalTestData.xls!O1288` =SIGN(N9:O10) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| ISNUMBER | 12 | fail 12 | `eval/FormulaEvalTestData.xls!P828` =ISNUMBER(G7:J7) → Excel false, SpreadJS "#SPILL!" |
| TREND | 12 | fail 12 | `functions/Trend.xls!B91` =TREND(F2:G5,F9:I10) → Excel "#REF!", SpreadJS "#N/A" |
| DATE | 10 | fail 10 | `eval/FormulaEvalTestData.xls!E328` =DATE(1900, 1, 32) → Excel 32, SpreadJS 33 |
| UPPER | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O1476` =UPPER(N7:P8) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| SUMX2PY2 | 10 | fail 10 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", SpreadJS 0 |
| LOWER | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O908` =LOWER(N8:O10) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| LN | 10 | fail 10 | `eval/FormulaEvalTestData.xls!P880` =LN(O878:O881) → Excel 1.8112511328273146, SpreadJS "#SPILL!" |
| ISERROR | 10 | fail 10 | `eval/FormulaEvalTestData.xls!L812` =ISERROR(H812) → Excel false, SpreadJS true |
| DEGREES | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O384` =DEGREES(N8:O9) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| VLOOKUP | 10 | fail 10 | `functions/LookupFunctionsTestCaseData.xls!B21` = VLOOKUP(F21, J$2:K$13, G21,H21) → Excel "#REF!", SpreadJS "#VALUE!" |
| ATANH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!N172` =ATANH(M9:O10) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| ATAN | 8 | fail 8 | `eval/FormulaEvalTestData.xls!O164` =ATAN(N162:N165) → Excel 0, SpreadJS "#SPILL!" |
| ASINH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!P160` =ASINH(O9:P11) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| ATAN2 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!L168` =ATAN2(K9:L10,2) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| ABS | 8 | fail 8 | `eval/FormulaEvalTestData.xls!O96` =ABS(N88:N96) → Excel 345, SpreadJS "#SPILL!" |
| SUMSQ | 8 | fail 8 | `eval/FormulaEvalTestData.xls!K1372` =SUMSQ(E13:E16) → Excel 0, SpreadJS 2.210000000001 |
| ROUNDDOWN | 8 | fail 8 | `eval/FormulaEvalTestData.xls!P1224` =ROUNDDOWN(K8:L9,2) → Excel "#VALUE!", SpreadJS "#SPILL!" |
| ROUNDUP | 8 | fail 8 | `eval/FormulaEvalTestData.xls!N1228` =ROUNDUP(K1227:K1229,1) → Excel 3, SpreadJS "#SPILL!" |
