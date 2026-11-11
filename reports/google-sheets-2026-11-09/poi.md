# poi vs google-sheets@2026-11-09 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 76. Fails where the case's own cell matches but its spill does not: 32.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2252 | 199 | 0 | 0 | 2 | 91.9% |
| functions | 693 | 586 | 105 | 0 | 0 | 2 | 84.8% |
| matrix | 34 | 8 | 26 | 0 | 0 | 0 | 23.5% |
| multisheet | 2 | 0 | 0 | 2 | 0 | 0 | 0.0% |
| **total** | 3182 | 2846 | 330 | 2 | 0 | 4 | 89.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2846 | 330 | 2 | 0 | 4 | 89.6% |
| **total** | 3182 | 2846 | 330 | 2 | 0 | 4 | 89.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 58 | fail 58 | `eval/FormulaEvalTestData.xls!E768` =INDEX(B14:C15,D7,2) → Excel "#VALUE!", Google Sheets "#REF!" |
| SUMX2PY2 | 12 | fail 12 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", Google Sheets 0 |
| VLOOKUP | 11 | fail 11 | `functions/LookupFunctionsTestCaseData.xls!B28` = VLOOKUP(F28, J$2:K$13, G28,H28) → Excel "#N/A", Google Sheets "#VALUE!" |
| TREND | 11 | fail 11 | `functions/Trend.xls!B90` =TREND(,G2:G10) → Excel "#VALUE!", Google Sheets "#REF!" |
| DATE | 10 | fail 10 | `eval/FormulaEvalTestData.xls!E328` =DATE(1900, 1, 32) → Excel 32, Google Sheets 33 |
| DGET | 8 | fail 8 | `functions/DGet.xls!B110` =DGET($G$108:$H$118,"Two",O108:O109) → Excel "#VALUE!", Google Sheets "" |
| MMULT | 8 | fail 8 | `matrix/MatrixFormulaEvalTestData.xlsx!H118` =MMULT(B118:C119,E118:F119) → Excel 5, Google Sheets 5 |
| TRANSPOSE | 7 | fail 7 | `matrix/MatrixFormulaEvalTestData.xlsx!H134` =TRANSPOSE({1,2;3,4}) → Excel 1, Google Sheets 1 |
| T | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O1392` =T(J1391:J1393) → Excel "", Google Sheets "non-numeric" |
| MODE | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", Google Sheets 2 |
| SUMXMY2 | 6 | fail 6 | `eval/FormulaEvalTestData.xls!J1384` =SUMXMY2(E15:E16,C15:C16) → Excel "#DIV/0!", Google Sheets 0 |
| SUMX2MY2 | 6 | fail 6 | `eval/FormulaEvalTestData.xls!H1376` =SUMX2MY2(B8:B10,G8:I8) → Excel 285138.753998, Google Sheets "#N/A" |
| CONCAT | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O268` =CONCAT(K8:L9,K10) → Excel "22.999933.54", Google Sheets "#VALUE!" |
| CONCATENATE | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K272` =CONCATENATE(K8:L9,K10) → Excel "#VALUE!", Google Sheets "22.999933.54" |
| MINVERSE | 6 | fail 6 | `matrix/MatrixFormulaEvalTestData.xlsx!H114` =MINVERSE(B114:C115) → Excel -3, Google Sheets -3 |
| AND | 5 | fail 5 | `eval/FormulaEvalTestData.xls!J136` =AND("1",1) → Excel true, Google Sheets "#VALUE!" |
| ROUND | 5 | fail 5 | `eval/FormulaEvalTestData.xls!L1220` =ROUND(C7,E7) → Excel "#VALUE!", Google Sheets 0 |
| SUM | 5 | fail 3, unsupported 2 | `eval/FormulaEvalTestData.xls!E47` =SUM(H:H H8:I8) → Excel 2.999, Google Sheets "#ERROR!" |
| INDIRECT | 5 | fail 5 | `functions/IndirectFunctionTestCaseData.xls!B15` =INDIRECT("Indirect"&"!"&"H:H",E15) → Excel 0, Google Sheets "#REF!" |
| DAY | 4 | fail 4 | `eval/FormulaEvalTestData.xls!E348` =DAY(1) → Excel 1, Google Sheets 31 |
| ASINH | 4 | fail 4 | `eval/FormulaEvalTestData.xls!K160` =ASINH("") → Excel "#VALUE!", Google Sheets 0 |
| AVEDEV | 4 | fail 4 | `eval/FormulaEvalTestData.xls!F176` =AVEDEV(D7:E9) → Excel "#NUM!", Google Sheets "#DIV/0!" |
| SUBSTITUTE | 4 | fail 4 | `eval/FormulaEvalTestData.xls!V1352` =SUBSTITUTE("ABCBBDBEB", "B", "x", 0) → Excel "#VALUE!", Google Sheets "AxCxxDxEx" |
| WEEKNUM | 4 | skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", Google Sheets "#VALUE!" |
| POWER | 4 | fail 4 | `eval/FormulaEvalTestData.xls!K1104` =POWER(E7,C9) → Excel "#VALUE!", Google Sheets 0 |
