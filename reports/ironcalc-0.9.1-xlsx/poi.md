# poi vs ironcalc@0.9.1 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 40. Fails where the case's own cell matches but its spill does not: 20.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2322 | 129 | 0 | 0 | 2 | 94.7% |
| functions | 693 | 610 | 81 | 0 | 0 | 2 | 88.3% |
| matrix | 34 | 15 | 19 | 0 | 0 | 0 | 44.1% |
| multisheet | 2 | 0 | 2 | 0 | 0 | 0 | 0.0% |
| **total** | 3182 | 2947 | 231 | 0 | 0 | 4 | 92.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2947 | 231 | 0 | 0 | 4 | 92.7% |
| **total** | 3182 | 2947 | 231 | 0 | 0 | 4 | 92.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| DGET | 17 | fail 17 | `functions/DGet.xls!B110` =DGET($G$108:$H$118,"Two",O108:O109) → Excel "#VALUE!", IronCalc 0 |
| SUM | 16 | fail 16 | `eval/FormulaEvalTestData.xls!N1044` =SUM(OFFSET(K7:L8, 0,0,)) → Excel 7.0009, IronCalc "#REF!" |
| VLOOKUP | 15 | fail 15 | `functions/LookupFunctionsTestCaseData.xls!B28` = VLOOKUP(F28, J$2:K$13, G28,H28) → Excel "#N/A", IronCalc "#VALUE!" |
| INDEX | 14 | fail 14 | `functions/IndexFunctionTestCaseData.xls!B95` =INDEX(oneByOne,E95) → Excel "m", IronCalc "#VALUE!" |
| SUMX2PY2 | 12 | fail 12 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", IronCalc 0 |
| SUMX2MY2 | 12 | fail 12 | `eval/FormulaEvalTestData.xls!K1376` =SUMX2MY2(G8,H8) → Excel -7.992000000000001, IronCalc "#VALUE!" |
| DATE | 10 | fail 10 | `eval/FormulaEvalTestData.xls!E328` =DATE(1900, 1, 32) → Excel 32, IronCalc 33 |
| SUMXMY2 | 10 | fail 10 | `eval/FormulaEvalTestData.xls!J1384` =SUMXMY2(E15:E16,C15:C16) → Excel "#DIV/0!", IronCalc 2.2100200001 |
| ROUND | 10 | fail 10 | `functions/Trend.xls!B102` =ROUND(I55,6) → Excel 23.119431, IronCalc "#REF!" |
| TREND | 7 | fail 7 | `functions/Trend.xls!B88` =TREND(F13:F16) → Excel "#VALUE!", IronCalc 64.79999999999998 |
| AVEDEV | 6 | fail 6 | `eval/FormulaEvalTestData.xls!L176` =AVEDEV(TRUE, FALSE) → Excel 0.5, IronCalc "#DIV/0!" |
| T | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O1392` =T(J1391:J1393) → Excel "", IronCalc "non-numeric" |
| RIGHT | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K1208` =RIGHT("a string", TRUE) → Excel "g", IronCalc "#VALUE!" |
| LEFT | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K856` =LEFT("a string", TRUE) → Excel "a", IronCalc "#VALUE!" |
| TRANSPOSE | 5 | fail 5 | `matrix/MatrixFormulaEvalTestData.xlsx!H134` =TRANSPOSE({1,2;3,4}) → Excel 1, IronCalc 1 |
| DAY | 4 | fail 4 | `eval/FormulaEvalTestData.xls!E348` =DAY(1) → Excel 1, IronCalc 31 |
| SUMSQ | 4 | fail 4 | `eval/FormulaEvalTestData.xls!M1372` =SUMSQ(TRUE, TRUE, FALSE) → Excel 2, IronCalc 0 |
| VAR | 4 | fail 4 | `eval/FormulaEvalTestData.xls!K1488` =VAR(TRUE, FALSE, TRUE) → Excel 0.33333333333333337, IronCalc "#DIV/0!" |
| VARP | 4 | fail 4 | `eval/FormulaEvalTestData.xls!K1496` =VARP(TRUE, FALSE, TRUE) → Excel 0.2222222222222222, IronCalc "#DIV/0!" |
| MODE | 4 | fail 4 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", IronCalc 2 |
| WEEKNUM | 4 | skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", IronCalc "#VALUE!" |
| STDEV | 4 | fail 4 | `eval/FormulaEvalTestData.xls!K1328` =STDEV(TRUE, FALSE, TRUE) → Excel 0.5773502691896258, IronCalc "#DIV/0!" |
| DEVSQ | 4 | fail 4 | `eval/FormulaEvalTestData.xls!I408` =DEVSQ("1","2",2) → Excel 0.6666666666666666, IronCalc 0 |
| YEAR | 4 | fail 4 | `eval/FormulaEvalTestData.xls!F1540` =YEAR(0) → Excel 1900, IronCalc "#NUM!" |
| MONTH | 4 | fail 4 | `eval/FormulaEvalTestData.xls!F972` =MONTH(0) → Excel 1, IronCalc "#NUM!" |
