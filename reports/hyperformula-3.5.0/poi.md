# poi vs hyperformula@3.5.0 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 26. Fails where the case's own cell matches but its spill does not: 20.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2043 | 249 | 159 | 0 | 2 | 83.4% |
| functions | 693 | 389 | 110 | 192 | 0 | 2 | 56.3% |
| matrix | 34 | 4 | 18 | 12 | 0 | 0 | 11.8% |
| multisheet | 2 | 0 | 1 | 1 | 0 | 0 | 0.0% |
| **total** | 3182 | 2436 | 378 | 364 | 0 | 4 | 76.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2436 | 378 | 364 | 0 | 4 | 76.7% |
| **total** | 3182 | 2436 | 378 | 364 | 0 | 4 | 76.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 81 | fail 8, unsupported 73 | `functions/IndexFunctionTestCaseData.xls!B114` =INDEX($I$114:$J$114,,F114) → Excel "b", HyperFormula "#VALUE!" |
| ROUND | 36 | fail 17, unsupported 19 | `eval/FormulaEvalTestData.xls!L1220` =ROUND(C7,E7) → Excel "#VALUE!", HyperFormula 0 |
| VLOOKUP | 36 | fail 31, unsupported 5 | `eval/FormulaEvalTestData.xls!G1512` =VLOOKUP(-1,$B$6:$H$12,4, TRUE) → Excel "#N/A", HyperFormula "#NAME?" |
| SUM | 29 | fail 17, unsupported 12 | `eval/FormulaEvalTestData.xls!N1044` =SUM(OFFSET(K7:L8, 0,0,)) → Excel 7.0009, HyperFormula "#ERROR!" |
| LOOKUP | 25 | unsupported 25 | `eval/FormulaEvalTestData.xls!E904` =LOOKUP("DOUBLE",C6:G6,C9:G9) → Excel 1.1, HyperFormula "#NAME?" |
| DOLLAR | 22 | unsupported 22 | `eval/FormulaEvalTestData.xls!E436` =DOLLAR(B8) → Excel "$1.00", HyperFormula "#NAME?" |
| DGET | 22 | fail 20, unsupported 2 | `functions/DGet.xls!B110` =DGET($G$108:$H$118,"Two",O108:O109) → Excel "#VALUE!", HyperFormula 0 |
| MODE | 20 | unsupported 20 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", HyperFormula "#NAME?" |
| FIXED | 20 | unsupported 20 | `functions/57003-FixedFunctionTestCaseData.xls!B18` =FIXED(1,2,"invalid") → Excel "#VALUE!", HyperFormula "#NAME?" |
| HLOOKUP | 18 | fail 10, unsupported 8 | `eval/FormulaEvalTestData.xls!G740` =HLOOKUP("STRING",$B$6:$H$12,6, TRUE) → Excel " beginsWithSpace", HyperFormula "#NAME?" |
| INDIRECT | 17 | unsupported 17 | `eval/FormulaEvalTestData.xls!F772` =INDIRECT(D10) → Excel "#REF!", HyperFormula "#NAME?" |
| ERROR.TYPE | 16 | unsupported 16 | `eval/FormulaEvalTestData.xls!F476` =ERROR.TYPE(AA9) → Excel 3, HyperFormula "#NAME?" |
| AND | 15 | fail 7, unsupported 8 | `eval/FormulaEvalTestData.xls!F136` =AND(E7) → Excel "#VALUE!", HyperFormula false |
| TREND | 15 | unsupported 15 | `functions/Trend.xls!B88` =TREND(F13:F16) → Excel "#VALUE!", HyperFormula "#NAME?" |
| MATCH | 14 | fail 14 | `functions/MatchFunctionTestCaseData.xls!B56` =MATCH(E56,H$3:H$8,F56) → Excel 3, HyperFormula "#N/A" |
| OR | 13 | fail 5, unsupported 8 | `eval/FormulaEvalTestData.xls!I1056` =OR(H13:J13) → Excel "#VALUE!", HyperFormula false |
| CONCAT | 12 | fail 2, unsupported 10 | `eval/FormulaEvalTestData.xls!D268` =CONCAT(B7,B15) → Excel "0-1", HyperFormula "#NAME?" |
| DATE | 10 | fail 10 | `eval/FormulaEvalTestData.xls!E328` =DATE(1900, 1, 32) → Excel 32, HyperFormula 33 |
| ISREF | 8 | fail 8 | `eval/FormulaEvalTestData.xls!J836` =ISREF(D7) → Excel true, HyperFormula false |
| ISBLANK | 7 | fail 7 | `eval/FormulaEvalTestData.xls!K804` =ISBLANK(K13:L14) → Excel false, HyperFormula "#VALUE!" |
| TRANSPOSE | 7 | fail 7 | `matrix/MatrixFormulaEvalTestData.xlsx!H134` =TRANSPOSE({1,2;3,4}) → Excel 1, HyperFormula 1 |
| AVEDEV | 6 | fail 4, unsupported 2 | `eval/FormulaEvalTestData.xls!L176` =AVEDEV(TRUE, FALSE) → Excel 0.5, HyperFormula "#NAME?" |
| T | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O1392` =T(J1391:J1393) → Excel "", HyperFormula "non-numeric" |
| SUMX2PY2 | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", HyperFormula 0 |
| PRODUCT | 6 | fail 6 | `eval/FormulaEvalTestData.xls!L1120` =PRODUCT(D7:D10) → Excel 0, HyperFormula 1 |
