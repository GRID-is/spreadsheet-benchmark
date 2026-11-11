# poi vs mog@1.0.0 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 40. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2055 | 394 | 2 | 0 | 2 | 83.8% |
| functions | 693 | 590 | 101 | 0 | 0 | 2 | 85.4% |
| matrix | 34 | 27 | 7 | 0 | 0 | 0 | 79.4% |
| multisheet | 2 | 1 | 1 | 0 | 0 | 0 | 50.0% |
| **total** | 3182 | 2673 | 503 | 2 | 0 | 4 | 84.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2673 | 503 | 2 | 0 | 4 | 84.1% |
| **total** | 3182 | 2673 | 503 | 2 | 0 | 4 | 84.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| VLOOKUP | 36 | fail 36 | `functions/LookupFunctionsTestCaseData.xls!B103` =VLOOKUP("*"&F103&"~?~?",J38:K52,G103,H103) → Excel "#N/A", Mog "something12" |
| HLOOKUP | 17 | fail 17 | `functions/LookupFunctionsTestCaseData.xls!B63` = HLOOKUP("*"&F63&"*", J$25:P$26, G63,H63) → Excel "something6", Mog "#N/A" |
| DGET | 13 | fail 13 | `functions/DGet.xls!B110` =DGET($G$108:$H$118,"Two",O108:O109) → Excel "#VALUE!", Mog 0 |
| TANH | 12 | fail 12 | `eval/FormulaEvalTestData.xls!N1400` =TANH(I8:K8) → Excel "#VALUE!", Mog 0.9981778976111987 |
| SUMX2PY2 | 12 | fail 12 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", Mog "#VALUE!" |
| TREND | 12 | fail 12 | `functions/Trend.xls!B88` =TREND(F13:F16) → Excel "#VALUE!", Mog 62.5 |
| ROUND | 11 | fail 11 | `eval/FormulaEvalTestData.xls!N1220` =ROUND(M9:O9,K10) → Excel 345, Mog "#VALUE!" |
| SUMX2MY2 | 10 | fail 10 | `eval/FormulaEvalTestData.xls!J1376` =SUMX2MY2(H8:H10,G7:G8) → Excel "#N/A", Mog "#VALUE!" |
| LOWER | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O908` =LOWER(N8:O10) → Excel "#VALUE!", Mog "" |
| ATAN2 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!L168` =ATAN2(K9:L10,2) → Excel "#VALUE!", Mog 0.5880026035475675 |
| ABS | 8 | fail 8 | `eval/FormulaEvalTestData.xls!O96` =ABS(N88:N96) → Excel 345, Mog "#VALUE!" |
| UPPER | 8 | fail 8 | `eval/FormulaEvalTestData.xls!O1476` =UPPER(N7:P8) → Excel "#VALUE!", Mog "BLANKSINCL" |
| ROUNDDOWN | 8 | fail 8 | `eval/FormulaEvalTestData.xls!P1224` =ROUNDDOWN(K8:L9,2) → Excel "#VALUE!", Mog 2 |
| ROUNDUP | 8 | fail 8 | `eval/FormulaEvalTestData.xls!N1228` =ROUNDUP(K1227:K1229,1) → Excel 3, Mog 0 |
| MOD | 8 | fail 8 | `eval/FormulaEvalTestData.xls!N964` =MOD(700,M9:O9) → Excel 10, Mog "#VALUE!" |
| SUMXMY2 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!J1384` =SUMXMY2(E15:E16,C15:C16) → Excel "#DIV/0!", Mog 8.8400400001 |
| COS | 8 | fail 8 | `eval/FormulaEvalTestData.xls!O284` =COS(K8:K10) → Excel "#VALUE!", Mog -0.4161468365471424 |
| ACOSH | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O108` =ACOSH(N9:O10) → Excel "#VALUE!", Mog 6.536689497185612 |
| SIN | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M1292` =SIN(K8:M8) → Excel 0.1411200080598672, Mog 0.9092974268256817 |
| SINH | 6 | fail 6 | `eval/FormulaEvalTestData.xls!N1296` =SINH(L1295:L1297) → Excel 1.4650188248182272, Mog 0 |
| MIN | 6 | fail 6 | `eval/FormulaEvalTestData.xls!L940` =MIN(G9,G8,G7,F10,F12,E16,"-20") → Excel -20, Mog 0.00001 |
| LN | 6 | fail 6 | `eval/FormulaEvalTestData.xls!P880` =LN(O878:O881) → Excel 1.8112511328273146, Mog "#NUM!" |
| INT | 6 | fail 6 | `eval/FormulaEvalTestData.xls!N788` =INT(M786:M789) → Excel 3, Mog 0 |
| DOLLAR | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M436` =DOLLAR(K9:L10) → Excel "#VALUE!", Mog "$3.00" |
| FLOOR | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M564` =FLOOR(G7:H8,H10) → Excel "#VALUE!", Mog 0 |
