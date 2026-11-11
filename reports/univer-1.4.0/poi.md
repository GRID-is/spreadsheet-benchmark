# poi vs univer@1.4.0 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 175. Fails where the case's own cell matches but its spill does not: 27.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 1968 | 481 | 2 | 0 | 2 | 80.3% |
| functions | 693 | 501 | 106 | 84 | 0 | 2 | 72.5% |
| matrix | 34 | 10 | 24 | 0 | 0 | 0 | 29.4% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 2481 | 611 | 86 | 0 | 4 | 78.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2481 | 611 | 86 | 0 | 4 | 78.1% |
| **total** | 3182 | 2481 | 611 | 86 | 0 | 4 | 78.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 87 | fail 14, unsupported 73 | `eval/FormulaEvalTestData.xls!E768` =INDEX(B14:C15,D7,2) → Excel "#VALUE!", Univer "#SPILL!" |
| ROUND | 23 | fail 23 | `eval/FormulaEvalTestData.xls!N1220` =ROUND(M9:O9,K10) → Excel 345, Univer "#SPILL!" |
| VLOOKUP | 21 | fail 21 | `functions/LookupFunctionsTestCaseData.xls!B28` = VLOOKUP(F28, J$2:K$13, G28,H28) → Excel "#N/A", Univer "#VALUE!" |
| DOLLAR | 16 | fail 16 | `eval/FormulaEvalTestData.xls!E436` =DOLLAR(B8) → Excel "$1.00", Univer "$1.00 " |
| TANH | 14 | fail 14 | `eval/FormulaEvalTestData.xls!N1400` =TANH(I8:K8) → Excel "#VALUE!", Univer 0.998177897611 |
| ROUNDDOWN | 14 | fail 14 | `eval/FormulaEvalTestData.xls!P1224` =ROUNDDOWN(K8:L9,2) → Excel "#VALUE!", Univer "#SPILL!" |
| SUM | 13 | fail 2, unsupported 11 | `eval/FormulaEvalTestData.xls!E75` =SUM(J7:(I8)) → Excel 15, Univer "#VALUE!" |
| MOD | 12 | fail 12 | `eval/FormulaEvalTestData.xls!N964` =MOD(700,M9:O9) → Excel 10, Univer "#VALUE!" |
| LOOKUP | 11 | fail 11 | `functions/LookupFunctionsTestCaseData.xls!B27` = LOOKUP(F27, J$2:S$2,I$2:I$11) → Excel "r3", Univer 30 |
| TREND | 11 | fail 11 | `functions/Trend.xls!B90` =TREND(,G2:G10) → Excel "#VALUE!", Univer "#REF!" |
| COS | 10 | fail 10 | `eval/FormulaEvalTestData.xls!F284` =COS(E15) → Excel 0.5403023058681398, Univer "#VALUE!" |
| ATAN | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O164` =ATAN(N162:N165) → Excel 0, Univer "#SPILL!" |
| ABS | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O96` =ABS(N88:N96) → Excel 345, Univer "#SPILL!" |
| UPPER | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O1476` =UPPER(N7:P8) → Excel "#VALUE!", Univer "#SPILL!" |
| SIN | 10 | fail 10 | `eval/FormulaEvalTestData.xls!P1292` =SIN(E15) → Excel 0.8414709848078965, Univer "#VALUE!" |
| ROUNDUP | 10 | fail 10 | `eval/FormulaEvalTestData.xls!N1228` =ROUNDUP(K1227:K1229,1) → Excel 3, Univer "#SPILL!" |
| ODD | 10 | fail 10 | `eval/FormulaEvalTestData.xls!N1040` =ODD(K8:L9) → Excel "#VALUE!", Univer "#SPILL!" |
| LOWER | 10 | fail 10 | `eval/FormulaEvalTestData.xls!O908` =LOWER(N8:O10) → Excel "#VALUE!", Univer "#SPILL!" |
| ISERROR | 10 | fail 10 | `eval/FormulaEvalTestData.xls!L812` =ISERROR(H812) → Excel false, Univer true |
| ISBLANK | 9 | fail 9 | `eval/FormulaEvalTestData.xls!K804` =ISBLANK(K13:L14) → Excel false, Univer "#SPILL!" |
| DGET | 9 | fail 9 | `functions/DGet.xls!B124` =DGET($G$122:$I$124,"1",$K$122:$K$123) → Excel "x", Univer "a" |
| ATANH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!N172` =ATANH(M9:O10) → Excel "#VALUE!", Univer "#SPILL!" |
| ASINH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!P160` =ASINH(O9:P11) → Excel "#VALUE!", Univer "#SPILL!" |
| ACOSH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!O108` =ACOSH(N9:O10) → Excel "#VALUE!", Univer "#SPILL!" |
| ATAN2 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!L168` =ATAN2(K9:L10,2) → Excel "#VALUE!", Univer "#SPILL!" |
