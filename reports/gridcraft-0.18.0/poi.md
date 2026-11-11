# poi vs gridcraft@0.18.0 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 16. Fails where the case's own cell matches but its spill does not: 6.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2201 | 250 | 0 | 0 | 2 | 89.8% |
| functions | 693 | 634 | 57 | 0 | 0 | 2 | 91.8% |
| matrix | 34 | 34 | 0 | 0 | 0 | 0 | 100.0% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 2871 | 307 | 0 | 0 | 4 | 90.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2871 | 307 | 0 | 0 | 4 | 90.3% |
| **total** | 3182 | 2871 | 307 | 0 | 0 | 4 | 90.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ROUND | 20 | fail 20 | `eval/FormulaEvalTestData.xls!N1220` =ROUND(M9:O9,K10) → Excel 345, GridCraft "#VALUE!" |
| ISERROR | 10 | fail 10 | `eval/FormulaEvalTestData.xls!L812` =ISERROR(H812) → Excel false, GridCraft true |
| TREND | 10 | fail 10 | `functions/Trend.xls!B86` =TREND(testArea,G2:G10,H2:H10,FALSE) → Excel 32.013139756562545, GridCraft "#VALUE!" |
| COLUMN | 8 | fail 8 | `eval/FormulaEvalTestData.xls!E256` =COLUMN(1:2) → Excel 1, GridCraft "#VALUE!" |
| TANH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!N1400` =TANH(I8:K8) → Excel "#VALUE!", GridCraft 0.9981778976111987 |
| ROW | 8 | fail 8 | `eval/FormulaEvalTestData.xls!K1232` =ROW(D13:E17) → Excel 13, GridCraft "#VALUE!" |
| INDEX | 8 | fail 8 | `functions/IndexFunctionTestCaseData.xls!B118` =INDEX($I$118:$I$119,,F118) → Excel "a", GridCraft "#VALUE!" |
| VLOOKUP | 8 | fail 8 | `functions/LookupFunctionsTestCaseData.xls!B28` = VLOOKUP(F28, J$2:K$13, G28,H28) → Excel "#N/A", GridCraft "#VALUE!" |
| ATANH | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M172` =ATANH(L171:L173) → Excel 0, GridCraft "#VALUE!" |
| UPPER | 6 | fail 6 | `eval/FormulaEvalTestData.xls!R1476` =UPPER(misc!R1000:R2000) → Excel "", GridCraft "#VALUE!" |
| T | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O1392` =T(J1391:J1393) → Excel "", GridCraft "#VALUE!" |
| MOD | 6 | fail 6 | `eval/FormulaEvalTestData.xls!N964` =MOD(700,M9:O9) → Excel 10, GridCraft "#VALUE!" |
| ODD | 6 | fail 6 | `eval/FormulaEvalTestData.xls!L1040` =ODD(K10:M10) → Excel 5, GridCraft "#VALUE!" |
| LN | 6 | fail 6 | `eval/FormulaEvalTestData.xls!P880` =LN(O878:O881) → Excel 1.8112511328273146, GridCraft "#VALUE!" |
| INDIRECT | 6 | fail 6 | `functions/IndirectFunctionTestCaseData.xls!B15` =INDIRECT("Indirect"&"!"&"H:H",E15) → Excel 0, GridCraft "#VALUE!" |
| ATAN | 4 | fail 4 | `eval/FormulaEvalTestData.xls!O164` =ATAN(N162:N165) → Excel 0, GridCraft "#VALUE!" |
| ASINH | 4 | fail 4 | `eval/FormulaEvalTestData.xls!O160` =ASINH(L158:L161) → Excel 0.7949577687638787, GridCraft "#VALUE!" |
| ASIN | 4 | fail 4 | `eval/FormulaEvalTestData.xls!O156` =ASIN(M154:M157) → Excel 0, GridCraft "#VALUE!" |
| ACOSH | 4 | fail 4 | `eval/FormulaEvalTestData.xls!M108` =ACOSH(K8:N8) → Excel 1.762747174039086, GridCraft "#VALUE!" |
| ATAN2 | 4 | fail 4 | `eval/FormulaEvalTestData.xls!J168` =ATAN2(I167:I169,I11:K11) → Excel 0.0006687785691780711, GridCraft "#VALUE!" |
| ABS | 4 | fail 4 | `eval/FormulaEvalTestData.xls!O96` =ABS(N88:N96) → Excel 345, GridCraft "#VALUE!" |
| ACOS | 4 | fail 4 | `eval/FormulaEvalTestData.xls!O104` =ACOS(L102:L105) → Excel 1.5707963267948966, GridCraft "#VALUE!" |
| TAN | 4 | fail 4 | `eval/FormulaEvalTestData.xls!L1396` =TAN(J1396:J1397) → Excel 74.68593339876537, GridCraft "#VALUE!" |
| SQRT | 4 | fail 4 | `eval/FormulaEvalTestData.xls!O1320` =SQRT(J1319:J1321) → Excel 1, GridCraft "#VALUE!" |
| SIN | 4 | fail 4 | `eval/FormulaEvalTestData.xls!M1292` =SIN(K8:M8) → Excel 0.1411200080598672, GridCraft "#VALUE!" |
