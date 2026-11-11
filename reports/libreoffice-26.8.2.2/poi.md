# poi vs libreoffice@26.8.2.2 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 142. Fails where the case's own cell matches but its spill does not: 6.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2132 | 307 | 12 | 0 | 2 | 87.0% |
| functions | 693 | 590 | 100 | 1 | 0 | 2 | 85.4% |
| matrix | 34 | 34 | 0 | 0 | 0 | 0 | 100.0% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 2758 | 407 | 13 | 0 | 4 | 86.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2758 | 407 | 13 | 0 | 4 | 86.8% |
| **total** | 3182 | 2758 | 407 | 13 | 0 | 4 | 86.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| VLOOKUP | 45 | fail 44, unsupported 1 | `functions/LookupFunctionsTestCaseData.xls!B103` =VLOOKUP("*"&F103&"~?~?",J38:K52,G103,H103) → Excel "#N/A", LibreOffice "something10" |
| DGET | 20 | fail 20 | `functions/DGet.xls!B113` =DGET($G$108:$H$118,"Two",U108:U109) → Excel "#NUM!", LibreOffice "#VALUE!" |
| ACOSH | 14 | fail 14 | `eval/FormulaEvalTestData.xls!E108` =ACOSH(B12) → Excel "#NUM!", LibreOffice "#VALUE!" |
| DATE | 12 | fail 12 | `eval/FormulaEvalTestData.xls!E328` =DATE(1900, 1, 32) → Excel 32, LibreOffice 33 |
| SUMX2PY2 | 12 | fail 12 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", LibreOffice 0 |
| LN | 12 | fail 12 | `eval/FormulaEvalTestData.xls!N880` =LN(K8:N8) → Excel "#NUM!", LibreOffice "#VALUE!" |
| HLOOKUP | 11 | fail 11 | `functions/LookupFunctionsTestCaseData.xls!B47` = HLOOKUP(F47, J$25:P$26, G47,H47) → Excel "#N/A", LibreOffice "something3" |
| SUMX2MY2 | 10 | fail 8, unsupported 2 | `eval/FormulaEvalTestData.xls!M1376` =SUMX2MY2(AA10:AA11,AA11:AA12) → Excel "#REF!", LibreOffice "#NAME?" |
| MODE | 10 | fail 10 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", LibreOffice 2 |
| LOG | 10 | fail 10 | `eval/FormulaEvalTestData.xls!L884` =LOG(F15,F16) → Excel "#NUM!", LibreOffice "#VALUE!" |
| MEDIAN | 10 | fail 10 | `eval/FormulaEvalTestData.xls!J928` =MEDIAN("2",2) → Excel 2, LibreOffice "#VALUE!" |
| ATANH | 8 | fail 8 | `eval/FormulaEvalTestData.xls!E172` =ATANH(C9) → Excel "#NUM!", LibreOffice "#VALUE!" |
| SUMXMY2 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!J1384` =SUMXMY2(E15:E16,C15:C16) → Excel "#DIV/0!", LibreOffice 8.8400400001 |
| YEAR | 8 | fail 6, unsupported 2 | `eval/FormulaEvalTestData.xls!H1540` =YEAR(#REF!) → Excel "#REF!", LibreOffice "#NAME?" |
| LOG10 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!J888` =LOG10(B18) → Excel "#VALUE!", LibreOffice 0 |
| LARGE | 8 | fail 8 | `eval/FormulaEvalTestData.xls!J848` =LARGE(D7:D8,1) → Excel "#NUM!", LibreOffice "#VALUE!" |
| DEVSQ | 6 | fail 6 | `eval/FormulaEvalTestData.xls!E408` =DEVSQ(B7:B9,C7:C9,B15:B17) → Excel 5.837134000085715, LibreOffice 6.3422131112 |
| DAY | 6 | fail 6 | `eval/FormulaEvalTestData.xls!E348` =DAY(1) → Excel 1, LibreOffice 31 |
| CONCATENATE | 6 | fail 6 | `eval/FormulaEvalTestData.xls!E272` =CONCATENATE(B16,B18) → Excel "TRUETRUE", LibreOffice "1TRUE" |
| T | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O1392` =T(J1391:J1393) → Excel "", LibreOffice "non-numeric" |
| MIN | 6 | fail 6 | `eval/FormulaEvalTestData.xls!L940` =MIN(G9,G8,G7,F10,F12,E16,"-20") → Excel -20, LibreOffice "#VALUE!" |
| WEEKNUM | 6 | fail 2, skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", LibreOffice "#VALUE!" |
| MONTH | 6 | fail 6 | `eval/FormulaEvalTestData.xls!G972` =MONTH(-10) → Excel "#NUM!", LibreOffice 12 |
| FACT | 6 | fail 6 | `eval/FormulaEvalTestData.xls!M516` =FACT(B14) → Excel "#NUM!", LibreOffice "#VALUE!" |
| COMBIN | 6 | fail 6 | `eval/FormulaEvalTestData.xls!J264` =COMBIN(-23,12) → Excel "#NUM!", LibreOffice "#VALUE!" |
