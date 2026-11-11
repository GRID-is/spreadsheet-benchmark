# poi vs gnumeric@1.12.62 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 65. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2327 | 122 | 2 | 0 | 2 | 94.9% |
| functions | 693 | 606 | 85 | 0 | 0 | 2 | 87.7% |
| matrix | 34 | 28 | 6 | 0 | 0 | 0 | 82.4% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 2963 | 213 | 2 | 0 | 4 | 93.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 2963 | 213 | 2 | 0 | 4 | 93.2% |
| **total** | 3182 | 2963 | 213 | 2 | 0 | 4 | 93.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 30 | fail 30 | `functions/IndexFunctionTestCaseData.xls!B88` =INDEX(oneByOne,E88,F88) → Excel "#VALUE!", Gnumeric "#REF!" |
| VLOOKUP | 22 | fail 22 | `functions/LookupFunctionsTestCaseData.xls!B103` =VLOOKUP("*"&F103&"~?~?",J38:K52,G103,H103) → Excel "#N/A", Gnumeric "something13" |
| DGET | 15 | fail 15 | `functions/DGet.xls!B125` =DGET($G$122:$I$124,"2",$K$122:$K$123) → Excel "#VALUE!", Gnumeric "#NUM!" |
| FACT | 12 | fail 12 | `eval/FormulaEvalTestData.xls!N516` =FACT(E17) → Excel 2, Gnumeric 5.999924633498387 |
| SUMXMY2 | 8 | fail 8 | `eval/FormulaEvalTestData.xls!J1384` =SUMXMY2(E15:E16,C15:C16) → Excel "#DIV/0!", Gnumeric "#VALUE!" |
| T | 6 | fail 6 | `eval/FormulaEvalTestData.xls!O1392` =T(J1391:J1393) → Excel "", Gnumeric "non-numeric" |
| SUMX2PY2 | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K1380` =SUMX2PY2(E7:E10,D7:D10) → Excel "#DIV/0!", Gnumeric "#VALUE!" |
| WEEKNUM | 6 | fail 2, skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", Gnumeric "#VALUE!" |
| SUMX2MY2 | 6 | fail 6 | `eval/FormulaEvalTestData.xls!J1376` =SUMX2MY2(H8:H10,G7:G8) → Excel "#N/A", Gnumeric "#VALUE!" |
| CONCATENATE | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K272` =CONCATENATE(K8:L9,K10) → Excel "#VALUE!", Gnumeric "22.999933.54" |
| HLOOKUP | 5 | fail 5 | `eval/FormulaEvalTestData.xls!G740` =HLOOKUP("STRING",$B$6:$H$12,6, TRUE) → Excel " beginsWithSpace", Gnumeric "blanksIncl" |
| YEAR | 4 | fail 4 | `eval/FormulaEvalTestData.xls!G1540` =YEAR(-10) → Excel "#NUM!", Gnumeric 1899 |
| MONTH | 4 | fail 4 | `eval/FormulaEvalTestData.xls!G972` =MONTH(-10) → Excel "#NUM!", Gnumeric 12 |
| DATE | 4 | fail 4 | `eval/FormulaEvalTestData.xls!I328` =DATE(1900, 1, 22222) → Excel 22222, Gnumeric 22223 |
| LARGE | 4 | fail 4 | `eval/FormulaEvalTestData.xls!H848` =LARGE("2.3",1) → Excel 2.3, Gnumeric "#NUM!" |
| DAY | 4 | fail 4 | `eval/FormulaEvalTestData.xls!G348` =DAY(C16) → Excel "#NUM!", Gnumeric 29 |
| COUNT | 2 | fail 2 | `eval/FormulaEvalTestData.xls!F292` =COUNT(2, "A",  "",#REF!, #DIV/0!) → Excel 1, Gnumeric "#VALUE!" |
| COUNTIF | 2 | fail 2 | `eval/FormulaEvalTestData.xls!E304` =COUNTIF(AA7:AA13, 1/0) → Excel 1, Gnumeric "#DIV/0!" |
| CHOOSE | 2 | fail 2 | `eval/FormulaEvalTestData.xls!E244` =CHOOSE(#REF!, 1) → Excel "#REF!", Gnumeric "#VALUE!" |
| ADDRESS | 2 | fail 2 | `eval/FormulaEvalTestData.xls!F128` =ADDRESS(2,3,3,TRUE,"[Book1]Sheet1") → Excel "[Book1]Sheet1!$C2", Gnumeric "'[Book1]Sheet1'!$C2" |
| ATAN2 | 2 | fail 2 | `eval/FormulaEvalTestData.xls!H168` =ATAN2(B18,B16) → Excel "#VALUE!", Gnumeric 0.7853981633974483 |
| ACOSH | 2 | fail 2 | `eval/FormulaEvalTestData.xls!H108` =ACOSH(B18) → Excel "#VALUE!", Gnumeric 0 |
| TANH | 2 | fail 2 | `eval/FormulaEvalTestData.xls!L1400` =TANH(B18) → Excel "#VALUE!", Gnumeric 0.7615941559557649 |
| MODE | 2 | fail 2 | `eval/FormulaEvalTestData.xls!M968` =MODE(B9,2,D7) → Excel "#VALUE!", Gnumeric 2 |
| SEARCH | 2 | fail 2 | `eval/FormulaEvalTestData.xls!H1256` =SEARCH(B8,C8) → Excel 7, Gnumeric 1 |
