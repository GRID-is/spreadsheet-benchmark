# poi vs poi@6.0.0 (poi)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 7. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 47 of 3178 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| eval | 2453 | 2431 | 20 | 0 | 0 | 2 | 99.2% |
| functions | 693 | 672 | 19 | 0 | 0 | 2 | 97.3% |
| matrix | 34 | 34 | 0 | 0 | 0 | 0 | 100.0% |
| multisheet | 2 | 2 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 3182 | 3139 | 39 | 0 | 0 | 4 | 98.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 3182 | 3139 | 39 | 0 | 0 | 4 | 98.8% |
| **total** | 3182 | 3139 | 39 | 0 | 0 | 4 | 98.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IMAGINARY | 7 | fail 7 | `functions/ImRealFunctionTestCaseData.xls!B8` =IMAGINARY("14-0.88j") → Excel -0.88, Apache POI "-0.88" |
| MODE | 6 | fail 6 | `eval/FormulaEvalTestData.xls!K968` =MODE("1",1,#DIV/0!,2,3) → Excel "#DIV/0!", Apache POI "#VALUE!" |
| IMREAL | 5 | fail 5 | `functions/ImRealFunctionTestCaseData.xls!B9` =IMREAL("0+j") → Excel 0, Apache POI "0" |
| WEEKNUM | 4 | skip 4 | `eval/FormulaEvalTestData.xls!J1524` =WEEKNUM("99/9/2008",-1) → Excel "#VALUE!", Apache POI "#VALUE!" |
| DGET | 4 | fail 4 | `functions/DGet.xls!B95` =DGET($G$94:$N$95,"One",$V$94:$W$95) → Excel "a", Apache POI "#VALUE!" |
| CODE | 3 | fail 3 | `functions/CodeFunctionTestCaseData.xls!B7` =CODE("!") → Excel 33, Apache POI "33" |
| ADDRESS | 2 | fail 2 | `eval/FormulaEvalTestData.xls!F128` =ADDRESS(2,3,3,TRUE,"[Book1]Sheet1") → Excel "[Book1]Sheet1!$C2", Apache POI "'[Book1]Sheet1'!$C2" |
| AND | 2 | fail 2 | `eval/FormulaEvalTestData.xls!J136` =AND("1",1) → Excel true, Apache POI "#VALUE!" |
| SUMPRODUCT | 2 | fail 2 | `eval/FormulaEvalTestData.xls!J1368` =SUMPRODUCT(G1370,I1370) → Excel 0, Apache POI "#VALUE!" |
| MINA | 2 | fail 2 | `eval/FormulaEvalTestData.xls!J944` =MINA(F12:F14,F10) → Excel 0, Apache POI 0.00001 |
| DOLLAR | 2 | fail 2 | `eval/FormulaEvalTestData.xls!L436` =DOLLAR(J434:J437) → Excel "$1.00", Apache POI "#VALUE!" |
| SUM | 1 | fail 1 | `eval/FormulaEvalTestData.xls!E75` =SUM(J7:(I8)) → Excel 15, Apache POI "#VALUE!" |
