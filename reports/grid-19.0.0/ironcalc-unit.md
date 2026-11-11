# ironcalc-unit vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 2. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 4 | 0 | 0 | 0 | 0 | 100.0% |
| dynamic_evaluation | 73 | 73 | 0 | 0 | 0 | 0 | 100.0% |
| engineering | 102 | 102 | 0 | 0 | 0 | 0 | 100.0% |
| general | 1128 | 1107 | 2 | 0 | 0 | 19 | 99.8% |
| logical | 47 | 47 | 0 | 0 | 0 | 0 | 100.0% |
| lookup_and_reference | 132 | 127 | 5 | 0 | 0 | 0 | 96.2% |
| mathematical | 31 | 31 | 0 | 0 | 0 | 0 | 100.0% |
| statistical | 272 | 272 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 1789 | 1763 | 7 | 0 | 0 | 19 | 99.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1763 | 7 | 0 | 0 | 19 | 99.6% |
| **total** | 1789 | 1763 | 7 | 0 | 0 | 19 | 99.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", GRID "#VALUE!" |
| XMATCH | 5 | fail 5 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("[A-Z][a-z]*day",A1:A3,3) → Excel 1, GRID "#VALUE!" |
| TODAY | 3 | skip 3 | `general/today.xlsx!A1` =TODAY() → Excel 46330, GRID 46335 |
| WEEKNUM | 2 | skip 2 | `general/date_and_time.xlsx!A2` =WEEKNUM(46029, -2) → Excel "#NUM!", GRID "#NUM!" |
| ECMA.CEILING | 2 | skip 2 | `general/ecma_ceiling.xlsx!A1` =ECMA.CEILING(4.3,2) → Excel 6, GRID 6 |
| NOW | 2 | skip 2 | `general/today.xlsx!A2` =NOW() → Excel 46330.77078611111, GRID 46335.06662734954 |
| LET | 1 | fail 1 | `general/fn_lambda.xlsx!A1` =LET(x, 1, y, LAMBDA(a,b, a + b), y(x, 22)) → Excel "#NAME?", GRID 23 |
| TEXT | 1 | skip 1 | `general/today.xlsx!A2` =TEXT(A1, "yyyy/m/d") → Excel "2026/11/4", GRID "2026/11/9" |
| TAN | 1 | fail 1 | `general/trigonometric.xlsx!A1` =TAN(PI()/2) → Excel 16324552277619072, GRID 16331239353195370 |
