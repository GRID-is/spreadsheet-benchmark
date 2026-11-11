# ironcalc-unit vs gnumeric@1.12.62 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 50. Fails where the case's own cell matches but its spill does not: 5.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 18 | 2 | 53 | 0 | 0 | 24.7% |
| engineering | 102 | 90 | 12 | 0 | 0 | 0 | 88.2% |
| general | 1128 | 769 | 252 | 88 | 0 | 19 | 69.3% |
| logical | 47 | 38 | 9 | 0 | 0 | 0 | 80.9% |
| lookup_and_reference | 132 | 94 | 6 | 32 | 0 | 0 | 71.2% |
| mathematical | 31 | 28 | 3 | 0 | 0 | 0 | 90.3% |
| statistical | 272 | 237 | 23 | 12 | 0 | 0 | 87.1% |
| **total** | 1789 | 1274 | 311 | 185 | 0 | 19 | 72.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1274 | 311 | 185 | 0 | 19 | 72.0% |
| **total** | 1789 | 1274 | 311 | 185 | 0 | 19 | 72.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 44 | unsupported 44 | `dynamic_evaluation/bug_hunt.xlsx!A1` =SEQUENCE(COUNTBLANK(C2:C3)+1) → Excel 1, Gnumeric "#NAME?" |
| LAMBDA | 19 | fail 19 | `general/fn_lambda.xlsx!A1` =LAMBDA(x,y, x + y)(1, 2) → Excel 3, Gnumeric "#\"_xlfn.LAMBDA(_xlpm.x,_xlpm.y, _xlpm.x + _xlpm.y)(1, 2)\"" |
| FILTER | 15 | unsupported 15 | `general/spill_functions.xlsx!C1` =FILTER(A1:A3,B1:B3) → Excel "Alice", Gnumeric "#NAME?" |
| SUM | 13 | fail 7, unsupported 6 | `dynamic_evaluation/known_gaps.xlsx!B1` =SUM(OFFSET(D1,1,0,2,1)) → Excel 5, Gnumeric "#NAME?" |
| WEEKNUM | 11 | fail 9, skip 2 | `general/date_and_time.xlsx!A14` =WEEKNUM(44561,0) → Excel "#NUM!", Gnumeric "#VALUE!" |
| LET | 11 | fail 11 | `general/fn_lambda.xlsx!A1` =LET(x, 1, y, LAMBDA(a,b, a + b), y(x, 22)) → Excel "#NAME?", Gnumeric "#\"_xlfn.LET(_xlpm.x, 1, _xlpm.y, _xlfn.LAMBDA(_xlpm.a,_xlpm.b, _xlpm.a + _xlpm.b), y(_xlpm.x, 22))\"" |
| ENCODEURL | 10 | fail 10 | `general/fn_encodeurl.xlsx!A3` =ENCODEURL("a b+c&d=e?f/g:h") → Excel "a%20b%2Bc%26d%3De%3Ff%2Fg%3Ah", Gnumeric "a%20b+c&d=e%3Ff/g:h" |
| POWER | 9 | fail 9 | `general/fn_power.xlsx!A5` =POWER(-1, 1/7) → Excel -1, Gnumeric "#NUM!" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", Gnumeric "#NAME?" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, Gnumeric "#NAME?" |
| MMULT | 7 | fail 7 | `general/fn_mmult.xlsx!F1` =MMULT(A1:B1, D1:D2) → Excel "#VALUE!", Gnumeric 1 |
| IFERROR | 6 | fail 6 | `logical/fn_iferror_array.xlsx!A1` =IFERROR({1,#DIV/0!}, {10,20,30}) → Excel 1, Gnumeric "#VALUE!" |
| DECIMAL | 6 | fail 6 | `general/base_decimal.xlsx!A2` =DECIMAL(A3, 2) → Excel 0, Gnumeric "#NUM!" |
| DATEVALUE | 6 | fail 6 | `general/datevalue_rollover.xlsx!A2` =DATEVALUE("2026-01-01 25:00:00") → Excel 46024, Gnumeric "#VALUE!" |
| TAKE | 6 | unsupported 6 | `general/fn_drop_take.xlsx!B1` =TAKE(A1:A3,2) → Excel 1, Gnumeric "#NAME?" |
| MAX | 6 | fail 1, unsupported 5 | `general/fn_max_min.xlsx!C1` =MAX(FILTER(A1:A3, A1:A3>15, "")) → Excel 30, Gnumeric "#NAME?" |
| SORT | 6 | unsupported 6 | `general/spill_functions.xlsx!B1` =SORT(A1:A3) → Excel 1, Gnumeric "#NAME?" |
| IF | 5 | fail 4, unsupported 1 | `dynamic_evaluation/bug_hunt.xlsx!E3` =IF(INDIRECT("C3"),"yes","no") → Excel "yes", Gnumeric "#NAME?" |
| NORM.S.DIST | 5 | unsupported 5 | `statistical/fn_norm_dist.xlsx!A5` =NORM.S.DIST(0.2, FALSE) → Excel 0.3910426939754559, Gnumeric "#NAME?" |
| BASE | 5 | fail 5 | `general/base_decimal.xlsx!A2` =BASE(5, 2, 255.9) → Excel "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000101", Gnumeric "#NUM!" |
| FLOOR.MATH | 5 | unsupported 5 | `general/ceiling_floor.xlsx!B10` =FLOOR.MATH(5.2) → Excel 5, Gnumeric "#NAME?" |
| WEEKDAY | 5 | fail 5 | `general/date_and_time.xlsx!A9` =WEEKDAY(-1) → Excel "#NUM!", Gnumeric 7 |
| BYCOL | 5 | fail 4, unsupported 1 | `general/fn_bycol_byrow.xlsx!A4` =BYCOL(A1:C2, LAMBDA(col, SUM(col))) → Excel 5, Gnumeric "#\"_xlfn.BYCOL(A1:C2, _xlfn.LAMBDA(_xlpm.col, SUM(_xlpm.col)))\"" |
| BYROW | 5 | fail 4, unsupported 1 | `general/fn_bycol_byrow.xlsx!D1` =BYROW(A1:B3, LAMBDA(row, SUM(row))) → Excel 3, Gnumeric "#\"_xlfn.BYROW(A1:B3, _xlfn.LAMBDA(_xlpm.row, SUM(_xlpm.row)))\"" |
| MAKEARRAY | 5 | fail 5 | `general/fn_bycol_byrow.xlsx!A1` =MAKEARRAY(3, 3, LAMBDA(x,y, x * y)) → Excel 1, Gnumeric "#\"_xlfn.MAKEARRAY(3, 3, _xlfn.LAMBDA(_xlpm.x,_xlpm.y, _xlpm.x * _xlpm.y))\"" |
