# ironcalc-unit vs spread-sheets@19.2.5 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 21. Fails where the case's own cell matches but its spill does not: 7.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 4 | 0 | 0 | 0 | 0 | 100.0% |
| dynamic_evaluation | 73 | 70 | 3 | 0 | 0 | 0 | 95.9% |
| engineering | 102 | 86 | 16 | 0 | 0 | 0 | 84.3% |
| general | 1128 | 939 | 164 | 6 | 0 | 19 | 84.7% |
| logical | 47 | 46 | 1 | 0 | 0 | 0 | 97.9% |
| lookup_and_reference | 132 | 124 | 8 | 0 | 0 | 0 | 93.9% |
| mathematical | 31 | 30 | 1 | 0 | 0 | 0 | 96.8% |
| statistical | 272 | 250 | 22 | 0 | 0 | 0 | 91.9% |
| **total** | 1789 | 1549 | 215 | 6 | 0 | 19 | 87.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1549 | 215 | 6 | 0 | 19 | 87.5% |
| **total** | 1789 | 1549 | 215 | 6 | 0 | 19 | 87.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SUM | 15 | fail 15 | `general/arrays_and_errors.xlsx!A6` =SUM(IFERROR(A1:C1/NA(), 2)) → Excel 2, SpreadJS 6 |
| MAX | 9 | fail 9 | `general/fn_max_min.xlsx!A5` =MAX("3", 2) → Excel 3, SpreadJS "#VALUE!" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", SpreadJS "#VALUE!" |
| WEEKNUM | 6 | fail 4, skip 2 | `general/date_and_time.xlsx!A2` =WEEKNUM(0) → Excel 0, SpreadJS 52 |
| MIN | 6 | fail 6 | `general/fn_max_min.xlsx!A6` =MIN("-3", 2) → Excel -3, SpreadJS "#VALUE!" |
| CHISQ.TEST | 5 | fail 5 | `statistical/fn_chisq_test.xlsx!G6` =CHISQ.TEST({1,2,3}, {3;3;4}) → Excel 0.38353157287631073, SpreadJS "#N/A" |
| DAY | 5 | fail 5 | `general/date_and_time.xlsx!A1` =DAY(TRUE) → Excel 1, SpreadJS 31 |
| XMATCH | 4 | fail 4 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("[A-Z][a-z]*day",A1:A3,3) → Excel 1, SpreadJS "#N/A" |
| DSUM | 4 | fail 4 | `general/database.xlsx!A12` =DSUM(A1:C4, C1, A6:A7) → Excel 0, SpreadJS 3000 |
| DPRODUCT | 4 | fail 4 | `general/database.xlsx!A13` =DPRODUCT(A1:C4, C1, A6:A7) → Excel 0, SpreadJS 1890000 |
| DSTDEVP | 4 | fail 4 | `general/database.xlsx!A20` =DSTDEVP(A1:C4, C1, A6:A7) → Excel "#DIV/0!", SpreadJS 600 |
| DCOUNT | 4 | fail 4 | `general/database.xlsx!A15` =DCOUNT(A1:C4, C1, A6:A7) → Excel 0, SpreadJS 2 |
| DVARP | 4 | fail 4 | `general/database.xlsx!A18` =DVARP(A1:C4, C1, A6:A7) → Excel "#DIV/0!", SpreadJS 360000 |
| DVAR | 4 | fail 4 | `general/database.xlsx!A17` =DVAR(A1:C4, C1, A6:A7) → Excel "#DIV/0!", SpreadJS 720000 |
| DCOUNTA | 4 | fail 4 | `general/database.xlsx!A16` =DCOUNTA(A1:C4, C1, A6:A7) → Excel 0, SpreadJS 2 |
| DSTDEV | 4 | fail 4 | `general/database.xlsx!A19` =DSTDEV(A1:C4, C1, A6:A7) → Excel "#DIV/0!", SpreadJS 848.528137423857 |
| DMAX | 4 | fail 4 | `general/database.xlsx!A10` =DMAX(A1:C4, C1, A6:A7) → Excel 0, SpreadJS 2100 |
| DMIN | 4 | fail 4 | `general/database.xlsx!A9` =DMIN(A1:C4, C1, A6:A7) → Excel 0, SpreadJS 900 |
| DAVERAGE | 4 | fail 4 | `general/database.xlsx!A11` =DAVERAGE(A1:C4, C1, A6:A7) → Excel "#DIV/0!", SpreadJS 1500 |
| EOMONTH | 4 | fail 4 | `general/date_and_time.xlsx!A4` =EOMONTH(DATE(2023,3,2), 1000000000000000) → Excel "#NUM!", SpreadJS "/OADate(NaN)/" |
| INDIRECT | 4 | fail 4 | `general/fn_indirect.xlsx!E7` =INDIRECT("R[2147483647]C[2147483647]", FALSE) → Excel "#REF!", SpreadJS "#VALUE!" |
| MMULT | 4 | fail 4 | `general/fn_mmult.xlsx!G1` =MMULT(A1:B2, D1:E2) → Excel "#VALUE!", SpreadJS 5 |
| PRODUCT | 4 | fail 4 | `general/fn_product.xlsx!A1` =PRODUCT("1") → Excel 1, SpreadJS "#VALUE!" |
| WEEKDAY | 4 | fail 4 | `general/weekday_arg_order.xlsx!A1` =WEEKDAY("1/1/10000", 1/0) → Excel "#VALUE!", SpreadJS "#DIV/0!" |
| SEQUENCE | 3 | fail 3 | `dynamic_evaluation/bug_hunt.xlsx!B1` =SEQUENCE(A1) → Excel "#CALC!", SpreadJS "#VALUE!" |
