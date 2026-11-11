# ironcalc-unit vs google-sheets@2026-11-09 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 1. Fails where both answered an error, but a different one: 36. Fails where the case's own cell matches but its spill does not: 6.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 40 | 30 | 3 | 0 | 0 | 54.8% |
| engineering | 102 | 100 | 2 | 0 | 0 | 0 | 98.0% |
| general | 1128 | 932 | 125 | 52 | 0 | 19 | 84.0% |
| logical | 47 | 44 | 3 | 0 | 0 | 0 | 93.6% |
| lookup_and_reference | 132 | 109 | 8 | 15 | 0 | 0 | 82.6% |
| mathematical | 31 | 28 | 3 | 0 | 0 | 0 | 90.3% |
| statistical | 272 | 263 | 9 | 0 | 0 | 0 | 96.7% |
| **total** | 1789 | 1516 | 184 | 70 | 0 | 19 | 85.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1516 | 184 | 70 | 0 | 19 | 85.6% |
| **total** | 1789 | 1516 | 184 | 70 | 0 | 19 | 85.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 21 | fail 20, unsupported 1 | `dynamic_evaluation/bug_hunt.xlsx!E1` =SEQUENCE(INDIRECT("C2:C3")) → Excel "#SPILL!", Google Sheets 1 |
| INDIRECT | 15 | fail 15 | `general/fn_indirect.xlsx!E1` =INDIRECT("B2") → Excel 0, Google Sheets "" |
| FILTER | 15 | unsupported 15 | `general/spill_functions.xlsx!C1` =FILTER(A1:A3,B1:B3) → Excel "Alice", Google Sheets "#NAME?" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", Google Sheets "#N/A" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, Google Sheets "#NAME?" |
| CONCAT | 8 | fail 8 | `general/fn_concatenate.xlsx!D3` =CONCAT(A1:A3, "!") → Excel "Hello my World!", Google Sheets "World!" |
| SUM | 6 | fail 4, unsupported 2 | `dynamic_evaluation/known_gaps.xlsx!A1` =SUM(C1#) → Excel 6, Google Sheets "#NAME?" |
| WEEKNUM | 6 | fail 4, skip 2 | `general/date_and_time.xlsx!A2` =WEEKNUM(0) → Excel 0, Google Sheets 52 |
| TAKE | 6 | unsupported 6 | `general/fn_drop_take.xlsx!B1` =TAKE(A1:A3,2) → Excel 1, Google Sheets "#NAME?" |
| SORT | 6 | unsupported 6 | `general/spill_functions.xlsx!B1` =SORT(A1:A3) → Excel 1, Google Sheets "#NAME?" |
| XMATCH | 5 | fail 5 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("[A-Z][a-z]*day",A1:A3,3) → Excel 1, Google Sheets "#VALUE!" |
| DAY | 5 | fail 5 | `general/date_and_time.xlsx!A1` =DAY(TRUE) → Excel 1, Google Sheets 31 |
| DROP | 5 | unsupported 5 | `general/fn_drop_take.xlsx!B1` =DROP(A1:A3,1) → Excel 2, Google Sheets "#NAME?" |
| MAX | 5 | fail 2, unsupported 3 | `general/fn_max_min.xlsx!A1` =MAX(LEFT("abc", 1)) → Excel "#VALUE!", Google Sheets 0 |
| ODDFPRICE | 5 | unsupported 5 | `general/fn_oddfprice.xlsx!C3` =ODDFPRICE(A1,A2,A3,A4,A5,A6,A7,A8,A9) → Excel 113.50846650690974, Google Sheets "#NAME?" |
| EXPAND | 4 | unsupported 4 | `lookup_and_reference/fn_expand.xlsx!C1` =EXPAND(A1:A2,4) → Excel 1, Google Sheets "#NAME?" |
| DECIMAL | 4 | fail 4 | `general/base_decimal.xlsx!A2` =DECIMAL(REPT("0",255)&"1", 2) → Excel "#VALUE!", Google Sheets "#NUM!" |
| DATEVALUE | 4 | fail 4 | `general/datevalue_rollover.xlsx!A4` =DATEVALUE("2026-01-01T24:00:00") → Excel "#VALUE!", Google Sheets 46024 |
| CHOOSE | 4 | fail 4 | `general/fn_choose.xlsx!A4` =CHOOSE(2, 42) → Excel "#VALUE!", Google Sheets "#NUM!" |
| POWER | 4 | fail 4 | `general/fn_power.xlsx!A1` =POWER(0, 0) → Excel "#NUM!", Google Sheets 1 |
| WEEKDAY | 4 | fail 4 | `general/weekday_arg_order.xlsx!A1` =WEEKDAY("1/1/10000", 1/0) → Excel "#VALUE!", Google Sheets "#DIV/0!" |
| IF | 3 | fail 3 | `logical/fn_if_array.xlsx!A1` =IF({TRUE,FALSE}, {1,2,3}, {"A","B","C"}) → Excel 1, Google Sheets 1 |
| CHISQ.TEST | 3 | fail 3 | `statistical/fn_chisq_test.xlsx!G6` =CHISQ.TEST({1,2,3}, {3;3;4}) → Excel 0.38353157287631073, Google Sheets "#N/A" |
| YEAR | 3 | fail 3 | `general/date_and_time.xlsx!A1` =YEAR(TRUE) → Excel 1900, Google Sheets 1899 |
| MONTH | 3 | fail 3 | `general/date_and_time.xlsx!A1` =MONTH(TRUE) → Excel 1, Google Sheets 12 |
