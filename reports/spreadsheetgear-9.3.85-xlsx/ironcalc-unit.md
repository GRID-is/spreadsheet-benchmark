# ironcalc-unit vs spreadsheetgear@9.3.85 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 2. Fails where both answered an error, but a different one: 13. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1772 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 18 | 1 | 54 | 0 | 0 | 24.7% |
| engineering | 102 | 98 | 4 | 0 | 0 | 0 | 96.1% |
| general | 1128 | 759 | 40 | 122 | 190 | 17 | 68.3% |
| logical | 47 | 47 | 0 | 0 | 0 | 0 | 100.0% |
| lookup_and_reference | 132 | 78 | 0 | 54 | 0 | 0 | 59.1% |
| mathematical | 31 | 31 | 0 | 0 | 0 | 0 | 100.0% |
| statistical | 272 | 262 | 10 | 0 | 0 | 0 | 96.3% |
| **total** | 1789 | 1293 | 59 | 230 | 190 | 17 | 73.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1293 | 59 | 230 | 190 | 17 | 73.0% |
| **total** | 1789 | 1293 | 59 | 230 | 190 | 17 | 73.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 44 | unsupported 44 | `dynamic_evaluation/bug_hunt.xlsx!A1` =SEQUENCE(COUNTBLANK(C2:C3)+1) → Excel 1, SpreadsheetGear "#NAME?" |
| WEEKNUM | 24 | error 24 | `general/date_and_time.xlsx!A14` =WEEKNUM(44561,0) → Excel "#NUM!", SpreadsheetGear null |
| XMATCH | 22 | unsupported 22 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("banana",A1:A3) → Excel 2, SpreadsheetGear "#NAME?" |
| LAMBDA | 19 | error 19 | `general/fn_lambda.xlsx!A1` =LAMBDA(x,y, x + y)(1, 2) → Excel 3, SpreadsheetGear null |
| FILTER | 15 | unsupported 15 | `general/spill_functions.xlsx!C1` =FILTER(A1:A3,B1:B3) → Excel "Alice", SpreadsheetGear "#NAME?" |
| WORKDAY | 14 | error 14 | `general/date_and_time.xlsx!A11` =WORKDAY(1,1) → Excel 2, SpreadsheetGear null |
| WEEKDAY | 12 | error 12 | `general/date_and_time.xlsx!A6` =WEEKDAY(44557,2) → Excel 1, SpreadsheetGear null |
| WORKDAY.INTL | 11 | error 11 | `general/date_and_time.xlsx!A10` =WORKDAY.INTL(44560,1,0) → Excel "#NUM!", SpreadsheetGear null |
| LET | 11 | error 11 | `general/fn_lambda.xlsx!A1` =LET(x, 1, y, LAMBDA(a,b, a + b), y(x, 22)) → Excel "#NAME?", SpreadsheetGear null |
| DAYS360 | 10 | error 10 | `general/date_and_time.xlsx!A9` =DAYS360(1,2,FALSE) → Excel 1, SpreadsheetGear null |
| EDATE | 10 | error 10 | `general/date_and_time.xlsx!A9` =EDATE(-1,1) → Excel "#NUM!", SpreadsheetGear null |
| ISOWEEKNUM | 9 | error 9 | `general/date_and_time.xlsx!A5` =ISOWEEKNUM(44564) → Excel 1, SpreadsheetGear null |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", SpreadsheetGear "#NAME?" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, SpreadsheetGear "#NAME?" |
| CONCAT | 8 | unsupported 8 | `general/fn_concatenate.xlsx!D3` =CONCAT(A1:A3, "!") → Excel "Hello my World!", SpreadsheetGear "#NAME?" |
| YEAR | 7 | error 7 | `general/date_and_time.xlsx!A2` =YEAR(27819) → Excel 1976, SpreadsheetGear null |
| MONTH | 7 | error 7 | `general/date_and_time.xlsx!A2` =MONTH(27819) → Excel 2, SpreadsheetGear null |
| DAY | 7 | error 7 | `general/date_and_time.xlsx!A2` =DAY(27819) → Excel 29, SpreadsheetGear null |
| YEARFRAC | 7 | error 7 | `general/date_and_time.xlsx!A6` =YEARFRAC(44197,44562,1) → Excel 1, SpreadsheetGear null |
| EOMONTH | 7 | error 7 | `general/date_and_time.xlsx!A2` =EOMONTH(46000, 3) → Excel 46112, SpreadsheetGear null |
| SUM | 6 | unsupported 6 | `dynamic_evaluation/known_gaps.xlsx!B1` =SUM(OFFSET(D1,1,0,2,1)) → Excel 5, SpreadsheetGear "#NAME?" |
| DATE | 6 | error 6 | `general/date_and_time.xlsx!C2` =DATE(10000, 5, 5) → Excel "#NUM!", SpreadsheetGear null |
| TIME | 6 | error 6 | `general/date_and_time.xlsx!A6` =TIME(12, 30, -1) → Excel 0.5208217592592592, SpreadsheetGear null |
| HOUR | 6 | error 6 | `general/date_and_time.xlsx!A5` =HOUR("abc") → Excel "#VALUE!", SpreadsheetGear null |
| MINUTE | 6 | error 6 | `general/date_and_time.xlsx!A5` =MINUTE("abc") → Excel "#VALUE!", SpreadsheetGear null |

## Errors

- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/date_and_time.xlsx`: spreadsheetgear: System.AggregateException: One or more errors occurred. (Index was outside the bounds of the array.) ---> System.IndexOutOfRangeException: Index was outside the bounds of the array.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `general/fn_let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
