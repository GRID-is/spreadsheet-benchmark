# ironcalc-unit vs ironcalc@0.9.1 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 21. Fails where the case's own cell matches but its spill does not: 3.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 4 | 0 | 0 | 0 | 0 | 100.0% |
| dynamic_evaluation | 73 | 51 | 22 | 0 | 0 | 0 | 69.9% |
| engineering | 102 | 90 | 12 | 0 | 0 | 0 | 88.2% |
| general | 1128 | 1003 | 106 | 0 | 0 | 19 | 90.4% |
| logical | 47 | 42 | 5 | 0 | 0 | 0 | 89.4% |
| lookup_and_reference | 132 | 131 | 1 | 0 | 0 | 0 | 99.2% |
| mathematical | 31 | 30 | 1 | 0 | 0 | 0 | 96.8% |
| statistical | 272 | 267 | 5 | 0 | 0 | 0 | 98.2% |
| **total** | 1789 | 1618 | 152 | 0 | 0 | 19 | 91.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1618 | 152 | 0 | 0 | 19 | 91.4% |
| **total** | 1789 | 1618 | 152 | 0 | 0 | 19 | 91.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 11 | fail 11 | `dynamic_evaluation/bug_hunt.xlsx!A3` =SEQUENCE(INDIRECT("C1:C3")) → Excel 1, IronCalc 1 |
| DECIMAL | 10 | fail 10 | `general/base_decimal.xlsx!A1` =DECIMAL("", 2) → Excel 0, IronCalc "#VALUE!" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", IronCalc "#VALUE!" |
| SUM | 7 | fail 7 | `general/arrays_and_errors.xlsx!A6` =SUM(IFERROR(A1:C1/NA(), 2)) → Excel 2, IronCalc 6 |
| ENCODEURL | 7 | fail 7 | `general/fn_encodeurl.xlsx!A5` =ENCODEURL("~!*'()") → Excel "%7E%21%2A%27%28%29", IronCalc "~%21%2A%27%28%29" |
| WEEKNUM | 6 | fail 4, skip 2 | `general/date_and_time.xlsx!A2` =WEEKNUM(0) → Excel 0, IronCalc "#NUM!" |
| NOT | 5 | fail 5 | `logical/fn_not_array.xlsx!A1` =NOT({TRUE;FALSE}) → Excel false, IronCalc "#N/IMPL!" |
| BASE | 5 | fail 5 | `general/base_decimal.xlsx!A3` =BASE(5, 2, 256) → Excel "#NUM!", IronCalc "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000101" |
| DATEDIF | 5 | fail 5 | `general/datedif_leap_month_end.xlsx!A1` =DATEDIF("29/2/2020", "28/2/2021", "YD") → Excel "#VALUE!", IronCalc 0 |
| DGET | 4 | fail 4 | `general/database.xlsx!A14` =DGET(A1:C4, C1, B6:B7) → Excel 1200, IronCalc "#VALUE!" |
| DAY | 4 | fail 4 | `general/date_and_time.xlsx!A1` =DAY(TRUE) → Excel 1, IronCalc 31 |
| YEAR | 3 | fail 3 | `general/date_and_time.xlsx!A1` =YEAR(TRUE) → Excel 1900, IronCalc 1899 |
| MONTH | 3 | fail 3 | `general/date_and_time.xlsx!A1` =MONTH(TRUE) → Excel 1, IronCalc 12 |
| CONCAT | 3 | fail 3 | `general/fn_concatenate.xlsx!B3` =CONCAT("x", {1,2}, A1:A2, UPPER(D1:D3)) → Excel "x12", IronCalc "#N/IMPL!" |
| TODAY | 3 | skip 3 | `general/today.xlsx!A1` =TODAY() → Excel 46330, IronCalc 46335 |
| DEVSQ | 2 | fail 2 | `statistical/fn_devsq.xlsx!A3` =DEVSQ(1, ) → Excel 0.5, IronCalc 0 |
| WORKDAY | 2 | fail 2 | `general/date_and_time.xlsx!A6` =WORKDAY(44560,3,B1:B3) → Excel 44566, IronCalc "#VALUE!" |
| ISOWEEKNUM | 2 | fail 2 | `general/date_and_time.xlsx!A3` =ISOWEEKNUM(0) → Excel 52, IronCalc "#NUM!" |
| WEEKDAY | 2 | fail 2 | `general/date_and_time.xlsx!A1` =WEEKDAY(0) → Excel 7, IronCalc "#NUM!" |
| DATEVALUE | 2 | fail 2 | `general/datevalue_rollover.xlsx!A4` =DATEVALUE("2026-01-01T24:00:00") → Excel "#VALUE!", IronCalc 46024 |
| ECMA.CEILING | 2 | skip 2 | `general/ecma_ceiling.xlsx!A1` =ECMA.CEILING(4.3,2) → Excel 6, IronCalc 6 |
| FORMULATEXT | 2 | fail 2 | `general/fn_formulatext.xlsx!A2` =FORMULATEXT(D1:E1) → Excel "#N/A", IronCalc "#ERROR!" |
| LET | 2 | fail 2 | `general/fn_lambda.xlsx!A1` =LET(x, 1, y, LAMBDA(a,b, a + b), y(x, 22)) → Excel "#NAME?", IronCalc 23 |
| MAX | 2 | fail 2 | `general/fn_row_column.xlsx!A2` =MAX((A3:J3<>"")*COLUMN(A3:J3)) → Excel 0, IronCalc 7 |
| TYPE | 2 | fail 2 | `general/fn_type.xlsx!A2` =TYPE(A1:C30) → Excel 0, IronCalc 64 |
