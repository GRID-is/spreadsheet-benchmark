# ironcalc-unit vs mog@1.0.0 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 19. Fails where the case's own cell matches but its spill does not: 19.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1772 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 2 | 2 | 0 | 0 | 0 | 50.0% |
| dynamic_evaluation | 73 | 61 | 12 | 0 | 0 | 0 | 83.6% |
| engineering | 102 | 90 | 12 | 0 | 0 | 0 | 88.2% |
| general | 1128 | 761 | 186 | 4 | 160 | 17 | 68.5% |
| logical | 47 | 35 | 12 | 0 | 0 | 0 | 74.5% |
| lookup_and_reference | 132 | 115 | 17 | 0 | 0 | 0 | 87.1% |
| mathematical | 31 | 30 | 1 | 0 | 0 | 0 | 96.8% |
| statistical | 272 | 256 | 16 | 0 | 0 | 0 | 94.1% |
| **total** | 1789 | 1350 | 258 | 4 | 160 | 17 | 76.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1350 | 258 | 4 | 160 | 17 | 76.2% |
| **total** | 1789 | 1350 | 258 | 4 | 160 | 17 | 76.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| WEEKNUM | 24 | error 24 | `general/date_and_time.xlsx!A14` =WEEKNUM(44561,0) → Excel "#NUM!", Mog null |
| SUM | 14 | fail 14 | `general/arrays_and_errors.xlsx!A6` =SUM(IFERROR(A1:C1/NA(), 2)) → Excel 2, Mog 6 |
| WORKDAY | 14 | error 14 | `general/date_and_time.xlsx!A11` =WORKDAY(1,1) → Excel 2, Mog null |
| WEEKDAY | 12 | error 12 | `general/date_and_time.xlsx!A6` =WEEKDAY(44557,2) → Excel 1, Mog null |
| WORKDAY.INTL | 12 | fail 1, error 11 | `general/date_and_time.xlsx!A10` =WORKDAY.INTL(44560,1,0) → Excel "#NUM!", Mog null |
| DAYS360 | 10 | error 10 | `general/date_and_time.xlsx!A9` =DAYS360(1,2,FALSE) → Excel 1, Mog null |
| EDATE | 10 | error 10 | `general/date_and_time.xlsx!A9` =EDATE(-1,1) → Excel "#NUM!", Mog null |
| MAX | 10 | fail 10 | `general/fn_max_min.xlsx!A1` =MAX({-1,"a",TRUE,-7}) → Excel -1, Mog 1 |
| DECIMAL | 9 | fail 9 | `general/base_decimal.xlsx!A1` =DECIMAL("", 2) → Excel 0, Mog "#NUM!" |
| ISOWEEKNUM | 9 | error 9 | `general/date_and_time.xlsx!A5` =ISOWEEKNUM(44564) → Excel 1, Mog null |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", Mog "#VALUE!" |
| DB | 8 | fail 8 | `general/fn_financial.xlsx!B4` =DB(A2,A3,A4,4,7) → Excel 120410.63551274998, Mog "#VALUE!" |
| SEQUENCE | 7 | fail 7 | `dynamic_evaluation/bug_hunt.xlsx!A3` =SEQUENCE(INDIRECT("C1:C3")) → Excel 1, Mog "#VALUE!" |
| YEAR | 7 | error 7 | `general/date_and_time.xlsx!A2` =YEAR(27819) → Excel 1976, Mog null |
| MONTH | 7 | error 7 | `general/date_and_time.xlsx!A2` =MONTH(27819) → Excel 2, Mog null |
| DAY | 7 | error 7 | `general/date_and_time.xlsx!A2` =DAY(27819) → Excel 29, Mog null |
| YEARFRAC | 7 | error 7 | `general/date_and_time.xlsx!A6` =YEARFRAC(44197,44562,1) → Excel 1, Mog null |
| EOMONTH | 7 | error 7 | `general/date_and_time.xlsx!A2` =EOMONTH(46000, 3) → Excel 46112, Mog null |
| IFERROR | 6 | fail 6 | `logical/fn_iferror_array.xlsx!A1` =IFERROR({1,#DIV/0!,3}, {10,20,30}) → Excel 1, Mog 1 |
| DATE | 6 | error 6 | `general/date_and_time.xlsx!C2` =DATE(10000, 5, 5) → Excel "#NUM!", Mog null |
| TIME | 6 | error 6 | `general/date_and_time.xlsx!A6` =TIME(12, 30, -1) → Excel 0.5208217592592592, Mog null |
| HOUR | 6 | error 6 | `general/date_and_time.xlsx!A5` =HOUR("abc") → Excel "#VALUE!", Mog null |
| MINUTE | 6 | error 6 | `general/date_and_time.xlsx!A5` =MINUTE("abc") → Excel "#VALUE!", Mog null |
| SECOND | 6 | error 6 | `general/date_and_time.xlsx!A5` =SECOND("abc") → Excel "#VALUE!", Mog null |
| DATEVALUE | 6 | fail 6 | `general/datevalue_rollover.xlsx!A2` =DATEVALUE("2026-01-01 25:00:00") → Excel 46024, Mog "#VALUE!" |

## Errors

- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `general/date_and_time.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68063678) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
