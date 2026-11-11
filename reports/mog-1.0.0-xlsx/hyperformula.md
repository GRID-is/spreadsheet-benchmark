# hyperformula vs mog@1.0.0 (hyperformula)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-lacks-function 241, volatile 10. Fails where the engine's error came from another cell: 2. Fails where both answered an error, but a different one: 85. Fails where the case's own cell matches but its spill does not: 2.
Functions outside Excel's list (22): ARRAYFORMULA, ARRAY_CONSTRAIN, COUNTUNIQUE, HF.ADD, HF.CONCAT, HF.DIVIDE, HF.EQ, HF.GT, HF.GTE, HF.LT, HF.LTE, HF.MINUS, HF.MULTIPLY, HF.NE, HF.POW, HF.UMINUS, HF.UNARY_PERCENT, HF.UPLUS, INTERVAL, ISBINARY, MAXPOOL, MEDIANPOOL. Never-supported functions present: ARRAY_CONSTRAIN.
Upstream literal differs from Excel's value in 677 of 4491 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 4166 | 3473 | 421 | 0 | 31 | 241 | 88.5% |
| other | 576 | 529 | 37 | 0 | 0 | 10 | 93.5% |
| **total** | 4742 | 4002 | 458 | 0 | 31 | 251 | 89.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 4742 | 4002 | 458 | 0 | 31 | 251 | 89.1% |
| **total** | 4742 | 4002 | 458 | 0 | 31 | 251 | 89.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| WEEKNUM | 31 | error 31 | `functions/function-weeknum.xlsx!A1` =WEEKNUM("foo", 1) → Excel "#VALUE!", Mog null |
| XLOOKUP | 30 | fail 30 | `functions/function-xlookup.xlsx!A6` =XLOOKUP(1, B1:E1, B2:C3) → Excel "#VALUE!", Mog "#N/A" |
| HF.EQ | 17 | skip 17 | `functions/function-hfeq.xlsx!A1` =HF.EQ(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.GT | 17 | skip 17 | `functions/function-hfgt.xlsx!A1` =HF.GT(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.GTE | 17 | skip 17 | `functions/function-hfgte.xlsx!A1` =HF.GTE(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.LT | 17 | skip 17 | `functions/function-hflt.xlsx!A1` =HF.LT(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.LTE | 17 | skip 17 | `functions/function-hflte.xlsx!A1` =HF.LTE(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.NE | 17 | skip 17 | `functions/function-hfne.xlsx!A1` =HF.NE(1) → Excel "#NAME?", Mog "#NAME?" |
| ADDRESS | 14 | fail 14 | `functions/function-address.xlsx!A5` =ADDRESS(1,1,4, TRUE(), "") → Excel "!A1", Mog "A1" |
| INTERVAL | 13 | skip 13 | `functions/function-interval.xlsx!B1` =INTERVAL("12/30/2018") → Excel "#NAME?", Mog "#NAME?" |
| VLOOKUP | 13 | fail 13 | `functions/function-vlookup.xlsx!A1` =VLOOKUP(D1:E1, A2:B3, 2, TRUE()) → Excel "#VALUE!", Mog "#N/A" |
| HF.DIVIDE | 12 | skip 12 | `functions/function-hfdivide.xlsx!A1` =HF.DIVIDE(1) → Excel "#NAME?", Mog "#NAME?" |
| COUNTUNIQUE | 11 | skip 11 | `functions/function-countunique.xlsx!A1` =COUNTUNIQUE() → Excel "#NAME?", Mog "#NAME?" |
| HF.ADD | 11 | skip 11 | `functions/function-hfadd.xlsx!A1` =HF.ADD(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.MINUS | 11 | skip 11 | `functions/function-hfminus.xlsx!A1` =HF.MINUS(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.MULTIPLY | 11 | skip 11 | `functions/function-hfmultiply.xlsx!A1` =HF.MULTIPLY(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.UMINUS | 11 | skip 11 | `functions/function-hfuminus.xlsx!A1` =HF.UMINUS() → Excel "#NAME?", Mog "#NAME?" |
| SUMIF | 11 | fail 11 | `functions/function-sumif.xlsx!A4` =SUMIF(A1:A3, "=1", B1:B3) → Excel 0, Mog 6 |
| SUMPRODUCT | 11 | fail 11 | `functions/function-sumproduct.xlsx!A2` =SUMPRODUCT(A1,B1) → Excel 0, Mog 42 |
| WORKDAY.INTL | 11 | fail 11 | `functions/function-workday.intl.xlsx!A3` =WORKDAY.INTL(0, 1, "1010102") → Excel "#VALUE!", Mog "#NUM!" |
| HF.POW | 10 | skip 10 | `functions/function-hfpow.xlsx!A1` =HF.POW(1) → Excel "#NAME?", Mog "#NAME?" |
| HF.UPLUS | 10 | skip 10 | `functions/function-hfuplus.xlsx!A1` =HF.UPLUS() → Excel "#NAME?", Mog "#NAME?" |
| ISOWEEKNUM | 10 | fail 10 | `functions/function-isoweeknum.xlsx!A1` =ISOWEEKNUM(-1) → Excel "#NUM!", Mog 1 |
| NETWORKDAYS.INTL | 10 | fail 10 | `functions/function-networkdays.intl.xlsx!A3` =NETWORKDAYS.INTL(0, 1, "1010102") → Excel "#VALUE!", Mog "#NUM!" |
| HF.CONCAT | 9 | skip 9 | `functions/function-hfconcat.xlsx!A1` =HF.CONCAT(1) → Excel "#NAME?", Mog "#NAME?" |

## Errors

- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `functions/function-weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68051412) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
