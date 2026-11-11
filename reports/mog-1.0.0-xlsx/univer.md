# univer vs mog@1.0.0 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 216. Fails where the case's own cell matches but its spill does not: 103.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 157 | 18 | 0 | 0 | 81 | 89.7% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 103 | 4 | 0 | 0 | 0 | 96.3% |
| database | 182 | 144 | 38 | 0 | 0 | 0 | 79.1% |
| date | 378 | 222 | 103 | 2 | 16 | 35 | 64.7% |
| engineering | 572 | 363 | 179 | 0 | 0 | 30 | 67.0% |
| financial | 800 | 581 | 170 | 3 | 46 | 0 | 72.6% |
| information | 156 | 107 | 8 | 0 | 0 | 41 | 93.0% |
| logical | 107 | 86 | 15 | 5 | 0 | 1 | 81.1% |
| lookup | 625 | 233 | 121 | 0 | 264 | 7 | 37.7% |
| math | 788 | 634 | 109 | 2 | 0 | 43 | 85.1% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 995 | 238 | 11 | 18 | 41 | 78.8% |
| text | 549 | 445 | 73 | 2 | 0 | 29 | 85.6% |
| web | 8 | 6 | 0 | 2 | 0 | 0 | 75.0% |
| **total** | 5856 | 4076 | 1076 | 27 | 344 | 333 | 73.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 4076 | 1076 | 27 | 344 | 333 | 73.8% |
| **total** | 5856 | 4076 | 1076 | 27 | 344 | 333 | 73.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 170 | fail 9, error 161 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, Mog 101 |
| INDEX | 101 | error 101 | `lookup/index.xlsx!Z1` =INDEX(A1:B1,1) → Excel "Tom", Mog null |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", Mog 46296 |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", Mog "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, Mog 46335 |
| ODDLPRICE | 23 | error 23 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, Mog null |
| ODDLYIELD | 23 | error 23 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, Mog null |
| NETWORKDAYS.INTL | 22 | fail 21, unsupported 1 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =NETWORKDAYS.INTL(44613,44613,"0000000",A1:A54) → Excel "#VALUE!", Mog 0 |
| WORKDAY.INTL | 21 | fail 20, unsupported 1 | `date/workday-intl.xlsx!Z1` =WORKDAY.INTL(45000,1,"1") → Excel "#VALUE!", Mog 45001 |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", Mog "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", Mog "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", Mog 44734.5869212963 |
| TREND | 19 | fail 17, unsupported 2 | `statistical/trend.xlsx!Z1` =TREND(A1:A6) → Excel 12452.380952380969, Mog 12452.380952380969 |
| VLOOKUP | 18 | fail 18 | `lookup/vlookup.xlsx!Z1` =VLOOKUP(,A1:B2,2,0) → Excel "Zero", Mog "Blank" |
| GROWTH | 18 | fail 16, unsupported 2 | `statistical/growth.xlsx!Z1` =GROWTH(A1:A6) → Excel 32618.203773539757, Mog 32618.20377353983 |
| WEEKNUM | 16 | error 16 | `date/weeknum.xlsx!Z1` =WEEKNUM(43832,2) → Excel 1, Mog null |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", Mog "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", Mog "#NAME?" |
| LOGEST | 14 | fail 12, unsupported 2 | `statistical/logest.xlsx!Z1` =LOGEST(A1:A4,A6:A9,TRUE,TRUE) → Excel 1.578436529997654, Mog 1.578436529997654 |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", Mog true |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, Mog 0.9109940930349454 |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", Mog true |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", Mog true |
| CHOOSE | 11 | fail 11 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =CHOOSE(A1:B1,A3:A7,A9:A13) → Excel "Red", Mog "#VALUE!" |
| NETWORKDAYS | 11 | fail 11 | `date/networkdays.xlsx!Z1` =NETWORKDAYS(0,0) → Excel 0, Mog "#NUM!" |

## Errors

- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68089501) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlprice.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `financial/oddlyield.xlsx`: timeout after 55000 ms
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/index.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097275) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 3, cols: 6 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `lookup/xlookup.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68097936) panicked at compute/core/crates/types/value-types/src/cell_array.rs:64:35:
CellArray::new: invalid shape: ShapeMismatch { data_len: 10, cols: 3 }
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
