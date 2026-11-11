# ironcalc-unit vs fast-formula-parser@1.0.19 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 4. Fails where both answered an error, but a different one: 55. Fails where the case's own cell matches but its spill does not: 49.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 10 | 26 | 32 | 5 | 0 | 13.7% |
| engineering | 102 | 77 | 15 | 10 | 0 | 0 | 75.5% |
| general | 1128 | 443 | 232 | 434 | 0 | 19 | 39.9% |
| logical | 47 | 16 | 31 | 0 | 0 | 0 | 34.0% |
| lookup_and_reference | 132 | 42 | 16 | 74 | 0 | 0 | 31.8% |
| mathematical | 31 | 23 | 4 | 4 | 0 | 0 | 74.2% |
| statistical | 272 | 238 | 9 | 25 | 0 | 0 | 87.5% |
| **total** | 1789 | 849 | 337 | 579 | 5 | 19 | 48.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 849 | 337 | 579 | 5 | 19 | 48.0% |
| **total** | 1789 | 849 | 337 | 579 | 5 | 19 | 48.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 44 | fail 1, unsupported 40, error 3 | `dynamic_evaluation/bug_hunt.xlsx!A1` =SEQUENCE(COUNTBLANK(C2:C3)+1) → Excel 1, fast-formula-parser "#NAME?" |
| MAX | 30 | fail 2, unsupported 28 | `general/fn_max_min.xlsx!A5` =MAX(C1:C5) → Excel 0, fast-formula-parser "#NAME?" |
| INDIRECT | 28 | unsupported 28 | `general/fn_indirect.xlsx!E5` =INDIRECT("R2C2") → Excel "#REF!", fast-formula-parser "#NAME?" |
| SUM | 26 | fail 6, unsupported 18, error 2 | `dynamic_evaluation/known_gaps.xlsx!B1` =SUM(OFFSET(D1,1,0,2,1)) → Excel 5, fast-formula-parser "#NAME?" |
| XMATCH | 22 | unsupported 22 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("banana",A1:A3) → Excel 2, fast-formula-parser "#NAME?" |
| LAMBDA | 19 | unsupported 19 | `general/fn_lambda.xlsx!A1` =LAMBDA(x,y, x + y)(1, 2) → Excel 3, fast-formula-parser "#NAME?" |
| POWER | 19 | fail 19 | `general/fn_power.xlsx!A7` =POWER(B7, 2) → Excel 0, fast-formula-parser "#ERROR!" |
| ACCRINTM | 18 | unsupported 18 | `general/fn_accrintm.xlsx!A1` =ACCRINTM(39539, 39614, 0.1, 1000, 3) → Excel 20.54794520547945, fast-formula-parser "#NAME?" |
| MIN | 18 | unsupported 18 | `general/fn_max_min.xlsx!A4` =MIN(-1, -5) → Excel -5, fast-formula-parser "#NAME?" |
| IF | 16 | fail 13, unsupported 3 | `dynamic_evaluation/bug_hunt.xlsx!E3` =IF(INDIRECT("C3"),"yes","no") → Excel "yes", fast-formula-parser "#NAME?" |
| FILTER | 15 | unsupported 15 | `general/spill_functions.xlsx!C1` =FILTER(A1:A3,B1:B3) → Excel "Alice", fast-formula-parser "#NAME?" |
| ACCRINT | 14 | unsupported 14 | `general/fn_accrint.xlsx!A1` =ACCRINT(39508, 39691, 39569, 0.1, 1000, 2, 0) → Excel 16.666666666666664, fast-formula-parser "#NAME?" |
| IFERROR | 13 | fail 12, unsupported 1 | `logical/fn_iferror_array.xlsx!A1` =IFERROR({1,#DIV/0!,3}, "x") → Excel 1, fast-formula-parser 1 |
| WEEKNUM | 13 | fail 11, skip 2 | `general/date_and_time.xlsx!A14` =WEEKNUM(44561,0) → Excel "#NUM!", fast-formula-parser 53 |
| LOOKUP | 12 | unsupported 12 | `lookup_and_reference/fn_lookup.xlsx!B1` =LOOKUP(20,A1:A3) → Excel 20, fast-formula-parser "#NAME?" |
| CHOOSE | 12 | unsupported 12 | `general/fn_choose.xlsx!A3` =CHOOSE(1, 42) → Excel 42, fast-formula-parser "#NAME?" |
| CONVERT | 10 | unsupported 10 | `engineering/convert.xlsx!A8` =CONVERT(B1,"mi","km") → Excel 9.656064, fast-formula-parser "#NAME?" |
| WEEKDAY | 10 | fail 9, unsupported 1 | `general/date_and_time.xlsx!A9` =WEEKDAY(-1) → Excel "#NUM!", fast-formula-parser "#VALUE!" |
| INDEX | 9 | fail 7, unsupported 2 | `lookup_and_reference/fn_index.xlsx!A1` =INDEX({1,2,3,4},2) → Excel 2, fast-formula-parser "#REF!" |
| MATCH | 9 | unsupported 9 | `lookup_and_reference/fn_match.xlsx!A1` =MATCH(69,{24,43.5,52.8,69,269,387,770},0) → Excel 4, fast-formula-parser "#NAME?" |
| LET | 9 | unsupported 9 | `general/fn_lambda.xlsx!C9` =LET(x, A1, MAX(x)) → Excel 0, fast-formula-parser "#NAME?" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, fast-formula-parser "#NAME?" |
| DECIMAL | 8 | fail 8 | `general/base_decimal.xlsx!A1` =DECIMAL("", 2) → Excel 0, fast-formula-parser "#NUM!" |
| ENCODEURL | 8 | fail 8 | `general/fn_encodeurl.xlsx!A5` =ENCODEURL("~!*'()") → Excel "%7E%21%2A%27%28%29", fast-formula-parser "~!*'()" |

## Errors

- `dynamic_evaluation/range_reads.xlsx`: timeout after 60000 ms
- `dynamic_evaluation/range_reads.xlsx`: timeout after 60000 ms
- `dynamic_evaluation/range_reads.xlsx`: timeout after 60000 ms
- `dynamic_evaluation/range_reads.xlsx`: timeout after 60000 ms
- `dynamic_evaluation/range_reads.xlsx`: timeout after 60000 ms
