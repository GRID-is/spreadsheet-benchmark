# ironcalc-unit vs hyperformula@3.5.0 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 49. Fails where the case's own cell matches but its spill does not: 10.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 24 | 44 | 5 | 0 | 0 | 32.9% |
| engineering | 102 | 72 | 8 | 22 | 0 | 0 | 70.6% |
| general | 1128 | 610 | 245 | 254 | 0 | 19 | 55.0% |
| logical | 47 | 30 | 3 | 14 | 0 | 0 | 63.8% |
| lookup_and_reference | 132 | 38 | 21 | 73 | 0 | 0 | 28.8% |
| mathematical | 31 | 17 | 2 | 12 | 0 | 0 | 54.8% |
| statistical | 272 | 189 | 9 | 74 | 0 | 0 | 69.5% |
| **total** | 1789 | 980 | 336 | 454 | 0 | 19 | 55.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 980 | 336 | 454 | 0 | 19 | 55.4% |
| **total** | 1789 | 980 | 336 | 454 | 0 | 19 | 55.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDIRECT | 28 | unsupported 28 | `general/fn_indirect.xlsx!E5` =INDIRECT("R2C2") → Excel "#REF!", HyperFormula "#NAME?" |
| XMATCH | 22 | unsupported 22 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("banana",A1:A3) → Excel 2, HyperFormula "#NAME?" |
| SEQUENCE | 20 | fail 17, unsupported 3 | `dynamic_evaluation/bug_hunt.xlsx!A3` =SEQUENCE(INDIRECT("C1:C3")) → Excel 1, HyperFormula "#NAME?" |
| SUM | 20 | fail 10, unsupported 10 | `dynamic_evaluation/known_gaps.xlsx!A1` =SUM(C1#) → Excel 6, HyperFormula "#ERROR!" |
| LAMBDA | 19 | fail 18, unsupported 1 | `general/fn_lambda.xlsx!A1` =LAMBDA(x,y, x + y)(1, 2) → Excel 3, HyperFormula "#ERROR!" |
| ACCRINTM | 18 | unsupported 18 | `general/fn_accrintm.xlsx!A1` =ACCRINTM(39539, 39614, 0.1, 1000, 3) → Excel 20.54794520547945, HyperFormula "#NAME?" |
| ENCODEURL | 18 | unsupported 18 | `general/fn_encodeurl.xlsx!A3` =ENCODEURL("a b+c&d=e?f/g:h") → Excel "a%20b%2Bc%26d%3De%3Ff%2Fg%3Ah", HyperFormula "#NAME?" |
| IF | 17 | fail 2, unsupported 15 | `dynamic_evaluation/bug_hunt.xlsx!E3` =IF(INDIRECT("C3"),"yes","no") → Excel "yes", HyperFormula "#NAME?" |
| INDEX | 14 | fail 12, unsupported 2 | `lookup_and_reference/fn_index.xlsx!A1` =INDEX({1,2,3,4},2) → Excel 2, HyperFormula "#NUM!" |
| ACCRINT | 14 | unsupported 14 | `general/fn_accrint.xlsx!A1` =ACCRINT(39508, 39691, 39569, 0.1, 1000, 2, 0) → Excel 16.666666666666664, HyperFormula "#NAME?" |
| POWER | 13 | fail 12, unsupported 1 | `general/fn_power.xlsx!A6` =POWER("3", TRUE) → Excel 3, HyperFormula "#NAME?" |
| LOOKUP | 12 | unsupported 12 | `lookup_and_reference/fn_lookup.xlsx!B1` =LOOKUP(20,A1:A3) → Excel 20, HyperFormula "#NAME?" |
| FILTER | 12 | fail 12 | `general/spill_functions.xlsx!C1` =FILTER(A1,B1) → Excel "#CALC!", HyperFormula "#N/A" |
| CONVERT | 10 | unsupported 10 | `engineering/convert.xlsx!A8` =CONVERT(B1,"mi","km") → Excel 9.656064, HyperFormula "#NAME?" |
| DAYS360 | 9 | unsupported 9 | `general/date_and_time.xlsx!A9` =DAYS360(1,2,FALSE) → Excel 1, HyperFormula "#NAME?" |
| LET | 9 | fail 1, unsupported 8 | `general/fn_lambda.xlsx!C9` =LET(x, A1, MAX(x)) → Excel 0, HyperFormula "#NAME?" |
| MAX | 9 | fail 6, unsupported 3 | `general/fn_max_min.xlsx!A1` =MAX({-1,"a",TRUE,-7}) → Excel -1, HyperFormula "#NAME?" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", HyperFormula "#NAME?" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, HyperFormula "#NAME?" |
| DATEVALUE | 8 | fail 8 | `general/datevalue_rollover.xlsx!A3` =DATEVALUE("2026-01-01 23:59:59") → Excel 46023, HyperFormula "#VALUE!" |
| CONCAT | 8 | fail 1, unsupported 7 | `general/fn_concatenate.xlsx!D3` =CONCAT(A1:A3, "!") → Excel "Hello my World!", HyperFormula "#NAME?" |
| PRODUCT | 8 | fail 4, unsupported 4 | `general/fn_product.xlsx!A1` =PRODUCT(B1:D1) → Excel 0, HyperFormula 1 |
| HLOOKUP | 7 | fail 3, unsupported 4 | `lookup_and_reference/fn_hlookup.xlsx!E1` =HLOOKUP("apple",A1:C2,2,FALSE) → Excel 2, HyperFormula "#NAME?" |
| VLOOKUP | 7 | fail 3, unsupported 4 | `lookup_and_reference/fn_vlookup.xlsx!D1` =VLOOKUP("apple",A1:B3,2,FALSE) → Excel 2, HyperFormula "#NAME?" |
| BINOM.DIST.RANGE | 7 | unsupported 7 | `statistical/fn_binom.xlsx!A3` =BINOM.DIST.RANGE(60, 1.2, 45, 50) → Excel "#NUM!", HyperFormula "#NAME?" |
