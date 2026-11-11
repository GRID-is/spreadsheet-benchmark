# ironcalc-unit vs poi@6.0.0 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 18. Fails where the case's own cell matches but its spill does not: 7.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 18 | 1 | 54 | 0 | 0 | 24.7% |
| engineering | 102 | 28 | 6 | 68 | 0 | 0 | 27.5% |
| general | 1128 | 682 | 155 | 272 | 0 | 19 | 61.5% |
| logical | 47 | 44 | 3 | 0 | 0 | 0 | 93.6% |
| lookup_and_reference | 132 | 84 | 17 | 31 | 0 | 0 | 63.6% |
| mathematical | 31 | 14 | 0 | 17 | 0 | 0 | 45.2% |
| statistical | 272 | 62 | 12 | 198 | 0 | 0 | 22.8% |
| **total** | 1789 | 932 | 198 | 640 | 0 | 19 | 52.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 932 | 198 | 640 | 0 | 19 | 52.7% |
| **total** | 1789 | 932 | 198 | 640 | 0 | 19 | 52.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 44 | unsupported 44 | `dynamic_evaluation/bug_hunt.xlsx!A1` =SEQUENCE(COUNTBLANK(C2:C3)+1) → Excel 1, Apache POI "#NAME?" |
| ACCRINTM | 18 | unsupported 18 | `general/fn_accrintm.xlsx!A1` =ACCRINTM(39539, 39614, 0.1, 1000, 3) → Excel 20.54794520547945, Apache POI "#NAME?" |
| ENCODEURL | 18 | unsupported 18 | `general/fn_encodeurl.xlsx!A3` =ENCODEURL("a b+c&d=e?f/g:h") → Excel "a%20b%2Bc%26d%3De%3Ff%2Fg%3Ah", Apache POI "#NAME?" |
| LAMBDA | 18 | fail 18 | `general/fn_lambda.xlsx!A1` =LAMBDA(x,y, x + y)(1, 2) → Excel 3, Apache POI "#VALUE!" |
| FILTER | 15 | unsupported 15 | `general/spill_functions.xlsx!C1` =FILTER(A1:A3,B1:B3) → Excel "Alice", Apache POI "#NAME?" |
| CHISQ.TEST | 14 | unsupported 14 | `statistical/fn_chisq_test.xlsx!C1` =CHISQ.TEST(A2:A10, B2:B10) → Excel 0.9971295382255398, Apache POI "#NAME?" |
| DECIMAL | 14 | unsupported 14 | `general/base_decimal.xlsx!A1` =DECIMAL("", 2) → Excel 0, Apache POI "#NAME?" |
| ACCRINT | 14 | unsupported 14 | `general/fn_accrint.xlsx!A1` =ACCRINT(39508, 39691, 39569, 0.1, 1000, 2, 0) → Excel 16.666666666666664, Apache POI "#NAME?" |
| XOR | 14 | unsupported 14 | `general/fn_or_xor.xlsx!A11` =XOR(1,0,0) → Excel true, Apache POI "#NAME?" |
| SUM | 12 | fail 7, unsupported 5 | `dynamic_evaluation/known_gaps.xlsx!B1` =SUM(OFFSET(D1,1,0,2,1)) → Excel 5, Apache POI "#NAME?" |
| GAUSS | 11 | unsupported 11 | `statistical/fn_gauss.xlsx!A6` =GAUSS(1) → Excel 0.34134474606854304, Apache POI "#NAME?" |
| BASE | 11 | unsupported 11 | `general/base_decimal.xlsx!A3` =BASE(5, 2, 256) → Excel "#NUM!", Apache POI "#NAME?" |
| LET | 11 | fail 11 | `general/fn_lambda.xlsx!A1` =LET(x, 1, y, LAMBDA(a,b, a + b), y(x, 22)) → Excel "#NAME?", Apache POI "#VALUE!" |
| CONVERT | 10 | unsupported 10 | `engineering/convert.xlsx!A8` =CONVERT(B1,"mi","km") → Excel 9.656064, Apache POI "#NAME?" |
| ISOWEEKNUM | 10 | unsupported 10 | `general/date_and_time.xlsx!A5` =ISOWEEKNUM(44564) → Excel 1, Apache POI "#NAME?" |
| DATEDIF | 10 | unsupported 10 | `general/datedif_leap_month_end.xlsx!A1` =DATEDIF("29/2/2020", "28/2/2021", "YD") → Excel "#VALUE!", Apache POI "#NAME?" |
| POWER | 10 | fail 10 | `general/fn_power.xlsx!A5` =POWER(-1, 1/7) → Excel -1, Apache POI "#NUM!" |
| CHISQ.DIST | 9 | unsupported 9 | `statistical/fn_chisq.xlsx!A6` =CHISQ.DIST(-1, 4, TRUE) → Excel "#NUM!", Apache POI "#NAME?" |
| CHISQ.INV | 9 | unsupported 9 | `statistical/fn_chisq.xlsx!A6` =CHISQ.INV(1.1, 4) → Excel "#NUM!", Apache POI "#NAME?" |
| CHISQ.INV.RT | 9 | unsupported 9 | `statistical/fn_chisq.xlsx!A6` =CHISQ.INV.RT(1.1, 4) → Excel "#NUM!", Apache POI "#NAME?" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", Apache POI "#NAME?" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, Apache POI "#NAME?" |
| BETA.DIST | 8 | unsupported 8 | `statistical/fn_beta_dist.xlsx!A1` =BETA.DIST(0.234, 2, 2.5, , 0.15, 1.2) → Excel 0.5882886668596937, Apache POI "#NAME?" |
| CHISQ.DIST.RT | 8 | unsupported 8 | `statistical/fn_chisq.xlsx!A5` =CHISQ.DIST.RT(-1, 4) → Excel "#NUM!", Apache POI "#NAME?" |
| DB | 8 | unsupported 8 | `general/fn_financial.xlsx!B4` =DB(A2,A3,A4,4,7) → Excel 120410.63551274998, Apache POI "#NAME?" |
