# ironcalc-unit vs libreoffice@26.8.2.2 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 2, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 197. Fails where the case's own cell matches but its spill does not: 16.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1770 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 46 | 24 | 3 | 0 | 0 | 63.0% |
| engineering | 102 | 79 | 20 | 3 | 0 | 0 | 77.5% |
| general | 1128 | 803 | 266 | 40 | 0 | 19 | 72.4% |
| logical | 47 | 30 | 17 | 0 | 0 | 0 | 63.8% |
| lookup_and_reference | 132 | 112 | 11 | 9 | 0 | 0 | 84.8% |
| mathematical | 31 | 26 | 5 | 0 | 0 | 0 | 83.9% |
| statistical | 272 | 170 | 95 | 7 | 0 | 0 | 62.5% |
| **total** | 1789 | 1266 | 442 | 62 | 0 | 19 | 71.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1266 | 442 | 62 | 0 | 19 | 71.5% |
| **total** | 1789 | 1266 | 442 | 62 | 0 | 19 | 71.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 18 | fail 17, unsupported 1 | `dynamic_evaluation/bug_hunt.xlsx!A3` =SEQUENCE(INDIRECT("C1:C3")) → Excel 1, LibreOffice 1 |
| WEEKNUM | 18 | fail 16, skip 2 | `general/date_and_time.xlsx!A14` =WEEKNUM(44561,0) → Excel "#NUM!", LibreOffice "#VALUE!" |
| LAMBDA | 18 | fail 17, unsupported 1 | `general/fn_lambda.xlsx!A1` =LAMBDA(x,y, x + y)(1, 2) → Excel 3, LibreOffice "#VALUE!" |
| ACCRINT | 13 | fail 13 | `general/fn_accrint.xlsx!A2` =ACCRINT(DATE(2008, 3, 5), 39691, 39569, 0.1, 1000, 2, 0, FAL → Excel 15.555555555555555, LibreOffice "#VALUE!" |
| ENCODEURL | 11 | fail 11 | `general/fn_encodeurl.xlsx!A4` =ENCODEURL("") → Excel "", LibreOffice "#VALUE!" |
| SUM | 10 | fail 8, unsupported 2 | `dynamic_evaluation/known_gaps.xlsx!A1` =SUM(C1#) → Excel 6, LibreOffice "#NAME?" |
| IFERROR | 9 | fail 9 | `logical/fn_iferror_array.xlsx!A1` =IFERROR({1,#DIV/0!,3}, "x") → Excel 1, LibreOffice "#N/A" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", LibreOffice "#VALUE!" |
| TRIMRANGE | 8 | unsupported 8 | `lookup_and_reference/fn_trimrange.xlsx!C1` =TRIMRANGE(A1:A4) → Excel 10, LibreOffice "#NAME?" |
| CHISQ.TEST | 8 | fail 8 | `statistical/fn_chisq_test.xlsx!C1` =CHISQ.TEST(A2:A10, B2:B10) → Excel 0.9971295382255398, LibreOffice "#VALUE!" |
| ACCRINTM | 8 | fail 8 | `general/fn_accrintm.xlsx!A1` =ACCRINTM(39614, 39539, 0.1, 1000, 3) → Excel "#NUM!", LibreOffice "#VALUE!" |
| BINOM.DIST.RANGE | 7 | unsupported 7 | `statistical/fn_binom.xlsx!A3` =BINOM.DIST.RANGE(60, 1.2, 45, 50) → Excel "#NUM!", LibreOffice "#NAME?" |
| DECIMAL | 7 | fail 7 | `general/base_decimal.xlsx!A4` =DECIMAL("", 99) → Excel "#NUM!", LibreOffice "#VALUE!" |
| EXACT | 7 | fail 7 | `general/fn_exact.xlsx!B4` =EXACT(TRUE, "TRUE") → Excel true, LibreOffice false |
| BASE | 6 | fail 6 | `general/base_decimal.xlsx!A3` =BASE(5, 2, 256) → Excel "#NUM!", LibreOffice "0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000101" |
| DAY | 6 | fail 6 | `general/date_and_time.xlsx!A1` =DAY(TRUE) → Excel 1, LibreOffice 31 |
| WEEKDAY | 6 | fail 6 | `general/date_and_time.xlsx!A8` =WEEKDAY(44561,0) → Excel "#NUM!", LibreOffice "#VALUE!" |
| MAX | 6 | fail 4, unsupported 2 | `general/fn_max_min.xlsx!C1` =MAX(A1:A6) → Excel -2, LibreOffice 1 |
| NOT | 5 | fail 5 | `logical/fn_not_array.xlsx!A1` =NOT({TRUE;FALSE}) → Excel false, LibreOffice false |
| CHISQ.INV | 5 | fail 5 | `statistical/fn_chisq.xlsx!A6` =CHISQ.INV(1.1, 4) → Excel "#NUM!", LibreOffice "#VALUE!" |
| CHISQ.INV.RT | 5 | fail 5 | `statistical/fn_chisq.xlsx!A6` =CHISQ.INV.RT(1.1, 4) → Excel "#NUM!", LibreOffice "#VALUE!" |
| POISSON.DIST | 5 | fail 5 | `statistical/fn_poisson.xlsx!A6` =POISSON.DIST(3, -2, TRUE) → Excel "#NUM!", LibreOffice "#VALUE!" |
| WEIBULL.DIST | 5 | fail 5 | `statistical/fn_weibull.xlsx!A6` =WEIBULL.DIST(1, 0, 1, TRUE) → Excel "#NUM!", LibreOffice "#VALUE!" |
| DATEVALUE | 5 | fail 5 | `general/datevalue_rollover.xlsx!A4` =DATEVALUE("2026-01-01T24:00:00") → Excel "#VALUE!", LibreOffice 46024 |
| BYCOL | 5 | unsupported 5 | `general/fn_bycol_byrow.xlsx!A4` =BYCOL(A1:C2, LAMBDA(col, SUM(col))) → Excel 5, LibreOffice "#NAME?" |
