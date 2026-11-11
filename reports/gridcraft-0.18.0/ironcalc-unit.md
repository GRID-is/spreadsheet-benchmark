# ironcalc-unit vs gridcraft@0.18.0 (ironcalc-unit)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-unstable 5, excel-lacks-function 2, volatile 14, volatile-precedent 1. Fails where the engine's error came from another cell: 1. Fails where both answered an error, but a different one: 32. Fails where the case's own cell matches but its spill does not: 16.
Functions outside Excel's list (1): ECMA.CEILING. Never-supported functions present: none.
Upstream literal differs from Excel's value in 277 of 1767 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| array_formulas | 4 | 0 | 4 | 0 | 0 | 0 | 0.0% |
| dynamic_evaluation | 73 | 51 | 22 | 0 | 0 | 0 | 69.9% |
| engineering | 102 | 95 | 7 | 0 | 0 | 0 | 93.1% |
| general | 1128 | 1009 | 97 | 0 | 0 | 22 | 91.2% |
| logical | 47 | 39 | 8 | 0 | 0 | 0 | 83.0% |
| lookup_and_reference | 132 | 125 | 7 | 0 | 0 | 0 | 94.7% |
| mathematical | 31 | 30 | 1 | 0 | 0 | 0 | 96.8% |
| statistical | 272 | 263 | 9 | 0 | 0 | 0 | 96.7% |
| **total** | 1789 | 1612 | 155 | 0 | 0 | 22 | 91.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1789 | 1612 | 155 | 0 | 0 | 22 | 91.2% |
| **total** | 1789 | 1612 | 155 | 0 | 0 | 22 | 91.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SEQUENCE | 22 | fail 22 | `dynamic_evaluation/bug_hunt.xlsx!A3` =SEQUENCE(INDIRECT("C1:C3")) → Excel 1, GridCraft 1 |
| SUM | 9 | fail 9 | `general/arrays_and_errors.xlsx!A6` =SUM(IFERROR(A1:C1/NA(), 2)) → Excel 2, GridCraft 6 |
| POWER | 9 | fail 9 | `general/fn_power.xlsx!A5` =POWER(-1, 1/7) → Excel -1, GridCraft "#NUM!" |
| RANDARRAY | 9 | skip 9 | `general/fn_randarray.xlsx!A1` =RANDARRAY(1,1,10,5) → Excel "#VALUE!", GridCraft "#VALUE!" |
| IFERROR | 6 | fail 6 | `logical/fn_iferror_array.xlsx!A1` =IFERROR({1,#DIV/0!,3}, {10,20,30}) → Excel 1, GridCraft 1 |
| XMATCH | 5 | fail 5 | `lookup_and_reference/fn_xmatch.xlsx!C1` =XMATCH("[A-Z][a-z]*day",A1:A3,3) → Excel 1, GridCraft "#VALUE!" |
| WEEKNUM | 5 | skip 5 | `general/date_and_time.xlsx!A3` =WEEKNUM(46023, 34903) → Excel "#NUM!", GridCraft "#NUM!" |
| INDIRECT | 4 | fail 4 | `general/fn_indirect.xlsx!E6` =INDIRECT("R2C2",) → Excel 0, GridCraft "#REF!" |
| MAX | 4 | fail 4 | `general/fn_max_min.xlsx!C2` =MAX(@A1:A3) → Excel 0, GridCraft "#VALUE!" |
| CHISQ.TEST | 3 | fail 3 | `statistical/fn_chisq_test.xlsx!G6` =CHISQ.TEST({1,2,3}, {3;3;4}) → Excel 0.38353157287631073, GridCraft "#N/A" |
| BASE | 3 | fail 3 | `general/base_decimal.xlsx!A1` =BASE(-0.5, 2) → Excel "#NUM!", GridCraft "0" |
| ODDFPRICE | 3 | fail 3 | `general/fn_oddfprice.xlsx!C3` =ODDFPRICE(A1,A2,A3,A4,A5,A6,A7,A8,A9) → Excel 113.50846650690974, GridCraft 113.60572709276762 |
| FILTER | 3 | fail 3 | `general/spill_functions.xlsx!C1` =FILTER(A1,B1) → Excel "#CALC!", GridCraft "#VALUE!" |
| TODAY | 3 | skip 3 | `general/today.xlsx!A1` =TODAY() → Excel 46330, GridCraft 46336 |
| IFNA | 2 | fail 2 | `logical/fn_iferror_array.xlsx!A1` =IFNA({1,#N/A,3}, {10,20,30}) → Excel 1, GridCraft 1 |
| DEVSQ | 2 | fail 2 | `statistical/fn_devsq.xlsx!A3` =DEVSQ(1, ) → Excel 0.5, GridCraft 0 |
| DECIMAL | 2 | fail 2 | `general/base_decimal.xlsx!A2` =DECIMAL(REPT("0",255)&"1", 2) → Excel "#VALUE!", GridCraft "#NUM!" |
| TIMEVALUE | 2 | fail 2 | `general/datevalue_timevalue_spaces.xlsx!A1` =TIMEVALUE("4 : 35") → Excel 0.1909722222222222, GridCraft "#VALUE!" |
| ECMA.CEILING | 2 | skip 2 | `general/ecma_ceiling.xlsx!A1` =ECMA.CEILING(4.3,2) → Excel 6, GridCraft 6 |
| CONCATENATE | 2 | fail 2 | `general/fn_concatenate.xlsx!C2` =CONCATENATE(A1:A3, "!") → Excel " my !", GridCraft "Hello!" |
| ENCODEURL | 2 | fail 2 | `general/fn_encodeurl.xlsx!A5` =ENCODEURL("~!*'()") → Excel "%7E%21%2A%27%28%29", GridCraft "~%21%2A%27%28%29" |
| FORMULATEXT | 2 | fail 2 | `general/fn_formulatext.xlsx!B1` =FORMULATEXT(@A:A) → Excel "=1 +  2", GridCraft "#N/A" |
| LET | 2 | fail 2 | `general/fn_lambda.xlsx!A1` =LET(x, 1, y, LAMBDA(a,b, a + b), y(x, 22)) → Excel "#NAME?", GridCraft 23 |
| MIN | 2 | fail 2 | `general/fn_max_min.xlsx!A5` =MIN(1,,2) → Excel 0, GridCraft 1 |
| TYPE | 2 | fail 2 | `general/fn_type.xlsx!A2` =TYPE(A1:C30) → Excel 0, GridCraft 64 |
