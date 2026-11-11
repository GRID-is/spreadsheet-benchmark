# ironcalc vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, excel-lacks-function 9. Fails where the engine's error came from another cell: 2. Fails where both answered an error, but a different one: 1. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (5): COUNTWORDS, GELU, HELLO, TEMPERATURE, XYZ. Never-supported functions present: none.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| (root) | 708 | 705 | 2 | 1 | 0 | 0 | 99.6% |
| DATABASE | 566 | 566 | 0 | 0 | 0 | 0 | 100.0% |
| DATE_AND_TIME | 2623 | 2620 | 0 | 0 | 0 | 3 | 100.0% |
| ENGINEERING | 2976 | 2976 | 0 | 0 | 0 | 0 | 100.0% |
| FINANCIAL | 3086 | 3086 | 0 | 0 | 0 | 0 | 100.0% |
| INFORMATION | 409 | 407 | 1 | 1 | 0 | 0 | 99.5% |
| LOGICAL | 533 | 526 | 0 | 1 | 0 | 6 | 99.8% |
| LOOKUP_AND_REFERENCE | 1047 | 1040 | 7 | 0 | 0 | 0 | 99.3% |
| MATH_AND_TRIGONOMETRY | 56100 | 56098 | 0 | 0 | 0 | 2 | 100.0% |
| STATISTICAL | 3184 | 3184 | 0 | 0 | 0 | 0 | 100.0% |
| TEXT | 1457 | 1457 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 72689 | 72665 | 10 | 3 | 0 | 11 | 100.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 72689 | 72665 | 10 | 3 | 0 | 11 | 100.0% |
| **total** | 72689 | 72665 | 10 | 3 | 0 | 11 | 100.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| TEMPERATURE | 4 | skip 4 | `LOGICAL/LAMBDA.xlsx!B6` =Temperature(A6) → Excel 37.77777777777778, GRID 37.77777777777778 |
| XMATCH | 3 | fail 3 | `LOOKUP_AND_REFERENCE/XMATCH.xlsx!D6` =XMATCH("[A-Z][a-z]+y",A2:A9,3,1) → Excel 4, GRID "#VALUE!" |
| LAMBDA | 2 | fail 1, unsupported 1 | `INFORMATION/ISOMITTED.xlsx!A6` =LAMBDA(a,_xlop.b,IF(ISOMITTED(b),a/2,a*b))(A5) → Excel 6, GRID "#VALUE!" |
| INDEX | 2 | fail 2 | `LOOKUP_AND_REFERENCE/XMATCH.xlsx!E7` =INDEX(A2:A9,D7) → Excel "Francialy", GRID "#VALUE!" |
| GELU | 2 | skip 2 | `MATH_AND_TRIGONOMETRY/POWER.xlsx!A3` =gelu(A1:C1) → Excel 0.8411919906082768, GRID 0.8411919906082768 |
| NOW | 1 | skip 1 | `DATE_AND_TIME/NOW_TODAY.xlsx!A2` =NOW() → Excel 46330.771844097224, GRID 46335.065448622685 |
| TODAY | 1 | skip 1 | `DATE_AND_TIME/NOW_TODAY.xlsx!B2` =TODAY() → Excel 46330, GRID 46335 |
| HELLO | 1 | skip 1 | `DATE_AND_TIME/TIME_HOUR_MINUTE_SECOND.xlsx!C21` =hello() → Excel "#NAME?", GRID "#NAME?" |
| XYZ | 1 | skip 1 | `LOGICAL/IFERROR.xlsx!E20` =XYZ(1) → Excel "#NAME?", GRID "#NAME?" |
| LET | 1 | unsupported 1 | `LOGICAL/LAMBDA.xlsx!F4` =LET(area, LAMBDA(x,y,x*y), _xlpm.area(D1,E1)) → Excel 42, GRID "#NAME?" |
| COUNTWORDS | 1 | skip 1 | `LOGICAL/LAMBDA.xlsx!F2` =CountWords(F1) → Excel 9, GRID 9 |
| COLUMN | 1 | fail 1 | `LOOKUP_AND_REFERENCE/ROW_COLUM.xlsx!L7` =COLUMN(INDIRECT(J7)) → Excel "#REF!", GRID 23 |
| ROW | 1 | fail 1 | `LOOKUP_AND_REFERENCE/ROW_COLUM.xlsx!K7` =ROW(INDIRECT(J7)) → Excel "#REF!", GRID 3 |
| IFERROR | 1 | fail 1 | `named_ranges.xlsx!G60` =IFERROR(IF(E60=F60,"PASS","FAIL"),"FAIL") → Excel "PASS", GRID "FAIL" |
| COUNTIF | 1 | fail 1 | `named_ranges.xlsx!B2` =COUNTIF(G5:G61,"PASS")&" / "&ROWS(G5:G61) → Excel "57 / 57", GRID "56 / 57" |
