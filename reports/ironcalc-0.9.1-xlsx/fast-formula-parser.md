# fast-formula-parser vs ironcalc@0.9.1 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 40. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 134 | 21 | 0 | 0 | 3 | 86.5% |
| engineering | 376 | 320 | 56 | 0 | 0 | 0 | 85.1% |
| information | 58 | 56 | 2 | 0 | 0 | 0 | 96.6% |
| logical | 37 | 31 | 6 | 0 | 0 | 0 | 83.8% |
| math | 272 | 250 | 20 | 0 | 0 | 2 | 92.6% |
| operators | 56 | 50 | 6 | 0 | 0 | 0 | 89.3% |
| reference | 88 | 68 | 20 | 0 | 0 | 0 | 77.3% |
| statistical | 262 | 0 | 0 | 0 | 262 | 0 | 0.0% |
| text | 81 | 62 | 14 | 5 | 0 | 0 | 76.5% |
| trigonometry | 82 | 76 | 6 | 0 | 0 | 0 | 92.7% |
| **total** | 1470 | 1047 | 151 | 5 | 262 | 5 | 71.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1047 | 151 | 5 | 262 | 5 | 71.5% |
| **total** | 1470 | 1047 | 151 | 5 | 262 | 5 | 71.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| HYPGEOM.DIST | 18 | error 18 | `statistical.xlsx!AA139` =HYPGEOM.DIST(1,8,8,7,FALSE) → Excel "#NUM!", IronCalc null |
| COUNTIF | 17 | error 17 | `statistical.xlsx!AA256` =COUNTIF(B1:E1,"TRUE") → Excel 2, IronCalc null |
| SUM | 11 | fail 8, error 3 | `math.xlsx!AA234` =SUM((A1:C1, C1:E1)) → Excel 18, IronCalc "#ERROR!" |
| DATEVALUE | 10 | fail 10 | `date.xlsx!AA45` =DATEVALUE("12/3") → Excel 46359, IronCalc "#VALUE!" |
| NORM.INV | 8 | error 8 | `statistical.xlsx!AA169` =NORM.INV(0.9, 40, -1.5) → Excel "#NUM!", IronCalc null |
| F.INV.RT | 8 | error 8 | `statistical.xlsx!AA64` =F.INV.RT(0.01, 1000000000000, 4) → Excel "#NUM!", IronCalc null |
| VLOOKUP | 7 | fail 7 | `reference.xlsx!AA81` =VLOOKUP("b", {"c";TRUE;"abc";65;63;61;"b";"a";TRUE;FALSE}, 1 → Excel "b", IronCalc "a" |
| HLOOKUP | 7 | fail 7 | `reference.xlsx!AA32` =HLOOKUP(FALSE, {"c",TRUE,"abc",65,63,61,"b","a",TRUE,FALSE}, → Excel "#N/A", IronCalc "a" |
| AVERAGEIF | 7 | error 7 | `statistical.xlsx!AA233` =AVERAGEIF(B8:B11,"<23000") → Excel 14000, IronCalc null |
| WEIBULL.DIST | 7 | error 7 | `statistical.xlsx!AA216` =WEIBULL.DIST(105, -20, 100,FALSE) → Excel "#NUM!", IronCalc null |
| T.INV.2T | 7 | error 7 | `statistical.xlsx!AA206` =T.INV.2T(0.546449, 60.1) → Excel 0.606533075825759, IronCalc null |
| T.INV | 7 | error 7 | `statistical.xlsx!AA199` =T.INV(0.75, 2.1) → Excel 0.8164965809277259, IronCalc null |
| F.INV | 7 | error 7 | `statistical.xlsx!AA56` =F.INV(0.01,  6,  0.4) → Excel "#NUM!", IronCalc null |
| F.DIST | 7 | error 7 | `statistical.xlsx!AA42` =F.DIST(-0.5,6,4,TRUE) → Excel "#NUM!", IronCalc null |
| COUNT | 6 | error 6 | `statistical.xlsx!AA243` =COUNT(A2:E2) → Excel 3, IronCalc null |
| AVERAGEA | 6 | error 6 | `statistical.xlsx!AA227` =AVERAGEA({TRUE,1,2,"12"}, 3) → Excel 1.5, IronCalc null |
| NEGBINOM.DIST | 6 | error 6 | `statistical.xlsx!AA154` =NEGBINOM.DIST(10, 5, -0.25,FALSE) → Excel "#NUM!", IronCalc null |
| F.DIST.RT | 6 | error 6 | `statistical.xlsx!AA50` =F.DIST.RT(15.6, 6, 0.4) → Excel "#NUM!", IronCalc null |
| IMREAL | 5 | fail 5 | `engineering.xlsx!AA285` =IMREAL("+10+i") → Excel 10, IronCalc "#NUM!" |
| IMPRODUCT | 5 | fail 5 | `engineering.xlsx!AA272` =IMPRODUCT("3+4i","5-3j", "5-3j") → Excel "#VALUE!", IronCalc "#ERROR!" |
| T.DIST | 5 | error 5 | `statistical.xlsx!AA188` =T.DIST(60,1,TRUE) → Excel 0.9946953263673767, IronCalc null |
| NORM.S.INV | 5 | error 5 | `statistical.xlsx!AA177` =NORM.S.INV(1.1) → Excel "#NUM!", IronCalc null |
| NORM.DIST | 5 | error 5 | `statistical.xlsx!AA162` =NORM.DIST(42, 40, -1.1, TRUE) → Excel "#NUM!", IronCalc null |
| GAMMA.INV | 5 | error 5 | `statistical.xlsx!AA101` =GAMMA.INV(0.5,9,-0.2) → Excel "#NUM!", IronCalc null |
| GAMMA.DIST | 5 | error 5 | `statistical.xlsx!AA93` =GAMMA.DIST(10.00001131, -9, 2, TRUE) → Excel "#NUM!", IronCalc null |

## Errors

- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
- `statistical.xlsx`: timeout after 55000 ms
