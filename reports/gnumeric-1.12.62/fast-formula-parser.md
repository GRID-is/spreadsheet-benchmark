# fast-formula-parser vs gnumeric@1.12.62 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 23. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 128 | 27 | 0 | 0 | 3 | 82.6% |
| engineering | 376 | 290 | 86 | 0 | 0 | 0 | 77.1% |
| information | 58 | 54 | 4 | 0 | 0 | 0 | 93.1% |
| logical | 37 | 32 | 5 | 0 | 0 | 0 | 86.5% |
| math | 272 | 199 | 33 | 38 | 0 | 2 | 73.7% |
| operators | 56 | 55 | 1 | 0 | 0 | 0 | 98.2% |
| reference | 88 | 84 | 4 | 0 | 0 | 0 | 95.5% |
| statistical | 262 | 214 | 41 | 7 | 0 | 0 | 81.7% |
| text | 81 | 63 | 13 | 5 | 0 | 0 | 77.8% |
| trigonometry | 82 | 72 | 10 | 0 | 0 | 0 | 87.8% |
| **total** | 1470 | 1191 | 224 | 50 | 0 | 5 | 81.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1191 | 224 | 50 | 0 | 5 | 81.3% |
| **total** | 1470 | 1191 | 224 | 50 | 0 | 5 | 81.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| DATEDIF | 19 | fail 19 | `date.xlsx!AA26` =DATEDIF("9/15/2001","8/15/2003","YD") → Excel 334, Gnumeric "#VALUE!" |
| FLOOR.MATH | 14 | unsupported 14 | `math.xlsx!AA92` =FLOOR.MATH(5.5,-2) → Excel 4, Gnumeric "#NAME?" |
| BITAND | 7 | fail 7 | `engineering.xlsx!AA75` =BITAND(1, 5.2) → Excel "#NUM!", Gnumeric 1 |
| FLOOR.PRECISE | 7 | unsupported 7 | `math.xlsx!AA96` =FLOOR.PRECISE(3.2, 1) → Excel 3, Gnumeric "#NAME?" |
| BITXOR | 6 | fail 6 | `engineering.xlsx!AA112` =BITXOR(5.5,3) → Excel "#NUM!", Gnumeric 6 |
| BITOR | 6 | fail 6 | `engineering.xlsx!AA91` =BITOR(3,-1) → Excel "#NUM!", Gnumeric "#VALUE!" |
| ISO.CEILING | 6 | unsupported 6 | `math.xlsx!AA125` =ISO.CEILING(4.3,-2) → Excel 6, Gnumeric "#NAME?" |
| CEILING.PRECISE | 6 | unsupported 6 | `math.xlsx!AA33` =CEILING.PRECISE(4.3,-2) → Excel 6, Gnumeric "#NAME?" |
| BESSELJ | 5 | fail 5 | `engineering.xlsx!AA13` =BESSELJ(1.9, {2,3}) → Excel 0.3299258286697852, Gnumeric 0.32992572769238726 |
| CEILING.MATH | 5 | unsupported 5 | `math.xlsx!AA26` =CEILING.MATH(6.7) → Excel 7, Gnumeric "#NAME?" |
| BASE | 5 | fail 5 | `math.xlsx!AA15` =BASE(7,2,0) → Excel "111", Gnumeric "#NUM!" |
| F.DIST | 5 | fail 5 | `statistical.xlsx!AA42` =F.DIST(-0.5,6,4,TRUE) → Excel "#NUM!", Gnumeric 0 |
| NUMBERVALUE | 5 | fail 5 | `text.xlsx!AA56` =NUMBERVALUE("3.5%",,) → Excel "#VALUE!", Gnumeric 0.035 |
| BITRSHIFT | 4 | fail 4 | `engineering.xlsx!AA108` =BITRSHIFT(9999999999,-50) → Excel 0, Gnumeric 10375167641554780000 |
| BITLSHIFT | 4 | fail 4 | `engineering.xlsx!AA83` =BITLSHIFT(4.1,2) → Excel "#NUM!", Gnumeric 16 |
| CONCATENATE | 4 | fail 4 | `text.xlsx!AA20` =CONCATENATE({0,2,3}, 1, "A", TRUE, -12) → Excel "01ATRUE-12", Gnumeric "0231ATRUE-12" |
| IMCOSH | 3 | fail 3 | `engineering.xlsx!AA237` =IMCOSH("4+3j") → Excel "-27.0349456030742+3.85115333481178j", Gnumeric "-27.034945603074224+3.8511533348117775j" |
| IMCOS | 3 | fail 3 | `engineering.xlsx!AA229` =IMCOS(1) → Excel "0.54030230586814", Gnumeric "0.5403023058681398" |
| DELTA | 3 | fail 3 | `engineering.xlsx!AA160` =DELTA(5,FALSE) → Excel "#VALUE!", Gnumeric 0 |
| BESSELY | 3 | fail 3 | `engineering.xlsx!AA22` =BESSELY(2.5, 1.9) → Excel 0.14591813750831284, Gnumeric -0.33778084608575304 |
| BESSELK | 3 | fail 3 | `engineering.xlsx!AA17` =BESSELK(TRUE, 1) → Excel "#VALUE!", Gnumeric 0.6019072301972346 |
| BESSELI | 3 | fail 3 | `engineering.xlsx!AA3` =BESSELI(1.5, 1.9) → Excel 0.981666428475166, Gnumeric 0.3828606192896613 |
| ISREF | 3 | fail 3 | `information.xlsx!AA41` =ISREF(XYZ1) → Excel false, Gnumeric "#\"ISREF(XYZ1)\"" |
| SUM | 3 | fail 3 | `math.xlsx!AA239` =SUM("1", {"1"},) → Excel 1, Gnumeric 2 |
| LCM | 3 | fail 3 | `math.xlsx!AA137` =LCM(5, 2, (A1, A2)) → Excel "#VALUE!", Gnumeric 100000 |
