# fast-formula-parser vs poi@6.0.0 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 13. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 92 | 33 | 30 | 0 | 3 | 59.4% |
| engineering | 376 | 70 | 33 | 273 | 0 | 0 | 18.6% |
| information | 58 | 43 | 3 | 12 | 0 | 0 | 74.1% |
| logical | 37 | 22 | 6 | 9 | 0 | 0 | 59.5% |
| math | 272 | 206 | 22 | 42 | 0 | 2 | 76.3% |
| operators | 56 | 51 | 5 | 0 | 0 | 0 | 91.1% |
| reference | 88 | 74 | 14 | 0 | 0 | 0 | 84.1% |
| statistical | 262 | 85 | 23 | 154 | 0 | 0 | 32.4% |
| text | 81 | 55 | 11 | 15 | 0 | 0 | 67.9% |
| trigonometry | 82 | 49 | 5 | 28 | 0 | 0 | 59.8% |
| **total** | 1470 | 747 | 155 | 563 | 0 | 5 | 51.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 747 | 155 | 563 | 0 | 5 | 51.0% |
| **total** | 1470 | 747 | 155 | 563 | 0 | 5 | 51.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| OCT2HEX | 26 | unsupported 26 | `engineering.xlsx!AA353` =OCT2HEX(100, 4) → Excel "0040", Apache POI "#NAME?" |
| DATEDIF | 20 | unsupported 20 | `date.xlsx!AA26` =DATEDIF("9/15/2001","8/15/2003","YD") → Excel 334, Apache POI "#NAME?" |
| OCT2BIN | 18 | unsupported 18 | `engineering.xlsx!AA337` =OCT2BIN(7771) → Excel "#NUM!", Apache POI "#NAME?" |
| HYPGEOM.DIST | 18 | unsupported 18 | `statistical.xlsx!AA139` =HYPGEOM.DIST(1,8,8,7,FALSE) → Excel "#NUM!", Apache POI "#NAME?" |
| BITAND | 16 | unsupported 16 | `engineering.xlsx!AA73` =BITAND(1,FALSE) → Excel 0, Apache POI "#NAME?" |
| BIN2HEX | 16 | unsupported 16 | `engineering.xlsx!AA35` =BIN2HEX(11111011) → Excel "FB", Apache POI "#NAME?" |
| IMREAL | 13 | fail 13 | `engineering.xlsx!AA281` =IMREAL("6-9i") → Excel 6, Apache POI "6" |
| BASE | 13 | unsupported 13 | `math.xlsx!AA15` =BASE(7,2,0) → Excel "111", Apache POI "#NAME?" |
| BITRSHIFT | 12 | unsupported 12 | `engineering.xlsx!AA105` =BITRSHIFT(9999999999,31) → Excel 4, Apache POI "#NAME?" |
| DATEVALUE | 11 | fail 11 | `date.xlsx!AA51` =DATEVALUE("12/3/14 12:00 AM") → Excel 41976, Apache POI "#VALUE!" |
| IMARGUMENT | 11 | unsupported 11 | `engineering.xlsx!AA221` =IMARGUMENT("52") → Excel 0, Apache POI "#NAME?" |
| BIN2OCT | 11 | unsupported 11 | `engineering.xlsx!AA55` =BIN2OCT(1001,"-7") → Excel "#NUM!", Apache POI "#NAME?" |
| IMAGINARY | 10 | fail 10 | `engineering.xlsx!AA209` =IMAGINARY("3+i") → Excel 1, Apache POI "1" |
| BITOR | 10 | unsupported 10 | `engineering.xlsx!AA97` =BITOR(60,FALSE) → Excel 60, Apache POI "#NAME?" |
| BITLSHIFT | 10 | unsupported 10 | `engineering.xlsx!AA82` =BITLSHIFT(9,-3) → Excel 1, Apache POI "#NAME?" |
| SUM | 10 | fail 7, unsupported 3 | `math.xlsx!AA234` =SUM((A1:C1, C1:E1)) → Excel 18, Apache POI "#VALUE!" |
| XOR | 9 | unsupported 9 | `logical.xlsx!AA36` =XOR("Test", "TRUE") → Excel true, Apache POI "#NAME?" |
| IMDIV | 8 | unsupported 8 | `engineering.xlsx!AA254` =IMDIV("52i+10",1) → Excel "#NUM!", Apache POI "#NAME?" |
| BITXOR | 8 | unsupported 8 | `engineering.xlsx!AA111` =BITXOR(24,3) → Excel 27, Apache POI "#NAME?" |
| F.INV.RT | 8 | unsupported 8 | `statistical.xlsx!AA64` =F.INV.RT(0.01, 1000000000000, 4) → Excel "#NUM!", Apache POI "#NAME?" |
| NETWORKDAYS.INTL | 7 | unsupported 7 | `date.xlsx!AA101` =NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31), "1") → Excel "#VALUE!", Apache POI "#NAME?" |
| IMPRODUCT | 7 | unsupported 7 | `engineering.xlsx!AA271` =IMPRODUCT("3+4i","5-3i") → Excel "27+11i", Apache POI "#NAME?" |
| GESTEP | 7 | unsupported 7 | `engineering.xlsx!AA169` =GESTEP(5,5) → Excel 1, Apache POI "#NAME?" |
| DEC2OCT | 7 | unsupported 7 | `engineering.xlsx!AA149` =DEC2OCT(-536874912,3) → Excel "#NUM!", Apache POI "#NAME?" |
| BESSELY | 7 | unsupported 7 | `engineering.xlsx!AA22` =BESSELY(2.5, 1.9) → Excel 0.14591813750831284, Apache POI "#NAME?" |
