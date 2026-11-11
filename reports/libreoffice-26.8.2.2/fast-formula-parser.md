# fast-formula-parser vs libreoffice@26.8.2.2 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 213. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 124 | 31 | 0 | 0 | 3 | 80.0% |
| engineering | 376 | 209 | 142 | 25 | 0 | 0 | 55.6% |
| information | 58 | 54 | 4 | 0 | 0 | 0 | 93.1% |
| logical | 37 | 31 | 6 | 0 | 0 | 0 | 83.8% |
| math | 272 | 228 | 42 | 0 | 0 | 2 | 84.4% |
| operators | 56 | 45 | 9 | 2 | 0 | 0 | 80.4% |
| reference | 88 | 58 | 29 | 1 | 0 | 0 | 65.9% |
| statistical | 262 | 143 | 117 | 2 | 0 | 0 | 54.6% |
| text | 81 | 65 | 11 | 5 | 0 | 0 | 80.2% |
| trigonometry | 82 | 69 | 13 | 0 | 0 | 0 | 84.1% |
| **total** | 1470 | 1026 | 404 | 35 | 0 | 5 | 70.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1026 | 404 | 35 | 0 | 5 | 70.0% |
| **total** | 1470 | 1026 | 404 | 35 | 0 | 5 | 70.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| YEARFRAC | 12 | fail 12 | `date.xlsx!AA149` =YEARFRAC("2020/2/1","2022/1/1", 1) → Excel 1.916058394160584, LibreOffice "#VALUE!" |
| OCT2HEX | 12 | fail 12 | `engineering.xlsx!AA371` =OCT2HEX(100, 0) → Excel "#NUM!", LibreOffice "#VALUE!" |
| VLOOKUP | 12 | fail 12 | `reference.xlsx!AA91` =VLOOKUP("*e", {"abc";"sd";"qwe"}, 1,FALSE) → Excel "qwe", LibreOffice "#N/A" |
| HLOOKUP | 12 | fail 12 | `reference.xlsx!AA43` =HLOOKUP("s?", {"abc","sd","qwe"}, 1,FALSE) → Excel "sd", LibreOffice "#N/A" |
| HYPGEOM.DIST | 12 | fail 12 | `statistical.xlsx!AA139` =HYPGEOM.DIST(1,8,8,7,FALSE) → Excel "#NUM!", LibreOffice "#VALUE!" |
| BITAND | 8 | fail 8 | `engineering.xlsx!AA77` =BITAND("TRUE",7) → Excel "#VALUE!", LibreOffice 1 |
| OCT2BIN | 7 | fail 7 | `engineering.xlsx!AA337` =OCT2BIN(7771) → Excel "#NUM!", LibreOffice "#VALUE!" |
| DATEVALUE | 6 | fail 6 | `date.xlsx!AA46` =DATEVALUE("3-Dec") → Excel 46359, LibreOffice "#VALUE!" |
| BITXOR | 6 | fail 6 | `engineering.xlsx!AA112` =BITXOR(5.5,3) → Excel "#NUM!", LibreOffice 6 |
| BITOR | 6 | fail 6 | `engineering.xlsx!AA91` =BITOR(3,-1) → Excel "#NUM!", LibreOffice "#VALUE!" |
| BIN2HEX | 6 | fail 6 | `engineering.xlsx!AA38` =BIN2HEX("11111011",-4) → Excel "#NUM!", LibreOffice "#VALUE!" |
| NORM.INV | 6 | fail 6 | `statistical.xlsx!AA169` =NORM.INV(0.9, 40, -1.5) → Excel "#NUM!", LibreOffice "#VALUE!" |
| F.INV.RT | 6 | fail 6 | `statistical.xlsx!AA64` =F.INV.RT(0.01, 1000000000000, 4) → Excel "#NUM!", LibreOffice "#VALUE!" |
| OCT2DEC | 5 | fail 5 | `engineering.xlsx!AA346` =OCT2DEC(12345671234) → Excel "#NUM!", LibreOffice "#VALUE!" |
| IMREAL | 5 | fail 5 | `engineering.xlsx!AA284` =IMREAL("+10+2y") → Excel "#NUM!", LibreOffice "#VALUE!" |
| IMDIV | 5 | fail 5 | `engineering.xlsx!AA254` =IMDIV("52i+10",1) → Excel "#NUM!", LibreOffice "#VALUE!" |
| IMCOSH | 5 | unsupported 5 | `engineering.xlsx!AA236` =IMCOSH(FALSE) → Excel "#VALUE!", LibreOffice "#NAME?" |
| BITLSHIFT | 5 | fail 5 | `engineering.xlsx!AA83` =BITLSHIFT(4.1,2) → Excel "#NUM!", LibreOffice 16 |
| BIN2OCT | 5 | fail 5 | `engineering.xlsx!AA55` =BIN2OCT(1001,"-7") → Excel "#NUM!", LibreOffice "#VALUE!" |
| BESSELJ | 5 | fail 5 | `engineering.xlsx!AA13` =BESSELJ(1.9, {2,3}) → Excel 0.3299258286697852, LibreOffice 0.329925727692387 |
| SUM | 5 | fail 4, unsupported 1 | `math.xlsx!AA239` =SUM("1", {"1"},) → Excel 1, LibreOffice "#VALUE!" |
| BASE | 5 | fail 5 | `math.xlsx!AA9` =BASE(-1,2) → Excel "#NUM!", LibreOffice "#VALUE!" |
| AVERAGEA | 5 | fail 5 | `statistical.xlsx!AA227` =AVERAGEA({TRUE,1,2,"12"}, 3) → Excel 1.5, LibreOffice 1.4 |
| WEIBULL.DIST | 5 | fail 5 | `statistical.xlsx!AA216` =WEIBULL.DIST(105, -20, 100,FALSE) → Excel "#NUM!", LibreOffice "#VALUE!" |
| F.INV | 5 | fail 5 | `statistical.xlsx!AA56` =F.INV(0.01,  6,  0.4) → Excel "#NUM!", LibreOffice "#VALUE!" |
