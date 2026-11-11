# fast-formula-parser vs univer@1.4.0 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 4. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 144 | 11 | 0 | 0 | 3 | 92.9% |
| engineering | 376 | 359 | 17 | 0 | 0 | 0 | 95.5% |
| information | 58 | 55 | 3 | 0 | 0 | 0 | 94.8% |
| logical | 37 | 25 | 12 | 0 | 0 | 0 | 67.6% |
| math | 272 | 250 | 14 | 6 | 0 | 2 | 92.6% |
| operators | 56 | 48 | 5 | 3 | 0 | 0 | 85.7% |
| reference | 88 | 80 | 8 | 0 | 0 | 0 | 90.9% |
| statistical | 262 | 231 | 31 | 0 | 0 | 0 | 88.2% |
| text | 81 | 70 | 11 | 0 | 0 | 0 | 86.4% |
| trigonometry | 82 | 79 | 3 | 0 | 0 | 0 | 96.3% |
| **total** | 1470 | 1341 | 115 | 9 | 0 | 5 | 91.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1341 | 115 | 9 | 0 | 5 | 91.5% |
| **total** | 1470 | 1341 | 115 | 9 | 0 | 5 | 91.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| SUM | 7 | fail 6, unsupported 1 | `math.xlsx!AA234` =SUM((A1:C1, C1:E1)) → Excel 18, Univer "#VALUE!" |
| ISO.CEILING | 6 | unsupported 6 | `math.xlsx!AA125` =ISO.CEILING(4.3,-2) → Excel 6, Univer "#NAME?" |
| COMPLEX | 5 | fail 5 | `engineering.xlsx!AA129` =COMPLEX(70,0,"j") → Excel "70", Univer 70 |
| HYPGEOM.DIST | 5 | fail 5 | `statistical.xlsx!AA129` =HYPGEOM.DIST(1,4,0,20,TRUE) → Excel 1, Univer "#NUM!" |
| YEAR | 4 | fail 2, skip 2 | `date.xlsx!AA103` =YEAR(NOW()) → Excel 2026, Univer 2026 |
| DATEVALUE | 4 | fail 4 | `date.xlsx!AA49` =DATEVALUE("Dec-3 11:11") → Excel 46359, Univer "#VALUE!" |
| XOR | 4 | fail 4 | `logical.xlsx!AA36` =XOR("Test", "TRUE") → Excel true, Univer "#VALUE!" |
| OR | 4 | fail 4 | `logical.xlsx!AA29` =OR((A2:C2, A3)) → Excel true, Univer "#VALUE!" |
| AND | 4 | fail 4 | `logical.xlsx!AA8` =AND((A2:C2, A3)) → Excel true, Univer "#VALUE!" |
| FREQUENCY | 4 | fail 4 | `statistical.xlsx!AA84` =FREQUENCY({1,2,3,4,5},{1,5,11}) → Excel 1, Univer "#SPILL!" |
| MMULT | 3 | fail 3 | `math.xlsx!AA156` =MMULT({1,3;7,2}, {2,0;0,2}) → Excel 2, Univer "#SPILL!" |
| COUNT | 3 | fail 3 | `statistical.xlsx!AA243` =COUNT(A2:E2) → Excel 3, Univer 4 |
| AVERAGEA | 3 | fail 3 | `statistical.xlsx!AA227` =AVERAGEA({TRUE,1,2,"12"}, 3) → Excel 1.5, Univer 3.8 |
| AVEDEV | 3 | fail 3 | `statistical.xlsx!AA220` =AVEDEV({4,1,6,7,5,4,3},(A7,B7),"1") → Excel 1.8, Univer "#VALUE!" |
| YEARFRAC | 2 | fail 2 | `date.xlsx!AA147` =YEARFRAC("2020/2/5","2021/2/1", 1) → Excel 0.9890710382513661, Univer 0.990424076607 |
| IMCONJUGATE | 2 | fail 2 | `engineering.xlsx!AA228` =IMCONJUGATE("3") → Excel "3", Univer 3 |
| IMCOS | 2 | fail 2 | `engineering.xlsx!AA229` =IMCOS(1) → Excel "0.54030230586814", Univer 0.540302305868 |
| ISREF | 2 | fail 2 | `information.xlsx!AA42` =ISREF(A1:XYZ1) → Excel false, Univer true |
| VLOOKUP | 2 | fail 2 | `reference.xlsx!AA77` =VLOOKUP(TRUE, {"c";FALSE;"abc";65;63;61;"b";"a";FALSE;TRUE}, → Excel true, Univer false |
| AREAS | 2 | fail 2 | `reference.xlsx!AA11` =AREAS((B2:D4,E5,F6:I9)) → Excel 3, Univer "#VALUE!" |
| COUNTIF | 2 | fail 2 | `statistical.xlsx!AA260` =COUNTIF(C2:E2, "{1,2}") → Excel 1, Univer 0 |
| NUMBERVALUE | 2 | fail 2 | `text.xlsx!AA50` =NUMBERVALUE("$3 50") → Excel 350, Univer "#VALUE!" |
| SEARCH | 2 | fail 2 | `text.xlsx!AA68` =SEARCH("c?e", "abcdef") → Excel 3, Univer "#VALUE!" |
| DBCS | 2 | fail 2 | `text.xlsx!AA23` =DBCS(",｡") → Excel ",｡", Univer "，｡" |
| DOLLAR | 2 | fail 2 | `text.xlsx!AA25` =DOLLAR(12345.67) → Excel "$12,345.67", Univer "$12,345.67 " |
