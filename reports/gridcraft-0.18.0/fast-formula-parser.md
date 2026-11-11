# fast-formula-parser vs gridcraft@0.18.0 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 3. Fails where the case's own cell matches but its spill does not: 13.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 151 | 4 | 0 | 0 | 3 | 97.4% |
| engineering | 376 | 346 | 30 | 0 | 0 | 0 | 92.0% |
| information | 58 | 58 | 0 | 0 | 0 | 0 | 100.0% |
| logical | 37 | 32 | 5 | 0 | 0 | 0 | 86.5% |
| math | 272 | 258 | 12 | 0 | 0 | 2 | 95.6% |
| operators | 56 | 53 | 3 | 0 | 0 | 0 | 94.6% |
| reference | 88 | 82 | 6 | 0 | 0 | 0 | 93.2% |
| statistical | 262 | 242 | 20 | 0 | 0 | 0 | 92.4% |
| text | 81 | 77 | 4 | 0 | 0 | 0 | 95.1% |
| trigonometry | 82 | 76 | 6 | 0 | 0 | 0 | 92.7% |
| **total** | 1470 | 1375 | 90 | 0 | 0 | 5 | 93.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1375 | 90 | 0 | 0 | 5 | 93.9% |
| **total** | 1470 | 1375 | 90 | 0 | 0 | 5 | 93.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| HYPGEOM.DIST | 5 | fail 5 | `statistical.xlsx!AA129` =HYPGEOM.DIST(1,4,0,20,TRUE) → Excel 1, GridCraft "#NUM!" |
| DATEVALUE | 4 | fail 4 | `date.xlsx!AA49` =DATEVALUE("Dec-3 11:11") → Excel 46359, GridCraft 37956 |
| BESSELJ | 4 | fail 4 | `engineering.xlsx!AA13` =BESSELJ(1.9, {2,3}) → Excel 0.3299258286697852, GridCraft 0.3299257276923872 |
| LCM | 4 | fail 4 | `math.xlsx!AA137` =LCM(5, 2, (A1, A2)) → Excel "#VALUE!", GridCraft 100000 |
| GCD | 4 | fail 4 | `math.xlsx!AA118` =GCD(5, 2, {3,"7"}, TRUE) → Excel "#VALUE!", GridCraft 1 |
| FREQUENCY | 4 | fail 4 | `statistical.xlsx!AA84` =FREQUENCY({1,2,3,4,5},{1,5,11}) → Excel 1, GridCraft "#VALUE!" |
| CONCATENATE | 4 | fail 4 | `text.xlsx!AA20` =CONCATENATE({0,2,3}, 1, "A", TRUE, -12) → Excel "01ATRUE-12", GridCraft "01ATRUE-12" |
| DELTA | 3 | fail 3 | `engineering.xlsx!AA160` =DELTA(5,FALSE) → Excel "#VALUE!", GridCraft 0 |
| GROWTH | 3 | fail 3 | `statistical.xlsx!AA116` =GROWTH({33100,47300,69000,102000,150000,220000}, {11,12,13,1 → Excel 32618.203773539743, GridCraft 32618.203773539713 |
| YEAR | 2 | skip 2 | `date.xlsx!AA103` =YEAR(NOW()) → Excel 2026, GridCraft 2026 |
| IMPOWER | 2 | fail 2 | `engineering.xlsx!AA270` =IMPOWER("2+3j", TRUE) → Excel "#VALUE!", GridCraft "2+3j" |
| COMPLEX | 2 | fail 2 | `engineering.xlsx!AA131` =COMPLEX(3,FALSE,"j") → Excel "#VALUE!", GridCraft "3" |
| BESSELY | 2 | fail 2 | `engineering.xlsx!AA26` =BESSELY(2.5, {1,2}) → Excel 0.14591813750831284, GridCraft 0.1459181379667858 |
| BESSELK | 2 | fail 2 | `engineering.xlsx!AA17` =BESSELK(TRUE, 1) → Excel "#VALUE!", GridCraft 0.6019072301972342 |
| BESSELI | 2 | fail 2 | `engineering.xlsx!AA7` =BESSELI(1.5, {1,2}) → Excel 0.981666428475166, GridCraft 0.9816664285779074 |
| OR | 2 | fail 2 | `logical.xlsx!AA27` =OR("Test", "FALSE") → Excel false, GridCraft "#VALUE!" |
| AND | 2 | fail 2 | `logical.xlsx!AA6` =AND("Test", "FALSE") → Excel false, GridCraft "#VALUE!" |
| MMULT | 2 | fail 2 | `math.xlsx!AA156` =MMULT({1,3;7,2}, {2,0;0,2}) → Excel 2, GridCraft "#VALUE!" |
| VLOOKUP | 2 | fail 2 | `reference.xlsx!AA81` =VLOOKUP("b", {"c";TRUE;"abc";65;63;61;"b";"a";TRUE;FALSE}, 1 → Excel "b", GridCraft "a" |
| HLOOKUP | 2 | fail 2 | `reference.xlsx!AA38` =HLOOKUP(3, 3, 1,FALSE) → Excel "#N/A", GridCraft 3 |
| AVERAGEA | 2 | fail 2 | `statistical.xlsx!AA227` =AVERAGEA({TRUE,1,2,"12"}, 3) → Excel 1.5, GridCraft 1.4 |
| COUNTIF | 2 | fail 2 | `statistical.xlsx!AA259` =COUNTIF(B1:E1, {1,3,4}) → Excel 1, GridCraft 1 |
| SECH | 2 | fail 2 | `trigonometry.xlsx!AA60` =SECH(2^100) → Excel 0, GridCraft "#NUM!" |
| CSCH | 2 | fail 2 | `trigonometry.xlsx!AA53` =CSCH(2^100) → Excel 0, GridCraft "#NUM!" |
| COTH | 2 | fail 2 | `trigonometry.xlsx!AA48` =COTH(-2^100) → Excel 1, GridCraft "#NUM!" |
