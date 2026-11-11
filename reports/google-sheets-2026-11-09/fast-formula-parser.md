# fast-formula-parser vs google-sheets@2026-11-09 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 39. Fails where the case's own cell matches but its spill does not: 2.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 138 | 17 | 0 | 0 | 3 | 89.0% |
| engineering | 376 | 306 | 44 | 26 | 0 | 0 | 81.4% |
| information | 58 | 55 | 3 | 0 | 0 | 0 | 94.8% |
| logical | 37 | 26 | 11 | 0 | 0 | 0 | 70.3% |
| math | 272 | 251 | 19 | 0 | 0 | 2 | 93.0% |
| operators | 56 | 53 | 3 | 0 | 0 | 0 | 94.6% |
| reference | 88 | 76 | 11 | 1 | 0 | 0 | 86.4% |
| statistical | 262 | 238 | 24 | 0 | 0 | 0 | 90.8% |
| text | 81 | 53 | 13 | 15 | 0 | 0 | 65.4% |
| trigonometry | 82 | 72 | 10 | 0 | 0 | 0 | 87.8% |
| **total** | 1470 | 1268 | 155 | 42 | 0 | 5 | 86.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 1268 | 155 | 42 | 0 | 5 | 86.6% |
| **total** | 1470 | 1268 | 155 | 42 | 0 | 5 | 86.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| NUMBERVALUE | 12 | unsupported 12 | `text.xlsx!AA56` =NUMBERVALUE("3.5%",,) → Excel "#VALUE!", Google Sheets "#NAME?" |
| BESSELY | 7 | unsupported 7 | `engineering.xlsx!AA22` =BESSELY(2.5, 1.9) → Excel 0.14591813750831284, Google Sheets "#NAME?" |
| BESSELI | 7 | unsupported 7 | `engineering.xlsx!AA6` =BESSELI("a", 1) → Excel "#VALUE!", Google Sheets "#NAME?" |
| SUM | 7 | fail 7 | `math.xlsx!AA234` =SUM((A1:C1, C1:E1)) → Excel 18, Google Sheets "#ERROR!" |
| BESSELJ | 6 | unsupported 6 | `engineering.xlsx!AA12` =BESSELJ("a", 2) → Excel "#VALUE!", Google Sheets "#NAME?" |
| BESSELK | 6 | unsupported 6 | `engineering.xlsx!AA14` =BESSELK(1.5, 1) → Excel 0.2773878036322587, Google Sheets "#NAME?" |
| NETWORKDAYS.INTL | 5 | fail 5 | `date.xlsx!AA101` =NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31), "1") → Excel "#VALUE!", Google Sheets "#NUM!" |
| YEARFRAC | 5 | fail 5 | `date.xlsx!AA149` =YEARFRAC("2020/2/1","2022/1/1", 1) → Excel 1.916058394160584, Google Sheets 1.915300546 |
| LCM | 5 | fail 5 | `math.xlsx!AA137` =LCM(5, 2, (A1, A2)) → Excel "#VALUE!", Google Sheets "#ERROR!" |
| HYPGEOM.DIST | 5 | fail 5 | `statistical.xlsx!AA129` =HYPGEOM.DIST(1,4,0,20,TRUE) → Excel 1, Google Sheets "#NUM!" |
| CONCAT | 5 | fail 5 | `text.xlsx!AA12` =CONCAT(0, {1,2,3;5,6,7}) → Excel "0123567", Google Sheets "01" |
| OR | 4 | fail 4 | `logical.xlsx!AA29` =OR((A2:C2, A3)) → Excel true, Google Sheets "#ERROR!" |
| AND | 4 | fail 4 | `logical.xlsx!AA8` =AND((A2:C2, A3)) → Excel true, Google Sheets "#ERROR!" |
| FREQUENCY | 4 | fail 4 | `statistical.xlsx!AA84` =FREQUENCY({1,2,3,4,5},{1,5,11}) → Excel 1, Google Sheets "#REF!" |
| CONCATENATE | 4 | fail 4 | `text.xlsx!AA20` =CONCATENATE({0,2,3}, 1, "A", TRUE, -12) → Excel "01ATRUE-12", Google Sheets "0231ATRUE-12" |
| WORKDAY.INTL | 3 | fail 3 | `date.xlsx!AA130` =WORKDAY.INTL(DATE(2012,1,1),30,"1") → Excel "#VALUE!", Google Sheets "#NUM!" |
| DATEVALUE | 3 | fail 3 | `date.xlsx!AA57` =DATEVALUE("10000/12/1") → Excel "#VALUE!", Google Sheets 2958801 |
| IMPRODUCT | 3 | fail 3 | `engineering.xlsx!AA272` =IMPRODUCT("3+4i","5-3j", "5-3j") → Excel "#VALUE!", Google Sheets "#NUM!" |
| IMDIV | 3 | fail 3 | `engineering.xlsx!AA250` =IMDIV(TRUE,"52") → Excel "#VALUE!", Google Sheets "#NUM!" |
| DELTA | 3 | fail 3 | `engineering.xlsx!AA160` =DELTA(5,FALSE) → Excel "#VALUE!", Google Sheets 0 |
| ISREF | 3 | fail 3 | `information.xlsx!AA41` =ISREF(XYZ1) → Excel false, Google Sheets true |
| XOR | 3 | fail 3 | `logical.xlsx!AA36` =XOR("Test", "TRUE") → Excel true, Google Sheets "#VALUE!" |
| GCD | 3 | fail 3 | `math.xlsx!AA118` =GCD(5, 2, {3,"7"}, TRUE) → Excel "#VALUE!", Google Sheets 1 |
| VLOOKUP | 3 | fail 3 | `reference.xlsx!AA81` =VLOOKUP("b", {"c";TRUE;"abc";65;63;61;"b";"a";TRUE;FALSE}, 1 → Excel "b", Google Sheets "a" |
| HLOOKUP | 3 | fail 3 | `reference.xlsx!AA38` =HLOOKUP(3, 3, 1,FALSE) → Excel "#N/A", Google Sheets 3 |
