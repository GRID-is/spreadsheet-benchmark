# fast-formula-parser vs hyperformula@3.5.0 (fast-formula-parser)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 5. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 21. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 70 of 1465 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| date | 158 | 45 | 110 | 0 | 0 | 3 | 29.0% |
| engineering | 376 | 270 | 48 | 58 | 0 | 0 | 71.8% |
| information | 58 | 37 | 6 | 15 | 0 | 0 | 63.8% |
| logical | 37 | 21 | 11 | 5 | 0 | 0 | 56.8% |
| math | 272 | 234 | 27 | 9 | 0 | 2 | 86.7% |
| operators | 56 | 44 | 3 | 9 | 0 | 0 | 78.6% |
| reference | 88 | 26 | 33 | 29 | 0 | 0 | 29.5% |
| statistical | 262 | 173 | 13 | 76 | 0 | 0 | 66.0% |
| text | 81 | 44 | 4 | 33 | 0 | 0 | 54.3% |
| trigonometry | 82 | 75 | 7 | 0 | 0 | 0 | 91.5% |
| **total** | 1470 | 969 | 262 | 234 | 0 | 5 | 66.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1470 | 969 | 262 | 234 | 0 | 5 | 66.1% |
| **total** | 1470 | 969 | 262 | 234 | 0 | 5 | 66.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| DATEVALUE | 24 | fail 24 | `date.xlsx!AA51` =DATEVALUE("12/3/14 12:00 AM") → Excel 41976, HyperFormula 41710 |
| VLOOKUP | 24 | fail 12, unsupported 12 | `reference.xlsx!AA86` =VLOOKUP(3, {1;2;3}, 1,FALSE) → Excel 3, HyperFormula "#NAME?" |
| HLOOKUP | 24 | fail 12, unsupported 12 | `reference.xlsx!AA25` =HLOOKUP(3, {3,2,1}, 1) → Excel 1, HyperFormula "#VALUE!" |
| YEARFRAC | 19 | fail 19 | `date.xlsx!AA149` =YEARFRAC("2020/2/1","2022/1/1", 1) → Excel 1.916058394160584, HyperFormula "#VALUE!" |
| SUM | 18 | fail 14, unsupported 4 | `math.xlsx!AA239` =SUM("1", {"1"},) → Excel 1, HyperFormula 2 |
| DATEDIF | 17 | fail 17 | `date.xlsx!AA26` =DATEDIF("9/15/2001","8/15/2003","YD") → Excel 334, HyperFormula "#VALUE!" |
| HYPGEOM.DIST | 13 | fail 2, unsupported 11 | `statistical.xlsx!AA139` =HYPGEOM.DIST(1,8,8,7,FALSE) → Excel "#NUM!", HyperFormula "#NAME?" |
| NUMBERVALUE | 12 | unsupported 12 | `text.xlsx!AA56` =NUMBERVALUE("3.5%",,) → Excel "#VALUE!", HyperFormula "#NAME?" |
| DAYS360 | 11 | fail 11 | `date.xlsx!AA70` =DAYS360("2/1/2019", "3/31/2019", TRUE) → Excel 59, HyperFormula "#VALUE!" |
| GESTEP | 7 | unsupported 7 | `engineering.xlsx!AA169` =GESTEP(5,5) → Excel 1, HyperFormula "#NAME?" |
| TYPE | 7 | unsupported 7 | `information.xlsx!AA55` =TYPE("") → Excel 2, HyperFormula "#NAME?" |
| ERROR.TYPE | 7 | fail 1, unsupported 6 | `information.xlsx!AA1` =ERROR.TYPE(#NULL!) → Excel 1, HyperFormula "#ERROR!" |
| AVERAGEA | 6 | fail 2, unsupported 4 | `statistical.xlsx!AA227` =AVERAGEA({TRUE,1,2,"12"}, 3) → Excel 1.5, HyperFormula "#NAME?" |
| WEEKDAY | 5 | fail 5 | `date.xlsx!AA114` =WEEKDAY("2/14/2008", 2) → Excel 4, HyperFormula "#VALUE!" |
| NETWORKDAYS.INTL | 5 | fail 5 | `date.xlsx!AA101` =NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31), "1") → Excel "#VALUE!", HyperFormula "#NUM!" |
| NETWORKDAYS | 5 | fail 5 | `date.xlsx!AA89` =NETWORKDAYS("10/1/2012","3/1/2013") → Excel 110, HyperFormula 258 |
| DAYS | 5 | fail 5 | `date.xlsx!AA64` =DAYS("12/31/2011","1/1/2011") → Excel 364, HyperFormula "#VALUE!" |
| BITAND | 5 | fail 2, unsupported 3 | `engineering.xlsx!AA73` =BITAND(1,FALSE) → Excel 0, HyperFormula "#NAME?" |
| OR | 5 | fail 4, unsupported 1 | `logical.xlsx!AA28` =OR({0,1,0}, FALSE) → Excel true, HyperFormula "#NAME?" |
| AND | 5 | fail 4, unsupported 1 | `logical.xlsx!AA8` =AND((A2:C2, A3)) → Excel true, HyperFormula "#ERROR!" |
| GROWTH | 5 | unsupported 5 | `statistical.xlsx!AA116` =GROWTH({33100,47300,69000,102000,150000,220000}, {11,12,13,1 → Excel 32618.203773539743, HyperFormula "#NAME?" |
| FIXED | 5 | unsupported 5 | `text.xlsx!AA31` =FIXED(12345.64123213) → Excel "12,345.64", HyperFormula "#NAME?" |
| CONCAT | 5 | unsupported 5 | `text.xlsx!AA12` =CONCAT(0, {1,2,3;5,6,7}) → Excel "0123567", HyperFormula "#NAME?" |
| HOUR | 4 | fail 3, skip 1 | `date.xlsx!AA104` =HOUR(NOW()) → Excel 18, HyperFormula 4 |
| IMSUB | 4 | fail 2, unsupported 2 | `engineering.xlsx!AA311` =IMSUB(TRUE,"5+3i") → Excel "#VALUE!", HyperFormula "#NAME?" |
