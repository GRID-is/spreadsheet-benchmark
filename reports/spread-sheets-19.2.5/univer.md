# univer vs spread-sheets@19.2.5 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 146. Fails where the case's own cell matches but its spill does not: 194.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 171 | 4 | 0 | 0 | 81 | 97.7% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 92 | 15 | 0 | 0 | 0 | 86.0% |
| database | 182 | 135 | 47 | 0 | 0 | 0 | 74.2% |
| date | 378 | 241 | 100 | 2 | 0 | 35 | 70.3% |
| engineering | 572 | 356 | 186 | 0 | 0 | 30 | 65.7% |
| financial | 800 | 597 | 186 | 17 | 0 | 0 | 74.6% |
| information | 156 | 94 | 12 | 9 | 0 | 41 | 81.7% |
| logical | 107 | 101 | 0 | 5 | 0 | 1 | 95.3% |
| lookup | 625 | 530 | 88 | 0 | 0 | 7 | 85.8% |
| math | 788 | 606 | 136 | 3 | 0 | 43 | 81.3% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1016 | 238 | 8 | 0 | 41 | 80.5% |
| text | 549 | 436 | 67 | 17 | 0 | 29 | 83.8% |
| web | 8 | 8 | 0 | 0 | 0 | 0 | 100.0% |
| **total** | 5856 | 4383 | 1079 | 61 | 0 | 333 | 79.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 4383 | 1079 | 61 | 0 | 333 | 79.4% |
| **total** | 5856 | 4383 | 1079 | 61 | 0 | 333 | 79.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", SpreadJS "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", SpreadJS "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, SpreadJS 46335 |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", SpreadJS "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", SpreadJS "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", SpreadJS "#NAME?" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", SpreadJS "#NAME?" |
| NUMBERVALUE | 15 | unsupported 15 | `text/numbervalue.xlsx!Z1` =NUMBERVALUE("2.500,27",",",".") → Excel 2500.27, SpreadJS "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", SpreadJS "#NAME?" |
| GROWTH | 14 | fail 13, unsupported 1 | `statistical/growth.xlsx!Z1` =GROWTH(A1:A6) → Excel 32618.203773539757, SpreadJS 42583.630890438566 |
| NETWORKDAYS.INTL | 13 | fail 12, unsupported 1 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =NETWORKDAYS.INTL(44613,44613,"0000000",A1:A54) → Excel "#VALUE!", SpreadJS 0 |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", SpreadJS "#NAME?" |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, SpreadJS "#VALUE!" |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", SpreadJS "#NAME?" |
| XNPV | 12 | fail 5, unsupported 7 | `financial/xnpv.xlsx!Z1` =XNPV(TRUE,A1:E1,A3:E3) → Excel "#VALUE!", SpreadJS -2509.77362081224 |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", SpreadJS "#NAME?" |
| VLOOKUP | 12 | fail 12 | `lookup/vlookup.xlsx!Z1` =VLOOKUP(2,A1:B8,A10:B11,0) → Excel "#REF!", SpreadJS 2 |
| XLOOKUP | 12 | fail 12 | `lookup/xlookup.xlsx!Z1` =XLOOKUP(2,A1:A2,A4:A5,#REF!,-1) → Excel 0, SpreadJS "#REF!" |
| TREND | 12 | fail 11, unsupported 1 | `statistical/trend.xlsx!Z1` =TREND(A1:A6) → Excel 12452.380952380969, SpreadJS 33100 |
| WORKDAY.INTL | 11 | fail 10, unsupported 1 | `date/workday-intl.xlsx!Z1` =WORKDAY.INTL(45000,1,"00000|1") → Excel "#VALUE!", SpreadJS 45001 |
| XIRR | 11 | fail 4, unsupported 7 | `financial/xirr.xlsx!Z1` =XIRR(A1:E1,A3:E3,TRUE) → Excel "#VALUE!", SpreadJS 0.373362533518832 |
| UNIQUE | 11 | fail 11 | `lookup/unique.xlsx!Z1` =UNIQUE(A1:A5) → Excel "Apfel", SpreadJS "Apfel" |
| EOMONTH | 10 | fail 10 | `date/eomonth.xlsx!Z1` =EOMONTH("2011/1/1",TRUE) → Excel "#VALUE!", SpreadJS 40602 |
| LINEST | 10 | fail 8, unsupported 2 | `statistical/linest.xlsx!Z1` =LINEST(A1:A4,A6:A9,FALSE,TRUE) → Excel 0.40229885057471254, SpreadJS 0.40229885057471265 |
| RANK.AVG | 10 | fail 10 | `statistical/rank-avg.xlsx!Z1` =RANK.AVG(A1,A3:H3,0) → Excel 4.5, SpreadJS 5 |
