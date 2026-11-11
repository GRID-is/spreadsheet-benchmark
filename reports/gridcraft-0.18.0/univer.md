# univer vs gridcraft@0.18.0 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 131. Fails where the case's own cell matches but its spill does not: 87.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 174 | 1 | 0 | 0 | 81 | 99.4% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 102 | 5 | 0 | 0 | 0 | 95.3% |
| database | 182 | 123 | 59 | 0 | 0 | 0 | 67.6% |
| date | 378 | 309 | 32 | 2 | 0 | 35 | 90.1% |
| engineering | 572 | 492 | 50 | 0 | 0 | 30 | 90.8% |
| financial | 800 | 687 | 112 | 1 | 0 | 0 | 85.9% |
| information | 156 | 111 | 4 | 0 | 0 | 41 | 96.5% |
| logical | 107 | 97 | 9 | 0 | 0 | 1 | 91.5% |
| lookup | 625 | 523 | 95 | 0 | 0 | 7 | 84.6% |
| math | 788 | 702 | 36 | 7 | 0 | 43 | 94.2% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1132 | 120 | 10 | 0 | 41 | 89.7% |
| text | 549 | 488 | 30 | 2 | 0 | 29 | 93.8% |
| web | 8 | 6 | 2 | 0 | 0 | 0 | 75.0% |
| **total** | 5856 | 4946 | 555 | 22 | 0 | 333 | 89.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 4946 | 555 | 22 | 0 | 333 | 89.6% |
| **total** | 5856 | 4946 | 555 | 22 | 0 | 333 | 89.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", GridCraft "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", GridCraft "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, GridCraft 46336 |
| INDEX | 21 | fail 21 | `lookup/index.xlsx!Z1` =INDEX(A1:B1,A3:F4,) → Excel "Tom", GridCraft "#VALUE!" |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", GridCraft "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", GridCraft "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", GridCraft "#NAME?" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", GridCraft "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", GridCraft "#NAME?" |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", GridCraft "#NAME?" |
| VLOOKUP | 13 | fail 13 | `lookup/vlookup.xlsx!Z1` =VLOOKUP(2,A1:B8,A10:B11,0) → Excel "#REF!", GridCraft 2 |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, GridCraft 0.48863416842423746 |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", GridCraft "#NAME?" |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", GridCraft "#NAME?" |
| XLOOKUP | 10 | fail 10 | `lookup/xlookup.xlsx!Z1` =XLOOKUP(#N/A,A1:A2,A4:A5,"missing") → Excel "#N/A", GridCraft "missing" |
| ARRAY_CONSTRAIN | 9 | skip 9 | `array/array-constrain.xlsx!Z1` =ARRAY_CONSTRAIN(A1:C3,2,2) → Excel "#NAME?", GridCraft "#NAME?" |
| TO_DATE | 9 | skip 9 | `date/to-date.xlsx!Z1` =TO_DATE(40826.4375) → Excel "#NAME?", GridCraft "#NAME?" |
| ISEMAIL | 9 | skip 9 | `information/isemail.xlsx!Z1` =ISEMAIL("developer@univer.ai") → Excel "#NAME?", GridCraft "#NAME?" |
| AGGREGATE | 9 | fail 7, unsupported 2 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", GridCraft 1 |
| AVERAGE.WEIGHTED | 9 | skip 9 | `statistical/average-weighted.xlsx!Z1` =AVERAGE.WEIGHTED(1,3,2,4) → Excel "#NAME?", GridCraft "#NAME?" |
| LINEST | 9 | fail 6, unsupported 3 | `statistical/linest.xlsx!Z1` =LINEST(A1:E1,A3:E6,TRUE,TRUE) → Excel 0.2595190987624845, GridCraft 0.2595191113650799 |
| MARGINOFERROR | 9 | skip 9 | `statistical/marginoferror.xlsx!Z1` =MARGINOFERROR(A1:B2,0.1) → Excel "#NAME?", GridCraft "#NAME?" |
| IMCOTH | 8 | skip 8 | `engineering/imcoth.xlsx!Z1` =IMCOTH("5+12i") → Excel "#NAME?", GridCraft "#NAME?" |
| IMTANH | 8 | skip 8 | `engineering/imtanh.xlsx!Z1` =IMTANH("5+12i") → Excel "#NAME?", GridCraft "#NAME?" |
| TAKE | 8 | fail 8 | `lookup/take.xlsx!Z1` =TAKE(1,A1:A3,A5:B5) → Excel 1, GridCraft 1 |
