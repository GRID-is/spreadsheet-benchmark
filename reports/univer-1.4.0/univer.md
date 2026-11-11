# univer vs univer@1.4.0 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 27. Fails where the case's own cell matches but its spill does not: 53.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 157 | 18 | 0 | 0 | 81 | 89.7% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 102 | 5 | 0 | 0 | 0 | 95.3% |
| database | 182 | 180 | 2 | 0 | 0 | 0 | 98.9% |
| date | 378 | 320 | 21 | 2 | 0 | 35 | 93.3% |
| engineering | 572 | 480 | 62 | 0 | 0 | 30 | 88.6% |
| financial | 800 | 770 | 26 | 4 | 0 | 0 | 96.3% |
| information | 156 | 113 | 2 | 0 | 0 | 41 | 98.3% |
| logical | 107 | 94 | 7 | 5 | 0 | 1 | 88.7% |
| lookup | 625 | 591 | 27 | 0 | 0 | 7 | 95.6% |
| math | 788 | 711 | 34 | 0 | 0 | 43 | 95.4% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1202 | 56 | 4 | 0 | 41 | 95.2% |
| text | 549 | 453 | 67 | 0 | 0 | 29 | 87.1% |
| web | 8 | 7 | 1 | 0 | 0 | 0 | 87.5% |
| **total** | 5856 | 5180 | 328 | 15 | 0 | 333 | 93.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 5180 | 328 | 15 | 0 | 333 | 93.8% |
| **total** | 5856 | 5180 | 328 | 15 | 0 | 333 | 93.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", Univer 46296 |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", Univer false |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, Univer 46336 |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", Univer 2.34 |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", Univer 46297 |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", Univer 44734.58692129629 |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", Univer "一百二十三" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", Univer "1.11394335230684+0.510732572130908i" |
| LINEST | 14 | fail 13, unsupported 1 | `statistical/linest.xlsx!Z1` =LINEST(A1:A4,A6:A9,TRUE,TRUE) → Excel 1, Univer 1 |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", Univer true |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, Univer 0.707682300923 |
| SEQUENCE | 13 | fail 13 | `math/sequence.xlsx!Z1` =SEQUENCE(10000,10000) → Excel "#SPILL!", Univer "#VALUE!" |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", Univer true |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", Univer true |
| ARRAY_CONSTRAIN | 9 | skip 9 | `array/array-constrain.xlsx!Z1` =ARRAY_CONSTRAIN(A1:C3,2,2) → Excel "#NAME?", Univer 1 |
| TO_DATE | 9 | skip 9 | `date/to-date.xlsx!Z1` =TO_DATE(40826.4375) → Excel "#NAME?", Univer 40826.4375 |
| ISEMAIL | 9 | skip 9 | `information/isemail.xlsx!Z1` =ISEMAIL("developer@univer.ai") → Excel "#NAME?", Univer true |
| AVERAGE.WEIGHTED | 9 | skip 9 | `statistical/average-weighted.xlsx!Z1` =AVERAGE.WEIGHTED(1,3,2,4) → Excel "#NAME?", Univer 1.57142857143 |
| MARGINOFERROR | 9 | skip 9 | `statistical/marginoferror.xlsx!Z1` =MARGINOFERROR(A1:B2,0.1) → Excel "#NAME?", Univer 0.0881737585365 |
| IMCOTH | 8 | skip 8 | `engineering/imcoth.xlsx!Z1` =IMCOTH("5+12i") → Excel "#NAME?", Univer "1.00003851275523+0.0000822295549932979i" |
| IMTANH | 8 | skip 8 | `engineering/imtanh.xlsx!Z1` =IMTANH("5+12i") → Excel "#NAME?", Univer "0.99996148196703-0.000082223221029802i" |
| XIRR | 8 | fail 6, unsupported 2 | `financial/xirr.xlsx!Z1` =XIRR(1,A1:E1,0.1) → Excel "#NUM!", Univer "#N/A" |
| DOLLAR | 8 | fail 8 | `text/dollar.xlsx!Z1` =DOLLAR(1234.567,) → Excel "$1,235", Univer "$1,234.57 " |
| CHOOSE | 7 | fail 7 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =CHOOSE(A1:B1,A3:A7,A9:A13) → Excel "Red", Univer "Red" |
| NETWORKDAYS.INTL | 7 | fail 6, unsupported 1 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =NETWORKDAYS.INTL(44613,44613,"0000000",A1:A54) → Excel "#VALUE!", Univer 0 |
