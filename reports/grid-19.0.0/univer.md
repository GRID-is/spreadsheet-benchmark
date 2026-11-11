# univer vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 11.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 175 | 0 | 0 | 0 | 81 | 100.0% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 107 | 0 | 0 | 0 | 0 | 100.0% |
| database | 182 | 182 | 0 | 0 | 0 | 0 | 100.0% |
| date | 378 | 343 | 0 | 0 | 0 | 35 | 100.0% |
| engineering | 572 | 533 | 9 | 0 | 0 | 30 | 98.3% |
| financial | 800 | 800 | 0 | 0 | 0 | 0 | 100.0% |
| information | 156 | 114 | 1 | 0 | 0 | 41 | 99.1% |
| logical | 107 | 106 | 0 | 0 | 0 | 1 | 100.0% |
| lookup | 625 | 608 | 10 | 0 | 0 | 7 | 98.4% |
| math | 788 | 743 | 1 | 1 | 0 | 43 | 99.7% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1261 | 1 | 0 | 0 | 41 | 99.9% |
| text | 549 | 514 | 6 | 0 | 0 | 29 | 98.8% |
| web | 8 | 6 | 0 | 2 | 0 | 0 | 75.0% |
| **total** | 5856 | 5492 | 28 | 3 | 0 | 333 | 99.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 5492 | 28 | 3 | 0 | 333 | 99.4% |
| **total** | 5856 | 5492 | 28 | 3 | 0 | 333 | 99.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", GRID "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", GRID "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, GRID 46335 |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", GRID "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", GRID "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", GRID "#NAME?" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", GRID "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", GRID "#NAME?" |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", GRID "#NAME?" |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, GRID 0.3585084659034442 |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", GRID "#NAME?" |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", GRID "#NAME?" |
| ARRAY_CONSTRAIN | 9 | skip 9 | `array/array-constrain.xlsx!Z1` =ARRAY_CONSTRAIN(A1:C3,2,2) → Excel "#NAME?", GRID "#NAME?" |
| TO_DATE | 9 | skip 9 | `date/to-date.xlsx!Z1` =TO_DATE(40826.4375) → Excel "#NAME?", GRID "#NAME?" |
| ISEMAIL | 9 | skip 9 | `information/isemail.xlsx!Z1` =ISEMAIL("developer@univer.ai") → Excel "#NAME?", GRID "#NAME?" |
| AVERAGE.WEIGHTED | 9 | skip 9 | `statistical/average-weighted.xlsx!Z1` =AVERAGE.WEIGHTED(1,3,2,4) → Excel "#NAME?", GRID "#NAME?" |
| MARGINOFERROR | 9 | skip 9 | `statistical/marginoferror.xlsx!Z1` =MARGINOFERROR(A1:B2,0.1) → Excel "#NAME?", GRID "#NAME?" |
| IMCOTH | 8 | skip 8 | `engineering/imcoth.xlsx!Z1` =IMCOTH("5+12i") → Excel "#NAME?", GRID "#NAME?" |
| IMTANH | 8 | skip 8 | `engineering/imtanh.xlsx!Z1` =IMTANH("5+12i") → Excel "#NAME?", GRID "#NAME?" |
| MULTIPLY | 7 | skip 7 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MULTIPLY(A1:A3,A5:A7) → Excel "#NAME?", GRID "#NAME?" |
| ISURL | 7 | skip 7 | `information/isurl.xlsx!Z1` =ISURL("univer.ai") → Excel "#NAME?", GRID "#NAME?" |
| IMAGE | 7 | skip 7 | `lookup/image.xlsx!Z1` =IMAGE(A1) → Excel "#VALUE!", GRID "#NAME?" |
| RANDBETWEEN | 6 | skip 6 | `math/randbetween.xlsx!Z1` =RANDBETWEEN(1,100) → Excel 42, GRID 76 |
| NOW | 3 | skip 3 | `date/now.xlsx!Z1` =NOW() → Excel 46330.911217824076, GRID 46335.06848297454 |
| CUBE | 3 | skip 3 | `meta/cube.xlsx!Z1` =CUBE(A1:B2,A4:A5) → Excel "#NAME?", GRID "#NAME?" |
