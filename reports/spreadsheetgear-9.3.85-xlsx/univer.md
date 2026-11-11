# univer vs spreadsheetgear@9.3.85 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 84. Fails where the case's own cell matches but its spill does not: 47.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 163 | 0 | 12 | 0 | 81 | 93.1% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 103 | 4 | 0 | 0 | 0 | 96.3% |
| database | 182 | 160 | 22 | 0 | 0 | 0 | 87.9% |
| date | 378 | 322 | 19 | 2 | 0 | 35 | 93.9% |
| engineering | 572 | 507 | 35 | 0 | 0 | 30 | 93.5% |
| financial | 800 | 748 | 51 | 1 | 0 | 0 | 93.5% |
| information | 156 | 114 | 1 | 0 | 0 | 41 | 99.1% |
| logical | 107 | 72 | 0 | 34 | 0 | 1 | 67.9% |
| lookup | 625 | 256 | 36 | 326 | 0 | 7 | 41.4% |
| math | 788 | 712 | 23 | 10 | 0 | 43 | 95.6% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1148 | 88 | 26 | 0 | 41 | 91.0% |
| text | 549 | 352 | 24 | 144 | 0 | 29 | 67.7% |
| web | 8 | 7 | 1 | 0 | 0 | 0 | 87.5% |
| **total** | 5856 | 4664 | 304 | 555 | 0 | 333 | 84.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 4664 | 304 | 555 | 0 | 333 | 84.4% |
| **total** | 5856 | 4664 | 304 | 555 | 0 | 333 | 84.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 168 | unsupported 168 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, SpreadsheetGear "#NAME?" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, SpreadsheetGear 46335 |
| TEXTAFTER | 25 | unsupported 25 | `text/textafter.xlsx!Z1` =TEXTAFTER("***TRUETRUEFALSE","T") → Excel "RUETRUEFALSE", SpreadsheetGear "#NAME?" |
| TEXTBEFORE | 25 | unsupported 25 | `text/textbefore.xlsx!Z1` =TEXTBEFORE("***TRUETRUEFALSE","T") → Excel "***", SpreadsheetGear "#NAME?" |
| TEXTSPLIT | 25 | unsupported 25 | `text/textsplit.xlsx!Z1` =TEXTSPLIT("Dakota Lennon Sanchez"," ") → Excel "Dakota", SpreadsheetGear "#NAME?" |
| INDEX | 22 | fail 22 | `lookup/index.xlsx!Z1` =INDEX(A1:B1,A3:F4,) → Excel "Tom", SpreadsheetGear "Tom" |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| UNIQUE | 19 | unsupported 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =UNIQUE(A1:A12) → Excel "Red", SpreadsheetGear "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| EXPAND | 17 | unsupported 17 | `lookup/expand.xlsx!Z1` =EXPAND(A1:B2,3,3) → Excel 2, SpreadsheetGear "#NAME?" |
| TAKE | 16 | unsupported 16 | `lookup/take.xlsx!Z1` =TAKE(A1:B2,1,1) → Excel 2, SpreadsheetGear "#NAME?" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| SORT | 15 | fail 1, unsupported 14 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =SORT(A1:B5,2,-1) → Excel "Red", SpreadsheetGear "#NAME?" |
| DROP | 15 | unsupported 15 | `lookup/drop.xlsx!Z1` =DROP(A1:B2,1,1) → Excel 2, SpreadsheetGear "#NAME?" |
| REGEXEXTRACT | 15 | unsupported 15 | `text/regexextract.xlsx!Z1` =REGEXEXTRACT("abcdefg","c.*f") → Excel "cdef", SpreadsheetGear "#NAME?" |
| TEXTJOIN | 14 | unsupported 14 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TEXTJOIN(", ",TRUE,A1:A3) → Excel "B, 101, 102", SpreadsheetGear "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, SpreadsheetGear "#NAME?" |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| FILTER | 12 | unsupported 12 | `lookup/filter.xlsx!Z1` =FILTER(A1:C2,A4:A5) → Excel 1, SpreadsheetGear "#NAME?" |
