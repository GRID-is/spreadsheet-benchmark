# univer vs google-sheets@2026-11-09 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 259. Fails where the case's own cell matches but its spill does not: 113.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 153 | 19 | 3 | 0 | 81 | 87.4% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 102 | 5 | 0 | 0 | 0 | 95.3% |
| database | 182 | 143 | 39 | 0 | 0 | 0 | 78.6% |
| date | 378 | 234 | 107 | 2 | 0 | 35 | 68.2% |
| engineering | 572 | 366 | 140 | 36 | 0 | 30 | 67.5% |
| financial | 800 | 543 | 151 | 105 | 1 | 0 | 67.9% |
| information | 156 | 105 | 10 | 0 | 0 | 41 | 91.3% |
| logical | 107 | 84 | 19 | 3 | 0 | 1 | 79.2% |
| lookup | 625 | 341 | 197 | 79 | 1 | 7 | 55.2% |
| math | 788 | 614 | 70 | 61 | 0 | 43 | 82.4% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 1130 | 120 | 12 | 0 | 41 | 89.5% |
| text | 549 | 334 | 82 | 104 | 0 | 29 | 64.2% |
| web | 8 | 6 | 0 | 2 | 0 | 0 | 75.0% |
| **total** | 5856 | 4155 | 959 | 407 | 2 | 333 | 75.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 4155 | 959 | 407 | 2 | 333 | 75.2% |
| **total** | 5856 | 4155 | 959 | 407 | 2 | 333 | 75.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 61 | fail 61 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =INDEX(A1,1) → Excel 0, Google Sheets "" |
| XLOOKUP | 55 | fail 55 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, Google Sheets "#N/A" |
| AGGREGATE | 54 | unsupported 54 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", Google Sheets "#NAME?" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", Google Sheets 46296 |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", Google Sheets "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, Google Sheets 46334 |
| WORKDAY.INTL | 26 | fail 25, unsupported 1 | `date/workday-intl.xlsx!Z1` =WORKDAY.INTL(45000,1,"1") → Excel "#VALUE!", Google Sheets "#NUM!" |
| TEXTAFTER | 26 | fail 1, unsupported 25 | `text/textafter.xlsx!Z1` =TEXTAFTER("***TRUETRUEFALSE","T") → Excel "RUETRUEFALSE", Google Sheets "#NAME?" |
| TEXTBEFORE | 26 | fail 1, unsupported 25 | `text/textbefore.xlsx!Z1` =TEXTBEFORE("***TRUETRUEFALSE","T") → Excel "***", Google Sheets "#NAME?" |
| TEXTSPLIT | 25 | unsupported 25 | `text/textsplit.xlsx!Z1` =TEXTSPLIT("Dakota Lennon Sanchez"," ") → Excel "Dakota", Google Sheets "#NAME?" |
| ODDFPRICE | 23 | unsupported 23 | `financial/oddfprice.xlsx!Z1` =ODDFPRICE("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 113.59771747407883, Google Sheets "#NAME?" |
| ODDLPRICE | 22 | unsupported 22 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, Google Sheets "#NAME?" |
| ODDLYIELD | 22 | unsupported 22 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, Google Sheets "#NAME?" |
| ODDFYIELD | 21 | unsupported 21 | `financial/oddfyield.xlsx!Z1` =ODDFYIELD("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 0.10076644980091415, Google Sheets "#NAME?" |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", Google Sheets "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", Google Sheets "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", Google Sheets 44734.58692 |
| NETWORKDAYS.INTL | 19 | fail 18, unsupported 1 | `date/networkdays-intl.xlsx!Z1` =NETWORKDAYS.INTL(45000,45006,"1") → Excel "#VALUE!", Google Sheets "#NUM!" |
| EXPAND | 18 | fail 1, unsupported 17 | `lookup/expand.xlsx!Z1` =EXPAND(A1:B2,3,3) → Excel 2, Google Sheets "#NAME?" |
| TAKE | 17 | fail 1, unsupported 16 | `lookup/take.xlsx!Z1` =TAKE(A1:B2,1,1) → Excel 2, Google Sheets "#NAME?" |
| DROP | 16 | fail 1, unsupported 15 | `lookup/drop.xlsx!Z1` =DROP(A1:B2,1,1) → Excel 2, Google Sheets "#NAME?" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", Google Sheets "#NAME?" |
| SORT | 15 | fail 1, unsupported 14 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =SORT(A1:B5,2,-1) → Excel "Red", Google Sheets "#NAME?" |
| NUMBERVALUE | 15 | unsupported 15 | `text/numbervalue.xlsx!Z1` =NUMBERVALUE("2.500,27",",",".") → Excel 2500.27, Google Sheets "#NAME?" |
| CHOOSE | 14 | fail 14 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =CHOOSE(A1:B1,A3:A7,A9:A13) → Excel "Red", Google Sheets "Red" |

## Errors

- `financial/mirr.xlsx`: Google Sheets failed to export sheet 5 Mirr ReinvestRate === -1
- `lookup/unique.xlsx`: Google Sheets failed to export sheet 18 Unique ByCol and exactlyOnce
