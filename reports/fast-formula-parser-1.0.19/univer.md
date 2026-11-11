# univer vs fast-formula-parser@1.0.19 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 220. Fails where the case's own cell matches but its spill does not: 75.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 93 | 20 | 62 | 0 | 81 | 53.1% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 22 | 2 | 83 | 0 | 0 | 20.6% |
| database | 182 | 12 | 0 | 170 | 0 | 0 | 6.6% |
| date | 378 | 196 | 144 | 3 | 0 | 35 | 57.1% |
| engineering | 572 | 292 | 201 | 49 | 0 | 30 | 53.9% |
| financial | 800 | 62 | 0 | 738 | 0 | 0 | 7.8% |
| information | 156 | 73 | 14 | 28 | 0 | 41 | 63.5% |
| logical | 107 | 54 | 29 | 23 | 0 | 1 | 50.9% |
| lookup | 625 | 134 | 110 | 374 | 0 | 7 | 21.7% |
| math | 788 | 465 | 162 | 118 | 0 | 43 | 62.4% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 607 | 195 | 460 | 0 | 41 | 48.1% |
| text | 549 | 243 | 127 | 150 | 0 | 29 | 46.7% |
| web | 8 | 4 | 2 | 2 | 0 | 0 | 50.0% |
| **total** | 5856 | 2257 | 1006 | 2260 | 0 | 333 | 40.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 2257 | 1006 | 2260 | 0 | 333 | 40.9% |
| **total** | 5856 | 2257 | 1006 | 2260 | 0 | 333 | 40.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 168 | unsupported 168 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, fast-formula-parser "#NAME?" |
| INDEX | 61 | fail 58, unsupported 3 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =INDEX(A1,1) → Excel "#DIV/0!", fast-formula-parser "#VALUE!" |
| AGGREGATE | 54 | unsupported 54 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", fast-formula-parser "#NAME?" |
| CONVERT | 37 | unsupported 37 | `engineering/convert.xlsx!Z1` =CONVERT(1,"lbm","kg") → Excel 0.45359237, fast-formula-parser "#NAME?" |
| SUBTOTAL | 37 | unsupported 37 | `math/subtotal.xlsx!Z1` =SUBTOTAL(1,A1) → Excel "#DIV/0!", fast-formula-parser "#NAME?" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| MATCH | 29 | unsupported 29 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MATCH(9E+99,A1:A1000) → Excel 54, fast-formula-parser "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, fast-formula-parser 46335 |
| TEXTAFTER | 26 | fail 1, unsupported 25 | `text/textafter.xlsx!Z1` =TEXTAFTER("***TRUETRUEFALSE","T") → Excel "RUETRUEFALSE", fast-formula-parser "#NAME?" |
| TEXTBEFORE | 26 | fail 1, unsupported 25 | `text/textbefore.xlsx!Z1` =TEXTBEFORE("***TRUETRUEFALSE","T") → Excel "***", fast-formula-parser "#NAME?" |
| TEXTSPLIT | 25 | unsupported 25 | `text/textsplit.xlsx!Z1` =TEXTSPLIT("Dakota Lennon Sanchez"," ") → Excel "Dakota", fast-formula-parser "#NAME?" |
| HYPGEOMDIST | 24 | unsupported 24 | `compatibility/hypgeomdist.xlsx!Z1` =HYPGEOMDIST(1,4,8,20) → Excel 0.3632610939112486, fast-formula-parser "#NAME?" |
| DURATION | 23 | unsupported 23 | `financial/duration.xlsx!Z1` =DURATION("2018-7-1","2048-1-1",0.08,0.09,2,1) → Excel 10.919145281591932, fast-formula-parser "#NAME?" |
| ODDFPRICE | 23 | unsupported 23 | `financial/oddfprice.xlsx!Z1` =ODDFPRICE("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 113.59771747407883, fast-formula-parser "#NAME?" |
| AMORLINC | 22 | unsupported 22 | `financial/amorlinc.xlsx!Z1` =AMORLINC(2400,"2008/8/19","2008/12/31",300,1,0.15,0) → Excel 360, fast-formula-parser "#NAME?" |
| ODDLPRICE | 22 | unsupported 22 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, fast-formula-parser "#NAME?" |
| ODDLYIELD | 22 | unsupported 22 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, fast-formula-parser "#NAME?" |
| YIELD | 22 | unsupported 22 | `financial/yield.xlsx!Z1` =YIELD("2008-11-11","2021-3-1",0.0785,98.45,100,2,1) → Excel 0.0804842002951741, fast-formula-parser "#NAME?" |
| COUPPCD | 21 | unsupported 21 | `financial/couppcd.xlsx!Z1` =COUPPCD("2011-1-25","2011-11-15",2,1) → Excel 40497, fast-formula-parser "#NAME?" |
| ODDFYIELD | 21 | unsupported 21 | `financial/oddfyield.xlsx!Z1` =ODDFYIELD("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 0.10076644980091415, fast-formula-parser "#NAME?" |
| TREND | 21 | unsupported 21 | `statistical/trend.xlsx!Z1` =TREND(A1:F1,A3:F3,A5:B5,TRUE) → Excel 155404.5358352654, fast-formula-parser "#NAME?" |
| COUNTIF | 20 | fail 20 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COUNTIF(A1:A12,A14:A18) → Excel 4, fast-formula-parser 0 |
| SUMIFS | 20 | unsupported 20 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =SUMIFS(A1:A9,A11:A19,A21,A23:A31,A33) → Excel 0, fast-formula-parser "#NAME?" |
| ACCRINT | 20 | unsupported 20 | `financial/accrint.xlsx!Z1` =ACCRINT("2008/2/29","2008/8/31","2008/5/1",0.1,1000,2) → Excel 16.944444444444446, fast-formula-parser "#NAME?" |
