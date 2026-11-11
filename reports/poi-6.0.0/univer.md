# univer vs poi@6.0.0 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 159. Fails where the case's own cell matches but its spill does not: 34.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 151 | 16 | 8 | 0 | 81 | 86.3% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 47 | 9 | 51 | 0 | 0 | 43.9% |
| database | 182 | 130 | 52 | 0 | 0 | 0 | 71.4% |
| date | 378 | 136 | 120 | 87 | 0 | 35 | 39.7% |
| engineering | 572 | 118 | 25 | 399 | 0 | 30 | 21.8% |
| financial | 800 | 117 | 30 | 653 | 0 | 0 | 14.6% |
| information | 156 | 86 | 6 | 23 | 0 | 41 | 74.8% |
| logical | 107 | 64 | 20 | 22 | 0 | 1 | 60.4% |
| lookup | 625 | 338 | 119 | 161 | 0 | 7 | 54.7% |
| math | 788 | 466 | 102 | 177 | 0 | 43 | 62.6% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 567 | 128 | 567 | 0 | 41 | 44.9% |
| text | 549 | 228 | 67 | 225 | 0 | 29 | 43.8% |
| web | 8 | 1 | 0 | 7 | 0 | 0 | 12.5% |
| **total** | 5856 | 2449 | 694 | 2380 | 0 | 333 | 44.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 2449 | 694 | 2380 | 0 | 333 | 44.3% |
| **total** | 5856 | 2449 | 694 | 2380 | 0 | 333 | 44.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| XLOOKUP | 56 | fail 56 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, Apache POI "#VALUE!" |
| AGGREGATE | 54 | unsupported 54 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", Apache POI "#NAME?" |
| NETWORKDAYS.INTL | 40 | unsupported 40 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =NETWORKDAYS.INTL(44613,44613,"0000000",A1:A54) → Excel "#VALUE!", Apache POI "#NAME?" |
| CONVERT | 37 | unsupported 37 | `engineering/convert.xlsx!Z1` =CONVERT(1,"lbm","kg") → Excel 0.45359237, Apache POI "#NAME?" |
| DATEDIF | 35 | unsupported 35 | `date/datedif.xlsx!Z1` =DATEDIF("2011/1/29","2021/3/31","Y") → Excel 10, Apache POI "#NAME?" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", Apache POI "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", Apache POI "#NAME?" |
| WORKDAY.INTL | 30 | fail 29, unsupported 1 | `date/workday-intl.xlsx!Z1` =WORKDAY.INTL(45000,1,"1") → Excel "#VALUE!", Apache POI 45001 |
| INDEX | 30 | fail 30 | `lookup/index.xlsx!Z1` =INDEX(A1:B1,A3:F4,) → Excel "Tom", Apache POI "#VALUE!" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, Apache POI 46335 |
| HYPGEOM.DIST | 27 | unsupported 27 | `statistical/hypgeom-dist.xlsx!Z1` =HYPGEOM.DIST(1,4,8,20,TRUE) → Excel 0.46542827657378727, Apache POI "#NAME?" |
| TEXTAFTER | 25 | unsupported 25 | `text/textafter.xlsx!Z1` =TEXTAFTER("***TRUETRUEFALSE","T") → Excel "RUETRUEFALSE", Apache POI "#NAME?" |
| TEXTBEFORE | 25 | unsupported 25 | `text/textbefore.xlsx!Z1` =TEXTBEFORE("***TRUETRUEFALSE","T") → Excel "***", Apache POI "#NAME?" |
| TEXTSPLIT | 25 | unsupported 25 | `text/textsplit.xlsx!Z1` =TEXTSPLIT("Dakota Lennon Sanchez"," ") → Excel "Dakota", Apache POI "#NAME?" |
| HYPGEOMDIST | 24 | unsupported 24 | `compatibility/hypgeomdist.xlsx!Z1` =HYPGEOMDIST(1,4,8,20) → Excel 0.3632610939112486, Apache POI "#NAME?" |
| DURATION | 23 | unsupported 23 | `financial/duration.xlsx!Z1` =DURATION("2018-7-1","2048-1-1",0.08,0.09,2,1) → Excel 10.919145281591932, Apache POI "#NAME?" |
| ODDFPRICE | 23 | unsupported 23 | `financial/oddfprice.xlsx!Z1` =ODDFPRICE("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 113.59771747407883, Apache POI "#NAME?" |
| AMORLINC | 22 | unsupported 22 | `financial/amorlinc.xlsx!Z1` =AMORLINC(2400,"2008/8/19","2008/12/31",300,1,0.15,0) → Excel 360, Apache POI "#NAME?" |
| ODDLPRICE | 22 | unsupported 22 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, Apache POI "#NAME?" |
| ODDLYIELD | 22 | unsupported 22 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, Apache POI "#NAME?" |
| YIELD | 22 | unsupported 22 | `financial/yield.xlsx!Z1` =YIELD("2008-11-11","2021-3-1",0.0785,98.45,100,2,1) → Excel 0.0804842002951741, Apache POI "#NAME?" |
| COUPPCD | 21 | unsupported 21 | `financial/couppcd.xlsx!Z1` =COUPPCD("2011-1-25","2011-11-15",2,1) → Excel 40497, Apache POI "#NAME?" |
| ODDFYIELD | 21 | unsupported 21 | `financial/oddfyield.xlsx!Z1` =ODDFYIELD("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 0.10076644980091415, Apache POI "#NAME?" |
| ACCRINT | 20 | unsupported 20 | `financial/accrint.xlsx!Z1` =ACCRINT("2008/2/29","2008/8/31","2008/5/1",0.1,1000,2) → Excel 16.944444444444446, Apache POI "#NAME?" |
| COUPNUM | 20 | unsupported 20 | `financial/coupnum.xlsx!Z1` =COUPNUM("2011-1-25","2011-11-15",2,1) → Excel 2, Apache POI "#NAME?" |
