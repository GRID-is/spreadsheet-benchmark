# univer vs hyperformula@3.5.0 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 230. Fails where the case's own cell matches but its spill does not: 109.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 154 | 15 | 6 | 0 | 81 | 88.0% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 26 | 62 | 19 | 0 | 0 | 24.3% |
| database | 182 | 131 | 39 | 12 | 0 | 0 | 72.0% |
| date | 378 | 145 | 182 | 16 | 0 | 35 | 42.3% |
| engineering | 572 | 299 | 129 | 114 | 0 | 30 | 55.2% |
| financial | 800 | 224 | 69 | 507 | 0 | 0 | 28.0% |
| information | 156 | 78 | 10 | 27 | 0 | 41 | 67.8% |
| logical | 107 | 56 | 12 | 38 | 0 | 1 | 52.8% |
| lookup | 625 | 269 | 198 | 151 | 0 | 7 | 43.5% |
| math | 788 | 486 | 92 | 153 | 14 | 43 | 65.2% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 755 | 74 | 433 | 0 | 41 | 59.8% |
| text | 549 | 202 | 30 | 288 | 0 | 29 | 38.8% |
| web | 8 | 1 | 0 | 7 | 0 | 0 | 12.5% |
| **total** | 5856 | 2826 | 912 | 1771 | 14 | 333 | 51.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 2826 | 912 | 1771 | 14 | 333 | 51.2% |
| **total** | 5856 | 2826 | 912 | 1771 | 14 | 333 | 51.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| INDEX | 92 | fail 91, unsupported 1 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =INDEX(A1,1) → Excel 0, HyperFormula null |
| AGGREGATE | 54 | unsupported 54 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", HyperFormula "#NAME?" |
| XLOOKUP | 41 | fail 41 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, HyperFormula "#VALUE!" |
| CONVERT | 37 | fail 2, unsupported 35 | `engineering/convert.xlsx!Z1` =CONVERT(1,"lbm","kg") → Excel 0.45359237, HyperFormula "#NAME?" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", HyperFormula "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", HyperFormula "#NAME?" |
| DATEDIF | 31 | fail 30, unsupported 1 | `date/datedif.xlsx!Z1` =DATEDIF("2011/1/29","2021/3/31","Y") → Excel 10, HyperFormula "#VALUE!" |
| WORKDAY.INTL | 29 | fail 28, unsupported 1 | `date/workday-intl.xlsx!Z1` =WORKDAY.INTL(45000,1,"1") → Excel "#VALUE!", HyperFormula "#NUM!" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, HyperFormula 46335 |
| HYPGEOMDIST | 28 | fail 28 | `compatibility/hypgeomdist.xlsx!Z1` =HYPGEOMDIST(1,4,8,20) → Excel 0.3632610939112486, HyperFormula "#N/A" |
| TEXTAFTER | 26 | fail 1, unsupported 25 | `text/textafter.xlsx!Z1` =TEXTAFTER("***TRUETRUEFALSE","T") → Excel "RUETRUEFALSE", HyperFormula "#NAME?" |
| TEXTBEFORE | 26 | fail 1, unsupported 25 | `text/textbefore.xlsx!Z1` =TEXTBEFORE("***TRUETRUEFALSE","T") → Excel "***", HyperFormula "#NAME?" |
| NETWORKDAYS.INTL | 25 | fail 24, unsupported 1 | `date/networkdays-intl.xlsx!Z1` =NETWORKDAYS.INTL(45000,45006,"1") → Excel "#VALUE!", HyperFormula "#NUM!" |
| TEXTSPLIT | 25 | fail 1, unsupported 24 | `text/textsplit.xlsx!Z1` =TEXTSPLIT("Dakota Lennon Sanchez"," ") → Excel "Dakota", HyperFormula "#NAME?" |
| DURATION | 23 | unsupported 23 | `financial/duration.xlsx!Z1` =DURATION("2018-7-1","2048-1-1",0.08,0.09,2,1) → Excel 10.919145281591932, HyperFormula "#NAME?" |
| ODDFPRICE | 23 | unsupported 23 | `financial/oddfprice.xlsx!Z1` =ODDFPRICE("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 113.59771747407883, HyperFormula "#NAME?" |
| AMORLINC | 22 | unsupported 22 | `financial/amorlinc.xlsx!Z1` =AMORLINC(2400,"2008/8/19","2008/12/31",300,1,0.15,0) → Excel 360, HyperFormula "#NAME?" |
| ODDLPRICE | 22 | unsupported 22 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, HyperFormula "#NAME?" |
| ODDLYIELD | 22 | unsupported 22 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, HyperFormula "#NAME?" |
| YIELD | 22 | unsupported 22 | `financial/yield.xlsx!Z1` =YIELD("2008-11-11","2021-3-1",0.0785,98.45,100,2,1) → Excel 0.0804842002951741, HyperFormula "#NAME?" |
| COUPPCD | 21 | unsupported 21 | `financial/couppcd.xlsx!Z1` =COUPPCD("2011-1-25","2011-11-15",2,1) → Excel 40497, HyperFormula "#NAME?" |
| ODDFYIELD | 21 | unsupported 21 | `financial/oddfyield.xlsx!Z1` =ODDFYIELD("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 0.10076644980091415, HyperFormula "#NAME?" |
| TREND | 21 | unsupported 21 | `statistical/trend.xlsx!Z1` =TREND(A1:F1,A3:F3,A5:B5,TRUE) → Excel 155404.5358352654, HyperFormula "#NAME?" |
| ACCRINT | 20 | unsupported 20 | `financial/accrint.xlsx!Z1` =ACCRINT("2008/2/29","2008/8/31","2008/5/1",0.1,1000,2) → Excel 16.944444444444446, HyperFormula "#NAME?" |
| COUPNUM | 20 | unsupported 20 | `financial/coupnum.xlsx!Z1` =COUPNUM("2011-1-25","2011-11-15",2,1) → Excel 2, HyperFormula "#NAME?" |

## Errors

- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
- `math/sequence.xlsx`: timeout after 60000 ms
