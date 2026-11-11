# univer vs ironcalc@0.9.1 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 275. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 251. Fails where the case's own cell matches but its spill does not: 69.
Functions outside Excel's list (22): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 442 of 5530 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 153 | 22 | 0 | 0 | 81 | 87.4% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 85 | 22 | 0 | 0 | 0 | 79.4% |
| database | 182 | 155 | 27 | 0 | 0 | 0 | 85.2% |
| date | 378 | 226 | 115 | 2 | 0 | 35 | 65.9% |
| engineering | 572 | 442 | 100 | 0 | 0 | 30 | 81.5% |
| financial | 800 | 606 | 98 | 9 | 87 | 0 | 75.8% |
| information | 156 | 97 | 18 | 0 | 0 | 41 | 84.3% |
| logical | 107 | 91 | 10 | 5 | 0 | 1 | 85.8% |
| lookup | 625 | 441 | 177 | 0 | 7 | 0 | 70.6% |
| math | 788 | 578 | 62 | 54 | 51 | 43 | 77.6% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 910 | 304 | 10 | 38 | 41 | 72.1% |
| text | 549 | 417 | 95 | 8 | 0 | 29 | 80.2% |
| web | 8 | 5 | 1 | 2 | 0 | 0 | 62.5% |
| **total** | 5856 | 4206 | 1051 | 90 | 183 | 326 | 76.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 4206 | 1051 | 90 | 183 | 326 | 76.1% |
| **total** | 5856 | 4206 | 1051 | 90 | 183 | 326 | 76.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| AGGREGATE | 54 | unsupported 54 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", IronCalc "#NAME?" |
| XLOOKUP | 40 | fail 40 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, IronCalc "#N/A" |
| INDEX | 40 | fail 40 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =INDEX(A1,1) → Excel 0, IronCalc "#VALUE!" |
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", IronCalc "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", IronCalc "#NAME?" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, IronCalc 46335 |
| ODDFPRICE | 24 | error 24 | `financial/oddfprice.xlsx!Z1` =ODDFPRICE("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 113.59771747407883, IronCalc null |
| ODDLPRICE | 23 | error 23 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, IronCalc null |
| ODDLYIELD | 23 | error 23 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, IronCalc null |
| VLOOKUP | 21 | fail 21 | `lookup/vlookup.xlsx!Z1` =VLOOKUP(,A1:B2,2,0) → Excel "Zero", IronCalc "#N/A" |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", IronCalc "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", IronCalc "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", IronCalc "#NAME?" |
| RATE | 17 | error 17 | `financial/rate.xlsx!Z1` =RATE(48,-200,100000,0,0,0.1) → Excel -0.07263177007660084, IronCalc null |
| TREND | 17 | fail 15, unsupported 2 | `statistical/trend.xlsx!Z1` =TREND(A1:F1,A3:F3,A5:B5,TRUE) → Excel 155404.5358352654, IronCalc "#REF!" |
| GROWTH | 16 | fail 14, unsupported 2 | `statistical/growth.xlsx!Z1` =GROWTH(A1:F1,A3:F3,A5:B5,TRUE) → Excel 136608.33775488482, IronCalc "#REF!" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", IronCalc "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", IronCalc "#NAME?" |
| XIRR | 13 | fail 10, unsupported 3 | `financial/xirr.xlsx!Z1` =XIRR(A1:E1,A3:E3,TRUE) → Excel "#VALUE!", IronCalc 0.37336253351883164 |
| ISDATE | 13 | skip 13 | `information/isdate.xlsx!Z1` =ISDATE("2012-2-2") → Excel "#NAME?", IronCalc "#NAME?" |
| RANDARRAY | 13 | skip 13 | `math/randarray.xlsx!Z1` =RANDARRAY() → Excel 0.970600747571545, IronCalc 0.4131333115419603 |
| REGEXMATCH | 13 | skip 13 | `text/regexmatch.xlsx!Z1` =REGEXMATCH("abcdefg","c.*f") → Excel "#NAME?", IronCalc "#NAME?" |
| DATEDIF | 12 | fail 12 | `date/datedif.xlsx!Z1` =DATEDIF(0,60,"M") → Excel 1, IronCalc "#NUM!" |
| WORKDAY.INTL | 12 | fail 11, unsupported 1 | `date/workday-intl.xlsx!Z1` =WORKDAY.INTL(45000,0,"11111111") → Excel 45000, IronCalc "#VALUE!" |
| ISBETWEEN | 12 | skip 12 | `information/isbetween.xlsx!Z1` =ISBETWEEN(7.9,1.2,12.45) → Excel "#NAME?", IronCalc "#NAME?" |

## Errors

- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddfprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlprice.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/oddlyield.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `financial/rate.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `lookup/image.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mdeterm.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/minverse.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/mmult.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2my2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumx2py2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `math/sumxmy2.xlsx`: ironcalc: _ironcalc.WorkbookError: XML Error: RowError: invalid digit found in string
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv-rt.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/f-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv-2t.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
- `statistical/t-inv.xlsx`: timeout after 55000 ms
