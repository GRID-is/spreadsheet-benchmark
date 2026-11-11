# univer vs libreoffice@26.8.2.2 (univer)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 51, excel-lacks-function 282. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 694. Fails where the case's own cell matches but its spill does not: 260.
Functions outside Excel's list (23): ARRAY_CONSTRAIN, AVERAGE.WEIGHTED, COMPARE, CUBE, DIVIDED, EPOCHTODATE, FLATTEN, IMAGE, IMCOTH, IMLOG, IMTANH, ISBETWEEN, ISDATE, ISEMAIL, ISURL, MARGINOFERROR, MINUS, MULTIPLY, NUMBERSTRING, PLUS, REGEXMATCH, ROUNDBANK, TO_DATE. Never-supported functions present: ARRAY_CONSTRAIN, FILTERXML, IMAGE.
Upstream literal differs from Excel's value in 440 of 5523 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| __tests__ | 256 | 168 | 6 | 1 | 0 | 81 | 96.0% |
| array | 11 | 0 | 0 | 0 | 0 | 11 | -% |
| compatibility | 107 | 72 | 35 | 0 | 0 | 0 | 67.3% |
| database | 182 | 112 | 70 | 0 | 0 | 0 | 61.5% |
| date | 378 | 221 | 120 | 2 | 0 | 35 | 64.4% |
| engineering | 572 | 276 | 210 | 56 | 0 | 30 | 50.9% |
| financial | 800 | 394 | 405 | 1 | 0 | 0 | 49.3% |
| information | 156 | 90 | 25 | 0 | 0 | 41 | 78.3% |
| logical | 107 | 82 | 11 | 13 | 0 | 1 | 77.4% |
| lookup | 625 | 423 | 191 | 4 | 0 | 7 | 68.4% |
| math | 788 | 548 | 195 | 2 | 0 | 43 | 73.6% |
| meta | 14 | 0 | 0 | 0 | 0 | 14 | -% |
| statistical | 1303 | 838 | 396 | 28 | 0 | 41 | 66.4% |
| text | 549 | 361 | 115 | 44 | 0 | 29 | 69.4% |
| web | 8 | 6 | 2 | 0 | 0 | 0 | 75.0% |
| **total** | 5856 | 3591 | 1781 | 151 | 0 | 333 | 65.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 5856 | 3591 | 1781 | 151 | 0 | 333 | 65.0% |
| **total** | 5856 | 3591 | 1781 | 151 | 0 | 333 | 65.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| MINUS | 34 | skip 34 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =MINUS(A1,1) → Excel "#NAME?", LibreOffice "#NAME?" |
| COMPARE | 33 | skip 33 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =COMPARE(A1:A3,999) → Excel "#NAME?", LibreOffice "#NAME?" |
| XLOOKUP | 29 | fail 27, unsupported 2 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =XLOOKUP(0,0,A1:A9,"No Data") → Excel 101, LibreOffice "#VALUE!" |
| TODAY | 28 | skip 28 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =TODAY() → Excel 46330, LibreOffice 46335 |
| INDEX | 27 | fail 26, unsupported 1 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =INDEX(A1:A3,#REF!) → Excel "#REF!", LibreOffice "#NAME?" |
| AGGREGATE | 25 | fail 25 | `math/aggregate.xlsx!Z1` =AGGREGATE(1,,1) → Excel "#VALUE!", LibreOffice 1 |
| DATEDIF | 23 | fail 23 | `date/datedif.xlsx!Z1` =DATEDIF("2011/1/29","2021/3/31","Y") → Excel 10, LibreOffice "#VALUE!" |
| VLOOKUP | 23 | fail 23 | `lookup/vlookup.xlsx!Z1` =VLOOKUP(8,A1:B8,3,0) → Excel "#REF!", LibreOffice "#VALUE!" |
| ROUNDBANK | 20 | skip 20 | `math/roundbank.xlsx!Z1` =ROUNDBANK(2.344,2) → Excel "#NAME?", LibreOffice "#NAME?" |
| PLUS | 19 | skip 19 | `__tests__/nested-functions.spec.ts.xlsx!Z1` =PLUS(A1,1) → Excel "#NAME?", LibreOffice "#NAME?" |
| EPOCHTODATE | 19 | skip 19 | `date/epochtodate.xlsx!Z1` =EPOCHTODATE(1655906710) → Excel "#NAME?", LibreOffice "#NAME?" |
| AMORLINC | 17 | fail 17 | `financial/amorlinc.xlsx!Z1` =AMORLINC(2400,"2008/8/19","2008/12/31",300,1,0.15,0) → Excel 360, LibreOffice "#VALUE!" |
| DURATION | 17 | fail 17 | `financial/duration.xlsx!Z1` =DURATION("2018-7-1","2048-1-1",0.08,0.09,2,1) → Excel 10.919145281591932, LibreOffice 10.9215739665694 |
| ODDFPRICE | 17 | fail 17 | `financial/oddfprice.xlsx!Z1` =ODDFPRICE("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 113.59771747407883, LibreOffice "#VALUE!" |
| HLOOKUP | 16 | fail 16 | `lookup/hlookup.xlsx!Z1` =HLOOKUP(8,A1:H2,3,0) → Excel "#REF!", LibreOffice "#VALUE!" |
| NUMBERSTRING | 16 | skip 16 | `text/numberstring.xlsx!Z1` =NUMBERSTRING(123,1) → Excel "#N/A", LibreOffice "#NAME?" |
| ODDFYIELD | 15 | fail 15 | `financial/oddfyield.xlsx!Z1` =ODDFYIELD("2008-11-11","2021-3-1","2008-10-15","2009-3-1",0. → Excel 0.10076644980091415, LibreOffice "#VALUE!" |
| ODDLPRICE | 15 | fail 15 | `financial/oddlprice.xlsx!Z1` =ODDLPRICE("2008-11-11","2021-3-1","2008-10-15",0.0785,0.0625 → Excel 110.8745242841996, LibreOffice 110.875429505795 |
| ODDLYIELD | 15 | fail 15 | `financial/oddlyield.xlsx!Z1` =ODDLYIELD("2008-11-11","2021-3-1","2008-10-15",0.0785,84.5,1 → Excel 0.10707208890680206, LibreOffice 0.107079868587002 |
| XIRR | 15 | fail 15 | `financial/xirr.xlsx!Z1` =XIRR(A1:E1,A3:E3,0.1) → Excel "#NUM!", LibreOffice "#VALUE!" |
| XNPV | 15 | fail 15 | `financial/xnpv.xlsx!Z1` =XNPV(0.1,A1:E1,A3:D3) → Excel "#NUM!", LibreOffice "#VALUE!" |
| SUBTOTAL | 15 | fail 15 | `math/subtotal.xlsx!Z1` =SUBTOTAL(1,A1:B2,A4:F5) → Excel 11.157, LibreOffice 9.38083333333333 |
| BINOM.DIST.RANGE | 15 | unsupported 15 | `statistical/binom-dist-range.xlsx!Z1` =BINOM.DIST.RANGE(60,0.75,45,50) → Excel 0.5236297934718872, LibreOffice "#NAME?" |
| REGEXEXTRACT | 15 | unsupported 15 | `text/regexextract.xlsx!Z1` =REGEXEXTRACT("abcdefg","c.*f") → Excel "cdef", LibreOffice "#NAME?" |
| IMLOG | 14 | skip 14 | `engineering/imlog.xlsx!Z1` =IMLOG("5+12i") → Excel "#NAME?", LibreOffice "#NAME?" |
