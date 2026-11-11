# libreoffice vs poi@6.0.0 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 13. Fails where both answered an error, but a different one: 98. Fails where the case's own cell matches but its spill does not: 7.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 414 | 59 | 1461 | 0 | 64 | 21.4% |
| array | 205 | 116 | 14 | 75 | 0 | 0 | 56.6% |
| database | 100 | 78 | 20 | 2 | 0 | 0 | 78.0% |
| date_time | 485 | 175 | 57 | 149 | 0 | 104 | 45.9% |
| financial | 1075 | 304 | 91 | 680 | 0 | 0 | 28.3% |
| fods | 127 | 84 | 8 | 34 | 0 | 1 | 66.7% |
| information | 153 | 99 | 15 | 30 | 0 | 9 | 68.8% |
| lambda | 108 | 1 | 23 | 0 | 0 | 84 | 4.2% |
| logical | 126 | 106 | 3 | 14 | 0 | 3 | 86.2% |
| mathematical | 1579 | 901 | 59 | 550 | 0 | 69 | 59.7% |
| spreadsheet | 1756 | 1163 | 208 | 345 | 0 | 40 | 67.8% |
| statistical | 3385 | 1411 | 117 | 1635 | 0 | 222 | 44.6% |
| text | 963 | 427 | 141 | 385 | 0 | 10 | 44.8% |
| **total** | 12060 | 5279 | 815 | 5360 | 0 | 606 | 46.1% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 5279 | 815 | 5360 | 0 | 606 | 46.1% |
| **total** | 12060 | 5279 | 815 | 5360 | 0 | 606 | 46.1% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| BESSELI | 321 | unsupported 321 | `addin/besseli.xlsx!A296` =BESSELI($L20,X$1) → Excel 0.002181586982126512, Apache POI "#NAME?" |
| BESSELK | 321 | unsupported 321 | `addin/besselk.xlsx!A296` =BESSELK($L20,X$1) → Excel 17.07579022834011, Apache POI "#NAME?" |
| BESSELY | 321 | unsupported 321 | `addin/bessely.xlsx!A296` =BESSELY($L20,X$1) → Excel -56.31681116405603, Apache POI "#NAME?" |
| CONVERT | 284 | unsupported 284 | `addin/convert.xlsx!A151` =CONVERT(1,H98,H99) → Excel "#N/A", Apache POI "#NAME?" |
| AGGREGATE | 209 | unsupported 209 | `mathematical/aggregate.xlsx!A33` =AGGREGATE(2,5,O2:O9) → Excel 6, Apache POI "#NAME?" |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", Apache POI "#NAME?" |
| BINOMDIST | 126 | unsupported 126 | `statistical/binomdist.xlsx!A158` =BINOMDIST(G155,$H$74,$H$75,TRUE()) → Excel 3.832021367585077e-272, Apache POI "#NAME?" |
| DATEDIF | 111 | unsupported 111 | `date_time/datedif.xlsx!A99` =DATEDIF(F99,G99,H99) → Excel 0, Apache POI "#NAME?" |
| CHIDIST | 111 | unsupported 111 | `statistical/chidist.xlsx!A111` =CHIDIST(7.5,0) → Excel "#NUM!", Apache POI "#NAME?" |
| FISHER | 99 | unsupported 99 | `statistical/fisher.xlsx!A95` =FISHER(K51) → Excel -5.551115123125783e-17, Apache POI "#NAME?" |
| FISHERINV | 99 | unsupported 99 | `statistical/fisherinv.xlsx!A95` =FISHERINV(K51) → Excel -1.1102230246251565e-16, Apache POI "#NAME?" |
| CODE | 98 | fail 98 | `text/code.xlsx!A96` =CODE(I96) → Excel 127, Apache POI "127" |
| CLEAN | 88 | fail 1, unsupported 87 | `text/clean.xlsx!A91` =CLEAN(J91) → Excel "Test text", Apache POI "#NAME?" |
| ASC | 79 | unsupported 79 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", Apache POI "#NAME?" |
| VDB | 73 | unsupported 73 | `financial/vdb.xlsx!A73` =VDB(100000,20000,10,4,5,1,1) → Excel 6561.000000000002, Apache POI "#NAME?" |
| LOOKUP | 72 | fail 72 | `spreadsheet/lookup.xlsx!A713` =LOOKUP(4, I715:K715*2, I717:K717/3) → Excel 7.333333333333333, Apache POI "#VALUE!" |
| LET | 65 | fail 65 | `lambda/lambda.xlsx!A20` =LET(f, com.microsoft.lambda(x, x * x), _xlpm.f(4)) → Excel "#NAME?", Apache POI "#VALUE!" |
| GAMMALN.PRECISE | 60 | unsupported 60 | `fods/gammaln.precise.xlsx!A26` =GAMMALN.PRECISE(I25) → Excel 25.191221182738683, Apache POI "#NAME?" |
| PRICE | 53 | unsupported 53 | `financial/price.xlsx!A50` =PRICE(N45,N46,N47,N48,N49,N50,N51) → Excel 100.06206741734462, Apache POI "#NAME?" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", Apache POI "#VALUE!" |
| HYPGEOM.DIST | 41 | unsupported 41 | `statistical/hypgeom.dist.xlsx!A27` =HYPGEOM.DIST(2,2,190,100,0) → Excel "#NUM!", Apache POI "#NAME?" |
| GETPIVOTDATA | 40 | unsupported 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, Apache POI "#NAME?" |
| BESSELJ | 39 | fail 39 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, Apache POI 0.27871487343732737 |
| HARMEAN | 37 | unsupported 37 | `fods/reference_operators.xlsx!A87` =HARMEAN(B1:D5,B1:D5) → Excel 4.015686274509804, Apache POI "#NAME?" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", Apache POI "#NAME?" |
