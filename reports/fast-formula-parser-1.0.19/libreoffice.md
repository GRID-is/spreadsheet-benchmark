# libreoffice vs fast-formula-parser@1.0.19 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 11. Fails where both answered an error, but a different one: 73. Fails where the case's own cell matches but its spill does not: 82.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1531 | 261 | 142 | 0 | 64 | 79.2% |
| array | 205 | 71 | 129 | 5 | 0 | 0 | 34.6% |
| database | 100 | 5 | 0 | 95 | 0 | 0 | 5.0% |
| date_time | 485 | 289 | 91 | 1 | 0 | 104 | 75.9% |
| financial | 1075 | 18 | 2 | 1055 | 0 | 0 | 1.7% |
| fods | 127 | 51 | 1 | 15 | 59 | 1 | 40.5% |
| information | 153 | 101 | 10 | 33 | 0 | 9 | 70.1% |
| lambda | 108 | 11 | 9 | 4 | 0 | 84 | 45.8% |
| logical | 126 | 95 | 18 | 10 | 0 | 3 | 77.2% |
| mathematical | 1579 | 968 | 163 | 379 | 0 | 69 | 64.1% |
| spreadsheet | 1756 | 196 | 58 | 1454 | 8 | 40 | 11.4% |
| statistical | 3385 | 1156 | 269 | 1723 | 15 | 222 | 36.5% |
| text | 963 | 725 | 181 | 40 | 7 | 10 | 76.1% |
| **total** | 12060 | 5217 | 1192 | 4956 | 89 | 606 | 45.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 5217 | 1192 | 4956 | 89 | 606 | 45.5% |
| **total** | 12060 | 5217 | 1192 | 4956 | 89 | 606 | 45.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| LOOKUP | 729 | unsupported 729 | `spreadsheet/lookup.xlsx!A649` =LOOKUP(H607,MMULT(F570:F609,1),MMULT(G570:G609,1)) → Excel 400, fast-formula-parser "#NAME?" |
| POISSON | 329 | unsupported 329 | `statistical/poisson.xlsx!A328` =POISSON(0,0,0) → Excel 1, fast-formula-parser "#NAME?" |
| CONVERT | 284 | unsupported 284 | `addin/convert.xlsx!A151` =CONVERT(1,H98,H99) → Excel "#N/A", fast-formula-parser "#NAME?" |
| AGGREGATE | 209 | unsupported 209 | `mathematical/aggregate.xlsx!A33` =AGGREGATE(2,5,O2:O9) → Excel 6, fast-formula-parser "#NAME?" |
| ROUND | 200 | unsupported 200 | `financial/rate.xlsx!A230` =ROUND(RATE(F231,35,-250,0),6) → Excel "#NUM!", fast-formula-parser "#NAME?" |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| BINOMDIST | 126 | unsupported 126 | `statistical/binomdist.xlsx!A158` =BINOMDIST(G155,$H$74,$H$75,TRUE()) → Excel 3.832021367585077e-272, fast-formula-parser "#NAME?" |
| MATCH | 112 | fail 12, unsupported 100 | `spreadsheet/match.xlsx!A57` =MATCH(F43:F45,F43:F45,0) → Excel 1, fast-formula-parser "#NAME?" |
| CHIDIST | 111 | unsupported 111 | `statistical/chidist.xlsx!A111` =CHIDIST(7.5,0) → Excel "#NUM!", fast-formula-parser "#NAME?" |
| POISSON.DIST | 103 | fail 103 | `statistical/poisson.dist.xlsx!A317` =POISSON.DIST(U596,$V$1,TRUE()) → Excel 7.510390949975846e-301, fast-formula-parser "#VALUE!" |
| BESSELI | 101 | fail 101 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, fast-formula-parser 5.21223018483405e+41 |
| XLOOKUP | 97 | unsupported 97 | `spreadsheet/xlookup.xlsx!A127` =XLOOKUP(N25,K22:K26,L22:L26,XLOOKUP(N25,P22:P26,Q22:Q26)) → Excel 16, fast-formula-parser "#NAME?" |
| VDB | 73 | unsupported 73 | `financial/vdb.xlsx!A73` =VDB(100000,20000,10,4,5,1,1) → Excel 6561.000000000002, fast-formula-parser "#NAME?" |
| ASC | 64 | fail 64 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", fast-formula-parser "ｱﾞｲﾞｳﾞｴﾞｵﾞ" |
| PRICE | 53 | unsupported 53 | `financial/price.xlsx!A50` =PRICE(N45,N46,N47,N48,N49,N50,N51) → Excel 100.06206741734462, fast-formula-parser "#NAME?" |
| LET | 50 | fail 1, unsupported 49 | `lambda/map.xlsx!A14` =LET(double, com.microsoft.lambda(x, x * 2), com.microsoft.ma → Excel "#NAME?", fast-formula-parser "#NAME?" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", fast-formula-parser "#NAME?" |
| GETPIVOTDATA | 40 | unsupported 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, fast-formula-parser "#NAME?" |
| SUMPRODUCT | 39 | fail 39 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", fast-formula-parser 0 |
| RATE | 38 | unsupported 38 | `financial/rate.xlsx!A241` =RATE(0,-100,1000) → Excel "#NUM!", fast-formula-parser "#NAME?" |
| BESSELJ | 37 | fail 37 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, fast-formula-parser 0.27871487333368605 |
| SWITCH | 37 | unsupported 10, error 27 | `fods/Functions_Excel_2016.xlsx!A49` =SWITCH(Sheet2!G55,Sheet2!G48:G54,Sheet2!H48:H54) → Excel "#N/A", fast-formula-parser null |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", fast-formula-parser "#NAME?" |
| SUMIF | 34 | fail 34 | `mathematical/sumif.xlsx!A28` =SUMIF(Z2:Z6,">",AA2:AA6) → Excel 0, fast-formula-parser 1110 |
| COUNTIFS | 34 | unsupported 34 | `statistical/countifs.xlsx!A39` =COUNTIFS($I$38:$I$39,F39) → Excel 0, fast-formula-parser "#NAME?" |

## Errors

- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `fods/Functions_Excel_2016.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `spreadsheet/sheets.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `statistical/maxifs.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
- `text/concat.xlsx`: timeout after 60000 ms
