# libreoffice vs spread-sheets@19.2.5 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 21. Fails where both answered an error, but a different one: 61. Fails where the case's own cell matches but its spill does not: 6.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1355 | 579 | 0 | 0 | 64 | 70.1% |
| array | 205 | 185 | 20 | 0 | 0 | 0 | 90.2% |
| database | 100 | 78 | 22 | 0 | 0 | 0 | 78.0% |
| date_time | 485 | 344 | 37 | 0 | 0 | 104 | 90.3% |
| financial | 1075 | 948 | 127 | 0 | 0 | 0 | 88.2% |
| fods | 127 | 112 | 14 | 0 | 0 | 1 | 88.9% |
| information | 153 | 136 | 7 | 1 | 0 | 9 | 94.4% |
| lambda | 108 | 24 | 0 | 0 | 0 | 84 | 100.0% |
| logical | 126 | 123 | 0 | 0 | 0 | 3 | 100.0% |
| mathematical | 1579 | 1378 | 132 | 0 | 0 | 69 | 91.3% |
| spreadsheet | 1756 | 1617 | 95 | 4 | 0 | 40 | 94.2% |
| statistical | 3385 | 2768 | 383 | 12 | 0 | 222 | 87.5% |
| text | 963 | 839 | 105 | 9 | 0 | 10 | 88.0% |
| **total** | 12060 | 9907 | 1521 | 26 | 0 | 606 | 86.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 9907 | 1521 | 26 | 0 | 606 | 86.5% |
| **total** | 12060 | 9907 | 1521 | 26 | 0 | 606 | 86.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| BESSELJ | 182 | fail 182 | `addin/besselj.xlsx!A278` =BESSELJ($L2,X$1) → Excel 0.06623604861483785, SpreadJS "#NUM!" |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", SpreadJS "#NAME?" |
| BESSELK | 157 | fail 157 | `addin/besselk.xlsx!A278` =BESSELK($L2,X$1) → Excel 9.524759804034812e-45, SpreadJS 3.75811940104004e+25 |
| BESSELI | 122 | fail 122 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, SpreadJS "#NUM!" |
| POISSON.DIST | 102 | fail 102 | `statistical/poisson.dist.xlsx!A317` =POISSON.DIST(U596,$V$1,TRUE()) → Excel 7.510390949975846e-301, SpreadJS "#NUM!" |
| POISSON | 102 | fail 102 | `statistical/poisson.xlsx!A305` =POISSON(U584,$V$1,TRUE()) → Excel 3.0112553282088257e-307, SpreadJS "#NUM!" |
| ASC | 71 | fail 71 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", SpreadJS "ｱ゛ｲ゛ｳ゛ｴ゛ｵ゛" |
| AGGREGATE | 64 | fail 64 | `mathematical/aggregate.xlsx!A36` =AGGREGATE(2,8,O2:O9) → Excel "#VALUE!", SpreadJS 6 |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", SpreadJS "#NAME?" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", SpreadJS "#NAME?" |
| GETPIVOTDATA | 33 | fail 33 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, SpreadJS "#REF!" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", SpreadJS "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", SpreadJS "#NAME?" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", SpreadJS "#NAME?" |
| EUROCONVERT | 24 | fail 24 | `mathematical/euroconvert.xlsx!A21` =EUROCONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", SpreadJS 24518 |
| RATE | 18 | fail 18 | `financial/rate.xlsx!A242` =RATE(3,-10,900,0,1,0) → Excel -0.8882334627656058, SpreadJS "#NUM!" |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", SpreadJS "#NAME?" |
| CLEAN | 15 | fail 15 | `text/clean.xlsx!A93` =CLEAN(J93) → Excel "#N/A", SpreadJS "!$Eh1￾E﷐F�$" |
| CONVERT | 14 | fail 14 | `addin/convert.xlsx!A136` =CONVERT(1,H90,G91) → Excel 8.467866646237152e+47, SpreadJS 6.00548995272563e+37 |
| ISLEAPYEAR | 14 | skip 14 | `date_time/isleapyear.xlsx!A11` =ISLEAPYEAR(years) → Excel "#NAME?", SpreadJS "#NAME?" |
| ORG.LIBREOFFICE.FORECAST.ETS.MULT | 14 | skip 14 | `statistical/forecast.ets.mult.xlsx!A212` =ORG.LIBREOFFICE.FORECAST.ETS.MULT($BI$64:$BI100,$BH$5:$BH$64 → Excel "#NAME?", SpreadJS "#NAME?" |
| PERCENTRANK.INC | 14 | fail 14 | `statistical/percentrank.inc.xlsx!A25` =PERCENTRANK.INC(R2:R10,5.43,1) → Excel 0.3, SpreadJS 0.4 |
| ORG.LIBREOFFICE.RAWSUBTRACT | 13 | skip 13 | `mathematical/rawsubtract.xlsx!A10` =ORG.LIBREOFFICE.RAWSUBTRACT(0.987654321098765, 0.98765432109 → Excel "#NAME?", SpreadJS "#NAME?" |
| CHISQINV | 12 | skip 12 | `statistical/chisqinv.xlsx!A36` =CHISQINV(1,1) → Excel "#NAME?", SpreadJS "#NAME?" |
| FORECAST.ETS | 12 | unsupported 12 | `statistical/forecast.ets.add.xlsx!A115` =FORECAST.ETS(AJ64:AJ100,AI5:AI64,AJ5:AJ64,12,1,1) → Excel 66.5299281467668, SpreadJS "#NAME?" |
