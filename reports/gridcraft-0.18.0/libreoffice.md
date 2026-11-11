# libreoffice vs gridcraft@0.18.0 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 7. Fails where both answered an error, but a different one: 53. Fails where the case's own cell matches but its spill does not: 6.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1765 | 169 | 0 | 0 | 64 | 91.3% |
| array | 205 | 171 | 34 | 0 | 0 | 0 | 83.4% |
| database | 100 | 95 | 5 | 0 | 0 | 0 | 95.0% |
| date_time | 485 | 364 | 17 | 0 | 0 | 104 | 95.5% |
| financial | 1075 | 823 | 252 | 0 | 0 | 0 | 76.6% |
| fods | 127 | 114 | 12 | 0 | 0 | 1 | 90.5% |
| information | 153 | 139 | 5 | 0 | 0 | 9 | 96.5% |
| lambda | 108 | 24 | 0 | 0 | 0 | 84 | 100.0% |
| logical | 126 | 123 | 0 | 0 | 0 | 3 | 100.0% |
| mathematical | 1579 | 1454 | 56 | 0 | 0 | 69 | 96.3% |
| spreadsheet | 1756 | 1587 | 89 | 40 | 0 | 40 | 92.5% |
| statistical | 3385 | 3109 | 54 | 0 | 0 | 222 | 98.3% |
| text | 963 | 862 | 91 | 0 | 0 | 10 | 90.5% |
| **total** | 12060 | 10630 | 784 | 40 | 0 | 606 | 92.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 10630 | 784 | 40 | 0 | 606 | 92.8% |
| **total** | 12060 | 10630 | 784 | 40 | 0 | 606 | 92.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ROUND | 165 | fail 165 | `financial/rate.xlsx!A230` =ROUND(RATE(F231,35,-250,0),6) → Excel "#NUM!", GridCraft -1 |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", GridCraft "#NAME?" |
| BESSELI | 121 | fail 121 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, GridCraft 5.2121422658880535e+41 |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", GridCraft "#NAME?" |
| GETPIVOTDATA | 40 | unsupported 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, GridCraft "#NAME?" |
| JIS | 38 | fail 38 | `text/jis.xlsx!A42` =JIS(I2) → Excel "#NAME?", GridCraft "!\"#$%&'()*+,-./" |
| BESSELJ | 37 | fail 37 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, GridCraft 0.27871487343732737 |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", GridCraft "#NAME?" |
| LOOKUP | 32 | fail 32 | `spreadsheet/lookup.xlsx!A368` =LOOKUP(I365,F365:F369,G365:H369) → Excel "#N/A", GridCraft "Andy" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", GridCraft "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", GridCraft "#NAME?" |
| SUMPRODUCT | 25 | fail 25 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", GridCraft 0 |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", GridCraft "#NAME?" |
| EUROCONVERT | 24 | fail 24 | `mathematical/euroconvert.xlsx!A21` =EUROCONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", GridCraft 23964 |
| CLEAN | 23 | fail 23 | `text/clean.xlsx!A90` =CLEAN(J90) → Excel "Test text", GridCraft "Test text" |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", GridCraft "#NAME?" |
| ISLEAPYEAR | 14 | skip 14 | `date_time/isleapyear.xlsx!A11` =ISLEAPYEAR(years) → Excel "#NAME?", GridCraft "#NAME?" |
| ORG.LIBREOFFICE.FORECAST.ETS.MULT | 14 | skip 14 | `statistical/forecast.ets.mult.xlsx!A212` =ORG.LIBREOFFICE.FORECAST.ETS.MULT($BI$64:$BI100,$BH$5:$BH$64 → Excel "#NAME?", GridCraft "#NAME?" |
| ORG.LIBREOFFICE.RAWSUBTRACT | 13 | skip 13 | `mathematical/rawsubtract.xlsx!A10` =ORG.LIBREOFFICE.RAWSUBTRACT(0.987654321098765, 0.98765432109 → Excel "#NAME?", GridCraft "#NAME?" |
| CHISQINV | 12 | skip 12 | `statistical/chisqinv.xlsx!A36` =CHISQINV(1,1) → Excel "#NAME?", GridCraft "#NAME?" |
| ORG.OPENOFFICE.EASTERSUNDAY | 11 | skip 11 | `date_time/eastersunday.xlsx!A9` =ORG.OPENOFFICE.EASTERSUNDAY() → Excel "#NAME?", GridCraft "#NAME?" |
| AGGREGATE | 11 | fail 11 | `mathematical/aggregate.xlsx!A36` =AGGREGATE(2,8,O2:O9) → Excel "#VALUE!", GridCraft 6 |
| CONVERT | 10 | fail 10 | `addin/convert.xlsx!A112` =CONVERT(1,H33,G34) → Excel 624150964712.0417, GridCraft 624150907446.0763 |
| YEARS | 10 | skip 10 | `date_time/years.xlsx!A9` =YEARS(DATE(2014,1,15),DATE(2016,4,1),0) → Excel "#NAME?", GridCraft "#NAME?" |
| COM.MICROSOFT.ARRAYTOTEXT | 10 | skip 10 | `spreadsheet/arraytotext.xlsx!A9` =com.microsoft.arraytotext(J5) → Excel "#NAME?", GridCraft "#NAME?" |
