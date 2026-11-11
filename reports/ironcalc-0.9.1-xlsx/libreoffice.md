# libreoffice vs ironcalc@0.9.1 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 17. Fails where both answered an error, but a different one: 109. Fails where the case's own cell matches but its spill does not: 69.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1848 | 86 | 0 | 0 | 64 | 95.6% |
| array | 205 | 176 | 29 | 0 | 0 | 0 | 85.9% |
| database | 100 | 93 | 7 | 0 | 0 | 0 | 93.0% |
| date_time | 485 | 301 | 80 | 0 | 0 | 104 | 79.0% |
| financial | 1075 | 976 | 99 | 0 | 0 | 0 | 90.8% |
| fods | 127 | 116 | 10 | 0 | 0 | 1 | 92.1% |
| information | 153 | 135 | 9 | 0 | 0 | 9 | 93.8% |
| lambda | 108 | 19 | 5 | 0 | 0 | 84 | 79.2% |
| logical | 126 | 122 | 1 | 0 | 0 | 3 | 99.2% |
| mathematical | 1579 | 1244 | 57 | 209 | 0 | 69 | 82.4% |
| spreadsheet | 1756 | 1350 | 326 | 40 | 0 | 40 | 78.7% |
| statistical | 3385 | 2950 | 213 | 0 | 0 | 222 | 93.3% |
| text | 963 | 830 | 107 | 16 | 0 | 10 | 87.1% |
| **total** | 12060 | 10160 | 1029 | 265 | 0 | 606 | 88.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 10160 | 1029 | 265 | 0 | 606 | 88.7% |
| **total** | 12060 | 10160 | 1029 | 265 | 0 | 606 | 88.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| AGGREGATE | 209 | unsupported 209 | `mathematical/aggregate.xlsx!A33` =AGGREGATE(2,5,O2:O9) → Excel 6, IronCalc "#NAME?" |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", IronCalc "#NAME?" |
| LOOKUP | 95 | fail 95 | `spreadsheet/lookup.xlsx!A715` =LOOKUP(1,1) → Excel 1, IronCalc "#N/A" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", IronCalc "#NAME?" |
| GETPIVOTDATA | 40 | unsupported 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, IronCalc "#NAME?" |
| BESSELJ | 38 | fail 38 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, IronCalc 0.2787148734373273 |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", IronCalc "#NAME?" |
| CLEAN | 32 | fail 32 | `text/clean.xlsx!A93` =CLEAN(J93) → Excel "#N/A", IronCalc "!🀐$🀃E𯀁h🀠1򯿾E﷐F�$" |
| RATE | 30 | fail 30 | `financial/rate.xlsx!A242` =RATE(3,-10,900,0,1,0) → Excel -0.8882334627656058, IronCalc "#ERROR!" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", IronCalc "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", IronCalc "#NAME?" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", IronCalc "#NAME?" |
| CHOOSECOLS | 23 | fail 23 | `spreadsheet/choosecols.xlsx!A143` =CHOOSECOLS(J3:N8, -3, {-3},P2) → Excel 1, IronCalc 1 |
| MATCH | 23 | fail 23 | `spreadsheet/match.xlsx!A57` =MATCH(F43:F45,F43:F45,0) → Excel 1, IronCalc "#N/A" |
| ASC | 21 | fail 21 | `text/asc.xlsx!A45` =ASC(M6) → Excel "［￥］＾＿‘", IronCalc "[￥]^_‘" |
| NETWORKDAYS.INTL | 19 | fail 19 | `date_time/networkdays.intl.xlsx!A18` =NETWORKDAYS.INTL(DATE(2014,11,1),DATE(2014,11,30),14,{"2014- → Excel 23, IronCalc "#VALUE!" |
| SUMPRODUCT | 18 | fail 18 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", IronCalc 0 |
| SUM | 16 | fail 16 | `array/frequency.xlsx!A61` =SUM(ABS(FREQUENCY(K40:K59,L40:L46))) → Excel 20, IronCalc 0 |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", IronCalc "#NAME?" |
| BAHTTEXT | 16 | unsupported 16 | `text/bahttext.xlsx!A13` =BAHTTEXT(12) → Excel "สิบสองบาทถ้วน", IronCalc "#NAME?" |
| XLOOKUP | 15 | fail 15 | `spreadsheet/xlookup.xlsx!A134` =XLOOKUP("B",{"A";"B"},{1;2},"") → Excel 2, IronCalc "#N/A" |
| ISLEAPYEAR | 14 | skip 14 | `date_time/isleapyear.xlsx!A11` =ISLEAPYEAR(years) → Excel "#NAME?", IronCalc "#NAME?" |
| ORG.LIBREOFFICE.FORECAST.ETS.MULT | 14 | skip 14 | `statistical/forecast.ets.mult.xlsx!A212` =ORG.LIBREOFFICE.FORECAST.ETS.MULT($BI$64:$BI100,$BH$5:$BH$64 → Excel "#NAME?", IronCalc "#NAME?" |
| WORKDAY.INTL | 13 | fail 13 | `date_time/workday.intl.xlsx!A41` =WORKDAY.INTL(K$2,K$3,K41) → Excel "#NUM!", IronCalc 41950 |
| ORG.LIBREOFFICE.RAWSUBTRACT | 13 | skip 13 | `mathematical/rawsubtract.xlsx!A10` =ORG.LIBREOFFICE.RAWSUBTRACT(0.987654321098765, 0.98765432109 → Excel "#NAME?", IronCalc "#NAME?" |
