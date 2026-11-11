# libreoffice vs google-sheets@2026-11-09 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 18. Fails where both answered an error, but a different one: 82. Fails where the case's own cell matches but its spill does not: 86.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 588 | 62 | 1284 | 0 | 64 | 30.4% |
| array | 205 | 158 | 47 | 0 | 0 | 0 | 77.1% |
| database | 100 | 74 | 26 | 0 | 0 | 0 | 74.0% |
| date_time | 485 | 320 | 61 | 0 | 0 | 104 | 84.0% |
| financial | 1075 | 732 | 281 | 62 | 0 | 0 | 68.1% |
| fods | 127 | 104 | 20 | 2 | 0 | 1 | 82.5% |
| information | 153 | 134 | 9 | 1 | 0 | 9 | 93.1% |
| lambda | 108 | 6 | 18 | 0 | 0 | 84 | 25.0% |
| logical | 126 | 117 | 6 | 0 | 0 | 3 | 95.1% |
| mathematical | 1579 | 1253 | 48 | 209 | 0 | 69 | 83.0% |
| spreadsheet | 1756 | 1382 | 162 | 172 | 0 | 40 | 80.5% |
| statistical | 3385 | 3053 | 98 | 12 | 0 | 222 | 96.5% |
| text | 963 | 770 | 174 | 9 | 0 | 10 | 80.8% |
| **total** | 12060 | 8691 | 1012 | 1751 | 0 | 606 | 75.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 8691 | 1012 | 1751 | 0 | 606 | 75.9% |
| **total** | 12060 | 8691 | 1012 | 1751 | 0 | 606 | 75.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| BESSELI | 321 | unsupported 321 | `addin/besseli.xlsx!A296` =BESSELI($L20,X$1) → Excel 0.002181586982126512, Google Sheets "#NAME?" |
| BESSELJ | 321 | unsupported 321 | `addin/besselj.xlsx!A296` =BESSELJ($L20,X$1) → Excel 0.0005451544437832112, Google Sheets "#NAME?" |
| BESSELK | 321 | unsupported 321 | `addin/besselk.xlsx!A296` =BESSELK($L20,X$1) → Excel 17.07579022834011, Google Sheets "#NAME?" |
| BESSELY | 321 | unsupported 321 | `addin/bessely.xlsx!A296` =BESSELY($L20,X$1) → Excel -56.31681116405603, Google Sheets "#NAME?" |
| AGGREGATE | 209 | unsupported 209 | `mathematical/aggregate.xlsx!A33` =AGGREGATE(2,5,O2:O9) → Excel 6, Google Sheets "#NAME?" |
| ROUND | 168 | fail 168 | `financial/rate.xlsx!A230` =ROUND(RATE(F231,35,-250,0),6) → Excel "#NUM!", Google Sheets 0.14 |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", Google Sheets "#NAME?" |
| ASC | 75 | fail 75 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", Google Sheets "ｱﾞｲﾞｳﾞｴﾞｵﾞ" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", Google Sheets "#N/A" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", Google Sheets "#NAME?" |
| SUMPRODUCT | 34 | fail 34 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", Google Sheets 0 |
| CLEAN | 31 | fail 31 | `text/clean.xlsx!A87` =CLEAN(J87) → Excel "#N/A", Google Sheets "﷐Test text﷐" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", Google Sheets "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", Google Sheets "#NAME?" |
| ODDLYIELD | 25 | unsupported 25 | `financial/oddlyield.xlsx!A22` =ODDLYIELD("1999-04-20","1999-06-15", "1998-10-15", 0.0375, 9 → Excel "#N/A", Google Sheets "#NAME?" |
| PRICE | 25 | fail 25 | `financial/price.xlsx!A50` =PRICE(N45,N46,N47,N48,N49,N50,N51) → Excel 100.06206741734462, Google Sheets 100.140567 |
| TEXTAFTER | 25 | unsupported 25 | `spreadsheet/textafter.xlsx!A22` =TEXTAFTER(F3, ",", 3,,1) → Excel "", Google Sheets "#NAME?" |
| TEXTSPLIT | 25 | unsupported 25 | `spreadsheet/textsplit.xlsx!A43` =TEXTSPLIT(L8,",",,TRUE()) → Excel "One", Google Sheets "#NAME?" |
| CONVERT | 24 | fail 24 | `addin/convert.xlsx!A154` =CONVERT(1,H56,H57) → Excel 1.6343131085957202e-30, Google Sheets "#N/A" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", Google Sheets "#NAME?" |
| TEXTBEFORE | 24 | unsupported 24 | `spreadsheet/textbefore.xlsx!A22` =TEXTBEFORE(F3, ",", 3,,1) → Excel "Brown, Lucas, Manager:1234:5678", Google Sheets "#NAME?" |
| ODDLPRICE | 23 | unsupported 23 | `financial/oddlprice.xlsx!A21` =ODDLPRICE(DATE(1999,2,7),DATE(1999,6,15),DATE(1998,10,15), 0 → Excel 101.33333333333333, Google Sheets "#NAME?" |
| NETWORKDAYS.INTL | 21 | fail 21 | `date_time/networkdays.intl.xlsx!A18` =NETWORKDAYS.INTL(DATE(2014,11,1),DATE(2014,11,30),14,{"2014- → Excel 23, Google Sheets "#VALUE!" |
| CHOOSECOLS | 21 | fail 21 | `spreadsheet/choosecols.xlsx!A143` =CHOOSECOLS(J3:N8, -3, {-3},P2) → Excel 1, Google Sheets 1 |
| WORKDAY.INTL | 20 | fail 20 | `date_time/workday.intl.xlsx!A45` =WORKDAY.INTL(K$45,1,,K46) → Excel 42542, Google Sheets "#VALUE!" |
