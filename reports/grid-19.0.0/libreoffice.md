# libreoffice vs Apiary 19.0.0 (e018d83eb8c6a829)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 4. Fails where both answered an error, but a different one: 4. Fails where the case's own cell matches but its spill does not: 3.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1933 | 1 | 0 | 0 | 64 | 99.9% |
| array | 205 | 205 | 0 | 0 | 0 | 0 | 100.0% |
| database | 100 | 87 | 4 | 9 | 0 | 0 | 87.0% |
| date_time | 485 | 376 | 5 | 0 | 0 | 104 | 98.7% |
| financial | 1075 | 1065 | 10 | 0 | 0 | 0 | 99.1% |
| fods | 127 | 126 | 0 | 0 | 0 | 1 | 100.0% |
| information | 153 | 144 | 0 | 0 | 0 | 9 | 100.0% |
| lambda | 108 | 24 | 0 | 0 | 0 | 84 | 100.0% |
| logical | 126 | 123 | 0 | 0 | 0 | 3 | 100.0% |
| mathematical | 1579 | 1508 | 2 | 0 | 0 | 69 | 99.9% |
| spreadsheet | 1756 | 1703 | 13 | 0 | 0 | 40 | 99.2% |
| statistical | 3385 | 3161 | 2 | 0 | 0 | 222 | 99.9% |
| text | 963 | 932 | 21 | 0 | 0 | 10 | 97.8% |
| **total** | 12060 | 11387 | 58 | 9 | 0 | 606 | 99.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 11387 | 58 | 9 | 0 | 606 | 99.4% |
| **total** | 12060 | 11387 | 58 | 9 | 0 | 606 | 99.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", GRID "#NAME?" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", GRID "#NAME?" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", GRID "#NAME?" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", GRID "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", GRID "#NAME?" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", GRID "#NAME?" |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", GRID "#NAME?" |
| ISLEAPYEAR | 14 | skip 14 | `date_time/isleapyear.xlsx!A11` =ISLEAPYEAR(years) → Excel "#NAME?", GRID "#NAME?" |
| ORG.LIBREOFFICE.FORECAST.ETS.MULT | 14 | skip 14 | `statistical/forecast.ets.mult.xlsx!A212` =ORG.LIBREOFFICE.FORECAST.ETS.MULT($BI$64:$BI100,$BH$5:$BH$64 → Excel "#NAME?", GRID "#NAME?" |
| ORG.LIBREOFFICE.RAWSUBTRACT | 13 | skip 13 | `mathematical/rawsubtract.xlsx!A10` =ORG.LIBREOFFICE.RAWSUBTRACT(0.987654321098765, 0.98765432109 → Excel "#NAME?", GRID "#NAME?" |
| CHISQINV | 12 | skip 12 | `statistical/chisqinv.xlsx!A36` =CHISQINV(1,1) → Excel "#NAME?", GRID "#NAME?" |
| ORG.OPENOFFICE.EASTERSUNDAY | 11 | skip 11 | `date_time/eastersunday.xlsx!A9` =ORG.OPENOFFICE.EASTERSUNDAY() → Excel "#NAME?", GRID "#NAME?" |
| YEARS | 10 | skip 10 | `date_time/years.xlsx!A9` =YEARS(DATE(2014,1,15),DATE(2016,4,1),0) → Excel "#NAME?", GRID "#NAME?" |
| COM.MICROSOFT.ARRAYTOTEXT | 10 | skip 10 | `spreadsheet/arraytotext.xlsx!A9` =com.microsoft.arraytotext(J5) → Excel "#NAME?", GRID "#NAME?" |
| ORG.LIBREOFFICE.REGEX | 10 | skip 10 | `text/regex.xlsx!A9` =ORG.LIBREOFFICE.REGEX("The end", "^The", "AA","g") → Excel "#NAME?", GRID "#NAME?" |
| COM.MICROSOFT.BYROW | 9 | skip 9 | `lambda/byrow.xlsx!A17` =com.microsoft.byrow({"A","B","C";"D","E","F"}, com.microsoft → Excel "#NAME?", GRID "#NAME?" |
| COM.MICROSOFT.REDUCE | 9 | skip 9 | `lambda/reduce.xlsx!A11` =com.microsoft.reduce("", {"a","b","c"}, com.microsoft.concat → Excel "#NAME?", GRID "#NAME?" |
| REPLACE | 9 | fail 9 | `text/replace.xlsx!A21` =REPLACE(I12,1,3,"xYx") → Excel "xYxc😂d", GRID "xYx😂d" |
| _XLUDF.IMCSCH | 8 | skip 8 | `addin/imcsch.xlsx!A19` =_xludf.IMCSCH(4) → Excel "#NAME?", GRID "#NAME?" |
| _XLUDF.IMSEC | 8 | skip 8 | `addin/imsec.xlsx!A19` =_xludf.IMSEC(4) → Excel "#NAME?", GRID "#NAME?" |
| _XLUDF.IMSECH | 8 | skip 8 | `addin/imsech.xlsx!A19` =_xludf.IMSECH(4) → Excel "#NAME?", GRID "#NAME?" |
| DAYSINMONTH | 8 | skip 8 | `date_time/daysinmonth.xlsx!A7` =DAYSINMONTH("a") → Excel "#NAME?", GRID "#NAME?" |
| MONTHS | 8 | skip 8 | `date_time/months.xlsx!A7` =MONTHS(J7, K7, 1) → Excel "#NAME?", GRID "#NAME?" |
| WEEKSINYEAR | 8 | skip 8 | `date_time/weeksinyear.xlsx!A7` =WEEKSINYEAR(K3,) → Excel "#NAME?", GRID "#NAME?" |
| _XLUDF.IMCOSH | 7 | skip 7 | `addin/imcosh.xlsx!A6` =_xludf.IMCOSH(R1:R12) → Excel "#NAME?", GRID "#NAME?" |
