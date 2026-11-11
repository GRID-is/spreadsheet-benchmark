# libreoffice vs univer@1.4.0 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 14. Fails where both answered an error, but a different one: 30. Fails where the case's own cell matches but its spill does not: 66.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1705 | 229 | 0 | 0 | 64 | 88.2% |
| array | 205 | 169 | 36 | 0 | 0 | 0 | 82.4% |
| database | 100 | 84 | 6 | 10 | 0 | 0 | 84.0% |
| date_time | 485 | 288 | 72 | 21 | 0 | 104 | 75.6% |
| financial | 1075 | 1011 | 50 | 14 | 0 | 0 | 94.0% |
| fods | 127 | 106 | 20 | 0 | 0 | 1 | 84.1% |
| information | 153 | 134 | 8 | 2 | 0 | 9 | 93.1% |
| lambda | 108 | 8 | 16 | 0 | 0 | 84 | 33.3% |
| logical | 126 | 118 | 5 | 0 | 0 | 3 | 95.9% |
| mathematical | 1579 | 1393 | 91 | 26 | 0 | 69 | 92.3% |
| spreadsheet | 1756 | 1431 | 250 | 20 | 15 | 40 | 83.4% |
| statistical | 3385 | 2890 | 243 | 30 | 0 | 222 | 91.4% |
| text | 963 | 839 | 114 | 0 | 0 | 10 | 88.0% |
| **total** | 12060 | 10176 | 1140 | 123 | 15 | 606 | 88.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 10176 | 1140 | 123 | 15 | 606 | 88.8% |
| **total** | 12060 | 10176 | 1140 | 123 | 15 | 606 | 88.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", Univer "#NAME?" |
| POISSON.DIST | 102 | fail 102 | `statistical/poisson.dist.xlsx!A317` =POISSON.DIST(U596,$V$1,TRUE()) → Excel 7.510390949975846e-301, Univer null |
| POISSON | 102 | fail 102 | `statistical/poisson.xlsx!A305` =POISSON(U584,$V$1,TRUE()) → Excel 3.0112553282088257e-307, Univer null |
| BESSELI | 100 | fail 100 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, Univer 5.21223018483405e+41 |
| LOOKUP | 76 | fail 76 | `spreadsheet/lookup.xlsx!A715` =LOOKUP(1,1) → Excel 1, Univer "#VALUE!" |
| CONVERT | 56 | fail 56 | `addin/convert.xlsx!A156` =CONVERT(1,H27,H28) → Excel 0.00000263046611952862, Univer 2.19205509919e-7 |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", Univer "#VALUE!" |
| AGGREGATE | 39 | fail 39 | `mathematical/aggregate.xlsx!A27` =AGGREGATE(4,5,O2:O9) → Excel 34, Univer "#DIV/0!" |
| WORKDAY.INTL | 38 | fail 38 | `date_time/workday.intl.xlsx!A45` =WORKDAY.INTL(K$45,1,,K46) → Excel 42542, Univer "#NUM!" |
| CLEAN | 37 | fail 37 | `text/clean.xlsx!A93` =CLEAN(J93) → Excel "#N/A", Univer "!$Eh1￾E﷐F�$" |
| BESSELJ | 36 | fail 36 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, Univer 0.278714873334 |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", Univer "#NAME?" |
| GETPIVOTDATA | 31 | fail 31 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, Univer "#VALUE!" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", Univer "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", Univer "#NAME?" |
| SUMPRODUCT | 25 | fail 25 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", Univer 0 |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", Univer "#NAME?" |
| ISO.CEILING | 23 | unsupported 23 | `mathematical/iso.ceiling.xlsx!A24` =ISO.CEILING(-4.3) → Excel -4, Univer "#NAME?" |
| ASC | 21 | fail 21 | `text/asc.xlsx!A45` =ASC(M6) → Excel "［￥］＾＿‘", Univer "[￥]^_‘" |
| MATCH | 20 | fail 20 | `spreadsheet/match.xlsx!A144` =MATCH($D$150,$150:$150,-1) → Excel "#N/A", Univer 4 |
| REPLACEB | 18 | fail 18 | `text/replaceb.xlsx!A36` =REPLACEB("🙁🙂A",3,2,"!") → Excel "🙁!A", Univer "🙁!\ude42A" |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", Univer "#NAME?" |
| HYPERLINK | 16 | fail 16 | `spreadsheet/hyperlink.xlsx!A32` =HYPERLINK("#Sheet2.A1","Jump to cell A1 on Sheet2") → Excel "Jump to cell A1 on Sheet2", Univer null |
| TAKE | 15 | error 15 | `spreadsheet/take.xlsx!A60` =TAKE(J3:L8, 2, 2) → Excel "January", Univer null |
| TEXTSPLIT | 15 | fail 15 | `spreadsheet/textsplit.xlsx!A44` =TEXTSPLIT(L7, ,M2:M5) → Excel "", Univer "" |

## Errors

- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
- `spreadsheet/take.xlsx`: worker crashed: Cannot read properties of undefined (reading 'map')
