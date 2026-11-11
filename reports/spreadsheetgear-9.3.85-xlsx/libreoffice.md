# libreoffice vs spreadsheetgear@9.3.85 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 452, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 42. Fails where the case's own cell matches but its spill does not: 1.
Functions outside Excel's list (28): B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2563 of 11525 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1673 | 261 | 0 | 5 | 59 | 86.3% |
| array | 205 | 196 | 9 | 0 | 0 | 0 | 95.6% |
| database | 100 | 91 | 0 | 9 | 0 | 0 | 91.0% |
| date_time | 485 | 346 | 35 | 0 | 0 | 104 | 90.8% |
| financial | 1075 | 1002 | 73 | 0 | 0 | 0 | 93.2% |
| fods | 127 | 67 | 0 | 59 | 0 | 1 | 53.2% |
| information | 153 | 140 | 0 | 4 | 0 | 9 | 97.2% |
| lambda | 108 | 14 | 0 | 0 | 61 | 33 | 18.7% |
| logical | 126 | 97 | 0 | 26 | 0 | 3 | 78.9% |
| mathematical | 1579 | 1505 | 5 | 0 | 0 | 69 | 99.7% |
| spreadsheet | 1756 | 1128 | 43 | 462 | 98 | 25 | 65.2% |
| statistical | 3385 | 3066 | 62 | 35 | 0 | 222 | 96.9% |
| text | 963 | 818 | 19 | 116 | 0 | 10 | 85.8% |
| **total** | 12060 | 10143 | 507 | 711 | 164 | 535 | 88.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 10143 | 507 | 711 | 164 | 535 | 88.0% |
| **total** | 12060 | 10143 | 507 | 711 | 164 | 535 | 88.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| BESSELJ | 134 | fail 134 | `addin/besselj.xlsx!A296` =BESSELJ($L20,X$1) → Excel 0.0005451544437832112, SpreadsheetGear 0.00007963559010670835 |
| BESSELI | 121 | fail 121 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, SpreadsheetGear 4.214423868887061e+41 |
| XLOOKUP | 97 | unsupported 97 | `spreadsheet/xlookup.xlsx!A127` =XLOOKUP(N25,K22:K26,L22:L26,XLOOKUP(N25,P22:P26,Q22:Q26)) → Excel 16, SpreadsheetGear "#NAME?" |
| ASC | 79 | unsupported 79 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", SpreadsheetGear "#NAME?" |
| LET | 64 | error 64 | `lambda/lambda.xlsx!A20` =LET(f, com.microsoft.lambda(x, x * x), _xlpm.f(4)) → Excel "#NAME?", SpreadsheetGear null |
| COM.MICROSOFT.LAMBDA | 47 | error 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", SpreadsheetGear null |
| GETPIVOTDATA | 40 | unsupported 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, SpreadsheetGear "#NAME?" |
| SWITCH | 37 | unsupported 37 | `fods/Functions_Excel_2016.xlsx!A49` =SWITCH(Sheet2!G55,Sheet2!G48:G54,Sheet2!H48:H54) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| XMATCH | 32 | unsupported 32 | `spreadsheet/xmatch.xlsx!A35` =XMATCH("chael",F$13:F$19,2,1) → Excel "#N/A", SpreadsheetGear "#NAME?" |
| CHOOSECOLS | 31 | unsupported 31 | `spreadsheet/choosecols.xlsx!A143` =CHOOSECOLS(J3:N8, -3, {-3},P2) → Excel 1, SpreadsheetGear "#NAME?" |
| CHOOSEROWS | 31 | unsupported 31 | `spreadsheet/chooserows.xlsx!A60` =CHOOSEROWS(J3:L8, -5, {-5,-5}) → Excel "February", SpreadsheetGear "#NAME?" |
| IFS | 28 | unsupported 28 | `fods/Functions_Excel_2016.xlsx!A32` =IFS(0,1/0,1/0,2) → Excel "#DIV/0!", SpreadsheetGear "#NAME?" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| TEXTAFTER | 25 | unsupported 25 | `spreadsheet/textafter.xlsx!A22` =TEXTAFTER(F3, ",", 3,,1) → Excel "", SpreadsheetGear "#NAME?" |
| TEXTSPLIT | 25 | unsupported 25 | `spreadsheet/textsplit.xlsx!A43` =TEXTSPLIT(L8,",",,TRUE()) → Excel "One", SpreadsheetGear "#NAME?" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| TEXTBEFORE | 24 | unsupported 24 | `spreadsheet/textbefore.xlsx!A22` =TEXTBEFORE(F3, ",", 3,,1) → Excel "Brown, Lucas, Manager:1234:5678", SpreadsheetGear "#NAME?" |
| LOOKUP | 23 | fail 23 | `spreadsheet/lookup.xlsx!A700` =LOOKUP(F701,{"A","B","C"},{"Andy"}) → Excel "#N/A", SpreadsheetGear "#VALUE!" |
| MAXIFS | 21 | unsupported 21 | `fods/Functions_Excel_2016.xlsx!A8` =MAXIFS(data!A1:A6,data!B1:B6,1) → Excel 91, SpreadsheetGear "#NAME?" |
| NETWORKDAYS.INTL | 20 | fail 20 | `date_time/networkdays.intl.xlsx!A18` =NETWORKDAYS.INTL(DATE(2014,11,1),DATE(2014,11,30),14,{"2014- → Excel 23, SpreadsheetGear 26 |
| ERROR.TYPE | 20 | error 20 | `lambda/isomitted.xlsx!A11` =ERROR.TYPE(F11) → Excel 5, SpreadsheetGear null |
| SORT | 20 | unsupported 20 | `spreadsheet/sort.xlsx!A96` =SORT($N$4:$P$11,{1,3},{1,0}) → Excel "#VALUE!", SpreadsheetGear "#NAME?" |

## Errors

- `addin/rot13.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `addin/rot13.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `addin/rot13.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `addin/rot13.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `addin/rot13.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/isomitted.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `lambda/lambda.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/error.type.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/errortype.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
- `spreadsheet/let.xlsx`: spreadsheetgear: System.IO.IOException: Corrupt OpenXML document.
