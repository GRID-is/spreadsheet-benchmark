# libreoffice vs gnumeric@1.12.62 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 9. Fails where both answered an error, but a different one: 137. Fails where the case's own cell matches but its spill does not: 2.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1561 | 301 | 0 | 72 | 64 | 80.7% |
| array | 205 | 187 | 18 | 0 | 0 | 0 | 91.2% |
| database | 100 | 81 | 19 | 0 | 0 | 0 | 81.0% |
| date_time | 485 | 339 | 42 | 0 | 0 | 104 | 89.0% |
| financial | 1075 | 720 | 348 | 7 | 0 | 0 | 67.0% |
| fods | 127 | 113 | 13 | 0 | 0 | 1 | 89.7% |
| information | 153 | 132 | 12 | 0 | 0 | 9 | 91.7% |
| lambda | 108 | 2 | 22 | 0 | 0 | 84 | 8.3% |
| logical | 126 | 112 | 11 | 0 | 0 | 3 | 91.1% |
| mathematical | 1579 | 931 | 134 | 445 | 0 | 69 | 61.7% |
| spreadsheet | 1756 | 1324 | 186 | 206 | 0 | 40 | 77.2% |
| statistical | 3385 | 2990 | 115 | 58 | 0 | 222 | 94.5% |
| text | 963 | 702 | 235 | 16 | 0 | 10 | 73.7% |
| **total** | 12060 | 9194 | 1456 | 732 | 72 | 606 | 80.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 9194 | 1456 | 732 | 72 | 606 | 80.3% |
| **total** | 12060 | 9194 | 1456 | 732 | 72 | 606 | 80.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| AGGREGATE | 209 | unsupported 209 | `mathematical/aggregate.xlsx!A33` =AGGREGATE(2,5,O2:O9) → Excel 6, Gnumeric "#NAME?" |
| ROUND | 165 | fail 165 | `financial/rate.xlsx!A230` =ROUND(RATE(F231,35,-250,0),6) → Excel "#NUM!", Gnumeric 0.14 |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", Gnumeric "#NAME?" |
| BESSELI | 123 | fail 123 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, Gnumeric 5.212142265888018e+41 |
| FLOOR.MATH | 95 | unsupported 95 | `mathematical/floor.xlsx!A59` =FLOOR.MATH(J59) → Excel -8, Gnumeric "#NAME?" |
| ASC | 75 | fail 75 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", Gnumeric "ｱﾞｲﾞｳﾞｴﾞｵﾞ" |
| CONVERT | 72 | fail 72 | `addin/convert.xlsx!A143` =CONVERT(1,H96,G97) → Excel "#N/A", Gnumeric "#NUM!" |
| LET | 65 | fail 65 | `lambda/lambda.xlsx!A20` =LET(f, com.microsoft.lambda(x, x * x), _xlpm.f(4)) → Excel "#NAME?", Gnumeric "#\"_xlfn.LET(_xlpm.f, com.microsoft.lambda(_xlpm.x, _xlpm.x * _xlpm.x), _xlpm.f(4))\"" |
| CEILING.MATH | 60 | unsupported 60 | `mathematical/ceiling.xlsx!A25` =CEILING.MATH(K$10,$L10,0) → Excel -23, Gnumeric "#NAME?" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", Gnumeric "#\"com.microsoft.lambda(_xlpm.a,_xlpm.b,_xlpm.c, IF(com.microsoft.isomitted(_xlpm.b), -1, _xlpm.b))(1, , 3)\"" |
| GETPIVOTDATA | 40 | fail 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, Gnumeric "#REF!" |
| BESSELJ | 39 | fail 39 | `addin/besselj.xlsx!A326` =BESSELJ(1,-9) → Excel "#NUM!", Gnumeric -5.2492501799118765e-9 |
| JIS | 38 | fail 38 | `text/jis.xlsx!A42` =JIS(I2) → Excel "#NAME?", Gnumeric "！”＃＄％＆’（）＊＋，－．／" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", Gnumeric "#NAME?" |
| CLEAN | 33 | fail 33 | `text/clean.xlsx!A93` =CLEAN(J93) → Excel "#N/A", Gnumeric "!🀐$🀃E𯀁h🀠1򯿾E﷐F�$" |
| CHOOSECOLS | 31 | unsupported 31 | `spreadsheet/choosecols.xlsx!A143` =CHOOSECOLS(J3:N8, -3, {-3},P2) → Excel 1, Gnumeric "#NAME?" |
| CHOOSEROWS | 31 | unsupported 31 | `spreadsheet/chooserows.xlsx!A60` =CHOOSEROWS(J3:L8, -5, {-5,-5}) → Excel "February", Gnumeric "#NAME?" |
| REPLACEB | 29 | fail 29 | `text/replaceb.xlsx!A36` =REPLACEB("🙁🙂A",3,2,"!") → Excel "🙁!A", Gnumeric "#VALUE!" |
| DEC2OCT | 28 | error 28 | `addin/dec2oct.xlsx!A21` =DEC2OCT(30, 4) → Excel "0036", Gnumeric null |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", Gnumeric "#NAME?" |
| DEC2HEX | 26 | error 26 | `addin/dec2hex.xlsx!A19` =DEC2HEX(30) → Excel "1E", Gnumeric null |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", Gnumeric "#NAME?" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", Gnumeric "#NAME?" |
| EUROCONVERT | 24 | fail 24 | `mathematical/euroconvert.xlsx!A21` =EUROCONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", Gnumeric 23964 |
| CEILING.PRECISE | 23 | unsupported 23 | `mathematical/ceiling.precise.xlsx!A24` =CEILING.PRECISE(-4.3) → Excel -4, Gnumeric "#NAME?" |

## Errors

- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2bin.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12480): CRITICAL **: 03:25:24.404: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2hex.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12483): CRITICAL **: 03:25:24.686: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
- `addin/dec2oct.xlsx`: ssconvert wrote nothing (killed by SIGSEGV): ** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_get_node: assertion 'parent || key' failed

** (ssconvert:12486): CRITICAL **: 03:25:24.968: go_conf_add_monitor: assertion 'node || key' fai
