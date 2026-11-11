# libreoffice vs hyperformula@3.5.0 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 14. Fails where both answered an error, but a different one: 75. Fails where the case's own cell matches but its spill does not: 31.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1529 | 252 | 153 | 0 | 64 | 79.1% |
| array | 205 | 68 | 97 | 40 | 0 | 0 | 33.2% |
| database | 100 | 66 | 24 | 10 | 0 | 0 | 66.0% |
| date_time | 485 | 156 | 204 | 21 | 0 | 104 | 40.9% |
| financial | 1075 | 608 | 51 | 416 | 0 | 0 | 56.6% |
| fods | 127 | 53 | 12 | 2 | 59 | 1 | 42.1% |
| information | 153 | 118 | 17 | 9 | 0 | 9 | 81.9% |
| lambda | 108 | 5 | 15 | 4 | 0 | 84 | 20.8% |
| logical | 126 | 117 | 6 | 0 | 0 | 3 | 95.1% |
| mathematical | 1579 | 1086 | 69 | 355 | 0 | 69 | 71.9% |
| spreadsheet | 1756 | 402 | 168 | 1146 | 0 | 40 | 23.4% |
| statistical | 3385 | 2570 | 312 | 281 | 0 | 222 | 81.3% |
| text | 963 | 642 | 62 | 249 | 0 | 10 | 67.4% |
| **total** | 12060 | 7420 | 1289 | 2686 | 59 | 606 | 64.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 7420 | 1289 | 2686 | 59 | 606 | 64.8% |
| **total** | 12060 | 7420 | 1289 | 2686 | 59 | 606 | 64.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| LOOKUP | 729 | unsupported 729 | `spreadsheet/lookup.xlsx!A649` =LOOKUP(H607,MMULT(F570:F609,1),MMULT(G570:G609,1)) → Excel 400, HyperFormula "#NAME?" |
| CONVERT | 284 | fail 2, unsupported 282 | `addin/convert.xlsx!A151` =CONVERT(1,H98,H99) → Excel "#N/A", HyperFormula "#NAME?" |
| AGGREGATE | 209 | unsupported 209 | `mathematical/aggregate.xlsx!A33` =AGGREGATE(2,5,O2:O9) → Excel 6, HyperFormula "#NAME?" |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", HyperFormula "#NAME?" |
| DATEDIF | 111 | fail 111 | `date_time/datedif.xlsx!A99` =DATEDIF(F99,G99,H99) → Excel 0, HyperFormula "#NUM!" |
| POISSON.DIST | 102 | fail 102 | `statistical/poisson.dist.xlsx!A317` =POISSON.DIST(U596,$V$1,TRUE()) → Excel 7.510390949975846e-301, HyperFormula "#NUM!" |
| POISSON | 102 | fail 102 | `statistical/poisson.xlsx!A305` =POISSON(U584,$V$1,TRUE()) → Excel 3.0112553282088257e-307, HyperFormula "#NUM!" |
| BESSELI | 100 | fail 100 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, HyperFormula 5.2122301848e+41 |
| ASC | 79 | unsupported 79 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", HyperFormula "#NAME?" |
| VDB | 73 | unsupported 73 | `financial/vdb.xlsx!A73` =VDB(100000,20000,10,4,5,1,1) → Excel 6561.000000000002, HyperFormula "#NAME?" |
| LET | 64 | fail 16, unsupported 48 | `lambda/lambda.xlsx!A20` =LET(f, com.microsoft.lambda(x, x * x), _xlpm.f(4)) → Excel "#NAME?", HyperFormula "#ERROR!" |
| PRICE | 53 | unsupported 53 | `financial/price.xlsx!A50` =PRICE(N45,N46,N47,N48,N49,N50,N51) → Excel 100.06206741734462, HyperFormula "#NAME?" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", HyperFormula "#ERROR!" |
| GETPIVOTDATA | 40 | unsupported 40 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, HyperFormula "#NAME?" |
| BESSELJ | 36 | fail 36 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, HyperFormula 0.27871487333 |
| SUMPRODUCT | 36 | fail 36 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", HyperFormula 0 |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", HyperFormula "#NAME?" |
| REPLACEB | 35 | unsupported 35 | `text/replaceb.xlsx!A32` =REPLACEB($K$1,-2,4,"ab") → Excel "#VALUE!", HyperFormula "#NAME?" |
| XMATCH | 32 | unsupported 32 | `spreadsheet/xmatch.xlsx!A35` =XMATCH("chael",F$13:F$19,2,1) → Excel "#N/A", HyperFormula "#NAME?" |
| CLEAN | 32 | fail 31, unsupported 1 | `text/clean.xlsx!A93` =CLEAN(J93) → Excel "#N/A", HyperFormula "#NAME?" |
| CHOOSECOLS | 31 | unsupported 31 | `spreadsheet/choosecols.xlsx!A143` =CHOOSECOLS(J3:N8, -3, {-3},P2) → Excel 1, HyperFormula "#NAME?" |
| CHOOSEROWS | 31 | unsupported 31 | `spreadsheet/chooserows.xlsx!A60` =CHOOSEROWS(J3:L8, -5, {-5,-5}) → Excel "February", HyperFormula "#NAME?" |
| XLOOKUP | 29 | fail 29 | `spreadsheet/xlookup.xlsx!A130` =XLOOKUP("Luchs",G$121:G$128,H$121:H$128, ,-1,-2) → Excel "Susanne", HyperFormula "First Name" |
| SUM | 28 | fail 16, unsupported 12 | `array/frequency.xlsx!A61` =SUM(ABS(FREQUENCY(K40:K59,L40:L46))) → Excel 20, HyperFormula "#NAME?" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", HyperFormula "#NAME?" |

## Errors

- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
- `fods/Functions_Excel_2016.xlsx`: Sheet size limit exceeded
