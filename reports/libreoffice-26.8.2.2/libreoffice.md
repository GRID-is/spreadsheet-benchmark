# libreoffice vs libreoffice@26.8.2.2 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, excel-unstable 2, environment 4, volatile 18. Fails where the engine's error came from another cell: 19. Fails where both answered an error, but a different one: 736. Fails where the case's own cell matches but its spill does not: 12.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2494 of 11454 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1661 | 273 | 0 | 0 | 64 | 85.9% |
| array | 205 | 170 | 35 | 0 | 0 | 0 | 82.9% |
| database | 100 | 82 | 9 | 9 | 0 | 0 | 82.0% |
| date_time | 485 | 332 | 49 | 0 | 0 | 104 | 87.1% |
| financial | 1075 | 580 | 495 | 0 | 0 | 0 | 54.0% |
| fods | 127 | 121 | 5 | 0 | 0 | 1 | 96.0% |
| information | 153 | 130 | 14 | 0 | 0 | 9 | 90.3% |
| lambda | 108 | 13 | 11 | 0 | 0 | 84 | 54.2% |
| logical | 126 | 118 | 5 | 0 | 0 | 3 | 95.9% |
| mathematical | 1579 | 1303 | 207 | 0 | 0 | 69 | 86.3% |
| spreadsheet | 1756 | 1623 | 87 | 6 | 0 | 40 | 94.6% |
| statistical | 3385 | 2739 | 424 | 0 | 0 | 222 | 86.6% |
| text | 963 | 721 | 216 | 16 | 0 | 10 | 75.7% |
| **total** | 12060 | 9593 | 1830 | 31 | 0 | 606 | 83.8% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 9593 | 1830 | 31 | 0 | 606 | 83.8% |
| **total** | 12060 | 9593 | 1830 | 31 | 0 | 606 | 83.8% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| ROUND | 165 | fail 165 | `financial/rate.xlsx!A230` =ROUND(RATE(F231,35,-250,0),6) → Excel "#NUM!", LibreOffice 0.14 |
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", LibreOffice 7.78167529704766e-9 |
| BESSELI | 123 | fail 123 | `addin/besseli.xlsx!A278` =BESSELI($L2,X$1) → Excel 5.212139801444066e+41, LibreOffice 5.21214226588802e+41 |
| CONVERT | 88 | fail 88 | `addin/convert.xlsx!A151` =CONVERT(1,H98,H99) → Excel "#N/A", LibreOffice 0.000160333493666367 |
| ASC | 75 | fail 75 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", LibreOffice "ｱﾞｲﾞｳﾞｴﾞｵﾞ" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", LibreOffice "#VALUE!" |
| BESSELJ | 39 | fail 39 | `addin/besselj.xlsx!A326` =BESSELJ(1,-9) → Excel "#NUM!", LibreOffice "#VALUE!" |
| JIS | 38 | fail 38 | `text/jis.xlsx!A42` =JIS(I2) → Excel "#NAME?", LibreOffice "！”＃＄％＆’（）＊＋，－．／" |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", LibreOffice 0 |
| PRICE | 35 | fail 35 | `financial/price.xlsx!A50` =PRICE(N45,N46,N47,N48,N49,N50,N51) → Excel 100.06206741734462, LibreOffice 100.140566955506 |
| SUMPRODUCT | 32 | fail 32 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", LibreOffice 0 |
| REPLACEB | 30 | fail 30 | `text/replaceb.xlsx!A35` =REPLACEB($K$29,5,25,"slow") → Excel "The slow", LibreOffice "#VALUE!" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", LibreOffice 0.864664716763387 |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", LibreOffice 1.235e+25 |
| LOOKUP | 26 | fail 26 | `spreadsheet/lookup.xlsx!A720` =LOOKUP(15000,{500;501;502;700;701;702;703;900;901;902;1600;1 → Excel 15000, LibreOffice "#N/A" |
| LET | 24 | fail 18, unsupported 6 | `lambda/lambda.xlsx!A20` =LET(f, com.microsoft.lambda(x, x * x), _xlpm.f(4)) → Excel "#NAME?", LibreOffice "#VALUE!" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", LibreOffice "#VALUE!" |
| EUROCONVERT | 24 | fail 24 | `mathematical/euroconvert.xlsx!A21` =EUROCONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", LibreOffice 23964 |
| ODDLYIELD | 18 | fail 18 | `financial/oddlyield.xlsx!A22` =ODDLYIELD("1999-04-20","1999-06-15", "1998-10-15", 0.0375, 9 → Excel "#N/A", LibreOffice "#VALUE!" |
| BINOM.INV | 17 | fail 17 | `statistical/binom.inv.xlsx!A25` =BINOM.INV(123,1,1) → Excel "#NUM!", LibreOffice 123 |
| ODDLPRICE | 16 | fail 16 | `financial/oddlprice.xlsx!A23` =ODDLPRICE(DATE(1999,2,7),DATE(1999,6,15),DATE(1998,10,15), 0 → Excel "#NUM!", LibreOffice "#VALUE!" |
| VDB | 16 | fail 16 | `financial/vdb.xlsx!A67` =VDB(G30,G31,G32,G54,G55) → Excel 5086.228480000002, LibreOffice 4777.21600000001 |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", LibreOffice 519 |
| PERCENTRANK.INC | 16 | fail 16 | `statistical/percentrank.inc.xlsx!A25` =PERCENTRANK.INC(R2:R10,5.43,1) → Excel 0.3, LibreOffice 0.4 |
| BAHTTEXT | 16 | unsupported 16 | `text/bahttext.xlsx!A13` =BAHTTEXT(12) → Excel "สิบสองบาทถ้วน", LibreOffice "#NAME?" |
