# libreoffice vs mog@1.0.0 (libreoffice)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: excel-read-as-user-function 59, excel-lacks-function 523, environment 4, volatile 18. Fails where the engine's error came from another cell: 10. Fails where both answered an error, but a different one: 69. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (35): APPLYTWICE, B, CHISQDIST, CHISQINV, COM.MICROSOFT.ARRAYTOTEXT, COM.MICROSOFT.BYCOL, COM.MICROSOFT.BYROW, COM.MICROSOFT.LAMBDA, COM.MICROSOFT.MAKEARRAY, COM.MICROSOFT.MAP, COM.MICROSOFT.REDUCE, COM.MICROSOFT.SCAN, DAYSINMONTH, DAYSINYEAR, F54, F55, ISLEAPYEAR, MONTHS, ORG.LIBREOFFICE.COLOR, ORG.LIBREOFFICE.FORECAST.ETS.MULT, ORG.LIBREOFFICE.FORECAST.ETS.STAT.MULT, ORG.LIBREOFFICE.RAWSUBTRACT, ORG.LIBREOFFICE.REGEX, ORG.LIBREOFFICE.ROUNDSIG, ORG.OPENOFFICE.CONVERT, ORG.OPENOFFICE.CURRENT, ORG.OPENOFFICE.EASTERSUNDAY, ORG.OPENOFFICE.ERRORTYPE, ORG.OPENOFFICE.STYLE, PLUS2, ROT, ROT13, WEEKS, WEEKSINYEAR, YEARS. Never-supported functions present: EUROCONVERT.
Upstream literal differs from Excel's value in 2496 of 11456 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| addin | 1998 | 1813 | 121 | 0 | 0 | 64 | 93.7% |
| array | 205 | 164 | 41 | 0 | 0 | 0 | 80.0% |
| database | 100 | 94 | 6 | 0 | 0 | 0 | 94.0% |
| date_time | 485 | 326 | 37 | 0 | 20 | 102 | 85.1% |
| financial | 1075 | 805 | 270 | 0 | 0 | 0 | 74.9% |
| fods | 127 | 109 | 14 | 3 | 0 | 1 | 86.5% |
| information | 153 | 140 | 4 | 0 | 0 | 9 | 97.2% |
| lambda | 108 | 22 | 0 | 2 | 0 | 84 | 91.7% |
| logical | 126 | 123 | 0 | 0 | 0 | 3 | 100.0% |
| mathematical | 1579 | 1349 | 161 | 0 | 0 | 69 | 89.3% |
| spreadsheet | 1756 | 1459 | 256 | 1 | 0 | 40 | 85.0% |
| statistical | 3385 | 3021 | 130 | 12 | 0 | 222 | 95.5% |
| text | 963 | 772 | 181 | 0 | 0 | 10 | 81.0% |
| **total** | 12060 | 10197 | 1221 | 18 | 20 | 604 | 89.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 12060 | 10197 | 1221 | 18 | 20 | 604 | 89.0% |
| **total** | 12060 | 10197 | 1221 | 18 | 20 | 604 | 89.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| B | 164 | skip 164 | `statistical/b.xlsx!A121` =B(J122,K122,L122,M122) → Excel "#NAME?", Mog "#NAME?" |
| ROUND | 150 | fail 150 | `financial/rate.xlsx!A222` =ROUND(RATE(F223,35,-250,0),6) → Excel "#NUM!", Mog 0.14 |
| LOOKUP | 87 | fail 87 | `spreadsheet/lookup.xlsx!A715` =LOOKUP(1,1) → Excel 1, Mog "#VALUE!" |
| ASC | 70 | fail 70 | `text/asc.xlsx!A72` =ASC(M33) → Excel "ア゛イ゛ウ゛エ゛オ゛", Mog "ｱ゛ｲ゛ｳ゛ｴ゛ｵ゛" |
| AGGREGATE | 68 | fail 68 | `mathematical/aggregate.xlsx!A183` =AGGREGATE( AC$13,$J25,$C$1:$C$11,2) → Excel 33, Mog 32 |
| CONVERT | 60 | fail 60 | `addin/convert.xlsx!A129` =CONVERT(1,H80,G79) → Excel 31557600, Mog "#N/A" |
| COM.MICROSOFT.LAMBDA | 47 | skip 47 | `lambda/isomitted.xlsx!A8` =com.microsoft.lambda(a,b,c, IF(com.microsoft.isomitted(b), - → Excel "#NAME?", Mog "#NAME?" |
| JIS | 38 | fail 38 | `text/jis.xlsx!A42` =JIS(I2) → Excel "#NAME?", Mog "！＂＃＄％＆＇（）＊＋，－．／" |
| BESSELJ | 36 | fail 36 | `addin/besselj.xlsx!A315` =BESSELJ($L15,Y$1) → Excel 0.2787146408460811, Mog 0.27871487333368605 |
| WEEKS | 36 | skip 36 | `date_time/weeks.xlsx!A31` =WEEKS(J31,K31,M31) → Excel "#NAME?", Mog "#NAME?" |
| GETPIVOTDATA | 33 | fail 33 | `spreadsheet/getpivotdata.xlsx!A34` =GETPIVOTDATA(I30,"Name['Wendy   ']") → Excel 21, Mog "#CALC!" |
| SUMPRODUCT | 32 | fail 32 | `array/sumproduct.xlsx!A36` =SUMPRODUCT(K40,J40*1) → Excel "#VALUE!", Mog 0 |
| CLEAN | 32 | fail 32 | `text/clean.xlsx!A93` =CLEAN(J93) → Excel "#N/A", Mog "!🀐$🀃E𯀁h🀠1򯿾E﷐F�$" |
| CHISQDIST | 28 | skip 28 | `statistical/chisqdist.xlsx!A23` =CHISQDIST(4,2,TRUE()) → Excel "#NAME?", Mog "#NAME?" |
| ORG.LIBREOFFICE.ROUNDSIG | 26 | skip 26 | `mathematical/roundsig.xlsx!A20` =ORG.LIBREOFFICE.ROUNDSIG(F20,G20) → Excel "#NAME?", Mog "#NAME?" |
| ORG.OPENOFFICE.CONVERT | 24 | skip 24 | `mathematical/convert_ooo.xlsx!A21` =ORG.OPENOFFICE.CONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", Mog "#NAME?" |
| EUROCONVERT | 24 | fail 24 | `mathematical/euroconvert.xlsx!A21` =EUROCONVERT(100,"EUR","SIT",0,3) → Excel "#NAME?", Mog 23964 |
| WEEKNUM | 20 | error 20 | `date_time/weeknum.xlsx!A19` =WEEKNUM($K$9,17) → Excel 31, Mog null |
| XLOOKUP | 18 | fail 18 | `spreadsheet/xlookup.xlsx!A124` =XLOOKUP($O$1,S$1:S$11,T$1:T$11,,-1,-2) → Excel "h", Mog "g" |
| ORG.OPENOFFICE.ERRORTYPE | 16 | skip 16 | `fods/reference_operators.xlsx!A79` =ORG.OPENOFFICE.ERRORTYPE(B1:D5) → Excel "#NAME?", Mog "#NAME?" |
| ISLEAPYEAR | 14 | skip 14 | `date_time/isleapyear.xlsx!A11` =ISLEAPYEAR(years) → Excel "#NAME?", Mog "#NAME?" |
| ORG.LIBREOFFICE.FORECAST.ETS.MULT | 14 | skip 14 | `statistical/forecast.ets.mult.xlsx!A212` =ORG.LIBREOFFICE.FORECAST.ETS.MULT($BI$64:$BI100,$BH$5:$BH$64 → Excel "#NAME?", Mog "#NAME?" |
| ORG.LIBREOFFICE.RAWSUBTRACT | 13 | skip 13 | `mathematical/rawsubtract.xlsx!A10` =ORG.LIBREOFFICE.RAWSUBTRACT(0.987654321098765, 0.98765432109 → Excel "#NAME?", Mog "#NAME?" |
| CHOOSECOLS | 12 | fail 12 | `spreadsheet/choosecols.xlsx!A143` =CHOOSECOLS(J3:N8, -3, {-3},P2) → Excel 1, Mog "#VALUE!" |
| CHOOSEROWS | 12 | fail 12 | `spreadsheet/chooserows.xlsx!A60` =CHOOSEROWS(J3:L8, -5, {-5,-5}) → Excel "February", Mog "#VALUE!" |

## Errors

- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
- `date_time/weeknum.xlsx`: mog --recalculate failed (exit status 1): thread 'compute-engine' (68071105) panicked at compute/core/crates/compute-functions/src/datetime/week.rs:210:59:
valid serial from NaiveDate
note: run with `RUST_BACKTRACE=1` environment variable to display a backtrace
engine shut down
