# gnumeric vs poi@6.0.0 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 10. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 590 | 20 | 335 | 0 | 5 | 62.4% |
| overview | 650 | 546 | 86 | 14 | 0 | 4 | 84.5% |
| yalta2008 | 369 | 164 | 15 | 190 | 0 | 0 | 44.4% |
| **total** | 1969 | 1300 | 121 | 539 | 0 | 9 | 66.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1300 | 121 | 539 | 0 | 9 | 66.3% |
| **total** | 1969 | 1300 | 121 | 539 | 0 | 9 | 66.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IF | 101 | fail 9, unsupported 92 | `functions/lookfuns.xls!F31` =IF(AND(J31=1,K31=2),1,0) → Excel 1, Apache POI 0 |
| DSUM | 22 | fail 22 | `overview/database.xlsx!B40` =DSUM($A$4:$C$10,"Value",$C$14:$C$15) → Excel 3, Apache POI 4 |
| IFERROR | 19 | fail 15, unsupported 4 | `overview/database.xlsx!D40` =IFERROR(ERROR.TYPE(B40)=ERROR.TYPE(C40),IFERROR(B40=C40,FALS → Excel true, Apache POI false |
| CHIINV | 18 | unsupported 18 | `functions/statfuns.xls!F28` =CHIINV(0.18,60) → Excel 69.86264562023686, Apache POI "#NAME?" |
| BETAINV | 17 | unsupported 17 | `functions/statfuns.xls!F25` =BETAINV(0.95,6.3,1) → Excel 0.9918912618659407, Apache POI "#NAME?" |
| BINOMDIST | 17 | unsupported 17 | `functions/statfuns.xls!F26` =BINOMDIST(0,2,0.8,1) → Excel 0.03999999999999999, Apache POI "#NAME?" |
| TINV | 16 | unsupported 16 | `functions/statfuns.xls!F89` =TINV(0.74,37) → Excel 0.334354486087895, Apache POI "#NAME?" |
| FINV | 15 | unsupported 15 | `functions/statfuns.xls!F39` =FINV(0.12,9,19) → Excel 1.8701696948810491, Apache POI "#NAME?" |
| SUM | 14 | fail 4, unsupported 10 | `functions/statfuns.xls!F56` =SUM(M56:N56) → Excel 2.0504867408866803, Apache POI "#NAME?" |
| HYPGEOMDIST | 13 | unsupported 13 | `functions/statfuns.xls!F51` =HYPGEOMDIST(2,3,4,300) → Excel 0.0003986442504096433, Apache POI "#NAME?" |
| GAMMADIST | 13 | unsupported 13 | `functions/statfuns.xls!F45` =GAMMADIST(2,1,5,1) → Excel 0.32967995396436073, Apache POI "#NAME?" |
| AND | 8 | fail 1, unsupported 7 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, Apache POI "#NAME?" |
| NORMSINV | 6 | fail 6 | `yalta2008/yalta2008.xls!C21` =NORMSINV(A21) → Excel -29.976284241123032, Apache POI null |
| TRIMMEAN | 3 | unsupported 3 | `functions/statfuns.xls!F91` =TRIMMEAN(B14:B17,0.9) → Excel 66.35, Apache POI "#NAME?" |
| TTEST | 3 | unsupported 3 | `functions/statfuns.xls!F92` =TTEST(A14:A17,B14:B17,2,3) → Excel 0.04429625303289446, Apache POI "#NAME?" |
| STEYX | 3 | unsupported 3 | `functions/statfuns.xls!F87` =STEYX(C14:C17,D14:D17) → Excel 28.622084863536475, Apache POI "#NAME?" |
| WEIBULL | 3 | unsupported 3 | `functions/statfuns.xls!F97` =WEIBULL(1,3,2,1) → Excel 0.11750309741540463, Apache POI "#NAME?" |
| ZTEST | 3 | unsupported 3 | `functions/statfuns.xls!F98` =ZTEST(E14:E17,7) → Excel 0.16195860992195008, Apache POI "#NAME?" |
| SKEW | 3 | unsupported 3 | `functions/statfuns.xls!F79` =SKEW(C14:C17) → Excel 1.7312236838648485, Apache POI "#NAME?" |
| PROB | 3 | unsupported 3 | `functions/statfuns.xls!F75` =PROB({0;1;2;3},{0.1;0.4;0.2;0.3},0 ) → Excel 0.1, Apache POI "#NAME?" |
| QUARTILE | 3 | unsupported 3 | `functions/statfuns.xls!F76` =QUARTILE(A14:A17,3) → Excel 39.325, Apache POI "#NAME?" |
| RSQ | 3 | unsupported 3 | `functions/statfuns.xls!F78` =RSQ(C14:C17,D14:D17) → Excel 0.18136567477867555, Apache POI "#NAME?" |
| PERMUT | 3 | unsupported 3 | `functions/statfuns.xls!F73` =PERMUT(9,2) → Excel 72, Apache POI "#NAME?" |
| NEGBINOMDIST | 3 | unsupported 3 | `functions/statfuns.xls!F65` =NEGBINOMDIST(3,8,0.245) → Excel 0.0006704258879152642, Apache POI "#NAME?" |
| MINA | 3 | fail 3 | `functions/statfuns.xls!F63` =MINA($C$14:$C$18) → Excel 11.5, Apache POI 0 |
