# gnumeric vs hyperformula@3.5.0 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 1. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 683 | 98 | 164 | 0 | 5 | 72.3% |
| overview | 650 | 397 | 167 | 82 | 0 | 4 | 61.5% |
| yalta2008 | 369 | 241 | 86 | 42 | 0 | 0 | 65.3% |
| **total** | 1969 | 1321 | 351 | 288 | 0 | 9 | 67.4% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1321 | 351 | 288 | 0 | 9 | 67.4% |
| **total** | 1969 | 1321 | 351 | 288 | 0 | 9 | 67.4% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| COUNTIF | 81 | fail 67, unsupported 14 | `overview/countif.xlsx!S13` =COUNTIF($A13,S$6) → Excel 1, HyperFormula 0 |
| IF | 59 | fail 21, unsupported 38 | `functions/lookfuns.xls!F39` =IF(AND(J39=F10,K39=G10),1,0) → Excel 1, HyperFormula 0 |
| IFERROR | 37 | fail 11, unsupported 26 | `overview/database.xlsx!D40` =IFERROR(ERROR.TYPE(B40)=ERROR.TYPE(C40),IFERROR(B40=C40,FALS → Excel true, HyperFormula false |
| INT | 31 | fail 31 | `functions/mathfuns.xls!D41` =INT(-7.5) → Excel -8, HyperFormula -7 |
| SUM | 23 | fail 6, unsupported 17 | `functions/statfuns.xls!F90` =SUM(K95:K98) → Excel 57.97039877766521, HyperFormula "#NAME?" |
| SUMPRODUCT | 17 | fail 15, unsupported 2 | `overview/countif.xlsx!S18` =SUMPRODUCT($U$7:$U$15,S7:S15) → Excel 64, HyperFormula 0 |
| HYPGEOMDIST | 13 | fail 13 | `functions/statfuns.xls!F51` =HYPGEOMDIST(2,3,4,300) → Excel 0.0003986442504096433, HyperFormula "#N/A" |
| AND | 10 | unsupported 10 | `functions/statfuns.xls!B43` =AND(J23=1,J24=1) → Excel true, HyperFormula "#NAME?" |
| POISSON | 10 | fail 10 | `yalta2008/yalta2008.xls!E22` =POISSON(A22,B22,C22) → Excel 0.500084104416326, HyperFormula "#NUM!" |
| DSUM | 10 | fail 8, unsupported 2 | `overview/database.xlsx!B40` =DSUM($A$4:$C$10,"Value",$C$14:$C$15) → Excel 3, HyperFormula 7 |
| SUMIF | 8 | fail 8 | `overview/ifs-funcs.xlsx!B27` =SUMIF($B$6:$B$12,"<>5",$C$6:$C$12) → Excel 3, HyperFormula 7 |
| NORMSINV | 6 | fail 6 | `yalta2008/yalta2008.xls!C21` =NORMSINV(A21) → Excel -29.976284241123032, HyperFormula -29.939570414 |
| FINV | 6 | fail 6 | `yalta2008/yalta2008.xls!E20` =FINV(A20,B20,C20) → Excel 4.052847345693522e+25, HyperFormula "#NUM!" |
| TINV | 4 | fail 4 | `yalta2008/yalta2008.xls!D20` =TINV(2*A20,B20) → Excel 318309886183.7907, HyperFormula 318316927900 |
| COLUMNS | 4 | fail 2, unsupported 2 | `overview/countif.xlsx!B2` =COLUMNS(B18:S20)-B1 → Excel 0, HyperFormula 3 |
| TRIMMEAN | 3 | unsupported 3 | `functions/statfuns.xls!F91` =TRIMMEAN(B14:B17,0.9) → Excel 66.35, HyperFormula "#NAME?" |
| PROB | 3 | unsupported 3 | `functions/statfuns.xls!F75` =PROB({0;1;2;3},{0.1;0.4;0.2;0.3},0 ) → Excel 0.1, HyperFormula "#NAME?" |
| RANK | 3 | unsupported 3 | `functions/statfuns.xls!F77` =RANK(21.5,A14:A17,0) → Excel 3, HyperFormula "#NAME?" |
| PERCENTRANK | 3 | unsupported 3 | `functions/statfuns.xls!F72` =PERCENTRANK(E14:E17,21.5,2) → Excel 0.63, HyperFormula "#NAME?" |
| PERMUT | 3 | unsupported 3 | `functions/statfuns.xls!F73` =PERMUT(9,2) → Excel 72, HyperFormula "#NAME?" |
| NEGBINOMDIST | 3 | fail 3 | `functions/statfuns.xls!F65` =NEGBINOMDIST(3,8,0.245) → Excel 0.0006704258879152642, HyperFormula "#N/A" |
| MODE | 3 | unsupported 3 | `functions/statfuns.xls!F64` =MODE(I14:I15) → Excel 2, HyperFormula "#NAME?" |
| LOGNORMDIST | 3 | fail 3 | `functions/statfuns.xls!F58` =LOGNORMDIST(0.2,0.44,23) → Excel 0.46449883604037734, HyperFormula "#N/A" |
| NORMSDIST | 3 | fail 3 | `functions/statfuns.xls!F68` =NORMSDIST(0.23) → Excel 0.5909541151420059, HyperFormula "#N/A" |
| INTERCEPT | 3 | unsupported 3 | `functions/statfuns.xls!F52` =INTERCEPT(C14:C17,D14:D17) → Excel 41.50774193548387, HyperFormula "#NAME?" |
