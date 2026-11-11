# gnumeric vs fast-formula-parser@1.0.19 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 3. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 461 | 72 | 412 | 0 | 5 | 48.8% |
| overview | 650 | 500 | 83 | 63 | 0 | 4 | 77.4% |
| yalta2008 | 369 | 253 | 0 | 116 | 0 | 0 | 68.6% |
| **total** | 1969 | 1214 | 155 | 591 | 0 | 9 | 61.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1214 | 155 | 591 | 0 | 9 | 61.9% |
| **total** | 1969 | 1214 | 155 | 591 | 0 | 9 | 61.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| COUNTIF | 55 | fail 55 | `overview/countif.xlsx!R15` =COUNTIF($A15,R$6) → Excel 1, fast-formula-parser 0 |
| DSUM | 35 | unsupported 35 | `functions/dbfuns.xls!D41` =DSUM($A$15:$D$22,"Salary",$A$11:$A$13) → Excel 170613, fast-formula-parser "#NAME?" |
| SUM | 28 | fail 28 | `functions/statfuns.xls!F90` =SUM(K95:K98) → Excel 57.97039877766521, fast-formula-parser 5.847891104764472 |
| CHIINV | 18 | unsupported 18 | `functions/statfuns.xls!F28` =CHIINV(0.18,60) → Excel 69.86264562023686, fast-formula-parser "#NAME?" |
| POISSON | 17 | unsupported 17 | `functions/statfuns.xls!F74` =POISSON(2,6,1) → Excel 0.06196880441665896, fast-formula-parser "#NAME?" |
| NORMSINV | 17 | unsupported 17 | `functions/statfuns.xls!F69` =NORMSINV(0.78) → Excel 0.772193214188685, fast-formula-parser "#NAME?" |
| BETAINV | 17 | unsupported 17 | `functions/statfuns.xls!F25` =BETAINV(0.95,6.3,1) → Excel 0.9918912618659407, fast-formula-parser "#NAME?" |
| BINOMDIST | 17 | unsupported 17 | `functions/statfuns.xls!F26` =BINOMDIST(0,2,0.8,1) → Excel 0.03999999999999999, fast-formula-parser "#NAME?" |
| TINV | 16 | unsupported 16 | `functions/statfuns.xls!F89` =TINV(0.74,37) → Excel 0.334354486087895, fast-formula-parser "#NAME?" |
| FINV | 15 | unsupported 15 | `functions/statfuns.xls!F39` =FINV(0.12,9,19) → Excel 1.8701696948810491, fast-formula-parser "#NAME?" |
| HYPGEOMDIST | 13 | unsupported 13 | `functions/statfuns.xls!F51` =HYPGEOMDIST(2,3,4,300) → Excel 0.0003986442504096433, fast-formula-parser "#NAME?" |
| GAMMADIST | 13 | unsupported 13 | `functions/statfuns.xls!F45` =GAMMADIST(2,1,5,1) → Excel 0.32967995396436073, fast-formula-parser "#NAME?" |
| SUMIF | 8 | fail 8 | `overview/ifs-funcs.xlsx!B20` =SUMIF($B$6:$B$12,">=2",$C$6:$C$12) → Excel "#DIV/0!", fast-formula-parser 2 |
| DAVERAGE | 5 | unsupported 5 | `functions/dbfuns.xls!F32` =DAVERAGE($A$15:$D$22,"Age",$A$11:$C$13) → Excel 35.666666666666664, fast-formula-parser "#NAME?" |
| DCOUNTA | 5 | unsupported 5 | `functions/dbfuns.xls!F34` =DCOUNTA($A$15:$D$22,"name",$C$11:$C$12) → Excel 4, fast-formula-parser "#NAME?" |
| DMAX | 5 | unsupported 5 | `functions/dbfuns.xls!F36` =DMAX($A$15:$D$22,"Age",$C$11:$C$12) → Excel 52, fast-formula-parser "#NAME?" |
| DVAR | 5 | unsupported 5 | `functions/dbfuns.xls!F42` =DVAR($A$15:$D$22,"Age",$C$11:$C$12) → Excel 160.33333333333334, fast-formula-parser "#NAME?" |
| DVARP | 5 | unsupported 5 | `functions/dbfuns.xls!F43` =DVARP($A$15:$D$22,"Age",$C$11:$C$12) → Excel 120.25, fast-formula-parser "#NAME?" |
| DPRODUCT | 5 | unsupported 5 | `functions/dbfuns.xls!F38` =DPRODUCT($A$15:$D$22,"Age",$C$11:$C$12) → Excel 1645696, fast-formula-parser "#NAME?" |
| DSTDEV | 5 | unsupported 5 | `functions/dbfuns.xls!F39` =DSTDEV($A$15:$D$22,"Age",$C$11:$C$12) → Excel 12.662279942148386, fast-formula-parser "#NAME?" |
| DSTDEVP | 5 | unsupported 5 | `functions/dbfuns.xls!F40` =DSTDEVP($A$15:$D$22,"Age",$C$11:$C$12) → Excel 10.965856099730654, fast-formula-parser "#NAME?" |
| DCOUNT | 5 | unsupported 5 | `functions/dbfuns.xls!F33` =DCOUNT($A$15:$D$22,"Age",$C$11:$C$12) → Excel 4, fast-formula-parser "#NAME?" |
| DMIN | 5 | unsupported 5 | `functions/dbfuns.xls!F37` =DMIN($A$15:$D$22,"Age",$C$11:$C$12) → Excel 23, fast-formula-parser "#NAME?" |
| IF | 4 | fail 4 | `functions/lookfuns.xls!F39` =IF(AND(J39=F10,K39=G10),1,0) → Excel 1, fast-formula-parser 0 |
| SUMIFS | 4 | unsupported 4 | `overview/ifs-funcs.xlsx!B31` =SUMIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel 3, fast-formula-parser "#NAME?" |
