# gnumeric vs univer@1.4.0 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 11. Fails where both answered an error, but a different one: 2. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 901 | 40 | 4 | 0 | 5 | 95.3% |
| overview | 650 | 478 | 168 | 0 | 0 | 4 | 74.0% |
| yalta2008 | 369 | 297 | 72 | 0 | 0 | 0 | 80.5% |
| **total** | 1969 | 1676 | 280 | 4 | 0 | 9 | 85.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1676 | 280 | 4 | 0 | 9 | 85.5% |
| **total** | 1969 | 1676 | 280 | 4 | 0 | 9 | 85.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IF | 47 | fail 47 | `functions/lookfuns.xls!F40` =IF(AND(J40=3,K40=3),1,0) → Excel 1, Univer 0 |
| DSUM | 26 | fail 26 | `overview/database.xlsx!B40` =DSUM($A$4:$C$10,"Value",$C$14:$C$15) → Excel 3, Univer 7 |
| IFERROR | 25 | fail 25 | `overview/database.xlsx!D40` =IFERROR(ERROR.TYPE(B40)=ERROR.TYPE(C40),IFERROR(B40=C40,FALS → Excel true, Univer false |
| SUM | 13 | fail 13 | `functions/statfuns.xls!D55` =SUM(M58:N62) → Excel 916.4461894098997, Univer 0.801644026329 |
| COUNTIF | 11 | fail 11 | `overview/countif.xlsx!S13` =COUNTIF($A13,S$6) → Excel 1, Univer 0 |
| POISSON | 10 | fail 10 | `yalta2008/yalta2008.xls!E22` =POISSON(A22,B22,C22) → Excel 0.500084104416326, Univer null |
| SUMPRODUCT | 7 | fail 7 | `overview/countif.xlsx!S18` =SUMPRODUCT($U$7:$U$15,S7:S15) → Excel 64, Univer 0 |
| NORMSINV | 6 | fail 6 | `yalta2008/yalta2008.xls!C21` =NORMSINV(A21) → Excel -29.976284241123032, Univer -29.9395704141 |
| SUMIF | 6 | fail 6 | `overview/ifs-funcs.xlsx!B27` =SUMIF($B$6:$B$12,"<>5",$C$6:$C$12) → Excel 3, Univer 7 |
| FINV | 5 | fail 5 | `yalta2008/yalta2008.xls!E20` =FINV(A20,B20,C20) → Excel 4.052847345693522e+25, Univer "#NUM!" |
| DAVERAGE | 4 | fail 4 | `functions/dbfuns.xls!F32` =DAVERAGE($A$15:$D$22,"Age",$A$11:$C$13) → Excel 35.666666666666664, Univer 37.5 |
| BINOMDIST | 4 | fail 4 | `yalta2008/yalta2008.xls!F23` =BINOMDIST(A23,B23,C23,D23) → Excel 0.9999983439256699, Univer null |
| AND | 4 | fail 4 | `yalta2008/yalta2008.xls!B1` =AND(K10:K65536) → Excel true, Univer false |
| XIRR | 3 | fail 3 | `functions/finfuns.xls!F69` =XIRR({-10000;2750;4290;3250;2742},{"1998/1/1";"1998/1/3";"19 → Excel 0.4230088663101197, Univer "#VALUE!" |
| AMORDEGRC | 3 | unsupported 3 | `functions/finfuns.xls!F23` =AMORDEGRC(2400,DATE(1998,8,19),DATE(1998,12,30),300,3,0.14,1 → Excel 310, Univer "#NAME?" |
| IMLN | 3 | fail 3 | `functions/engfuns.xls!F38` =IMLN("3.3-9.2j") → Excel "2.27972016714096-1.2263959247521j", Univer "2.27972016714096+1.2263959247521j" |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", Univer "#NAME?" |
| CELL | 3 | fail 3 | `functions/infofuns.xls!F26` =CELL("format",A14) → Excel "C0-", Univer "G" |
| AREAS | 3 | fail 3 | `functions/lookfuns.xls!F29` =AREAS((A4,C7,C8,E10)) → Excel 4, Univer "#VALUE!" |
| GETPIVOTDATA | 3 | fail 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, Univer "#VALUE!" |
| MINA | 2 | fail 2 | `functions/statfuns.xls!F63` =MINA($C$14:$C$18) → Excel 11.5, Univer 0 |
| DOLLARDE | 2 | fail 2 | `functions/finfuns.xls!F36` =DOLLARDE(1.4,28) → Excel 2.4285714285714284, Univer 2.428571 |
| ACCRINT | 2 | fail 2 | `functions/finfuns.xls!F21` =ACCRINT(DATE(1998,2,28),DATE(1998,8,31),DATE(1998,5,1),0.13, → Excel 44.77777777777778, Univer 45.5 |
| DOLLAR | 2 | fail 2 | `functions/textfuns.xls!D16` =DOLLAR(23.25,1) → Excel "$23.3", Univer "$23.3 " |
| IMPOWER | 2 | fail 2 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", Univer "-169691027.479611-208555947.133372j" |
