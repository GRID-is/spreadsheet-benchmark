# gnumeric vs spread-sheets@19.2.5 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 886 | 56 | 3 | 0 | 5 | 93.8% |
| overview | 650 | 478 | 168 | 0 | 0 | 4 | 74.0% |
| yalta2008 | 369 | 0 | 0 | 0 | 369 | 0 | 0.0% |
| **total** | 1969 | 1364 | 224 | 3 | 369 | 9 | 69.6% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1364 | 224 | 3 | 369 | 9 | 69.6% |
| **total** | 1969 | 1364 | 224 | 3 | 369 | 9 | 69.6% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IF | 121 | fail 5, error 116 | `yalta2008/yalta2008.xls!J22` =IF(ISERROR(F22),FALSE,ABS(H22-F22)<=I22) → Excel true, SpreadJS null |
| INT | 115 | error 115 | `yalta2008/yalta2008.xls!I21` =10^(INT(LOG10(H21))-5)/2 → Excel 5e-7, SpreadJS null |
| COUNTIF | 40 | fail 40 | `overview/countif.xlsx!D10` =COUNTIF($A10,D$6) → Excel 0, SpreadJS 1 |
| SUMPRODUCT | 16 | fail 16 | `overview/countif.xlsx!E18` =SUMPRODUCT($U$7:$U$15,E7:E15) → Excel 130, SpreadJS 146 |
| CHIINV | 15 | error 15 | `yalta2008/yalta2008.xls!D24` =CHIINV(A24,B24) → Excel 779.3526516082724, SpreadJS null |
| BINOMDIST | 14 | error 14 | `yalta2008/yalta2008.xls!F23` =BINOMDIST(A23,B23,C23,D23) → Excel 0.9999983439256699, SpreadJS null |
| POISSON | 14 | error 14 | `yalta2008/yalta2008.xls!E22` =POISSON(A22,B22,C22) → Excel 0.500084104416326, SpreadJS null |
| NORMSINV | 14 | error 14 | `yalta2008/yalta2008.xls!C21` =NORMSINV(A21) → Excel -29.976284241123032, SpreadJS null |
| BETAINV | 14 | error 14 | `yalta2008/yalta2008.xls!E22` =BETAINV(A22,B22,C22) → Excel 0.001755888516049827, SpreadJS null |
| DSUM | 14 | fail 14 | `overview/database.xlsx!B40` =DSUM($A$4:$C$10,"Value",$C$14:$C$15) → Excel 3, SpreadJS 8 |
| SUM | 13 | fail 13 | `functions/statfuns.xls!F56` =SUM(M56:N56) → Excel 2.0504867408866803, SpreadJS 1.05048674088668 |
| TINV | 13 | error 13 | `yalta2008/yalta2008.xls!D20` =TINV(2*A20,B20) → Excel 318309886183.7907, SpreadJS null |
| FINV | 12 | error 12 | `yalta2008/yalta2008.xls!E20` =FINV(A20,B20,C20) → Excel 4.052847345693522e+25, SpreadJS null |
| IFERROR | 11 | fail 11 | `overview/database.xlsx!D40` =IFERROR(ERROR.TYPE(B40)=ERROR.TYPE(C40),IFERROR(B40=C40,FALS → Excel true, SpreadJS false |
| HYPGEOMDIST | 10 | error 10 | `yalta2008/yalta2008.xls!G19` =HYPGEOMDIST(A19,B19,C19,D19) → Excel 1.6013706045729205e-280, SpreadJS null |
| GAMMADIST | 10 | error 10 | `yalta2008/yalta2008.xls!F19` =GAMMADIST(A19,B19,C19,D19) → Excel 0.726678321707702, SpreadJS null |
| AND | 9 | error 9 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, SpreadJS null |
| SUMIF | 5 | fail 5 | `overview/ifs-funcs.xlsx!B27` =SUMIF($B$6:$B$12,"<>5",$C$6:$C$12) → Excel 3, SpreadJS 8 |
| COLUMNS | 4 | fail 4 | `overview/countif.xlsx!B2` =COLUMNS(B18:S20)-B1 → Excel 0, SpreadJS 4 |
| SERIESSUM | 3 | fail 3 | `functions/mathfuns.xls!F64` =SERIESSUM(3,1,2.23,D14:D17) → Excel -42587.98578678404, SpreadJS -20067 |
| IMPOWER | 3 | fail 3 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", SpreadJS "-104445951.99999997-142049792.00000006i" |
| IMSUB | 3 | fail 3 | `functions/engfuns.xls!F46` =IMSUB("3-j","8-2j") → Excel "-5+j", SpreadJS "-5+i" |
| IMSUM | 3 | fail 3 | `functions/engfuns.xls!F47` =IMSUM("2+5j","1+j","4-7j") → Excel "7-j", SpreadJS "7-i" |
| IMPRODUCT | 3 | fail 3 | `functions/engfuns.xls!F42` =IMPRODUCT("2-j","4-2j","9-2j") → Excel "38-84j", SpreadJS "38-84i" |
| IMSQRT | 3 | fail 3 | `functions/engfuns.xls!F45` =IMSQRT("5-9j") → Excel "2.76546832751588-1.62721082545978j", SpreadJS "2.765468327515884-1.6272108254597806i" |

## Errors

- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
- `yalta2008/yalta2008.xls`: timeout after 60000 ms
