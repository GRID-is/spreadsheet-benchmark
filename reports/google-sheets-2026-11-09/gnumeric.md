# gnumeric vs google-sheets@2026-11-09 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 895 | 22 | 28 | 0 | 5 | 94.7% |
| overview | 650 | 610 | 36 | 0 | 0 | 4 | 94.4% |
| yalta2008 | 369 | 318 | 51 | 0 | 0 | 0 | 86.2% |
| **total** | 1969 | 1823 | 109 | 28 | 0 | 9 | 93.0% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1823 | 109 | 28 | 0 | 9 | 93.0% |
| **total** | 1969 | 1823 | 109 | 28 | 0 | 9 | 93.0% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IF | 30 | fail 30 | `functions/lookfuns.xls!F40` =IF(AND(J40=3,K40=3),1,0) → Excel 1, Google Sheets 0 |
| AND | 6 | fail 6 | `yalta2008/yalta2008.xls!B1` =AND(K10:K65536) → Excel true, Google Sheets false |
| SUM | 6 | fail 6 | `overview/database.xlsx!B1` =SUM(0+D$17:D$10001) → Excel 22, Google Sheets "#VALUE!" |
| COUNTIF | 5 | fail 5 | `overview/countif.xlsx!E12` =COUNTIF($A12,E$6) → Excel 1, Google Sheets 0 |
| TINV | 4 | fail 4 | `yalta2008/yalta2008.xls!D20` =TINV(2*A20,B20) → Excel 318309886183.7907, Google Sheets 318299258353 |
| FINV | 4 | fail 4 | `yalta2008/yalta2008.xls!E20` =FINV(A20,B20,C20) → Excel 4.052847345693522e+25, Google Sheets 9007200000000000 |
| SUMPRODUCT | 4 | fail 4 | `overview/countif.xlsx!E18` =SUMPRODUCT($J$7:$J$15,E7:E15) → Excel 509, Google Sheets 477 |
| ODDLPRICE | 3 | unsupported 3 | `functions/finfuns.xls!F53` =ODDLPRICE(DATE(1999,11,11),DATE(2012,3,1),DATE(1999,10,15),0 → Excel 167.40887498637008, Google Sheets "#NAME?" |
| ODDLYIELD | 3 | unsupported 3 | `functions/finfuns.xls!F54` =ODDLYIELD(DATE(1999,11,11),DATE(2012,3,1),DATE(1999,10,15),0 → Excel 38.28414544976381, Google Sheets "#NAME?" |
| ODDFYIELD | 3 | unsupported 3 | `functions/finfuns.xls!F52` =ODDFYIELD(DATE(1999,11,11),DATE(2012,3,1),DATE(1999,10,15),D → Excel 12.615078687777377, Google Sheets "#NAME?" |
| ODDFPRICE | 3 | unsupported 3 | `functions/finfuns.xls!F51` =ODDFPRICE(DATE(1999,11,11),DATE(2012,3,1),DATE(1999,10,15),D → Excel 113.65002161109066, Google Sheets "#NAME?" |
| AMORDEGRC | 3 | unsupported 3 | `functions/finfuns.xls!F23` =AMORDEGRC(2400,DATE(1998,8,19),DATE(1998,12,30),300,3,0.14,1 → Excel 310, Google Sheets "#NAME?" |
| IMLN | 3 | fail 3 | `functions/engfuns.xls!F38` =IMLN("3.3-9.2j") → Excel "2.27972016714096-1.2263959247521j", Google Sheets "2.27972016714096-1.2263959247521i" |
| BESSELY | 3 | unsupported 3 | `functions/engfuns.xls!F16` =BESSELY(6,9) → Excel -2.290660974154936, Google Sheets "#NAME?" |
| BESSELK | 3 | unsupported 3 | `functions/engfuns.xls!F15` =BESSELK(0.3,3) → Excel 292.9991961807509, Google Sheets "#NAME?" |
| BESSELI | 3 | unsupported 3 | `functions/engfuns.xls!F13` =BESSELI(10,9) → Excel 52.31929317876348, Google Sheets "#NAME?" |
| BESSELJ | 3 | unsupported 3 | `functions/engfuns.xls!F14` =BESSELJ(4,3) → Excel 0.4301714711533965, Google Sheets "#NAME?" |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", Google Sheets "#NAME?" |
| CELL | 3 | fail 3 | `functions/infofuns.xls!F26` =CELL("format",A14) → Excel "C0-", Google Sheets "#VALUE!" |
| AREAS | 3 | fail 2, unsupported 1 | `functions/lookfuns.xls!F29` =AREAS((A4,C7,C8,E10)) → Excel 4, Google Sheets "#ERROR!" |
| GETPIVOTDATA | 3 | fail 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, Google Sheets "#N/A" |
| IMPOWER | 2 | fail 2 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", Google Sheets "-169691027.47961-208555947.133372j" |
| NETWORKDAYS | 2 | fail 2 | `functions/datefuns.xls!F30` =NETWORKDAYS("1999/01/22","2001/05/21",@{"2000/03/23";"2000/0 → Excel 606, Google Sheets "#VALUE!" |
| CHIINV | 2 | fail 2 | `yalta2008/yalta2008.xls!D21` =CHIINV(A21,B21) → Excel 65.23863621336785, Google Sheets 65.23856632 |
| COLUMNS | 2 | fail 2 | `overview/countif.xlsx!B2` =COLUMNS(B18:G20)-B1 → Excel 0, Google Sheets 1 |
