# gnumeric vs ironcalc@0.9.1 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 9. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 924 | 18 | 3 | 0 | 5 | 97.8% |
| overview | 650 | 199 | 447 | 0 | 0 | 4 | 30.8% |
| yalta2008 | 369 | 321 | 48 | 0 | 0 | 0 | 87.0% |
| **total** | 1969 | 1444 | 513 | 3 | 0 | 9 | 73.7% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1444 | 513 | 3 | 0 | 9 | 73.7% |
| **total** | 1969 | 1444 | 513 | 3 | 0 | 9 | 73.7% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| COUNTIF | 300 | fail 300 | `overview/countif.xlsx!S14` =COUNTIF($A14,S$6) → Excel 0, IronCalc "#VALUE!" |
| SUMPRODUCT | 36 | fail 36 | `overview/countif.xlsx!S18` =SUMPRODUCT($U$7:$U$15,S7:S15) → Excel 64, IronCalc "#VALUE!" |
| IF | 21 | fail 21 | `functions/lookfuns.xls!F40` =IF(AND(J40=3,K40=3),1,0) → Excel 1, IronCalc "#VALUE!" |
| SUM | 13 | fail 13 | `overview/database.xlsx!B1` =SUM(0+D$17:D$10001) → Excel 22, IronCalc "#VALUE!" |
| IFERROR | 10 | fail 10 | `overview/database.xlsx!D27` =IFERROR(ERROR.TYPE(B27)=ERROR.TYPE(C27),IFERROR(B27=C27,FALS → Excel true, IronCalc false |
| DSUM | 9 | fail 9 | `overview/database.xlsx!B29` =DSUM($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", IronCalc 3 |
| AND | 7 | fail 7 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, IronCalc true |
| TINV | 6 | fail 6 | `yalta2008/yalta2008.xls!D20` =TINV(2*A20,B20) → Excel 318309886183.7907, IronCalc 92007338.29081053 |
| FINV | 5 | fail 5 | `yalta2008/yalta2008.xls!E20` =FINV(A20,B20,C20) → Excel 4.052847345693522e+25, IronCalc 4503599627370496 |
| COLUMNS | 4 | fail 4 | `overview/countif.xlsx!B2` =COLUMNS(B18:S20)-B1 → Excel 0, IronCalc "#VALUE!" |
| XIRR | 3 | fail 3 | `functions/finfuns.xls!F69` =XIRR({-10000;2750;4290;3250;2742},{"1998/1/1";"1998/1/3";"19 → Excel 0.4230088663101197, IronCalc "#VALUE!" |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", IronCalc "#N/IMPL!" |
| CELL | 3 | fail 3 | `functions/infofuns.xls!F26` =CELL("format",A14) → Excel "C0-", IronCalc "#VALUE!" |
| GETPIVOTDATA | 3 | unsupported 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, IronCalc "#NAME?" |
| IMSUM | 2 | fail 2 | `functions/engfuns.xls!F47` =IMSUM("2+5j","1+j","4-7j") → Excel "7-j", IronCalc "#ERROR!" |
| NETWORKDAYS | 2 | fail 2 | `functions/datefuns.xls!F30` =NETWORKDAYS("1999/01/22","2001/05/21",@{"2000/03/23";"2000/0 → Excel 606, IronCalc "#VALUE!" |
| AREAS | 2 | fail 2 | `functions/lookfuns.xls!F29` =AREAS((A4,C7,C8,E10)) → Excel 4, IronCalc "#ERROR!" |
| CHIINV | 2 | fail 2 | `yalta2008/yalta2008.xls!D21` =CHIINV(A21,B21) → Excel 65.23863621336785, IronCalc 65.2385663178878 |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", IronCalc "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", IronCalc "#NAME?" |
| SUMIF | 2 | fail 2 | `overview/ifs-funcs.xlsx!B22` =SUMIF($B$6:$B$14,A22,$C$6:$C$14) → Excel 32, IronCalc 34 |
| TEXT | 1 | fail 1 | `functions/textfuns.xls!D31` =TEXT("4/15/1999","mmmm, dd, yy") → Excel "April, 15, 99", IronCalc "4/15/1999" |
| IMPOWER | 1 | fail 1 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", IronCalc "-169691027.47961-208555947.133372j" |
| IMPRODUCT | 1 | fail 1 | `functions/engfuns.xls!F42` =IMPRODUCT("2-j","4-2j","9-2j") → Excel "38-84j", IronCalc "#ERROR!" |
| BESSELJ | 1 | fail 1 | `functions/engfuns.xls!D14` =BESSELJ(8,3) → Excel -0.29112524185253685, IronCalc -0.2911322070659522 |
