# gnumeric vs gridcraft@0.18.0 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 9. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 936 | 6 | 3 | 0 | 5 | 99.0% |
| overview | 650 | 617 | 29 | 0 | 0 | 4 | 95.5% |
| yalta2008 | 369 | 347 | 22 | 0 | 0 | 0 | 94.0% |
| **total** | 1969 | 1900 | 57 | 3 | 0 | 9 | 96.9% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1900 | 57 | 3 | 0 | 9 | 96.9% |
| **total** | 1969 | 1900 | 57 | 3 | 0 | 9 | 96.9% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IFERROR | 9 | fail 9 | `overview/database.xlsx!D27` =IFERROR(ERROR.TYPE(B27)=ERROR.TYPE(C27),IFERROR(B27=C27,FALS → Excel true, GridCraft false |
| IF | 7 | fail 7 | `yalta2008/yalta2008.xls!J10` =IF(ISERROR(F10),FALSE,ABS(H10-F10)<=I10) → Excel false, GridCraft true |
| AND | 5 | fail 5 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, GridCraft true |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", GridCraft "GridCraft" |
| GETPIVOTDATA | 3 | unsupported 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, GridCraft "#NAME?" |
| SUM | 3 | fail 3 | `overview/database.xlsx!B1` =SUM(0+D$17:D$10001) → Excel 22, GridCraft 13 |
| CELL | 2 | fail 2 | `functions/infofuns.xls!F26` =CELL("format",A14) → Excel "C0-", GridCraft "S0" |
| POISSON | 2 | fail 2 | `yalta2008/yalta2008.xls!E22` =POISSON(A22,B22,C22) → Excel 0.500084104416326, GridCraft 0.5008668054037354 |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", GridCraft "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", GridCraft "#NAME?" |
| YIELDDISC | 1 | fail 1 | `functions/finfuns.xls!D72` =YIELDDISC(DATE(1999,2,15),DATE(2007,11,15),0.0575,100,1) → Excel 198.68665411686433, GridCraft 198.5657617200789 |
| MDURATION | 1 | fail 1 | `functions/finfuns.xls!D46` =MDURATION(DATE(1999,1,1),DATE(2007,1,1),0.18,0.09,2,2) → Excel 4.94749398222136, GridCraft 4.950152142774259 |
| IMPOWER | 1 | fail 1 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", GridCraft "-169691027.47961-208555947.133372j" |
| BESSELJ | 1 | fail 1 | `functions/engfuns.xls!D14` =BESSELJ(8,3) → Excel -0.29112524185253685, GridCraft -0.2911322070659522 |
| NOW | 1 | skip 1 | `functions/datefuns.xls!B31` =NOW() → Excel 46330.76197175926, GridCraft 46336.792073460645 |
| TODAY | 1 | skip 1 | `functions/datefuns.xls!B35` =TODAY() → Excel 46330, GridCraft 46336 |
| CHIINV | 1 | fail 1 | `yalta2008/yalta2008.xls!D20` =CHIINV(A20,B20) → Excel 50.84417133244917, GridCraft 50.84412791181815 |
| TINV | 1 | fail 1 | `yalta2008/yalta2008.xls!D22` =TINV(2*A22,B22) → Excel 3.183098861837907e+99, GridCraft "#NUM!" |
| FINV | 1 | fail 1 | `yalta2008/yalta2008.xls!E21` =FINV(A21,B21,C21) → Excel 4.052847345693495e+199, GridCraft "#NUM!" |
| DVARP | 1 | fail 1 | `overview/database.xlsx!B31` =DVARP($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", GridCraft 0.25 |
| DVAR | 1 | fail 1 | `overview/database.xlsx!B30` =DVAR($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", GridCraft 0.5 |
| DPRODUCT | 1 | fail 1 | `overview/database.xlsx!B26` =DPRODUCT($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", GridCraft 2 |
| DSTDEVP | 1 | fail 1 | `overview/database.xlsx!B28` =DSTDEVP($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", GridCraft 0.5 |
| DSTDEV | 1 | fail 1 | `overview/database.xlsx!B27` =DSTDEV($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", GridCraft 0.7071067811865476 |
| DSUM | 1 | fail 1 | `overview/database.xlsx!B29` =DSUM($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", GridCraft 3 |
