# gnumeric vs mog@1.0.0 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 9. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 927 | 18 | 0 | 0 | 5 | 98.1% |
| overview | 650 | 555 | 91 | 0 | 0 | 4 | 85.9% |
| yalta2008 | 369 | 129 | 11 | 229 | 0 | 0 | 35.0% |
| **total** | 1969 | 1611 | 120 | 229 | 0 | 9 | 82.2% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1611 | 120 | 229 | 0 | 9 | 82.2% |
| **total** | 1969 | 1611 | 120 | 229 | 0 | 9 | 82.2% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| IF | 108 | fail 3, unsupported 105 | `yalta2008/yalta2008.xls!J22` =IF(ISERROR(F22),FALSE,ABS(H22-F22)<=I22) → Excel true, Mog "#NAME?" |
| INT | 106 | unsupported 106 | `yalta2008/yalta2008.xls!I21` =10^(INT(LOG10(H21))-5)/2 → Excel 5e-7, Mog "#NAME?" |
| IFERROR | 17 | fail 17 | `overview/database.xlsx!D27` =IFERROR(ERROR.TYPE(B27)=ERROR.TYPE(C27),IFERROR(B27=C27,FALS → Excel true, Mog false |
| SUM | 11 | fail 11 | `functions/statfuns.xls!F90` =SUM(K95:K98) → Excel 57.97039877766521, Mog 40.577538332046814 |
| SUMIF | 11 | fail 11 | `overview/ifs-funcs.xlsx!B29` =SUMIF($B$6:$B$14,A29,$C$6:$C$14) → Excel 15, Mog 0 |
| COUNTIF | 10 | fail 10 | `overview/countif.xlsx!F14` =COUNTIF($A14,F$6) → Excel 0, Mog 1 |
| AND | 9 | unsupported 9 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, Mog "#NAME?" |
| TINV | 6 | fail 6 | `yalta2008/yalta2008.xls!D20` =TINV(2*A20,B20) → Excel 318309886183.7907, Mog 92007338.2908108 |
| SUMPRODUCT | 6 | fail 6 | `overview/countif.xlsx!F18` =SUMPRODUCT($J$7:$J$15,F7:F15) → Excel 3, Mog 195 |
| DSUM | 5 | fail 5 | `overview/database.xlsx!B29` =DSUM($A$4:$C$10,"Value",$A$14:$A$15) → Excel "#DIV/0!", Mog 3 |
| VDB | 3 | fail 3 | `functions/finfuns.xls!F68` =VDB(2400,300,10,0,0.3234,2.5) → Excel 194.04000000000002, Mog 0 |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", Mog "Shortcut" |
| GETPIVOTDATA | 3 | fail 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, Mog "#VALUE!" |
| IMPOWER | 2 | fail 2 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", Mog "-169691027.47961-208555947.133372j" |
| COLUMNS | 2 | fail 2 | `overview/countif.xlsx!B2` =COLUMNS(B18:G20)-B1 → Excel 0, Mog 3 |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", Mog "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", Mog "#NAME?" |
| YIELDDISC | 1 | fail 1 | `functions/finfuns.xls!D72` =YIELDDISC(DATE(1999,2,15),DATE(2007,11,15),0.0575,100,1) → Excel 198.68665411686433, Mog 198.7017656664625 |
| MDURATION | 1 | fail 1 | `functions/finfuns.xls!D46` =MDURATION(DATE(1999,1,1),DATE(2007,1,1),0.18,0.09,2,2) → Excel 4.94749398222136, Mog 4.950152142774258 |
| IMSQRT | 1 | fail 1 | `functions/engfuns.xls!D45` =IMSQRT("3-2j") → Excel "1.81735402102397-0.550250522700337j", Mog "1.81735402102397-0.550250522700338j" |
| BESSELJ | 1 | fail 1 | `functions/engfuns.xls!D14` =BESSELJ(8,3) → Excel -0.29112524185253685, Mog -0.29113220692602515 |
| NOW | 1 | skip 1 | `functions/datefuns.xls!B31` =NOW() → Excel 46330.76197175926, Mog 46335.24663194444 |
| TODAY | 1 | skip 1 | `functions/datefuns.xls!B35` =TODAY() → Excel 46330, Mog 46335 |
| VLOOKUP | 1 | fail 1 | `functions/lookfuns.xls!B43` =VLOOKUP(4,A22:C25,2,TRUE) → Excel 7.2, Mog 5.9 |
| MATCH | 1 | fail 1 | `functions/lookfuns.xls!B38` =MATCH(5,D11:D20,1) → Excel 9, Mog 5 |
