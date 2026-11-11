# gnumeric vs libreoffice@26.8.2.2 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 909 | 35 | 1 | 0 | 5 | 96.2% |
| overview | 650 | 473 | 173 | 0 | 0 | 4 | 73.2% |
| yalta2008 | 369 | 353 | 16 | 0 | 0 | 0 | 95.7% |
| **total** | 1969 | 1735 | 224 | 1 | 0 | 9 | 88.5% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1735 | 224 | 1 | 0 | 9 | 88.5% |
| **total** | 1969 | 1735 | 224 | 1 | 0 | 9 | 88.5% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| COUNTIF | 35 | fail 35 | `overview/countif.xlsx!S9` =COUNTIF($A9,S$6) → Excel 0, LibreOffice 1 |
| SUMPRODUCT | 22 | fail 22 | `overview/countif.xlsx!S18` =SUMPRODUCT($U$7:$U$15,S7:S15) → Excel 64, LibreOffice 69 |
| IF | 15 | fail 15 | `yalta2008/yalta2008.xls!J10` =IF(ISERROR(F10),FALSE,ABS(H10-F10)<=I10) → Excel false, LibreOffice "#NUM!" |
| IFERROR | 14 | fail 14 | `overview/database.xlsx!D40` =IFERROR(ERROR.TYPE(B40)=ERROR.TYPE(C40),IFERROR(B40=C40,FALS → Excel true, LibreOffice false |
| SUM | 14 | fail 14 | `overview/database.xlsx!B1` =SUM(0+D$17:D$10001) → Excel 22, LibreOffice "#DIV/0!" |
| DSUM | 9 | fail 9 | `overview/database.xlsx!B40` =DSUM($A$4:$C$10,"Value",$C$14:$C$15) → Excel 3, LibreOffice 4 |
| SUMIF | 6 | fail 6 | `overview/ifs-funcs.xlsx!B27` =SUMIF($B$6:$B$12,"<>5",$C$6:$C$12) → Excel 3, LibreOffice 4 |
| XIRR | 3 | fail 3 | `functions/finfuns.xls!F69` =XIRR({-10000;2750;4290;3250;2742},{"1998/1/1";"1998/1/3";"19 → Excel 0.4230088663101197, LibreOffice "#VALUE!" |
| ODDFYIELD | 3 | fail 3 | `functions/finfuns.xls!F52` =ODDFYIELD(DATE(1999,11,11),DATE(2012,3,1),DATE(1999,10,15),D → Excel 12.615078687777377, LibreOffice "#VALUE!" |
| ODDFPRICE | 3 | fail 3 | `functions/finfuns.xls!F51` =ODDFPRICE(DATE(1999,11,11),DATE(2012,3,1),DATE(1999,10,15),D → Excel 113.65002161109066, LibreOffice "#VALUE!" |
| MDURATION | 3 | fail 3 | `functions/finfuns.xls!F46` =MDURATION(DATE(1999,1,1),DATE(2007,1,1),0.08,0.09,4,1) → Excel 5.784501720486185, LibreOffice 5.78509678900461 |
| NETWORKDAYS | 3 | fail 2, unsupported 1 | `functions/datefuns.xls!F30` =NETWORKDAYS("1999/01/22","2001/05/21",@{"2000/03/23";"2000/0 → Excel 606, LibreOffice "#NAME?" |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", LibreOffice "Windows (32-bit) NT 5.01" |
| VLOOKUP | 3 | fail 3 | `functions/lookfuns.xls!F43` =VLOOKUP(3,A22:C25,1,TRUE) → Excel 2.3, LibreOffice "#N/A" |
| GETPIVOTDATA | 3 | fail 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, LibreOffice "#REF!" |
| AND | 3 | fail 3 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, LibreOffice "#NUM!" |
| COLUMNS | 3 | fail 3 | `overview/countif.xlsx!B2` =COLUMNS(B18:S20)-B1 → Excel 0, LibreOffice 11 |
| DURATION | 2 | fail 2 | `functions/finfuns.xls!F38` =DURATION(DATE(1998,1,1),DATE(2006,1,1),0.18,0.09,2,1) → Excel 5.170131211421321, LibreOffice 5.17073966898141 |
| CONVERT | 2 | fail 2 | `functions/engfuns.xls!F21` =CONVERT(7.9,"cal","J") → Excel 33.07572, LibreOffice 33.0756792845004 |
| DATEVALUE | 2 | fail 2 | `functions/datefuns.xls!D22` =DATEVALUE("1929/02/28") → Excel 10652, LibreOffice "#VALUE!" |
| CELL | 2 | fail 2 | `functions/infofuns.xls!F26` =CELL("format",A14) → Excel "C0-", LibreOffice ",0-()" |
| INT | 2 | fail 2 | `yalta2008/yalta2008.xls!I10` =10^(INT(LOG10(H10))-5)/2 → Excel 0, LibreOffice "#NUM!" |
| POISSON | 2 | fail 2 | `yalta2008/yalta2008.xls!E22` =POISSON(A22,B22,C22) → Excel 0.500084104416326, LibreOffice "#N/A" |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", LibreOffice "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", LibreOffice "#NAME?" |
