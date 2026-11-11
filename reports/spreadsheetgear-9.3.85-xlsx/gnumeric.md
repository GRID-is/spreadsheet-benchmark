# gnumeric vs spreadsheetgear@9.3.85 (gnumeric)

Function list: GRID 19.0.0. Pass rate is over parity cases only. Skips: volatile 2, environment 3, excel-read-as-user-function 4. Fails where the engine's error came from another cell: 0. Fails where both answered an error, but a different one: 0. Fails where the case's own cell matches but its spill does not: 0.
Functions outside Excel's list (0): none. Never-supported functions present: none.
Upstream literal differs from Excel's value in 10 of 945 annotated parity cases.

| category | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| functions | 950 | 937 | 4 | 4 | 0 | 5 | 99.2% |
| overview | 650 | 646 | 0 | 0 | 0 | 4 | 100.0% |
| yalta2008 | 369 | 363 | 6 | 0 | 0 | 0 | 98.4% |
| **total** | 1969 | 1946 | 10 | 4 | 0 | 9 | 99.3% |

| saved by | cases | pass | fail | unsupported | error | skip | parity pass rate |
|---|---|---|---|---|---|---|---|
| Microsoft Excel 16.0 (20527) | 1969 | 1946 | 10 | 4 | 0 | 9 | 99.3% |
| **total** | 1969 | 1946 | 10 | 4 | 0 | 9 | 99.3% |

## Functions with most non-passing cases

| function | non-pass | statuses | example |
|---|---|---|---|
| BESSELJ | 3 | fail 3 | `functions/engfuns.xls!F14` =BESSELJ(4,3) → Excel 0.4301714711533965, SpreadsheetGear 0.28012474255025016 |
| INFO | 3 | skip 3 | `functions/infofuns.xls!F29` =INFO("osversion") → Excel "Windows (64-bit) NT 10.00", SpreadsheetGear "#VALUE!" |
| GETPIVOTDATA | 3 | unsupported 3 | `functions/dbfuns.xls!F44` =GETPIVOTDATA(F15:J16,"Jos2") → Excel 76446, SpreadsheetGear "#NAME?" |
| IF | 2 | fail 2 | `yalta2008/yalta2008.xls!J10` =IF(ISERROR(F10),FALSE,ABS(H10-F10)<=I10) → Excel false, SpreadsheetGear true |
| _XLUDF.MINIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B22` =_xludf.MINIFS($C$6:$C$12,$B$6:$B$12,">=2") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| _XLUDF.MAXIFS | 2 | skip 2 | `overview/ifs-funcs.xlsx!B30` =_xludf.MAXIFS($C$6:$C$12,$B$6:$B$12,"<>5") → Excel "#NAME?", SpreadsheetGear "#NAME?" |
| IMPOWER | 1 | fail 1 | `functions/engfuns.xls!F41` =IMPOWER("8-2j",9.2) → Excel "-169691027.479611-208555947.133373j", SpreadsheetGear "-169691027.47961-208555947.133372j" |
| NETWORKDAYS | 1 | unsupported 1 | `functions/datefuns.xls!F30` =NETWORKDAYS("1999/01/22","2001/05/21",@{"2000/03/23";"2000/0 → Excel 606, SpreadsheetGear "#NAME?" |
| NOW | 1 | skip 1 | `functions/datefuns.xls!B31` =NOW() → Excel 46330.76197175926, SpreadsheetGear 46335.70366520174 |
| TODAY | 1 | skip 1 | `functions/datefuns.xls!B35` =TODAY() → Excel 46330, SpreadsheetGear 46335 |
| AND | 1 | fail 1 | `yalta2008/yalta2008.xls!B1` =AND(J10:J65536) → Excel false, SpreadsheetGear true |
| POISSON | 1 | fail 1 | `yalta2008/yalta2008.xls!E23` =POISSON(A23,B23,C23) → Excel 0.500008410441739, SpreadsheetGear 0.5000077942065494 |
| CHIINV | 1 | fail 1 | `yalta2008/yalta2008.xls!D20` =CHIINV(A20,B20) → Excel 50.84417133244917, SpreadsheetGear 50.84418265206363 |
